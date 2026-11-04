#!/usr/bin/env python3
"""
Per-minute statistics (Avg, P50, P95, P99, Max) for the CSV produced by
Export-FoundryTokenMetrics.ps1, plus an optional Excel workbook.

Usage:
    python scripts/foundry_token_stats.py --run-dir output/foundry_20261103-120000Z
    python scripts/foundry_token_stats.py --run-dir output/tenantA --run-dir output/tenantB --out-dir output/all

Outputs (in the first --run-dir unless --out-dir is given). Every file carries the UTC timestamp of
this execution (<ts> = yyyyMMdd-HHmmssZ) so scheduled runs never overwrite each other:
    foundry-token-stats_<ts>.csv            one row per Level x group x Mode
    foundry-token-daily_<ts>.csv            per day peak/sum per deployment and model (UTC)
    foundry-deployment-limits_<ts>.csv      per-minute consumption vs deployment TPM/RPM limits
    foundry-quota-utilization_<ts>.csv      per-minute consumption vs subscription/region quota
    foundry-token-hourly_<ts>.csv           average/max per minute by UTC hour of day, per model
    foundry-token-report_<ts>.xlsx          (unless --no-excel) compact report, no charts:
                                            Summary, Models (one row per deployment), Quotas.
                                            Full detail stays in the CSVs above.

Semantics
---------
* Values are per-minute totals (Total aggregation at PT1M), i.e. tokens per minute (TPM).
  TPS = TPM / 60. Input (ProcessedPromptTokens) and output (GeneratedTokens) are always kept
  separate; TotalTokens(derived) is only informational.
* The extraction writes only non-zero minutes. A missing minute is a real zero ONLY when the
  manifest shows a successful call for that resource/chunk. Failed chunks are excluded from the
  denominator and reported through CompletenessPct.
* Modes (denominator N for Avg/percentiles):
    FullWindow  - every successfully observed minute of the window (zero-filled). PTU sizing view.
    Lifetime    - minutes between first and last active minute of the group (zero-filled).
    Active      - only minutes with usage > 0. Overstates steady demand; use with care.
* Percentiles use the nearest-rank method over the zero-filled distribution, computed
  analytically (zeros are never materialised).
* Utilization V2 (percent) is only reported at Deployment level and is never summed.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path

try:
    import duckdb
except ImportError:  # pragma: no cover
    sys.exit("duckdb is required: python -m pip install duckdb pandas openpyxl")

PERCENTILES = (0.50, 0.95, 0.99)
TOKEN_METRICS = ("ProcessedPromptTokens", "GeneratedTokens")
DERIVED = "TotalTokens(derived)"
UTIL = "AzureOpenAIProvisionedManagedUtilizationV2"
THROTTLED = "ThrottledRequests(429)"
CLIENT_ERR = "ClientErrors(4xx)"
SERVER_ERR = "ServerErrors(5xx)"
TTR = "AzureOpenAITimeToResponse"
TTLT = "AzureOpenAITTLTInMS"
TPS = "AzureOpenAITokenPerSecond"
AVAIL = "AzureOpenAIAvailabilityRate"
LATENCY_METRICS = (TTR, TTLT, TPS, AVAIL)
REQUEST_METRICS = ("AzureOpenAIRequests", THROTTLED, CLIENT_ERR, SERVER_ERR)
MINUTES_PER_MONTH = 30 * 24 * 60
EXCEL_MAX_ROWS = 1_000_000
DEP_COLS = {
    "TenantId": "VARCHAR", "SubscriptionId": "VARCHAR", "SubscriptionName": "VARCHAR", "Region": "VARCHAR",
    "ResourceGroup": "VARCHAR", "ResourceName": "VARCHAR", "Kind": "VARCHAR", "DeploymentName": "VARCHAR",
    "ModelFormat": "VARCHAR", "ModelName": "VARCHAR", "ModelVersion": "VARCHAR", "DeploymentSku": "VARCHAR",
    "Capacity": "DOUBLE", "TokensPerMinuteLimit": "DOUBLE", "RequestsPerMinuteLimit": "DOUBLE",
    "QuotaName": "VARCHAR", "QuotaMatch": "VARCHAR", "QuotaLimit": "DOUBLE", "QuotaCurrentValue": "DOUBLE",
    "ResourceId": "VARCHAR", "CapturedAt": "VARCHAR",
}
QUOTA_COLS = {
    "TenantId": "VARCHAR", "SubscriptionId": "VARCHAR", "SubscriptionName": "VARCHAR", "Region": "VARCHAR",
    "QuotaName": "VARCHAR", "DeploymentType": "VARCHAR", "QuotaModel": "VARCHAR", "CurrentValue": "DOUBLE",
    "Limit": "DOUBLE", "Unit": "VARCHAR", "AssignedPct": "DOUBLE", "UsedByDeployments": "VARCHAR",
    "CapturedAt": "VARCHAR",
}


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--run-dir", action="append", required=True, help="Run folder from the export script (repeatable)")
    p.add_argument("--out-dir", help="Output folder (default: first --run-dir)")
    p.add_argument("--no-excel", action="store_true", help="Skip the consolidated Excel report")
    p.add_argument("--excel", action="store_true", help=argparse.SUPPRESS)  # kept for compatibility (Excel is the default)
    p.add_argument("--top-n", type=int, default=10, help=argparse.SUPPRESS)  # deprecated no-op (Excel has no charts)
    p.add_argument("--db", default=":memory:", help="DuckDB database file (use a file for very large datasets)")
    p.add_argument("--prices", help="CSV with list prices per 1M tokens (see prices.example.csv); overrides --fetch-prices")
    p.add_argument("--fetch-prices", action="store_true",
                   help="Get candidate PAYG prices from the public Azure Retail Prices API (exact matches only)")
    p.add_argument("--currency", default="USD", help="Currency for --fetch-prices (default USD)")
    return p.parse_args()


def load(con: duckdb.DuckDBPyConnection, run_dirs: list[Path]) -> dict:
    windows = set()
    targets: set[str] = set()
    for d in run_dirs:
        run = json.loads((d / "run.json").read_text(encoding="utf-8-sig"))
        fp = run["Fingerprint"]
        windows.add((fp["StartTime"], fp["EndTime"]))
        tr = run.get("TargetResources") or []
        targets.update([tr] if isinstance(tr, str) else tr)
        for f in ("foundry-token-metrics.csv", "manifest.csv"):
            if not (d / f).exists():
                sys.exit(f"Missing {d / f}. Did the export finish?")
    if len(windows) != 1:
        sys.exit(f"All run folders must share the same window, found: {sorted(windows)}")
    start, end = windows.pop()

    csvs = [str(d / "foundry-token-metrics.csv") for d in run_dirs]
    mans = [str(d / "manifest.csv") for d in run_dirs]
    con.execute("SET TimeZone = 'UTC'")
    con.execute(
        """
        CREATE TABLE raw AS SELECT * FROM read_csv($files, header = true, auto_detect = false,
            timestampformat = '%Y-%m-%dT%H:%M:%SZ',
            columns = {
                'Timestamp': 'TIMESTAMP', 'SubscriptionId': 'VARCHAR', 'SubscriptionName': 'VARCHAR',
                'ResourceGroup': 'VARCHAR', 'ResourceName': 'VARCHAR', 'Region': 'VARCHAR',
                'DeploymentName': 'VARCHAR', 'ModelName': 'VARCHAR', 'ModelVersion': 'VARCHAR',
                'Metric': 'VARCHAR', 'Total': 'DOUBLE', 'TenantId': 'VARCHAR', 'Kind': 'VARCHAR',
                'Aggregation': 'VARCHAR', 'DeploymentSku': 'VARCHAR', 'DeploymentCapacity': 'DOUBLE',
                'ModelNameSource': 'VARCHAR', 'ResourceId': 'VARCHAR'})
        """,
        {"files": csvs},
    )
    con.execute(
        """
        CREATE TABLE manifest AS
        SELECT lower(ResourceId) AS rid, CAST(ChunkStart AS TIMESTAMP) AS cs, CAST(ChunkEnd AS TIMESTAMP) AS ce,
               Call, Status
        FROM read_csv($files, header = true, all_varchar = true)
        QUALIFY row_number() OVER (PARTITION BY lower(ResourceId), ChunkStart, Call ORDER BY RecordedAt DESC) = 1
        """,
        {"files": mans},
    )
    con.execute("CREATE TABLE win AS SELECT CAST($s AS TIMESTAMP) AS ws, CAST($e AS TIMESTAMP) AS we",
                {"s": start.replace("Z", ""), "e": end.replace("Z", "")})
    # Snapshot of deployments (rate limits, quota mapping) and subscription/region quota (optional files).
    for table, fname, cols in (("deps", "deployments.csv", DEP_COLS), ("quota", "quota.csv", QUOTA_COLS)):
        files = [str(d / fname) for d in run_dirs if (d / fname).exists() and (d / fname).stat().st_size > 0]
        if files:
            con.execute(f"CREATE TABLE {table} AS SELECT * FROM read_csv($f, header = true, auto_detect = false, columns = $c)",
                        {"f": files, "c": cols})
        else:
            con.execute(f"CREATE TABLE {table} ({', '.join(f'{k} {v}' for k, v in cols.items())})")
    con.execute("ALTER TABLE deps ADD COLUMN rid VARCHAR")
    con.execute("UPDATE deps SET rid = lower(ResourceId), SubscriptionId = lower(SubscriptionId)")
    con.execute("UPDATE quota SET SubscriptionId = lower(SubscriptionId)")
    # Same deployment captured by several runs: keep the latest snapshot.
    con.execute("CREATE OR REPLACE TABLE deps AS SELECT * FROM deps QUALIFY row_number() OVER (PARTITION BY rid, DeploymentName ORDER BY CapturedAt DESC) = 1")
    con.execute("CREATE OR REPLACE TABLE quota AS SELECT * FROM quota QUALIFY row_number() OVER (PARTITION BY SubscriptionId, Region, QuotaName ORDER BY CapturedAt DESC) = 1")
    w = con.execute("SELECT date_diff('minute', ws, we) FROM win").fetchone()[0]
    return {"start": start, "end": end, "window_minutes": w, "targets": sorted(targets)}


SEP = "chr(31)"


def key(level: str, *cols: str) -> str:
    """Collision-free group id (level prefix, NULL-safe, separator that never occurs in names)."""
    return f"concat_ws({SEP}, '{level}', " + ", ".join(f"coalesce(CAST({c} AS VARCHAR), '')" for c in cols) + ")"


def build(con: duckdb.DuckDBPyConnection) -> None:
    # Exact duplicates (e.g. resumed runs) collapse to one point per full identity (ModelName included).
    con.execute(
        """
        CREATE TABLE pts AS
        SELECT lower(ResourceId) AS rid, Timestamp AS ts, DeploymentName AS dep, ModelVersion AS ver, Region AS region,
               ModelName AS model, Metric AS metric, Aggregation AS agg, max(Total) AS v,
               any_value(SubscriptionName) AS sub, any_value(ResourceGroup) AS rg,
               any_value(ResourceName) AS res, any_value(DeploymentSku) AS sku, max(DeploymentCapacity) AS cap,
               string_agg(DISTINCT ModelNameSource, ',') AS model_src
        FROM raw
        GROUP BY rid, ts, dep, ver, region, model, metric, agg
        """
    )
    # Derived TotalTokens = prompt + generated (informational only).
    con.execute(
        f"""
        INSERT INTO pts
        SELECT rid, ts, dep, ver, region, model, '{DERIVED}', 'Total', sum(v), any_value(sub),
               any_value(rg), any_value(res), any_value(sku), max(cap), string_agg(DISTINCT model_src, ',')
        FROM pts WHERE metric IN {TOKEN_METRICS} AND agg = 'Total'
        GROUP BY rid, ts, dep, ver, region, model
        """
    )
    # Latency metrics are 1-minute averages/maxima: 0 ms or 0 tokens/s in a minute with requests is a missing
    # sample, not a measurement. 0 % availability is real and kept.
    con.execute(f"DELETE FROM pts WHERE metric IN ('{TTR}', '{TTLT}', '{TPS}') AND v = 0")
    call_of = (f"CASE WHEN metric IN {TOKEN_METRICS} OR metric = '{DERIVED}' THEN 'Tokens' "
               f"WHEN metric IN {REQUEST_METRICS} THEN 'Requests' "
               f"WHEN metric IN {LATENCY_METRICS} THEN 'Latency' "
               f"WHEN metric = '{UTIL}' THEN 'PtuUtilization' ELSE 'Unknown' END")
    k_dep = key("D", "rid", "region", "dep", "model", "ver", "metric", "agg")
    k_mr = key("MR", "model", "ver", "region", "metric")
    k_m = key("M", "model", "metric")
    k_all = key("A", "metric")
    k_dl = key("DL", "rid", "dep", "metric")
    k_q = key("Q", "subid", "qregion", "qname", "metric")
    tot = f"agg = 'Total' AND metric <> '{UTIL}'"
    # Deployment -> subscription/region quota (current snapshot). Deleted deployments stay visible as Unmapped.
    con.execute(
        """
        CREATE TABLE qmap AS
        SELECT p.rid, p.dep, split_part(p.rid, '/', 3) AS subid, coalesce(d.Region, '') AS qregion,
               coalesce(nullif(d.QuotaName, ''), CASE WHEN d.rid IS NULL THEN 'Unmapped(deployment not found)' ELSE 'Unmapped' END) AS qname
        FROM (SELECT DISTINCT rid, dep FROM pts) p LEFT JOIN deps d ON d.rid = p.rid AND d.DeploymentName = p.dep
        UNION
        SELECT rid, DeploymentName, split_part(rid, '/', 3), coalesce(Region, ''), coalesce(nullif(QuotaName, ''), 'Unmapped')
        FROM deps
        """
    )
    con.execute(f"CREATE TABLE pq AS SELECT p.*, q.subid, q.qregion, q.qname FROM pts p JOIN qmap q USING (rid, dep) WHERE {tot}")
    con.execute(
        f"""
        CREATE TABLE groups AS
        SELECT 'Deployment' AS level, {k_dep} AS gid,
               any_value(sub) AS SubscriptionName, any_value(rg) AS ResourceGroup, any_value(res) AS ResourceName,
               region AS Region, dep AS DeploymentName, any_value(sku) AS DeploymentSku, max(cap) AS DeploymentCapacity,
               model AS ModelName, ver AS ModelVersion, metric AS Metric, agg AS Aggregation,
               string_agg(DISTINCT model_src, ',') AS ModelNameSource, list(DISTINCT rid) AS rids, {call_of} AS call,
               split_part(rid, '/', 3) AS SubscriptionId, rid AS ResourceId, NULL AS QuotaName
        FROM pts GROUP BY rid, region, dep, model, ver, metric, agg
        UNION ALL
        SELECT 'DeploymentLimit', {k_dl}, any_value(sub), any_value(rg), any_value(res), string_agg(DISTINCT region, '|'),
               dep, any_value(sku), max(cap), string_agg(DISTINCT model, '|'), string_agg(DISTINCT ver, '|'), metric, 'Total',
               string_agg(DISTINCT model_src, ','), list(DISTINCT rid), {call_of}, split_part(rid, '/', 3), rid, NULL
        FROM pts WHERE {tot} GROUP BY rid, dep, metric
        UNION ALL
        SELECT 'Quota', {k_q}, any_value(sub), NULL, NULL, qregion, NULL, NULL, NULL, string_agg(DISTINCT model, '|'), NULL,
               metric, 'Total', string_agg(DISTINCT model_src, ','), list(DISTINCT rid), {call_of}, subid, NULL, qname
        FROM pq GROUP BY subid, qregion, qname, metric
        UNION ALL
        SELECT 'ModelRegion', {k_mr}, NULL, NULL, NULL, region, NULL, NULL, NULL, model, ver, metric, 'Total',
               string_agg(DISTINCT model_src, ','), list(DISTINCT rid), {call_of}, NULL, NULL, NULL
        FROM pts WHERE {tot} GROUP BY model, ver, region, metric
        UNION ALL
        SELECT 'Model', {k_m}, NULL, NULL, NULL, NULL, NULL, NULL, NULL, model, NULL, metric, 'Total',
               string_agg(DISTINCT model_src, ','), list(DISTINCT rid), {call_of}, NULL, NULL, NULL
        FROM pts WHERE {tot} GROUP BY model, metric
        UNION ALL
        SELECT 'All', {k_all}, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, metric, 'Total',
               NULL, list(DISTINCT rid), {call_of}, NULL, NULL, NULL
        FROM pts WHERE {tot} GROUP BY metric
        """
    )
    # 'All' must be complete for every target resource queried for that call, not only those with usage.
    con.execute(
        """
        UPDATE groups SET rids = (SELECT list(DISTINCT m.rid) FROM manifest m WHERE m.Call = groups.call)
        WHERE level = 'All'
        """
    )
    # A quota is shared by every deployment mapped to it (also idle ones): all their resources must be covered.
    con.execute(
        """
        UPDATE groups SET rids = (SELECT list(DISTINCT q.rid) FROM qmap q
                                  WHERE q.subid = groups.SubscriptionId AND q.qregion = groups.Region AND q.qname = groups.QuotaName)
        WHERE level = 'Quota'
        """
    )
    # Coverage: elementary segments between all chunk boundaries; a segment counts for a group only
    # if EVERY contributing resource has a 'Success' record for that call (Truncated/Failed never count).
    con.execute(
        """
        CREATE TABLE seg AS
        WITH b AS (SELECT cs AS t FROM manifest UNION SELECT ce FROM manifest
                   UNION SELECT ws FROM win UNION SELECT we FROM win)
        SELECT s, e FROM (SELECT t AS s, lead(t) OVER (ORDER BY t) AS e FROM b), win
        WHERE e IS NOT NULL AND s >= ws AND e <= we
        """
    )
    con.execute(
        """
        CREATE TABLE rcov AS
        SELECT DISTINCT m.rid, m.Call AS call, seg.s
        FROM seg JOIN manifest m ON m.Status IN ('Success', 'NotSupported') AND m.cs <= seg.s AND seg.e <= m.ce
        """
    )
    con.execute(
        """
        CREATE TABLE gcov AS
        WITH gr AS (SELECT gid, call, unnest(rids) AS rid FROM groups)
        SELECT gr.gid, seg.s, seg.e
        FROM gr CROSS JOIN seg
        LEFT JOIN rcov c ON c.rid = gr.rid AND c.call = gr.call AND c.s = seg.s
        GROUP BY gr.gid, seg.s, seg.e
        HAVING bool_and(c.rid IS NOT NULL)
        """
    )
    con.execute(
        """
        CREATE TABLE gmeta AS
        WITH gr AS (SELECT gid, call, unnest(rids) AS rid FROM groups)
        SELECT g.gid, coalesce(c.cov, 0) AS cov, coalesce(t.tr, false) AS any_truncated
        FROM groups g
        LEFT JOIN (SELECT gid, sum(date_diff('minute', s, e)) AS cov FROM gcov GROUP BY gid) c USING (gid)
        LEFT JOIN (SELECT gr.gid, true AS tr FROM gr JOIN manifest m ON m.rid = gr.rid AND m.Call = gr.call
                   WHERE m.Status = 'Truncated' GROUP BY gr.gid) t USING (gid)
        """
    )
    con.execute(
        f"""
        CREATE TABLE series_raw AS
        SELECT {k_dep} AS gid, ts, v FROM pts
        UNION ALL
        SELECT {k_dl}, ts, sum(v) FROM pts WHERE {tot} GROUP BY rid, dep, metric, ts
        UNION ALL
        SELECT {k_q}, ts, sum(v) FROM pq GROUP BY subid, qregion, qname, metric, ts
        UNION ALL
        SELECT {k_mr}, ts, sum(v) FROM pts WHERE agg = 'Total' AND metric <> '{UTIL}' GROUP BY model, ver, region, metric, ts
        UNION ALL
        SELECT {k_m}, ts, sum(v) FROM pts WHERE agg = 'Total' AND metric <> '{UTIL}' GROUP BY model, metric, ts
        UNION ALL
        SELECT {k_all}, ts, sum(v) FROM pts WHERE agg = 'Total' AND metric <> '{UTIL}' GROUP BY metric, ts
        """
    )
    # Keep only minutes fully covered for the group, so rollups never mix complete and partial totals.
    con.execute(
        """
        CREATE TABLE series AS
        SELECT r.gid, r.ts, r.v FROM series_raw r JOIN gcov c ON c.gid = r.gid AND r.ts >= c.s AND r.ts < c.e
        """
    )
    dropped = con.execute("SELECT (SELECT count(*) FROM series_raw) - (SELECT count(*) FROM series)").fetchone()[0]
    if dropped:
        print(f"NOTE: {dropped:,} group-minute value(s) outside fully covered segments were excluded.")


def compute_stats(con: duckdb.DuckDBPyConnection, window_minutes: int):
    con.execute(
        """
        CREATE TABLE gstats AS
        SELECT gid, count(*) AS k, sum(v) AS s, max(v) AS mx, arg_max(ts, v) AS mx_ts, min(ts) AS first_ts, max(ts) AS last_ts
        FROM series GROUP BY gid
        """
    )
    # Lifetime denominator: covered minutes inside [first_ts, last_ts + 1 min).
    con.execute(
        """
        CREATE TABLE life AS
        SELECT s.gid, sum(greatest(0, date_diff('minute', greatest(c.s, s.first_ts),
                                                least(c.e, s.last_ts + INTERVAL 1 MINUTE)))) AS life
        FROM gstats s JOIN gcov c USING (gid) GROUP BY s.gid
        """
    )
    con.execute("CREATE TABLE ranked AS SELECT gid, v, row_number() OVER (PARTITION BY gid ORDER BY v) AS rn FROM series")
    con.execute(
        """
        CREATE TABLE modes AS
        -- Groups whose usage fell only in uncovered segments are kept (k = 0) so they stay visible.
        SELECT g.*, gm.cov, gm.any_truncated, coalesce(s.k, 0) AS k, coalesce(s.s, 0) AS s, s.mx, s.mx_ts,
               s.first_ts, s.last_ts, m.mode,
               -- Latency/availability are averages of minutes with requests: never zero-filled.
               greatest(coalesce(s.k, 0), CASE WHEN g.call = 'Latency' THEN coalesce(s.k, 0)
                                                      WHEN m.mode = 'FullWindow' THEN gm.cov
                                                      WHEN m.mode = 'Lifetime' THEN coalesce(l.life, 0)
                                                      ELSE coalesce(s.k, 0) END) AS n
        FROM groups g JOIN gmeta gm USING (gid) LEFT JOIN gstats s USING (gid) LEFT JOIN life l USING (gid)
        CROSS JOIN (VALUES ('FullWindow'), ('Lifetime'), ('Active')) AS m(mode)
        """
    )
    pct_cols = []
    for q in PERCENTILES:
        name = f"P{int(round(q * 100))}"
        # nearest rank r = ceil(q*n); zeros occupy ranks 1..(n-k); nonzero value index = r-(n-k)
        con.execute(
            f"""
            CREATE TABLE {name} AS
            SELECT md.gid, md.mode, CASE WHEN md.n = 0 THEN NULL ELSE coalesce(r.v, 0) END AS {name}
            FROM modes md LEFT JOIN ranked r
              ON r.gid = md.gid AND r.rn = CAST(ceil({q} * md.n) AS BIGINT) - (md.n - md.k)
            """
        )
        pct_cols.append(name)
    joins = " ".join(f"JOIN {c} USING (gid, mode)" for c in pct_cols)
    return con.execute(
        f"""
        SELECT md.level AS Level, md.mode AS Mode, md.SubscriptionId, SubscriptionName, ResourceGroup, ResourceName, Region,
               DeploymentName, DeploymentSku, DeploymentCapacity, md.QuotaName, ModelName, ModelVersion, ModelNameSource,
               Metric, Aggregation,
               md.n AS Minutes, md.k AS ActiveMinutes,
               round(100.0 * md.cov / {window_minutes}, 2) AS CompletenessPct, md.any_truncated AS Truncated,
               md.s AS Sum, md.s / nullif(md.n, 0) AS Avg, {', '.join(pct_cols)}, CASE WHEN md.n = 0 THEN NULL ELSE coalesce(md.mx, 0) END AS Max,
               strftime(md.mx_ts, '%Y-%m-%dT%H:%M:%SZ') AS MaxTimestamp,
               strftime(md.first_ts, '%Y-%m-%dT%H:%M:%SZ') AS FirstActive,
               strftime(md.last_ts, '%Y-%m-%dT%H:%M:%SZ') AS LastActive, md.ResourceId
        FROM modes md {joins}
        ORDER BY CASE md.level WHEN 'All' THEN 0 WHEN 'Quota' THEN 1 WHEN 'Model' THEN 2 WHEN 'ModelRegion' THEN 3
                               WHEN 'DeploymentLimit' THEN 4 ELSE 5 END,
                 Metric, CASE md.mode WHEN 'FullWindow' THEN 0 WHEN 'Lifetime' THEN 1 ELSE 2 END, md.s DESC
        """
    ).df()


def daily(con: duckdb.DuckDBPyConnection):
    return con.execute(
        """
        SELECT g.level AS Level, CAST(date_trunc('day', s.ts) AS DATE) AS Day, g.SubscriptionName, g.ResourceName,
               g.Region, g.DeploymentName, g.ModelName, g.ModelVersion, g.Metric, g.Aggregation,
               sum(s.v) AS DaySum, max(s.v) AS PeakMinute, count(*) AS ActiveMinutes,
               strftime(arg_max(s.ts, s.v), '%H:%M') AS PeakAtUTC
        FROM series s JOIN groups g USING (gid)
        WHERE g.level IN ('Deployment', 'Model')
        GROUP BY ALL ORDER BY Level, Metric, Day
        """
    ).df()


def hourly(con: duckdb.DuckDBPyConnection):
    # Denominator = covered minutes falling in each UTC hour-of-day for that group.
    return con.execute(
        """
        WITH hrs AS (
            SELECT c.gid, c.s, c.e, unnest(generate_series(date_trunc('hour', c.s), c.e - INTERVAL 1 MINUTE, INTERVAL 1 HOUR)) AS h
            FROM gcov c JOIN groups g USING (gid) WHERE g.level = 'Model'
        ),
        den AS (
            SELECT gid, hour(h) AS HourUTC,
                   sum(date_diff('minute', greatest(h, s), least(h + INTERVAL 1 HOUR, e))) AS CoveredMinutes
            FROM hrs GROUP BY ALL
        ),
        num AS (
            SELECT s.gid, hour(s.ts) AS HourUTC, sum(s.v) AS SumV, max(s.v) AS MaxPerMinute
            FROM series s JOIN groups g USING (gid) WHERE g.level = 'Model' GROUP BY ALL
        )
        SELECT g.ModelName, g.Metric, d.HourUTC, d.CoveredMinutes,
               coalesce(n.SumV, 0) / d.CoveredMinutes AS AvgPerMinute, coalesce(n.MaxPerMinute, 0) AS MaxPerMinute
        FROM den d JOIN groups g USING (gid) LEFT JOIN num n USING (gid, HourUTC)
        WHERE d.CoveredMinutes > 0
        ORDER BY g.ModelName, g.Metric, d.HourUTC
        """
    ).df()


def _pivot(full, level: str, keys: list[str]):
    """One row per group with Tokens/Input/Output/Requests/Throttled stats side by side (FullWindow)."""
    f = full[full["Level"] == level]
    spec = [
        (DERIVED, "Tokens", ["P95", "P99", "Max", "MaxTimestamp", "CompletenessPct"]),
        ("ProcessedPromptTokens", "Input", ["P99", "Max", "Sum"]),
        ("GeneratedTokens", "Output", ["P99", "Max", "Sum"]),
        ("AzureOpenAIRequests", "Requests", ["P99", "Max", "Sum"]),
        (THROTTLED, "Throttled429", ["Sum", "ActiveMinutes"]),
        (CLIENT_ERR, "Errors4xx", ["Sum"]),
        (SERVER_ERR, "Errors5xx", ["Sum"]),
    ]
    out = None
    for metric, prefix, cols in spec:
        m = f[f["Metric"] == metric][keys + cols].rename(columns={c: f"{prefix}_{c}" for c in cols})
        out = m if out is None else out.merge(m, on=keys, how="outer")
    # Sparse extraction emits no 429/4xx/5xx rows when there were none: zero, not unknown, if requests were covered.
    for c in ("Throttled429_Sum", "Throttled429_ActiveMinutes", "Errors4xx_Sum", "Errors5xx_Sum"):
        out[c] = out[c].where(out[c].notna() | out["Requests_Max"].isna(), 0)
    return out


def _pct(num, den):
    import numpy as np
    return (100.0 * num / den.where(den > 0)).round(2).replace([np.inf, -np.inf], np.nan)


def limit_reports(con: duckdb.DuckDBPyConnection, stats, window_minutes: int):
    """Per-minute consumption vs deployment rate limits and vs subscription/region quota."""
    full = stats[stats["Mode"] == "FullWindow"]
    deps = con.execute("SELECT * FROM deps").df()
    quota = con.execute("SELECT * FROM quota").df()

    # Deployment: tokens/requests per minute vs the deployment's own TPM/RPM limit.
    p = _pivot(full, "DeploymentLimit", ["ResourceId", "DeploymentName"])
    names = full[full["Level"] == "DeploymentLimit"].groupby(["ResourceId", "DeploymentName"], as_index=False).agg(
        SubscriptionName=("SubscriptionName", "first"), ResourceName=("ResourceName", "first"),
        MetricModel=("ModelName", "first"))
    p = p.merge(names, on=["ResourceId", "DeploymentName"], how="left")
    d = deps.rename(columns={"rid": "ResourceIdLower"}).drop(columns=["ResourceId"])
    dep = d.merge(p, left_on=["ResourceIdLower", "DeploymentName"], right_on=["ResourceId", "DeploymentName"],
                  how="outer", suffixes=("", "_m"))
    for c in ("SubscriptionName", "ResourceName"):
        dep[c] = dep[c].fillna(dep[f"{c}_m"])
    dep["ModelName"] = dep["ModelName"].fillna(dep["MetricModel"])
    dep["QuotaName"] = dep["QuotaName"].fillna("Unmapped(deployment not found)")
    dep["ResourceId"] = dep["ResourceIdLower"].fillna(dep["ResourceId"])
    for s in ("P95", "P99", "Max"):
        dep[f"PctOfTPM_{s}"] = _pct(dep[f"Tokens_{s}"], dep["TokensPerMinuteLimit"])
    for s in ("P99", "Max"):
        dep[f"PctOfRPM_{s}"] = _pct(dep[f"Requests_{s}"], dep["RequestsPerMinuteLimit"])
    dep_cols = ["SubscriptionName", "Region", "ResourceName", "DeploymentName", "ModelName", "ModelVersion", "DeploymentSku",
                "Capacity", "TokensPerMinuteLimit", "RequestsPerMinuteLimit", "QuotaName",
                "Tokens_P95", "Tokens_P99", "Tokens_Max", "Tokens_MaxTimestamp", "PctOfTPM_P95", "PctOfTPM_P99", "PctOfTPM_Max",
                "Input_P99", "Input_Max", "Input_Sum", "Output_P99", "Output_Max", "Output_Sum",
                "Requests_P99", "Requests_Max", "Requests_Sum", "PctOfRPM_P99", "PctOfRPM_Max",
                "Throttled429_Sum", "Throttled429_ActiveMinutes", "Errors4xx_Sum", "Errors5xx_Sum",
                "Tokens_CompletenessPct", "ResourceId"]
    dep = dep[dep_cols].sort_values(["PctOfTPM_Max", "Tokens_Max"], ascending=False, na_position="last")
    dep = dep.reset_index(drop=True)
    dep = add_status(con, dep, window_minutes)

    # Quota: sum of all deployments sharing a subscription/region quota, per minute, vs the quota limit.
    # Capacity units are converted to TPM/RPM with the ratio observed on the deployments themselves.
    deps = deps.assign(TpmCap=deps["Capacity"].where(deps["TokensPerMinuteLimit"].notna()),
                       RpmCap=deps["Capacity"].where(deps["RequestsPerMinuteLimit"].notna()))
    agg = deps.groupby(["SubscriptionId", "Region", "QuotaName"], as_index=False).agg(
        Deployments=("DeploymentName", lambda x: ", ".join(sorted(x))),
        AssignedCapacity=("Capacity", "sum"), AssignedTPM=("TokensPerMinuteLimit", "sum"),
        AssignedRPM=("RequestsPerMinuteLimit", "sum"), TpmCap=("TpmCap", "sum"), RpmCap=("RpmCap", "sum"))
    agg["TokensPerUnit"] = agg["AssignedTPM"] / agg["TpmCap"].where(agg["TpmCap"] > 0)
    agg["RequestsPerUnit"] = agg["AssignedRPM"] / agg["RpmCap"].where(agg["RpmCap"] > 0)
    q = quota.merge(agg.drop(columns=["TpmCap", "RpmCap"]), on=["SubscriptionId", "Region", "QuotaName"], how="right")
    qp = _pivot(full, "Quota", ["SubscriptionId", "Region", "QuotaName"])
    q = q.merge(qp, on=["SubscriptionId", "Region", "QuotaName"], how="outer")
    sub_names = deps.groupby("SubscriptionId")["SubscriptionName"].first()
    q["SubscriptionName"] = q["SubscriptionName"].fillna(q["SubscriptionId"].map(sub_names))
    q["DeploymentType"] = q["DeploymentType"].fillna(q["QuotaName"].str.split(".").str[1])
    ptu = q["DeploymentType"].fillna("").str.contains("Provisioned")
    q["QuotaTPM"] = (q["Limit"] * q["TokensPerUnit"]).where(~ptu)
    q["QuotaRPM"] = (q["Limit"] * q["RequestsPerUnit"]).where(~ptu)
    q["AssignedPct"] = _pct(q["CurrentValue"], q["Limit"])
    # Quota CurrentValue covers the whole subscription/region; if discovered deployments don't add up to it,
    # some consumers of this quota were not extracted (filtered scope or no access).
    q["AllConsumersExtracted"] = ((q["AssignedCapacity"] - q["CurrentValue"]).abs() < 1e-6).astype(object)
    q.loc[q["CurrentValue"].isna() | q["AssignedCapacity"].isna(), "AllConsumersExtracted"] = None
    for s in ("P95", "P99", "Max"):
        q[f"PctOfQuotaTPM_{s}"] = _pct(q[f"Tokens_{s}"], q["QuotaTPM"])
        q[f"PctOfAssignedTPM_{s}"] = _pct(q[f"Tokens_{s}"], q["AssignedTPM"].where(~ptu))
    q["PctOfQuotaRPM_Max"] = _pct(q["Requests_Max"], q["QuotaRPM"])
    q_cols = ["SubscriptionName", "Region", "QuotaName", "DeploymentType", "Deployments",
              "Limit", "CurrentValue", "AssignedPct", "AssignedCapacity", "AllConsumersExtracted", "Unit", "TokensPerUnit", "QuotaTPM", "AssignedTPM",
              "Tokens_P95", "Tokens_P99", "Tokens_Max", "Tokens_MaxTimestamp",
              "PctOfQuotaTPM_P95", "PctOfQuotaTPM_P99", "PctOfQuotaTPM_Max",
              "PctOfAssignedTPM_P95", "PctOfAssignedTPM_P99", "PctOfAssignedTPM_Max",
              "RequestsPerUnit", "QuotaRPM", "Requests_P99", "Requests_Max", "PctOfQuotaRPM_Max",
              "Throttled429_Sum", "Throttled429_ActiveMinutes", "Tokens_CompletenessPct", "SubscriptionId"]
    q = q[q_cols].sort_values(["AssignedPct", "PctOfQuotaTPM_Max"], ascending=False, na_position="last")
    return dep, q


def add_status(con: duckdb.DuckDBPyConnection, dep, window_minutes: int):
    """Data status per deployment from manifest coverage (Success/NotSupported, idle pre-flight included).

    'No traffic (confirmed)' requires every minute of the window to be covered for both the Tokens and the
    Requests calls of the resource; absence of metrics alone is reported as 'No data (incomplete coverage)'.
    """
    cov = con.execute(
        """
        SELECT rid, call, sum(date_diff('minute', seg.s, seg.e)) AS m
        FROM rcov JOIN seg USING (s) WHERE call IN ('Tokens', 'Requests') GROUP BY rid, call
        """
    ).df()
    piv = cov.pivot_table(index="rid", columns="call", values="m", aggfunc="sum")
    rid = dep["ResourceId"].str.lower()
    for call in ("Tokens", "Requests"):
        s = piv[call] if call in piv.columns else None
        dep[f"{call}CoveragePct"] = (100.0 * rid.map(s).fillna(0) / window_minutes).round(2) if s is not None else 0.0

    ptu = dep["DeploymentSku"].fillna("").str.contains("Provisioned")
    no_tpm = dep["TokensPerMinuteLimit"].isna() & ~ptu
    unmapped = dep["QuotaName"].fillna("").str.startswith("Unmapped")
    no_data = dep["Tokens_Max"].isna() & dep["Requests_Max"].isna()
    full_cov = dep[["TokensCoveragePct", "RequestsCoveragePct"]].min(axis=1) >= 100
    st = dep["Tokens_Max"].isna().map({True: "Token metric not available", False: "With metrics"})
    st[no_tpm & ~dep["Tokens_Max"].isna()] = "No TPM limit (RPM only)"
    st[~full_cov] = "Incomplete coverage"
    st[no_data & full_cov] = "No traffic (confirmed)"
    st[no_data & ~full_cov] = "No data (incomplete coverage)"
    st[unmapped] = "Deployment deleted (Unmapped)"
    dep["DataStatus"] = st
    return dep


def latency_summary(con: duckdb.DuckDBPyConnection):
    """Per deployment: P95 of the 1-minute average TTFT/TTLT (minutes with requests only), availability min/avg.
    Minutes with several series (e.g. model version change) are first collapsed to one conservative value:
    the worst latency, the lowest availability and the average tokens/s."""
    return con.execute(
        f"""
        WITH m AS (
            SELECT rid, dep, ts, metric,
                   CASE WHEN metric = '{AVAIL}' THEN min(v) WHEN metric = '{TPS}' THEN avg(v) ELSE max(v) END AS v
            FROM pts WHERE metric IN {LATENCY_METRICS} AND agg = 'Average'
            GROUP BY rid, dep, ts, metric
        )
        SELECT rid, dep AS DeploymentName,
               quantile_disc(v, 0.95) FILTER (WHERE metric = '{TTR}') AS TTR_P95,
               quantile_disc(v, 0.95) FILTER (WHERE metric = '{TTLT}') AS TTLT_P95,
               quantile_disc(v, 0.50) FILTER (WHERE metric = '{TPS}') AS TPS_P50,
               min(v) FILTER (WHERE metric = '{AVAIL}') AS Avail_Min,
               avg(v) FILTER (WHERE metric = '{AVAIL}') AS Avail_Avg,
               count(DISTINCT ts) FILTER (WHERE metric = '{AVAIL}') AS LatencyMinutes
        FROM m GROUP BY rid, dep
        """
    ).df()


# ---------------------------------------------------------------------------------------------------------------
# Prices (list-price PAYG estimate)
# ---------------------------------------------------------------------------------------------------------------
PRICE_COLS = ["Model", "DeploymentType", "Region", "InputPer1M", "OutputPer1M", "Currency", "Source"]
RETAIL_API = "https://prices.azure.com/api/retail/prices"
# Non-standard tiers (cached, batch, priority, flex, short/long context, fine-tuning...) are never candidates.
_EXCLUDE = re.compile(r"\b(cd|cached|cache|batch|pp|pri|priority|longco|shortco|loco|shco|flex|fl|ft|fine|tuning|"
                      r"training|hosting|hosted|grdr|"
                      r"realtime|audio|image|img|vision|embedding|aud)\b", re.I)
_DIR = [("Input", re.compile(r"\b(inp|input)\b", re.I)), ("Output", re.compile(r"\b(outp|opt|output|out)\b", re.I))]
_TYPE = [("DataZone", re.compile(r"\b(data\s*zone|dzone|dz)\b", re.I)), ("Global", re.compile(r"\b(glbl|gl|global)\b", re.I)),
         ("Regional", re.compile(r"\b(regnl|rgnl|regional)\b", re.I))]


def norm_model(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def deployment_type(sku) -> str | None:
    s = str(sku or "")
    if "Provisioned" in s or "Batch" in s:
        return None
    if s.startswith("DataZone"):
        return "DataZone"
    if s.startswith("Global"):
        return "Global"
    if s == "Standard":
        return "Regional"
    return None


def parse_meter(item: dict):
    """(model_key, type, direction, price_per_1M) for unambiguous token meters, else None. Never guesses."""
    name = item.get("meterName") or ""
    if _EXCLUDE.search(name.replace("-", " ")):
        return None
    words = name.replace("-", " ")
    dirs = [d for d, rx in _DIR if rx.search(words)]
    types = [t for t, rx in _TYPE if rx.search(words)]
    if len(dirs) != 1 or len(types) != 1:
        return None
    m = re.search(r"(\d+)\s*([KM])", item.get("unitOfMeasure") or "", re.I)
    if not m:
        return None
    per = float(m.group(1)) * (1_000 if m.group(2).upper() == "K" else 1_000_000)
    first = min(rx.search(words).start() for _, rx in _DIR if rx.search(words))
    prefix = words[:first].strip()
    if not prefix:
        return None
    # Meters like "5.4 inp Gl" carry the family only in productName ("Azure OpenAI GPT5").
    if prefix[0].isdigit() and re.search(r"\bGPT", item.get("productName") or ""):
        prefix = "gpt " + prefix
    return norm_model(prefix), types[0], dirs[0], float(item["retailPrice"]) * 1_000_000 / per


def fetch_prices(regions: list[str], currency: str):
    """Candidate prices from the Azure Retail Prices API (public, unauthenticated). Ambiguous meters are dropped."""
    import urllib.parse
    import urllib.request
    import pandas as pd
    rows = []
    for region in sorted({r for r in regions if r}):
        flt = f"serviceName eq 'Foundry Models' and armRegionName eq '{region}' and type eq 'Consumption'"
        url = f"{RETAIL_API}?currencyCode='{currency}'&$filter={urllib.parse.quote(flt)}"
        while url:
            with urllib.request.urlopen(url, timeout=60) as r:
                page = json.loads(r.read().decode("utf-8"))
            for it in page.get("Items", []):
                p = parse_meter(it)
                if p:
                    rows.append({"key": p[0], "DeploymentType": p[1], "dir": p[2], "price": p[3], "Region": region,
                                 "Currency": it.get("currencyCode"), "meter": it.get("meterName"),
                                 "eff": it.get("effectiveStartDate", "")})
            url = page.get("NextPageLink")
    df = pd.DataFrame(rows, columns=["key", "DeploymentType", "dir", "price", "Region", "Currency", "meter", "eff"])
    if df.empty:
        return df
    # Latest effective price per key; several distinct prices for the same key = ambiguous -> dropped.
    df = df.sort_values("eff").groupby(["key", "DeploymentType", "dir", "Region", "eff"], as_index=False).agg(
        price=("price", lambda x: x.iloc[0] if x.nunique() == 1 else math.nan), meter=("meter", "first"),
        Currency=("Currency", "first"))
    df = df.groupby(["key", "DeploymentType", "dir", "Region"], as_index=False).last()
    return df.dropna(subset=["price"])


def resolve_prices(dep, prices_file: str | None, fetch: bool, currency: str, out: Path, ts: str):
    """Attach InputPer1M/OutputPer1M per deployment. User file wins; fetched candidates only on exact model match."""
    import pandas as pd
    d = dep.assign(_key=dep["ModelName"].map(norm_model), _type=dep["DeploymentSku"].map(deployment_type),
                   _region=dep["Region"].fillna("").str.lower())
    d["InputPer1M"] = math.nan
    d["OutputPer1M"] = math.nan
    d["Currency"] = None
    d["PriceSource"] = None
    if fetch:
        cand = fetch_prices(d["_region"].unique().tolist(), currency)
        if len(cand):
            wide = cand.pivot_table(index=["key", "DeploymentType", "Region"], columns="dir", values="price",
                                    aggfunc="first").reset_index()
            meters = cand.pivot_table(index=["key", "DeploymentType", "Region"], columns="dir", values="meter",
                                      aggfunc="first").reset_index()
            wide = wide.merge(meters, on=["key", "DeploymentType", "Region"], suffixes=("", "Meter"))
            for c in ("Input", "Output", "InputMeter", "OutputMeter"):
                if c not in wide.columns:
                    wide[c] = None
            hit = d[["_key", "_type", "_region", "ModelName"]].drop_duplicates().merge(
                wide, left_on=["_key", "_type", "_region"], right_on=["key", "DeploymentType", "Region"])
            cand_out = hit.rename(columns={"Input": "InputPer1M", "Output": "OutputPer1M"}).assign(
                Model=hit["ModelName"], Currency=currency, Source="RetailAPI (candidate, review)")
            cand_out[PRICE_COLS + ["InputMeter", "OutputMeter"]].to_csv(out / f"prices-candidates_{ts}.csv", index=False)
            for c, s in (("InputPer1M", "Input"), ("OutputPer1M", "Output")):
                m = d.merge(wide, left_on=["_key", "_type", "_region"], right_on=["key", "DeploymentType", "Region"],
                            how="left")[s]
                d[c] = m.values
            d.loc[d["InputPer1M"].notna() | d["OutputPer1M"].notna(), ["Currency", "PriceSource"]] = [currency, "RetailAPI"]
    if prices_file:
        p = pd.read_csv(prices_file, dtype=str).fillna("")
        missing = set(PRICE_COLS[:5]) - set(p.columns)
        if missing:
            sys.exit(f"{prices_file}: missing columns {sorted(missing)} (expected {PRICE_COLS})")
        for r in p.itertuples():
            if not (str(r.InputPer1M).strip() or str(r.OutputPer1M).strip()):
                continue
            sel =  (d["_key"] == norm_model(r.Model)) & (d["_type"] == r.DeploymentType) & \
                  ((r.Region in ("", "*")) | (d["_region"] == str(r.Region).lower()))
            for c in ("InputPer1M", "OutputPer1M"):
                v = getattr(r, c)
                if str(v).strip():
                    d.loc[sel, c] = float(v)
            d.loc[sel, "Currency"] = getattr(r, "Currency", "") or currency
            d.loc[sel, "PriceSource"] = "File"
    return d.drop(columns=["_key", "_type", "_region"])


def add_costs(dep, window_minutes: int):
    cov_min = dep["TokensCoveragePct"].fillna(0) / 100.0 * window_minutes
    dep["InputCost"] = dep["Input_Sum"].fillna(0) / 1e6 * dep["InputPer1M"]
    dep["OutputCost"] = dep["Output_Sum"].fillna(0) / 1e6 * dep["OutputPer1M"]
    dep["EstCost"] = dep["InputCost"] + dep["OutputCost"]
    dep["EstMonthlyCost"] = dep["EstCost"] / cov_min.where(cov_min > 0) * MINUTES_PER_MONTH
    cost_cols = ["InputCost", "OutputCost", "EstCost", "EstMonthlyCost"]
    dep[cost_cols] = dep[cost_cols].round(4)
    return dep


THRESHOLDS = {"high_pct": 80.0, "low_pct": 10.0, "quota_pct": 90.0, "ptu_low": 40.0, "ptu_high": 90.0, "avail_min": 99.0,
              "client_err_pct": 5.0}


def recommend(m, quota_util, th: dict = THRESHOLDS):
    """Rule-based, explainable recommendations (one string per deployment, rules joined with '; ')."""
    import pandas as pd
    q = quota_util[["SubscriptionId", "Region", "QuotaName", "AssignedPct", "PctOfQuotaTPM_Max"]]
    m = m.assign(SubscriptionId=m["ResourceId"].str.lower().str.split("/").str[2]).merge(
        q, on=["SubscriptionId", "Region", "QuotaName"], how="left")
    ptu = m["DeploymentSku"].fillna("").str.contains("Provisioned")
    thr = m["Throttled429_Sum"].fillna(0) > 0
    pct_max, pct_p99 = m["PctOfTPM_Max"], m["PctOfTPM_P99"]
    rules = [
        (m["DataStatus"] == "No traffic (confirmed)", "No traffic in window: deletion candidate (frees quota)"),
        (m["DataStatus"] == "No data (incomplete coverage)", "Incomplete data: re-run extraction before deciding"),
        (thr & (pct_max >= th["high_pct"]), "429 at TPM limit: increase capacity or spread load"),
        (thr & ~(pct_max >= th["high_pct"]), "429 below TPM limit: check RPM, bursts and max_tokens"),
        (~ptu & ~thr & (pct_max < 50) & (pct_p99 < th["low_pct"]), "Low use vs TPM limit: reduce capacity to free quota"),
        (ptu & (m["Util_P99"] > th["ptu_high"]), "PTU near saturation: add PTU or spillover"),
        (ptu & (m["Util_P99"] < th["ptu_low"]), "PTU under-utilized: reduce PTU"),
        ((m["AssignedPct"] > th["quota_pct"]) | (m["PctOfQuotaTPM_Max"] > th["quota_pct"]),
         f"Quota > {th['quota_pct']:.0f}%: request increase or use another region/Data Zone"),
        ((m["Errors5xx_Sum"].fillna(0) > 0) | (m["Avail_Min"] < th["avail_min"]), "Service errors: review 5xx/availability"),
        (_pct(m["Errors4xx_Sum"], m["Requests_Sum"]) > th["client_err_pct"],
         f"Client errors (4xx, not 429) > {th['client_err_pct']:.0f}% of requests: review callers"),
    ]
    rec = pd.Series([[] for _ in range(len(m))], index=m.index)
    for mask, text in rules:
        for i in m.index[mask.fillna(False)]:
            rec[i].append(text)
    deleted = m["DataStatus"] == "Deployment deleted (Unmapped)"
    m["Recommendation"] = ["; ".join(r) if r else ("" if dl else "OK") for r, dl in zip(rec, deleted)]
    return m.drop(columns=["SubscriptionId", "AssignedPct", "PctOfQuotaTPM_Max"])


def run_meta(run_dirs: list[Path]):
    """run.json metadata of each extraction folder (one row per run)."""
    import pandas as pd

    def _join(v):
        return v if isinstance(v, str) else ", ".join(v or [])

    runs = []
    for d in run_dirs:
        r = json.loads((d / "run.json").read_text(encoding="utf-8-sig"))
        fp = r.get("Fingerprint", {})
        runs.append({"RunName": d.name, "RunTimestamp": r.get("RunTimestamp", ""),
                     "Tenants": _join(fp.get("Tenants")), "FailedTenants": _join(r.get("FailedTenants")),
                     "TargetResources": len(r.get("TargetResources") or [])})
    return pd.DataFrame(runs)


NA = "N/A"
MODEL_COLS = {
    "SubscriptionName": "Subscription", "Region": "Region", "ResourceName": "Account", "DeploymentName": "Deployment",
    "ModelName": "Model", "ModelVersion": "Version", "DeploymentSku": "SKU", "Capacity": "Capacity",
    "TokensPerMinuteLimit": "TPM limit", "RequestsPerMinuteLimit": "RPM limit",
    "Input_P95": "Input TPM P95", "Input_P99": "Input TPM P99", "Output_P95": "Output TPM P95", "Output_P99": "Output TPM P99",
    "Tokens_P99": "TPM P99", "Tokens_Max": "TPM max", "Tokens_MaxTimestamp": "TPM peak (UTC)",
    "PctOfTPM_P99": "% TPM P99", "PctOfTPM_Max": "% TPM max",
    "Requests_P99": "RPM P99", "Requests_Max": "RPM max", "PctOfRPM_Max": "% RPM max",
    "Throttled429_Sum": "429 requests", "Errors4xx_Sum": "Other 4xx", "Errors5xx_Sum": "5xx", "ErrorRatePct": "Error rate %",
    "TTR_P95": "Time to response P95 ms (1-min avg)", "TTLT_P95": "Time to last token P95 ms (1-min avg)",
    "Avail_Min": "Availability min %",
    "Util_P95": "PTU utilization P95 %", "Util_P99": "PTU utilization P99 %", "Util_Max": "PTU utilization max %",
    "EstCost": "Est. cost (window)", "EstMonthlyCost": "Est. cost / 30 days", "Currency": "Currency",
    "TokensCoveragePct": "Coverage %", "DataStatus": "Data status", "Recommendation": "Recommendation",
}
QUOTA_XLS_COLS = {
    "SubscriptionName": "Subscription", "Region": "Region", "QuotaName": "Quota", "DeploymentType": "Type",
    "Deployments": "Deployments", "Limit": "Limit", "CurrentValue": "Assigned", "AssignedPct": "% assigned",
    "QuotaTPM": "Quota TPM", "Tokens_P99": "TPM P99", "Tokens_Max": "TPM max", "Tokens_MaxTimestamp": "TPM peak (UTC)",
    "PctOfQuotaTPM_P99": "% quota P99", "PctOfQuotaTPM_Max": "% quota max", "Throttled429_Sum": "429 requests",
    "Tokens_CompletenessPct": "Coverage %", "AllConsumersExtracted": "All consumers extracted",
}


def enrich_deployments(stats, dep_lim, lat, quota_util):
    """Per deployment: limits + input/output P95 + PTU utilization + latency/availability + recommendation."""
    full = stats[stats["Mode"] == "FullWindow"].copy()
    full["rid"] = full["ResourceId"].str.lower()
    keys = ["rid", "DeploymentName"]
    m = dep_lim.assign(rid=dep_lim["ResourceId"].str.lower())
    for metric, prefix in (("ProcessedPromptTokens", "Input"), ("GeneratedTokens", "Output")):
        s = full[(full["Level"] == "DeploymentLimit") & (full["Metric"] == metric)][keys + ["P95"]]
        m = m.merge(s.rename(columns={"P95": f"{prefix}_P95"}), on=keys, how="left")
    u = full[(full["Level"] == "Deployment") & (full["Metric"] == UTIL) & (full["Aggregation"] == "Maximum")]
    u = u.groupby(keys, as_index=False)[["P95", "P99", "Max"]].max()
    m = m.merge(u.rename(columns={c: f"Util_{c}" for c in ("P95", "P99", "Max")}), on=keys, how="left")
    m = m.merge(lat, on=keys, how="left")
    lat_cols = ["TTR_P95", "TTLT_P95", "TPS_P50", "Avail_Min", "Avail_Avg"]
    m[lat_cols] = m[lat_cols].astype(float).round(1)
    errs = m[["Throttled429_Sum", "Errors4xx_Sum", "Errors5xx_Sum"]].sum(axis=1, min_count=1)
    m["ErrorRatePct"] = _pct(errs, m["Requests_Sum"])
    return recommend(m, quota_util).drop(columns=["rid"])


def model_sheet(m):
    """Excel view: one row per deployment, minimum columns, ordered by pressure."""
    import pandas as pd
    ptu = m["DeploymentSku"].fillna("").str.contains("Provisioned")
    no_tpm = m["TokensPerMinuteLimit"].isna() & ~ptu
    idle = m["DataStatus"].isin(["No traffic (confirmed)", "No data (incomplete coverage)", "Deployment deleted (Unmapped)"])
    rank = pd.Series(1, index=m.index)
    rank[ptu] = 2
    rank[no_tpm] = 3
    rank[idle] = 4
    rank[m["Throttled429_Sum"].fillna(0) > 0] = 0
    m = m.assign(_rank=rank, _key=m["PctOfTPM_Max"].where(~ptu, m["Util_Max"]).fillna(m["PctOfRPM_Max"]))
    m = m.sort_values(["_rank", "_key", "Tokens_Max"], ascending=[True, False, False], na_position="last")

    out = m[list(MODEL_COLS)].astype(object)
    ptu, no_tpm = ptu.reindex(out.index), no_tpm.reindex(out.index)
    for c in ("PctOfTPM_P99", "PctOfTPM_Max"):
        out.loc[ptu | no_tpm, c] = NA
    for c in ("Util_P95", "Util_P99", "Util_Max"):
        out.loc[~ptu, c] = NA
    for c in ("EstCost", "EstMonthlyCost"):
        out.loc[ptu, c] = NA
    return out.rename(columns=MODEL_COLS).reset_index(drop=True)

def quota_sheet(quota_util):
    q = quota_util[list(QUOTA_XLS_COLS)].astype(object)
    ptu = quota_util["DeploymentType"].fillna("").str.contains("Provisioned").values
    for c in ("QuotaTPM", "PctOfQuotaTPM_P99", "PctOfQuotaTPM_Max"):
        q.loc[ptu, c] = NA
    q["AllConsumersExtracted"] = q["AllConsumersExtracted"].map({True: "Yes", False: "No"})
    return q.rename(columns=QUOTA_XLS_COLS).reset_index(drop=True)


def write_excel(path: Path, info: dict, runs, models, quotas, generated_at: str, ts: str, price_note: str) -> None:
    """Compact report: Summary + Models (one row per deployment) + Quotas. No charts; detail stays in the CSVs."""
    import pandas as pd
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    cov = pd.to_numeric(models["Coverage %"], errors="coerce")
    monthly = pd.to_numeric(models["Est. cost / 30 days"], errors="coerce")
    # Never add amounts in different currencies: one total per currency.
    by_cur = monthly.groupby(models["Currency"].fillna("?")).sum()
    cost_txt = "No prices supplied (--prices / --fetch-prices)" if monthly.isna().all() else \
        "; ".join(f"{v:,.2f} {c}" for c, v in by_cur.items()) + \
        f" ({int(monthly.notna().sum())} of {len(models)} deployments priced)"
    recs = models["Recommendation"].fillna("").str.split("; ").explode()
    rec_rows = [(f"  {k}", int(v)) for k, v in recs[recs.ne("") & recs.ne("OK")].value_counts().items()] or [("  None", "")]
    summary = [
        ("Report generated (UTC)", generated_at),
        ("Analysed window (UTC)", f"{info['start']} -> {info['end']} ({info['window_minutes']:,} min)"),
        ("Runs", ", ".join(f"{r.RunName}" for r in runs.itertuples())),
        ("Tenants", ", ".join(sorted(set(", ".join(runs["Tenants"]).split(", ")) - {""}))),
        ("Failed tenants", ", ".join(sorted(set(", ".join(runs["FailedTenants"]).split(", ")) - {""})) or "None"),
        ("Accounts with OpenAI metrics", len(info["targets"])),
        ("Deployments", len(models)),
        ("Deployments with 429", int((pd.to_numeric(models["429 requests"], errors="coerce") > 0).sum())),
        ("Minimum coverage %", "" if cov.isna().all() else float(cov.min())),
        ("Deployments with no traffic (confirmed)", int((models["Data status"] == "No traffic (confirmed)").sum())),
        ("Deployments with 5xx", int((pd.to_numeric(models["5xx"], errors="coerce") > 0).sum())),
        ("Est. cost / 30 days (PAYG list price)", cost_txt),
        ("Prices", price_note),
        ("", ""),
        ("Recommendations", ""),
        *rec_rows,
        ("", ""),
        ("How to read it", ""),
        ("TPM / RPM", "Tokens and requests per minute (Total aggregation at PT1M). Input and output are shown separately."),
        ("P95 / P99", "Computed over every observed minute of the window, including zero minutes."),
        ("% TPM / % RPM", "Relative to the deployment rate limit (current snapshot). N/A for PTU and for models without a TPM limit."),
        ("PTU utilization", "Provisioned-managed Utilization V2 (maximum per minute). Only applies to PTU deployments."),
        ("429", "Direct saturation signal: the rate limiter counts estimated tokens (prompt + max_tokens), "
                "so the real pressure can be higher than the % shown."),
        ("Quotas", "Per-minute sum of all deployments sharing a quota (subscription + region/data zone + type + model)."),
        ("Latency", "P95 over minutes with requests of the 1-minute AVERAGE (Azure Monitor does not expose per-request "
                    "percentiles). Time to response = time to first token for streaming requests. Availability = min "
                    "of the 1-minute availability rate. Errors: other 4xx exclude 429; error rate = (429+4xx+5xx)/requests."),
        ("Data status", "'No traffic (confirmed)' = every minute of the window was successfully queried for tokens and "
                        "requests and nothing was returned. 'No data (incomplete coverage)' = some chunks failed or were "
                        "not queried: absence of metrics is NOT proof of zero traffic."),
        ("Cost", "Estimate = input and output tokens x list PAYG price per 1M (--prices file or --fetch-prices candidates). "
                 "Cached-input discounts, batch, priority, fine-tuning and PTU commitments are NOT included. "
                 "'/ 30 days' extrapolates the covered minutes."),
        ("More detail", f"CSVs from this execution (_{ts}): foundry-token-stats, foundry-deployment-limits, "
                        "foundry-quota-utilization, foundry-token-daily, foundry-token-hourly; and in the run folder: "
                        "foundry-token-metrics (minutes), deployments, quota, resources, manifest."),
    ]
    if len(runs) > 1:
        summary.insert(5, ("Note", "Multiple runs: deployment/quota snapshots use the most recent one."))
    res = pd.DataFrame(summary, columns=["Item", "Value"])

    pct_cols = {c for df in (models, quotas) for c in df.columns if "%" in c}
    money_cols = {"Est. cost (window)", "Est. cost / 30 days"}
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        res.to_excel(xw, sheet_name="Summary", index=False)
        ws = xw.sheets["Summary"]
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 120
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
            if row[0].value == "How to read it":
                row[0].font = Font(bold=True)

        for name, df in (("Models", models), ("Quotas", quotas)):
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "B2" if name == "Models" else "A2"
            if len(df):
                t = Table(displayName=name, ref=f"A1:{get_column_letter(df.shape[1])}{len(df) + 1}")
                t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
                ws.add_table(t)
            for i, col in enumerate(df.columns, 1):
                letter = get_column_letter(i)
                width = max([len(str(col))] + [len(str(x)) for x in df[col].head(200) if x is not None]) + 2
                ws.column_dimensions[letter].width = min(45, max(10, width))
                ws.cell(row=1, column=i).alignment = Alignment(wrap_text=True, vertical="top")
                fmt = '0.00"%"' if col in pct_cols else "#,##0.00" if col in money_cols else "#,##0"
                for (c,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    if isinstance(c.value, (int, float)):
                        c.number_format = fmt
                    elif c.value == NA:
                        c.alignment = Alignment(horizontal="right")
            ws.row_dimensions[1].height = 32


def main() -> None:
    a = parse_args()
    run_dirs = [Path(d).resolve() for d in a.run_dir]
    out = Path(a.out_dir).resolve() if a.out_dir else run_dirs[0]
    out.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect(a.db)
    info = load(con, run_dirs)
    raw_rows = con.execute("SELECT count(*) FROM raw").fetchone()[0]
    print(f"Window {info['start']} -> {info['end']} ({info['window_minutes']} min); raw rows: {raw_rows:,}")
    build(con)
    stats = compute_stats(con, info["window_minutes"])
    day = daily(con)
    hour = hourly(con)
    dep_lim, quota_util = limit_reports(con, stats, info["window_minutes"])

    from datetime import datetime, timezone
    now = datetime.now(timezone.utc)
    ts = now.strftime("%Y%m%d-%H%M%SZ")
    generated_at = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    dep_lim = add_costs(resolve_prices(dep_lim, a.prices, a.fetch_prices, a.currency, out, ts), info["window_minutes"])
    dep_lim = enrich_deployments(stats, dep_lim, latency_summary(con), quota_util)
    in_p, out_p = dep_lim["InputPer1M"].notna(), dep_lim["OutputPer1M"].notna()
    payg = dep_lim["DeploymentSku"].map(deployment_type).notna()
    unpriced = sorted(set(dep_lim.loc[payg & ~in_p & ~out_p, "ModelName"].dropna()) |
                      {f"{m} (input or output price missing)" for m in dep_lim.loc[payg & (in_p ^ out_p), "ModelName"].dropna()})
    price_note = ("Not requested" if not (a.prices or a.fetch_prices) else
                  f"Sources: {', '.join(sorted(dep_lim['PriceSource'].dropna().unique())) or 'none matched'}. "
                  f"Unpriced models (add them to --prices): {', '.join(unpriced) or 'none'}")
    if unpriced and (a.prices or a.fetch_prices):
        print(f"Prices: no unambiguous price for {', '.join(unpriced)} (add them to a --prices file).")

    stats_path = out / f"foundry-token-stats_{ts}.csv"
    stats.to_csv(stats_path, index=False)
    day.to_csv(out / f"foundry-token-daily_{ts}.csv", index=False)
    hour.to_csv(out / f"foundry-token-hourly_{ts}.csv", index=False)
    dep_lim.to_csv(out / f"foundry-deployment-limits_{ts}.csv", index=False)
    quota_util.to_csv(out / f"foundry-quota-utilization_{ts}.csv", index=False)
    print(f"Stats: {len(stats):,} rows -> {stats_path}")
    print(f"Quota utilization: {len(quota_util):,} quota(s), {len(dep_lim):,} deployment(s)")
    if not con.execute("SELECT count(*) FROM deps").fetchone()[0]:
        print("WARNING: deployments.csv/quota.csv not found (run exported with an older script?): limits and quota % are empty.")

    inc = stats[(stats["Mode"] == "FullWindow") & (stats["CompletenessPct"] < 100)]
    if len(inc):
        print(f"WARNING: {len(inc)} FullWindow group(s) have CompletenessPct < 100 (failed chunks excluded).")

    if not a.no_excel:
        xlsx = out / f"foundry-token-report_{ts}.xlsx"
        write_excel(xlsx, info, run_meta(run_dirs), model_sheet(dep_lim), quota_sheet(quota_util), generated_at, ts, price_note)
        print(f"Excel: {xlsx}")


if __name__ == "__main__":
    main()
