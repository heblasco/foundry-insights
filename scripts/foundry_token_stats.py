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
    call_of = (f"CASE WHEN metric IN {TOKEN_METRICS} OR metric = '{DERIVED}' THEN 'Tokens' "
               f"WHEN metric IN ('AzureOpenAIRequests', '{THROTTLED}') THEN 'Requests' ELSE 'PtuUtilization' END")
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
        FROM seg JOIN manifest m ON m.Status = 'Success' AND m.cs <= seg.s AND seg.e <= m.ce
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
               greatest(coalesce(s.k, 0), CASE m.mode WHEN 'FullWindow' THEN gm.cov
                                                      WHEN 'Lifetime' THEN coalesce(l.life, 0)
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
        ("ProcessedPromptTokens", "Input", ["P99", "Max"]),
        ("GeneratedTokens", "Output", ["P99", "Max"]),
        ("AzureOpenAIRequests", "Requests", ["P99", "Max"]),
        (THROTTLED, "Throttled429", ["Sum", "ActiveMinutes"]),
    ]
    out = None
    for metric, prefix, cols in spec:
        m = f[f["Metric"] == metric][keys + cols].rename(columns={c: f"{prefix}_{c}" for c in cols})
        out = m if out is None else out.merge(m, on=keys, how="outer")
    # Sparse extraction emits no 429 rows when there was no throttling: zero, not unknown, if requests were covered.
    for c in ("Throttled429_Sum", "Throttled429_ActiveMinutes"):
        out[c] = out[c].where(out[c].notna() | out["Requests_Max"].isna(), 0)
    return out


def _pct(num, den):
    import numpy as np
    return (100.0 * num / den.where(den > 0)).round(2).replace([np.inf, -np.inf], np.nan)


def limit_reports(con: duckdb.DuckDBPyConnection, stats):
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
                "Input_P99", "Input_Max", "Output_P99", "Output_Max",
                "Requests_P99", "Requests_Max", "PctOfRPM_P99", "PctOfRPM_Max",
                "Throttled429_Sum", "Throttled429_ActiveMinutes", "Tokens_CompletenessPct", "ResourceId"]
    dep = dep[dep_cols].sort_values(["PctOfTPM_Max", "Tokens_Max"], ascending=False, na_position="last")

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
    "Throttled429_Sum": "429 requests",
    "Util_P95": "PTU utilization P95 %", "Util_P99": "PTU utilization P99 %", "Util_Max": "PTU utilization max %",
    "Tokens_CompletenessPct": "Coverage %", "Status": "Data status",
}
QUOTA_XLS_COLS = {
    "SubscriptionName": "Subscription", "Region": "Region", "QuotaName": "Quota", "DeploymentType": "Type",
    "Deployments": "Deployments", "Limit": "Limit", "CurrentValue": "Assigned", "AssignedPct": "% assigned",
    "QuotaTPM": "Quota TPM", "Tokens_P99": "TPM P99", "Tokens_Max": "TPM max", "Tokens_MaxTimestamp": "TPM peak (UTC)",
    "PctOfQuotaTPM_P99": "% quota P99", "PctOfQuotaTPM_Max": "% quota max", "Throttled429_Sum": "429 requests",
    "Tokens_CompletenessPct": "Coverage %", "AllConsumersExtracted": "All consumers extracted",
}


def model_sheet(stats, dep_lim):
    """One row per deployment: limits, per-minute input/output/total P95-P99-max, % of limit, 429, PTU utilization."""
    import pandas as pd
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

    ptu = m["DeploymentSku"].fillna("").str.contains("Provisioned")
    no_tpm = m["TokensPerMinuteLimit"].isna() & ~ptu
    unmapped = m["QuotaName"].fillna("").str.startswith("Unmapped")
    no_data = m["Tokens_Max"].isna() & m["Requests_Max"].isna()
    incomplete = m["Tokens_CompletenessPct"].fillna(100) < 100
    m["Status"] = "With metrics"
    m.loc[no_tpm, "Status"] = "No TPM limit (RPM only)"
    m.loc[m["Tokens_Max"].isna() & m["Requests_Max"].notna(), "Status"] = "Token metric not available"
    m.loc[incomplete, "Status"] = "Incomplete coverage"
    m.loc[no_data, "Status"] = "No traffic observed"
    m.loc[unmapped, "Status"] = "Deployment deleted (Unmapped)"

    rank = pd.Series(1, index=m.index)
    rank[ptu] = 2
    rank[no_tpm] = 3
    rank[no_data | unmapped] = 4
    rank[m["Throttled429_Sum"].fillna(0) > 0] = 0
    m["_rank"] = rank
    m["_key"] = m["PctOfTPM_Max"].where(~ptu, m["Util_Max"]).fillna(m["PctOfRPM_Max"])
    m = m.sort_values(["_rank", "_key", "Tokens_Max"], ascending=[True, False, False], na_position="last")

    out = m[list(MODEL_COLS)].astype(object)
    ptu, no_tpm = ptu.reindex(out.index), no_tpm.reindex(out.index)
    for c in ("PctOfTPM_P99", "PctOfTPM_Max"):
        out.loc[ptu | no_tpm, c] = NA
    for c in ("Util_P95", "Util_P99", "Util_Max"):
        out.loc[~ptu, c] = NA
    return out.rename(columns=MODEL_COLS).reset_index(drop=True)


def quota_sheet(quota_util):
    q = quota_util[list(QUOTA_XLS_COLS)].astype(object)
    ptu = quota_util["DeploymentType"].fillna("").str.contains("Provisioned").values
    for c in ("QuotaTPM", "PctOfQuotaTPM_P99", "PctOfQuotaTPM_Max"):
        q.loc[ptu, c] = NA
    q["AllConsumersExtracted"] = q["AllConsumersExtracted"].map({True: "Yes", False: "No"})
    return q.rename(columns=QUOTA_XLS_COLS).reset_index(drop=True)


def write_excel(path: Path, info: dict, runs, models, quotas, generated_at: str, ts: str) -> None:
    """Compact report: Summary + Models (one row per deployment) + Quotas. No charts; detail stays in the CSVs."""
    import pandas as pd
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    cov = pd.to_numeric(models["Coverage %"], errors="coerce")
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
        ("", ""),
        ("How to read it", ""),
        ("TPM / RPM", "Tokens and requests per minute (Total aggregation at PT1M). Input and output are shown separately."),
        ("P95 / P99", "Computed over every observed minute of the window, including zero minutes."),
        ("% TPM / % RPM", "Relative to the deployment rate limit (current snapshot). N/A for PTU and for models without a TPM limit."),
        ("PTU utilization", "Provisioned-managed Utilization V2 (maximum per minute). Only applies to PTU deployments."),
        ("429", "Direct saturation signal: the rate limiter counts estimated tokens (prompt + max_tokens), "
                "so the real pressure can be higher than the % shown."),
        ("Quotas", "Per-minute sum of all deployments sharing a quota (subscription + region/data zone + type + model)."),
        ("More detail", f"CSVs from this execution (_{ts}): foundry-token-stats, foundry-deployment-limits, "
                        "foundry-quota-utilization, foundry-token-daily, foundry-token-hourly; and in the run folder: "
                        "foundry-token-metrics (minutes), deployments, quota, resources, manifest."),
    ]
    if len(runs) > 1:
        summary.insert(5, ("Note", "Multiple runs: deployment/quota snapshots use the most recent one."))
    res = pd.DataFrame(summary, columns=["Item", "Value"])

    pct_cols = {c for df in (models, quotas) for c in df.columns if "%" in c}
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
                fmt = '0.00"%"' if col in pct_cols else "#,##0"
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
    dep_lim, quota_util = limit_reports(con, stats)

    from datetime import datetime, timezone
    now = datetime.now(timezone.utc)
    ts = now.strftime("%Y%m%d-%H%M%SZ")
    generated_at = now.strftime("%Y-%m-%dT%H:%M:%SZ")

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
        write_excel(xlsx, info, run_meta(run_dirs), model_sheet(stats, dep_lim), quota_sheet(quota_util), generated_at, ts)
        print(f"Excel: {xlsx}")


if __name__ == "__main__":
    main()
