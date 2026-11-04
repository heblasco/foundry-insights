# Foundry Insights – Azure AI Foundry / Azure OpenAI token consumption

This project extracts per-minute metrics from every accessible subscription and tenant and computes the statistics needed to size PTU (average, P95, P99 and maximum). It includes an Azure Monitor workbook and an Excel report.

| Component | File |
|---|---|
| One-command run (prerequisite checks + extraction + report) | `scripts\Run-FoundryInsights.ps1` |
| Extraction (PowerShell 7) | `scripts\Export-FoundryTokenMetrics.ps1` (+ `scripts\FoundryInsights.Common.ps1`) |
| Statistics, prices, recommendations and Excel (Python + DuckDB) | `scripts\foundry_token_stats.py` |
| Azure Monitor workbook | `workbooks\FoundryTokenUsage.workbook.json` |
| Price list template | `prices.example.csv` |
| Container image (PowerShell + Python + Azure CLI) | `Dockerfile` |

## Quick start

```powershell
pip install -r requirements.txt
az login --tenant <tenantId>
pwsh -File .\scripts\Run-FoundryInsights.ps1 -TenantId <tenantId> -Days 30 -FetchPrices
# -> .\output\foundry_<yyyyMMdd-HHmmss>Z\ with the extractor CSVs, the statistics CSVs and the Excel report
```

`Run-FoundryInsights.ps1` checks PowerShell 7, Python 3.10+ and the packages (`-InstallPythonDeps` installs them), then runs the extractor and `foundry_token_stats.py` on the new run folder. Options: `-TenantId`, `-SubscriptionId`, `-Days`, `-RunName`, `-OutputRoot`, `-AuthMode`, `-PricesFile`, `-FetchPrices`, `-NoExcel`, `-Python`, and `-ExportArgs` for any other extractor parameter: a hashtable when called in-process (`& .\scripts\Run-FoundryInsights.ps1 -ExportArgs @{ IncludePtuUtilization = $true; Parallel = 8 }`) or a JSON string with `pwsh -File` (`-ExportArgs '{"IncludePtuUtilization": true, "Parallel": 8}'`).

## Requirements

- PowerShell 7.2 or later (`pwsh`).
- One authentication provider (`-AuthMode`, default `Auto`):
  - `AzCli`: Azure CLI with a signed-in session (`az login`). Chosen by `Auto` when `az` is installed.
  - `AzPowerShell`: `Connect-AzAccount` (module `Az.Accounts`). Chosen by `Auto` when `az` is missing.
  - `ManagedIdentity`: Azure VM/VMSS (IMDS), App Service, Functions, Container Apps or ACI identity. Requires `-TenantId`; use `-ManagedIdentityClientId` for a user-assigned identity. Chosen by `Auto` only when neither of the above is available and `IDENTITY_ENDPOINT` is set.
  - `Auto` never switches provider after an authentication failure, so it never mixes identities.
- Python 3.10 or later: `pip install -r requirements.txt` (duckdb, pandas, openpyxl). Alternatively, use the container image (see [Container](#container)).
- Permissions on each subscription or resource:
  - **Reader**, for ARG, the account and the deployments.
  - **Monitoring Reader**, for the metrics.
  - Reader is enough to read platform metrics; Monitoring Reader is the explicit minimum role.
  - The fast `metrics:getBatch` API (default `-MetricsApi Batch`) needs the metrics read permission **at subscription scope**. Where it is denied (401/403/404) the script switches that subscription and region to the per-resource ARM metrics API automatically; results are the same, only slower.
  - Quota (`usages` and `models` per region) requires **Reader at subscription level**; without it, deployments are flagged `NoQuotaAccess` and the rest of the extraction continues.

## 1. Extraction

```powershell
# All visible tenants and subscriptions, last 30 days, 24 h chunks
pwsh -File .\scripts\Export-FoundryTokenMetrics.ps1 -IncludePtuUtilization

# 90 days in a specific tenant, with a fixed folder (allows resuming)
pwsh -File .\scripts\Export-FoundryTokenMetrics.ps1 -TenantId <tenantId> -Days 90 -ChunkHours 24 `
     -IncludePtuUtilization -OutputDirectory .\output\tenantA-90d

# Inventory only (no metrics)
pwsh -File .\scripts\Export-FoundryTokenMetrics.ps1 -InventoryOnly

# Periodic run (e.g. a daily scheduled task): each run creates its own timestamped folder
pwsh -File .\scripts\Export-FoundryTokenMetrics.ps1 -TenantId <tenantId> -Days 1 -IncludePtuUtilization -RunName mcap
# -> .\output\mcap_20261103-183206Z\
```

| Parameter | Default | Description |
|---|---|---|
| `-OutputDirectory` | `<OutputRoot>\<RunName>_<yyyyMMdd-HHmmss>Z` | Run folder. If omitted, a new one is created with the UTC timestamp of the run. Set it explicitly to resume a previous run. |
| `-OutputRoot` | `.\output` | Parent folder of the timestamped folders. |
| `-RunName` | `foundry` | Base folder name: it stays fixed and only the timestamp changes, which makes runs easy to inventory. |
| `-Days` | 30 | Window back from `-EndTime`. Maximum 93, the platform metrics retention. |
| `-StartTime` / `-EndTime` | start of the current UTC hour | Explicit UTC window: `[Start, End)`. Use the format `2026-10-05T00:00:00Z`. |
| `-ChunkHours` | 24 | Size of each query, between 1 and 168 (1 to 7 days). If Azure Monitor rejects it for size, it is split automatically. |
| `-TenantId` | all (ARM `GET /tenants`) | Tenants to process. Required with `-AuthMode ManagedIdentity`. |
| `-AuthMode` | `Auto` | `Auto`, `AzCli`, `AzPowerShell` or `ManagedIdentity` (see Requirements). |
| `-ManagedIdentityClientId` | — | Client ID of a user-assigned managed identity. |
| `-SubscriptionId` | all | Allow-list of subscriptions. |
| `-Kind` / `-IncludeAllKinds` | `OpenAI`, `AIServices` | Filter by `kind`. In every case the account is validated to expose OpenAI metrics; if it does not, it is flagged `NotOpenAI` in `resources.csv`. |
| `-IncludePtuUtilization` | off | Adds `Provisioned-managed Utilization V2` (Maximum and Average). |
| `-NoLatency` | off | Skips the latency/availability call (`AzureOpenAITimeToResponse`, `AzureOpenAITTLTInMS`, `AzureOpenAITokenPerSecond`, `AzureOpenAIAvailabilityRate`). Saves one call per batch and chunk. |
| `-Top` | 1000 | Series limit per call. If reached, the query is repeated per deployment to avoid truncation. |
| `-TokenTimeoutSeconds` | 60 | Maximum time to get a tenant token. Tenants that require interactive login (MFA/CA) are flagged as failed and the others continue. |
| `-MetricsApi` | `Batch` | `Batch` uses the regional `metrics:getBatch` data-plane API (up to 50 accounts of the same subscription and region per call). `Arm` forces the classic per-resource API. Both produce identical CSVs. |
| `-Parallel` | 12 | Concurrent workers (1–32). Requests are paced by shared rate limiters, so raising it only helps while latency, not the limiters, is the bottleneck. |
| `-NoSkipIdle` | off | Disables the idle pre-flight (see below) and extracts every chunk of every account. |
| `-MaxPointsPerResponse` | 500000 | Budget used to size getBatch batches (series × minutes per response). Oversized requests (413) are halved automatically. |
| `-MaxRetries` | 6 | Retries per request for 429/5xx/transport errors. |

### Performance and throttling

The extractor is built for large estates:

- **Batched metrics:** `metrics:getBatch` returns up to 50 accounts per call, and its limits are per subscription and region, separate from the ARM read budget.
- **Parallel workers** (`-Parallel`) share **adaptive rate limiters** (token buckets): global and per subscription/region for getBatch; global, per subscription and metrics-per-subscription for ARM; and Resource Graph. Each bucket increases its rate gradually after clean calls and halves it on `429`/`503`/`529`. A throttled response pauses every worker that shares the bucket (honouring `Retry-After`), and a low `x-ms-ratelimit-remaining-subscription-reads` header slows down the bucket before throttling happens.
- **Idle pre-flight:** one coarse daily query per batch over the whole window. Accounts with no tokens, requests (or PTU utilization) at all are recorded as `Success` / `Idle (pre-flight)` without querying every chunk. Any doubt (missing metric, error) keeps the account in the full extraction.
- **Latency only where there is traffic:** the latency/availability call is only made for accounts and chunks with requests, filtered to the deployments that had traffic.
- Discovery and quota calls also run in parallel (quota usages and the models catalog are separate units); tokens are refreshed in the background during long runs.
- **Token prefetch:** all tenant tokens are requested up front, a few `az` processes at a time, each with its own deadline. A tenant that fails is recorded once and not retried by every call.
- **Fast parsing and writing:** metrics responses are parsed by a compiled .NET (`System.Text.Json`) parser instead of `ConvertFrom-Json`, and CSV lines, manifest appends and the final merge use .NET streams. Workers reuse the library loaded in their runspace.
- **Cost-ordered work:** extraction units are scheduled from the most to the least expensive (deployments × calls), so large batches do not end up alone at the tail of the run.
- Requests are split (by time or by batch) **only** on explicit size errors; overload errors (`5xx`, `529`, transport) are retried with backoff, never multiplied.

`run.json` → `Performance` records the durations, the number of work units and batches, getBatch availability per subscription/region, and calls, throttles and rates per limiter. Reference: 90 days, 1 account, 10 deployments, with latency: **~1.4 min** (extraction 41 s) with the default `-Parallel 12`, vs ~4 min for the previous version and ~15 min for the original sequential one, with 0 throttling events. `-Parallel 12` is an empirically conservative default, not a guarantee: the shared limiters are what keep the request rate safe, so on large estates you can raise it and watch `Throttles` in `run.json`. When latency is not needed, `-NoLatency` cuts the number of calls by about a third.

The resolution is fixed at one minute (PT1M), because the P95/P99 per minute are the goal of the tool.

### Multi-tenant

The script gets one token per tenant from the selected provider (`-AuthMode`) and never mixes tenants in the same call.

If a tenant requires MFA or a specific login:

1. Run `az login --tenant <id>` (or `Connect-AzAccount -Tenant <id>`).
2. Run the script again **with the same `-OutputDirectory`**.

Failed tenants are recorded in `run.json` (`FailedTenants`).

### Resuming and idempotency

When the same folder is reused without window parameters, the original window is kept and only missing or failed chunks are queried. Each chunk is saved as an atomic file in `parts\`. If the window, chunk size or calls change, the script stops (fingerprint in `run.json`).

- A successful chunk whose part file has been deleted is extracted again.
- Only chunks whose calls all completed in the manifest are merged into the final CSV, so leftovers of an interrupted run never leak in.
- `run.json` → `Completeness` reports `Complete` or `Incomplete`, with the expected calls, missing or unsuccessful calls, and missing part files.
- The folder is locked (`.run.lock`) while a run is active, so two scheduled runs cannot write to it at the same time.

### Output files (run folder)

| File | Content |
|---|---|
| `foundry-token-metrics.csv` | Per-minute data (see columns) |
| `manifest.csv` | Every metrics call, with status `Success`, `NotSupported` (the account does not define those metrics), `Failed` or `Truncated`, number of series and rows. Idle accounts appear as `Success` with message `Idle (pre-flight)` |
| `resources.csv` | Accounts found, `kind`, capability status and deployments |
| `deployments.csv` | Snapshot of each deployment: SKU, capacity, `TokensPerMinuteLimit`/`RequestsPerMinuteLimit` limits (from `rateLimits`), associated quota (`QuotaName`, `QuotaMatch`, `QuotaLimit`, `QuotaCurrentValue`) |
| `quota.csv` | Snapshot of the `locations/{region}/usages` API per subscription and region: `QuotaName`, `DeploymentType`, `QuotaModel`, `CurrentValue` (assigned capacity), `Limit`, `AssignedPct`, `UsedByDeployments` |
| `run.json`, `run.log` | Parameters, fingerprint, `RunTimestamp` (start of the run, kept when resuming), `Performance`, `Completeness` and log |

`QuotaMatch` shows how the deployment was linked to its quota: `ModelsApi` (exact, via `locations/{region}/models`), `NameMatch` (normalized name or shared PTU pool), `Ambiguous`, `Unmapped` or `NoQuotaAccess` (subscription-level Reader missing for the usages API).

`AzureOpenAIRequests` is queried with the `StatusCode` dimension: the metric keeps the total number of requests, and three additional metrics hold the failures: `ThrottledRequests(429)` (rejected by the rate limit), `ClientErrors(4xx)` (other 4xx, without 429) and `ServerErrors(5xx)`.

Latency and availability (unless `-NoLatency`) are written as `AzureOpenAITimeToResponse`, `AzureOpenAITTLTInMS` (ms), `AzureOpenAITokenPerSecond` and `AzureOpenAIAvailabilityRate` (%), with `Aggregation` = `Average` and `Maximum` per minute. Azure Monitor only emits them for minutes with traffic.

The current schema is v5: folders produced by earlier versions cannot be resumed; use a new folder.

### CSV columns

The mandatory columns are, in this order: `Timestamp`, `SubscriptionId`, `SubscriptionName`, `ResourceGroup`, `ResourceName`, `Region`, `DeploymentName`, `ModelName`, `ModelVersion`, `Metric` and `Total`.

Additional columns:

| Column | Meaning |
|---|---|
| `TenantId`, `Kind`, `ResourceId` | Resource identification |
| `Aggregation` | `Total` for tokens, requests and errors. `Maximum` or `Average` for Utilization V2 and the latency/availability metrics, which **are never summed**. |
| `DeploymentSku`, `DeploymentCapacity` | Current deployment SKU (Standard, GlobalStandard, ProvisionedManaged…) |
| `ModelNameSource` | Source of `ModelName` (see the note below) |

Notes:

- **`ModelName`:** token metrics do not have the `ModelName` dimension, so it is taken, in order, from:
  1. `RequestsMetric`: the model that served `AzureOpenAIRequests` in **the same minute**, deployment, version and region.
  2. `RequestsChunk`: the only model seen for that deployment and version within the chunk.
  3. `CurrentDeployment`: the current deployment configuration, if consistent with the above.
  4. `Ambiguous` or `Unknown`: the model cannot be proven.
- **`Timestamp`** is UTC and marks the start of the minute. **`Total`** is tokens (or requests) in that minute, i.e. TPM.
- **Sparse format:** only minutes with a value greater than 0 are written. A missing minute is a real zero only if the chunk is `Success` in the manifest.
- **`Region`** is normalized (`swedencentral`), because Azure Monitor returns different formats depending on the metric.

## 2. Statistics and Excel report

```powershell
python .\scripts\foundry_token_stats.py --run-dir .\output\tenantA-90d
# Several tenants (with the same window) in a single report
python .\scripts\foundry_token_stats.py --run-dir .\output\tenantA --run-dir .\output\tenantB --out-dir .\output\all
# With estimated costs
python .\scripts\foundry_token_stats.py --run-dir .\output\tenantA-90d --fetch-prices
python .\scripts\foundry_token_stats.py --run-dir .\output\tenantA-90d --prices .\my-prices.csv
```

| Option | Description |
|---|---|
| `--run-dir` | Run folder from the extractor (repeatable). |
| `--out-dir` | Output folder (default: first `--run-dir`). |
| `--no-excel` | Skip the Excel report. |
| `--fetch-prices` | Get candidate pay-as-you-go prices from the public [Azure Retail Prices API](https://learn.microsoft.com/rest/api/cost-management/retail-prices/azure-retail-prices) (no authentication). |
| `--prices` | CSV with your prices per 1M tokens (see `prices.example.csv`). Overrides `--fetch-prices` for the rows it matches. |
| `--currency` | Currency for `--fetch-prices` (default `USD`). |

Every output file carries the UTC timestamp of the statistics run (`<ts>` = `yyyyMMdd-HHmmssZ`) while keeping its base name, so periodic runs never overwrite each other:

- `foundry-token-stats_<ts>.csv`
- `foundry-token-daily_<ts>.csv`
- `foundry-deployment-limits_<ts>.csv`: per deployment, P95/P99/Max tokens per minute against its TPM limit (`PctOfTPM_*`), requests against RPM, 429s (sum and throttled minutes), other 4xx, 5xx and `ErrorRatePct`, latency (`TTR_P95`, `TTLT_P95`, `TPS_P50`, `Avail_Min`, `Avail_Avg`), coverage (`TokensCoveragePct`, `RequestsCoveragePct`), `DataStatus`, prices and costs (`InputPer1M`, `OutputPer1M`, `EstCost`, `EstMonthlyCost`, `PriceSource`) and `Recommendation`.
- `prices-candidates_<ts>.csv` (with `--fetch-prices`): the Retail API price matched for each model + deployment type + region, with its meter names. Review it and copy rows to your `--prices` file to fix or complete prices.
- `foundry-quota-utilization_<ts>.csv`: per subscription + region + quota, assigned capacity against the limit (`AssignedPct`), quota TPM (`Limit × TokensPerUnit`) and per-minute P95/P99/Max consumption summing all deployments of that quota (`PctOfQuotaTPM_*`, `PctOfAssignedTPM_*`).
- `foundry-token-hourly_<ts>.csv`: profile by UTC hour of day (average and maximum per minute) per model.
- `foundry-token-report_<ts>.xlsx` (always generated; `--no-excel` skips it): **compact report, no charts**, with the minimum metrics per model. For more detail use the CSVs above and the extractor ones (`foundry-token-metrics`, `deployments`, `quota`, `resources`, `manifest`).

| Sheet | Content |
|---|---|
| Summary | Generation date, window, runs, tenants (and failed tenants), number of deployments, deployments with 429, with 5xx and with no traffic, estimated cost, price notes, recommendation counts, minimum coverage and reading notes |
| Models | One row per deployment: subscription, region, account, model, version, SKU, capacity, TPM/RPM limits, Input/Output TPM P95/P99, TPM P99/max and peak (UTC), % TPM P99/max, RPM P99/max, % RPM max, 429, other 4xx, 5xx, error rate, latency P95, availability min, PTU utilization P95/P99/max, estimated cost, coverage, **Data status** and **Recommendation** |
| Quotas | Per subscription + region + quota: limit, assigned, % assigned, quota TPM, TPM P99/max, % quota P99/max, 429, coverage and whether all consumers were extracted |

Notes on the Models sheet:

- Order: deployments with 429 first, then standard by % TPM max, PTU by utilization, models without a TPM limit and, finally, those without data.
- `N/A`: % TPM does not apply to PTU or to models without a TPM limit (e.g. image models, RPM only); PTU utilization does not apply to standard deployments. An empty cell means "no data", not zero.
- `Data status`:

  | Status | Meaning |
  |---|---|
  | `With metrics` | Traffic observed with full coverage. |
  | `No traffic (confirmed)` | No points, and **every** token and request call of the account succeeded for the whole window (`Success` or `NotSupported`). Absence of data is then a real zero. |
  | `No data (incomplete coverage)` | No points, but some calls failed or were truncated: it is **not** confirmed that there was no traffic. Re-run the extraction (same folder) before deciding. |
  | `Incomplete coverage` | Traffic observed, but part of the window is missing; percentiles only use covered minutes. |
  | `Token metric not available` | Requests but no token metrics (the model does not publish them). |
  | `No TPM limit (RPM only)` | E.g. image models. |
  | `Deployment deleted (Unmapped)` | Metrics for a deployment that no longer exists. |

- **Latency:** `Time to response P95` and `Time to last token P95` are the P95 **of the 1-minute averages** over the minutes with traffic (Azure Monitor does not expose per-request percentiles); `Availability min %` is the lowest 1-minute availability rate. Use them as trends, not SLAs.
- **Errors:** `Error rate %` = (429 + other 4xx + 5xx) / requests.

### Estimated cost

Costs are an **estimate at list price**, never a bill: they ignore discounts, EA/MCA agreements, cached-token pricing, Batch and PTU reservations.

- `--fetch-prices` queries the public Retail Prices API and keeps only **exact, unambiguous** matches for model + deployment type (`Global`, `DataZone`, `Regional`, derived from the SKU) + region + direction (input/output). Cached, batch, priority, flex, short/long-context, fine-tuning, realtime, audio, image and embedding meters are excluded. When several prices fit, none is chosen.
- PTU (`Provisioned*`) and Batch deployments are not priced per token and get no cost.
- `--prices` (template in `prices.example.csv`: `Model, DeploymentType, Region, InputPer1M, OutputPer1M, Currency, Source`) overrides fetched prices. `DeploymentType` must be `Global`, `DataZone` or `Regional`; `Region` `*` (or empty) applies the row to every region. Rows with no price are ignored.
- `prices.example.csv` ships with Standard list prices in USD, taken from the [Azure OpenAI pricing page](https://azure.microsoft.com/en-us/pricing/details/azure-openai/) and Retail Prices API in November 2026. Models with short/long-context tiers use the short-context price. Realtime and image models are left unpriced because token metrics do not split text, audio and image tokens. Review it against your agreement before use.
- `EstCost` = input tokens × input price + output tokens × output price over the covered minutes; `EstMonthlyCost` extrapolates it to 30 days.
- Models without a price are listed on the console and in the Summary sheet.

### Recommendations

Rule-based and explainable; several rules can apply to a deployment (joined with `;`). Thresholds are in `THRESHOLDS` in `foundry_token_stats.py`.

| Rule | Recommendation |
|---|---|
| `No traffic (confirmed)` | Deletion candidate (frees quota) |
| `No data (incomplete coverage)` | Re-run the extraction before deciding |
| 429 and TPM max ≥ 80% | Increase capacity or spread load |
| 429 and TPM max < 80% | Check RPM, bursts and `max_tokens` |
| Standard, no 429, TPM max < 50% and P99 < 10% | Reduce capacity to free quota |
| PTU utilization P99 > 90% / < 40% | Add PTU or spillover / reduce PTU |
| Quota assigned or quota TPM max > 90% | Request an increase or use another region/Data Zone |
| Any 5xx or availability min < 99% | Review service errors |
| Other 4xx > 5% of requests | Review callers (bad requests, auth, content filter…) |

Use windows of 7 days or more: on short windows the "low use" rule is noisy.

### Grouping levels

| Level | Groups by |
|---|---|
| Deployment | resource, region, deployment, model, version, metric and aggregation |
| DeploymentLimit | resource and deployment (sum of versions/models per minute; compared with the deployment rate limit) |
| Quota | subscription, region and quota (per-minute sum of all deployments consuming that quota) |
| ModelRegion | model, version, region and metric (sum of deployments per minute) |
| Model | model and metric |
| All | metric (global consumption per minute) |

### Modes (denominator for the average and percentiles)

| Mode | Use |
|---|---|
| **FullWindow** | Every observed minute of the window, zero-filled. **This is the reference mode for PTU.** |
| Lifetime | Minutes between the group's first and last active minute |
| Active | Only minutes with consumption; overstates sustained demand |

Percentiles use the nearest-rank method.

`CompletenessPct` is the percentage of the window in which **every** resource in the group has a `Success` call. Minutes from `Failed` or `Truncated` chunks are excluded from both numerator and denominator, never counted as zero. Rollups (Model or All) only use minutes covered for all their resources. `TotalTokens(derived)` is informational only: for sizing, **input and output are always analysed separately**.

### PTU guidance

- Use per-minute `P95` and `P99` of `ProcessedPromptTokens` and `GeneratedTokens` in FullWindow mode, at Deployment or ModelRegion level, together with `Max` and `MaxTimestamp`.
- To convert to TPS, divide TPM by 60. Enter input and output values separately in the Foundry capacity calculator.
- Once PTUs are purchased, monitor `Provisioned-managed Utilization V2` (`-IncludePtuUtilization`). The previous utilization metric is deprecated.

### Proximity to the quota limit

- Quota is assigned per **subscription + region + deployment type (Standard, GlobalStandard, DataZoneStandard, PTU…) + model**. Data Zone is also counted against the account region.
- `AssignedPct` measures **assigned** capacity (it can exceed 100% for some quotas), not consumption. Actual per-minute consumption is in `PctOfQuotaTPM_*` and `PctOfTPM_*`.
- `AllConsumersExtracted = False` means the capacity of the extracted deployments does not add up to the quota `CurrentValue`: there are consumers out of scope (filters or no access) and quota consumption % is underestimated.
- Capacity units are not always 1,000 TPM; `TokensPerUnit` is derived from the deployments' `rateLimits`. No quota TPM is computed for PTU.
- **Metrics count processed tokens, but the rate limiter counts estimated tokens (prompt + `max_tokens`)**, so the real pressure on the limit is higher than computed. The direct saturation signal is `ThrottledRequests(429)`.

## 3. Azure Monitor workbook

To import it:

1. Go to Azure Portal → Monitor → Workbooks → New.
2. Open the **Advanced Editor** (`</>`) and paste the content of `workbooks\FoundryTokenUsage.workbook.json`.
3. Click **Apply** and then **Save**.

Workbook content:

- Parameters: subscriptions, accounts (ARG, `kind` OpenAI or AIServices), time range and granularity (PT1M, PT5M or PT1H; PT1M by default).
- Account inventory.
- Input, output, requests, throttling (StatusCode 429) and Utilization V2 charts, split by deployment (top 25).
- **Per model / deployment detail:** tables with input, output and request totals per deployment, requests per model and 429s per deployment. After choosing an account and deployment, its limits are shown (model, version, SKU, capacity, TPM and RPM) along with charts of tokens per minute (input vs output), requests by `StatusCode` (429, other 4xx, 5xx), 429s, Utilization V2, latency (time to response and time to last token) and availability rate / tokens per second.
- Assigned quota vs limit table (ARM `usages` API) for the selected subscription and region.

Limits:

- Up to about 10,000 points per chart and 100 series. With PT1M, use ranges of 6 days or less (1 day recommended).
- Up to 200 resources per metrics query.

The workbook is meant for visual exploration. **P95/P99 over 30–90 days are obtained with the scripts.**

## Container

The `Dockerfile` bundles PowerShell 7, Python with the requirements and Azure CLI, and runs as a non-root user.

```bash
docker build -t foundry-insights .
# Behind a proxy that blocks PyPI, use your corporate package index:
#   docker build --build-arg PIP_INDEX_URL=https://<your-feed>/pypi/simple/ -t foundry-insights .
# Interactive Azure CLI login inside the container; results are written to ./output on the host
docker run --rm -it -v "$PWD/output:/app/output" foundry-insights \
  -c "az login --use-device-code --tenant <tenantId> && pwsh -File scripts/Run-FoundryInsights.ps1 -TenantId <tenantId> -Days 30 -FetchPrices"
```

For unattended runs (e.g. an Azure Container Apps job with a managed identity that has Reader + Monitoring Reader), use `-AuthMode ManagedIdentity` and mount or upload `/app/output`.

## Known limitations

- 93-day retention for platform metrics.
- ARM and Azure Monitor throttling: adaptive shared rate limiters, retries with `Retry-After` and exponential backoff.
- Deleted deployments keep their metric history, but `DeploymentSku` and `DeploymentCapacity` will be empty.
- If a single deployment exceeds `-Top` series within a chunk, the chunk is flagged `Truncated`. It is excluded from the statistics and retried on resume; you can increase `-Top` when resuming.
- If a resource is visible from several tenants, the first tenant in which Resource Graph returns it is used. If that tenant has no permission on the metrics, run the script with `-TenantId` pointing to the other tenant.
- Reusing a folder with a different scope (`-TenantId`, `-SubscriptionId`, `-Kind`) is blocked; use another folder.
