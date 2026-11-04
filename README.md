# Foundry Insights – Azure AI Foundry / Azure OpenAI token consumption

This project extracts per-minute metrics from every accessible subscription and tenant and computes the statistics needed to size PTU (average, P95, P99 and maximum). It includes an Azure Monitor workbook and an Excel report.

| Component | File |
|---|---|
| Extraction (PowerShell 7 + Azure CLI) | `scripts\Export-FoundryTokenMetrics.ps1` |
| Statistics and Excel (Python + DuckDB) | `scripts\foundry_token_stats.py` |
| Azure Monitor workbook | `workbooks\FoundryTokenUsage.workbook.json` |

## Requirements

- PowerShell 7.2 or later (`pwsh`) and Azure CLI with a signed-in session (`az login`).
- Python 3.10 or later: `pip install duckdb pandas openpyxl`.
- Permissions on each subscription or resource:
  - **Reader**, for ARG, the account and the deployments.
  - **Monitoring Reader**, for the metrics.
  - Reader is enough to read platform metrics; Monitoring Reader is the explicit minimum role.
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
| `-TenantId` | all (`az account list --all`) | Tenants to process. |
| `-SubscriptionId` | all | Allow-list of subscriptions. |
| `-Kind` / `-IncludeAllKinds` | `OpenAI`, `AIServices` | Filter by `kind`. In every case the account is validated to expose OpenAI metrics; if it does not, it is flagged `NotOpenAI` in `resources.csv`. |
| `-IncludePtuUtilization` | off | Adds `Provisioned-managed Utilization V2` (Maximum and Average). |
| `-Top` | 1000 | Series limit per call. If reached, the query is repeated per deployment to avoid truncation. |
| `-TokenTimeoutSeconds` | 60 | Maximum time to get a tenant token. Tenants that require interactive login (MFA/CA) are flagged as failed and the others continue. |

### Multi-tenant

The script gets one token per tenant (`az account get-access-token --tenant`) and never mixes tenants in the same call.

If a tenant requires MFA or a specific login:

1. Run `az login --tenant <id>`.
2. Run the script again **with the same `-OutputDirectory`**.

Failed tenants are recorded in `run.json` (`FailedTenants`).

### Resuming and idempotency

When the same folder is reused without window parameters, the original window is kept and only missing or failed chunks are queried. Each chunk is saved as an atomic file in `parts\`. If the window, chunk size or calls change, the script stops (fingerprint in `run.json`).

### Output files (run folder)

| File | Content |
|---|---|
| `foundry-token-metrics.csv` | Per-minute data (see columns) |
| `manifest.csv` | Every metrics call, with status `Success`, `Failed` or `Truncated`, number of series and rows |
| `resources.csv` | Accounts found, `kind`, capability status and deployments |
| `deployments.csv` | Snapshot of each deployment: SKU, capacity, `TokensPerMinuteLimit`/`RequestsPerMinuteLimit` limits (from `rateLimits`), associated quota (`QuotaName`, `QuotaMatch`, `QuotaLimit`, `QuotaCurrentValue`) |
| `quota.csv` | Snapshot of the `locations/{region}/usages` API per subscription and region: `QuotaName`, `DeploymentType`, `QuotaModel`, `CurrentValue` (assigned capacity), `Limit`, `AssignedPct`, `UsedByDeployments` |
| `run.json`, `run.log` | Parameters, fingerprint, `RunTimestamp` (start of the run, kept when resuming) and log |

`QuotaMatch` shows how the deployment was linked to its quota: `ModelsApi` (exact, via `locations/{region}/models`), `NameMatch` (normalized name or shared PTU pool), `Ambiguous`, `Unmapped` or `NoQuotaAccess` (subscription-level Reader missing for the usages API).

`AzureOpenAIRequests` is queried with the `StatusCode` dimension: the metric keeps the total number of requests and an additional `ThrottledRequests(429)` metric holds the requests rejected by the rate limit.

The current schema is v4: folders produced by earlier versions cannot be resumed; use a new folder.

### CSV columns

The mandatory columns are, in this order: `Timestamp`, `SubscriptionId`, `SubscriptionName`, `ResourceGroup`, `ResourceName`, `Region`, `DeploymentName`, `ModelName`, `ModelVersion`, `Metric` and `Total`.

Additional columns:

| Column | Meaning |
|---|---|
| `TenantId`, `Kind`, `ResourceId` | Resource identification |
| `Aggregation` | `Total` for tokens and requests. `Maximum` or `Average` for Utilization V2, which is a percentage and **is never summed**. |
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
```

Every output file carries the UTC timestamp of the statistics run (`<ts>` = `yyyyMMdd-HHmmssZ`) while keeping its base name, so periodic runs never overwrite each other:

- `foundry-token-stats_<ts>.csv`
- `foundry-token-daily_<ts>.csv`
- `foundry-deployment-limits_<ts>.csv`: per deployment, P95/P99/Max tokens per minute against its TPM limit (`PctOfTPM_*`), requests against RPM and number of 429s (sum and throttled minutes).
- `foundry-quota-utilization_<ts>.csv`: per subscription + region + quota, assigned capacity against the limit (`AssignedPct`), quota TPM (`Limit × TokensPerUnit`) and per-minute P95/P99/Max consumption summing all deployments of that quota (`PctOfQuotaTPM_*`, `PctOfAssignedTPM_*`).
- `foundry-token-hourly_<ts>.csv`: profile by UTC hour of day (average and maximum per minute) per model.
- `foundry-token-report_<ts>.xlsx` (always generated; `--no-excel` skips it): **compact report, no charts**, with the minimum metrics per model. For more detail use the CSVs above and the extractor ones (`foundry-token-metrics`, `deployments`, `quota`, `resources`, `manifest`).

| Sheet | Content |
|---|---|
| Summary | Generation date, window, runs, tenants (and failed tenants), number of deployments, deployments with 429, minimum coverage and reading notes |
| Models | One row per deployment: subscription, region, account, model, version, SKU, capacity, TPM/RPM limits, Input/Output TPM P95/P99, TPM P99/max and peak (UTC), % TPM P99/max, RPM P99/max, % RPM max, 429 requests, PTU utilization P95/P99/max, coverage and **Data status** |
| Quotas | Per subscription + region + quota: limit, assigned, % assigned, quota TPM, TPM P99/max, % quota P99/max, 429, coverage and whether all consumers were extracted |

Notes on the Models sheet:

- Order: deployments with 429 first, then standard by % TPM max, PTU by utilization, models without a TPM limit and, finally, those without data.
- `N/A`: % TPM does not apply to PTU or to models without a TPM limit (e.g. image models, RPM only); PTU utilization does not apply to standard deployments. An empty cell means "no data", not zero.
- `Data status`: `With metrics`, `No traffic observed`, `Token metric not available` (there are requests but the model does not publish token metrics), `No TPM limit (RPM only)`, `Incomplete coverage` or `Deployment deleted (Unmapped)`.

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
- **Per model / deployment detail:** tables with input, output and request totals per deployment, requests per model and 429s per deployment. After choosing an account and deployment, its limits are shown (model, version, SKU, capacity, TPM and RPM) along with charts of tokens per minute (input vs output), requests by `StatusCode`, 429s and Utilization V2.
- Assigned quota vs limit table (ARM `usages` API) for the selected subscription and region.

Limits:

- Up to about 10,000 points per chart and 100 series. With PT1M, use ranges of 6 days or less (1 day recommended).
- Up to 200 resources per metrics query.

The workbook is meant for visual exploration. **P95/P99 over 30–90 days are obtained with the scripts.**

## Known limitations

- 93-day retention for platform metrics.
- ARM throttling: retries with `Retry-After` and exponential backoff.
- Deleted deployments keep their metric history, but `DeploymentSku` and `DeploymentCapacity` will be empty.
- If a single deployment exceeds `-Top` series within a chunk, the chunk is flagged `Truncated`. It is excluded from the statistics and retried on resume; you can increase `-Top` when resuming.
- If a resource is visible from several tenants, the first tenant in which Resource Graph returns it is used. If that tenant has no permission on the metrics, run the script with `-TenantId` pointing to the other tenant.
- Reusing a folder with a different scope (`-TenantId`, `-SubscriptionId`, `-Kind`) is blocked; use another folder.
