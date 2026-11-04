#Requires -Version 7.2
<#
.SYNOPSIS
    Extracts per-minute Azure OpenAI / Foundry token and request metrics from every accessible
    subscription (across tenants) into a CSV file.

.DESCRIPTION
    1. Resolves tenants (default: every tenant known to Azure CLI) and gets tokens per tenant.
    2. Lists enabled subscriptions per tenant and finds Microsoft.CognitiveServices/accounts with
       Azure Resource Graph (queried per tenant, never mixing tenants in one call).
    3. Validates each account through its metric definitions (must expose the Azure OpenAI metrics)
       and reads deployments and the subscription/region quota (in parallel).
    4. Idle pre-flight (default): one daily-grain query per batch of up to 50 accounts over the whole
       window. Accounts whose tokens/requests are all zero are recorded as successful empty chunks
       and are not queried per minute.
    5. Queries Azure Monitor per time chunk at PT1M, in parallel, using the regional
       metrics:getBatch API (up to 50 accounts of the same subscription + region per call) and
       falling back to the per-resource ARM metrics API when batch is not permitted or not usable:
         - Tokens   : ProcessedPromptTokens, GeneratedTokens (Total) split by ModelDeploymentName,
                      ModelVersion, Region. These metrics have NO ModelName dimension.
         - Requests : AzureOpenAIRequests (Total) split by ModelDeploymentName, ModelName,
                      ModelVersion, Region, StatusCode. Used as the historical source of ModelName
                      and re-aggregated into ThrottledRequests(429), ClientErrors(4xx), ServerErrors(5xx).
         - Latency  : (default, -NoLatency to skip) time to response, time to last token, tokens per
                      second and availability rate (Average, Maximum).
         - PTU      : (optional) AzureOpenAIProvisionedManagedUtilizationV2 (Maximum, Average).
       Calls whose metrics an account does not define are recorded as 'NotSupported' for it.
       Every request goes through adaptive rate limiters (per subscription/region, per subscription
       and global) that slow down and pause all workers as soon as the service throttles.
    6. Writes one atomic part file per resource/chunk plus a manifest of every call, so a run can be
       resumed safely and statistics can tell real zeros apart from missing data.
    7. Merges the parts into <OutputDirectory>\foundry-token-metrics.csv.

    Only non-zero points are written (sparse). Minutes covered by a *successful* call with no
    row are real zeros; minutes of failed calls are unknown (see manifest.csv).

.PARAMETER OutputDirectory
    Run folder. Re-running with the same folder resumes (only missing/failed chunks are queried)
    and reuses the original time window. Default: <OutputRoot>\<RunName>_<yyyyMMdd-HHmmss>Z (UTC
    execution timestamp), so every scheduled execution gets its own folder.

.PARAMETER OutputRoot
    Parent folder for timestamped run folders when -OutputDirectory is not given (default .\output).

.PARAMETER RunName
    Base name of the timestamped run folder (default 'foundry').

.PARAMETER Days
    Look-back window ending at -EndTime (default 30). Azure Monitor keeps platform metrics 93 days.

.PARAMETER StartTime
    Explicit window start (UTC).

.PARAMETER EndTime
    Explicit window end (UTC, exclusive). Defaults to the start of the current UTC hour.

.PARAMETER ChunkHours
    Size of each metrics query window (1-168, default 24). Chunks are split automatically if the
    service rejects them as too large.

.PARAMETER TenantId
    Tenants to process. Default: every tenant returned by ARM GET /tenants for the signed-in identity
    (required with -AuthMode ManagedIdentity).

.PARAMETER SubscriptionId
    Optional subscription allow-list.

.PARAMETER Kind
    Account kinds considered (default OpenAI, AIServices). Use -IncludeAllKinds to test every
    Cognitive Services account; capability is always confirmed through metric definitions.

.PARAMETER IncludePtuUtilization
    Also extract AzureOpenAIProvisionedManagedUtilizationV2 (Maximum and Average per minute).

.PARAMETER NoLatency
    Skip the latency/availability call (AzureOpenAITimeToResponse, AzureOpenAITTLTInMS,
    AzureOpenAITokenPerSecond, AzureOpenAIAvailabilityRate; Average and Maximum per minute). Saves one
    metrics call per batch and chunk.

.PARAMETER Top
    Max time series per metrics call and resource. If a resource returns this many series, it is
    re-queried per deployment to avoid silent truncation. If a single deployment still reaches the
    limit the chunk is marked 'Truncated': it is excluded from statistics coverage and retried on
    resume (you may raise -Top when resuming).

.PARAMETER MetricsApi
    Batch (default): regional metrics:getBatch, falling back to ARM per subscription/region when it
    is not permitted (it needs Monitoring Reader or Reader at subscription scope). Arm: classic
    per-resource ARM metrics API only.

.PARAMETER Parallel
    Number of concurrent workers (1-32, default 6). Request rates are still bounded by the adaptive
    rate limiters, so raising it never bypasses throttling protection.

.PARAMETER NoSkipIdle
    Disable the idle pre-flight and query every account per minute.

.PARAMETER MaxPointsPerResponse
    Budget used to size getBatch calls (default 500000 data points per response). Batches are split
    automatically if the service still rejects them as too large.

.PARAMETER InventoryOnly
    Only discover subscriptions/resources/deployments (resources.csv), no metrics.

.PARAMETER AuthMode
    Token provider: Auto (default: Azure CLI if installed, else Az PowerShell, else managed identity when
    IDENTITY_ENDPOINT is set), AzCli (az login), AzPowerShell (Connect-AzAccount, needs Az.Accounts) or
    ManagedIdentity (IMDS or App Service/Container Apps identity endpoint; requires -TenantId). Auto never
    switches provider after an authentication failure.

.PARAMETER ManagedIdentityClientId
    Client id of a user-assigned managed identity (only with -AuthMode ManagedIdentity).

.EXAMPLE
    pwsh -File .\scripts\Export-FoundryTokenMetrics.ps1 -Days 90 -ChunkHours 24 -IncludePtuUtilization

.EXAMPLE
    pwsh -File .\scripts\Export-FoundryTokenMetrics.ps1 -TenantId <tenant> -SubscriptionId <sub> -Days 7 -OutputDirectory .\output\test
#>
[CmdletBinding()]
param(
    [string]   $OutputDirectory,
    [string]   $OutputRoot = (Join-Path (Get-Location) 'output'),
    [ValidatePattern('^[\w.-]+$')] [string] $RunName = 'foundry',
    [ValidateRange(1, 93)] [int] $Days = 30,
    [datetime] $StartTime,
    [datetime] $EndTime,
    [ValidateRange(1, 168)] [int] $ChunkHours = 24,
    [string[]] $TenantId,
    [string[]] $SubscriptionId,
    [string[]] $Kind = @('OpenAI', 'AIServices'),
    [switch]   $IncludeAllKinds,
    [switch]   $IncludePtuUtilization,
    [switch]   $NoLatency,
    [ValidateRange(10, 10000)] [int] $Top = 1000,
    [ValidateRange(1, 10)] [int] $MaxRetries = 6,
    [ValidateRange(10, 600)] [int] $TokenTimeoutSeconds = 60,
    [ValidateSet('Batch', 'Arm')] [string] $MetricsApi = 'Batch',
    [ValidateRange(1, 32)] [int] $Parallel = 6,
    [switch]   $NoSkipIdle,
    [ValidateRange(10000, 5000000)] [int] $MaxPointsPerResponse = 500000,
    [ValidateSet('Auto', 'AzCli', 'AzPowerShell', 'ManagedIdentity')] [string] $AuthMode = 'Auto',
    [string]   $ManagedIdentityClientId,
    [switch]   $InventoryOnly
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$CommonFile = Join-Path $PSScriptRoot 'FoundryInsights.Common.ps1'
. $CommonFile

$SchemaVersion = 5
$Interval = 'PT1M'
$MetricsApiMode = $MetricsApi
$MaxBatchResources = 50
$CsvColumns = @(
    'Timestamp', 'SubscriptionId', 'SubscriptionName', 'ResourceGroup', 'ResourceName', 'Region',
    'DeploymentName', 'ModelName', 'ModelVersion', 'Metric', 'Total',
    'TenantId', 'Kind', 'Aggregation', 'DeploymentSku', 'DeploymentCapacity', 'ModelNameSource', 'ResourceId'
)
$ManifestColumns = @('RecordedAt', 'ResourceId', 'ChunkStart', 'ChunkEnd', 'Call', 'Status', 'SeriesCount', 'Rows', 'Truncated', 'Message')

$Calls = [ordered]@{
    Tokens   = @{ Metrics = 'ProcessedPromptTokens,GeneratedTokens'; Aggregation = 'Total'; Aggs = @('total'); Dims = @('ModelDeploymentName', 'ModelVersion', 'Region') }
    Requests = @{ Metrics = 'AzureOpenAIRequests'; Aggregation = 'Total'; Aggs = @('total'); Dims = @('ModelDeploymentName', 'ModelName', 'ModelVersion', 'Region', 'StatusCode') }
}
# Requests split by StatusCode are re-aggregated: AzureOpenAIRequests = all codes, plus ThrottledRequests(429),
# ClientErrors(4xx) (other 4xx) and ServerErrors(5xx).
if (-not $NoLatency) {
    # Per-minute average and maximum per deployment, queried only for deployments with requests in the chunk
    # and written only for minutes with requests (zeros kept: 0 % availability is a real outage).
    $Calls['Latency'] = @{ Metrics = 'AzureOpenAITimeToResponse,AzureOpenAITTLTInMS,AzureOpenAITokenPerSecond,AzureOpenAIAvailabilityRate'
        Aggregation = 'Average,Maximum'; Aggs = @('average', 'maximum'); ActiveOnly = $true; Dims = @('ModelDeploymentName', 'ModelName', 'ModelVersion', 'Region') }
}
if ($IncludePtuUtilization) {
    $Calls['PtuUtilization'] = @{ Metrics = 'AzureOpenAIProvisionedManagedUtilizationV2'; Aggregation = 'Maximum,Average'; Aggs = @('maximum', 'average'); Dims = @('ModelDeploymentName', 'ModelName', 'ModelVersion', 'Region') }
}
$RequiredMetrics = @('ProcessedPromptTokens', 'GeneratedTokens', 'AzureOpenAIRequests')

# Shared (thread-safe) state, passed by reference to every worker runspace.
$TokenCache = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
$BatchCap = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()
$BatchCapReason = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()
$CanRefreshTokens = $true
# Initial / minimum / maximum requests per second. Limits are additive-increase, halve on throttling.
$Limiters = [FoundryInsights.LimiterSet]::new()
$Limiters.Configure('batch-global', 20, 1, 50)      # metrics:getBatch, all regions
$Limiters.Configure('batch', 4, 0.5, 10)            # metrics:getBatch per subscription + region
$Limiters.Configure('arm-global', 10, 1, 15)        # every ARM request
$Limiters.Configure('arm', 5, 0.5, 10)              # ARM per subscription
$Limiters.Configure('armmetrics', 1.5, 0.2, 3)      # ARM metrics API per subscription (~12k/h)
$Limiters.Configure('arg', 1, 0.2, 2)               # Resource Graph

#region helpers ---------------------------------------------------------------------------------
function Write-Log([string]$Message, [string]$Level = 'INFO') {
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level, $Message
    if ($Level -eq 'WARN') { Write-Warning $Message } else { Write-Host $line }
    Add-Content -LiteralPath $script:LogFile -Value $line -Encoding utf8NoBOM
}

function Invoke-ResourceGraph([string]$Tenant, [string[]]$Subscriptions, [string]$Query) {
    $rows = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Subscriptions.Count; $i += 1000) {
        $batch = $Subscriptions[$i..([math]::Min($i + 999, $Subscriptions.Count - 1))]
        $skip = $null
        do {
            $opts = @{ resultFormat = 'objectArray'; '$top' = 1000 }
            if ($skip) { $opts['$skipToken'] = $skip }
            $r = Invoke-Arm -Tenant $Tenant -Method POST -Uri '/providers/Microsoft.ResourceGraph/resources?api-version=2022-10-01' `
                -Body @{ subscriptions = @($batch); query = $Query; options = $opts }
            if ($r.data) { $rows.AddRange([object[]]$r.data) }
            $skip = if ($r.PSObject.Properties['$skipToken']) { $r.'$skipToken' } else { $null }
        } while ($skip)
    }
    return , $rows
}

function Get-NormName([string]$s) { ($s -replace '[^A-Za-z0-9]', '').ToLowerInvariant() }

# Maps a deployment to the subscription/region quota (usage) it consumes.
#   ModelsApi : exact usageName from the regional models catalog (model + version + SKU + kind)
#   NameMatch : fallback on '<prefix>.<Sku>.<model>' with punctuation-insensitive model name
function Resolve-QuotaName($Dep, [string]$AccountKind, $Models, $Usages) {
    $isFt = $Dep.Model -match '\.ft-|-ft-|finetune'
    foreach ($strictKind in $true, $false) {
        $names = @(@(foreach ($m in $Models) {
                if ((Get-Prop $m 'model.name') -ne $Dep.Model) { continue }
                $v = [string](Get-Prop $m 'model.version')
                if ($Dep.Version -and $v -and $v -ne $Dep.Version) { continue }
                if ($strictKind -and (Get-Prop $m 'kind') -ne $AccountKind) { continue }
                foreach ($sku in @(Get-Prop $m 'model.skus')) {
                    if (-not $sku -or $sku.name -ne $Dep.Sku -or -not (Get-Prop $sku 'usageName')) { continue }
                    if (($sku.usageName -match '-finetune$') -ne $isFt) { continue }
                    $sku.usageName
                }
            }) | Sort-Object -Unique)
        if ($names.Count -eq 1) { return @($names[0], 'ModelsApi') }
        if ($names.Count -gt 1) { return @(($names -join '|'), 'Ambiguous') }
    }
    $nm = Get-NormName $Dep.Model
    $hits = @($Usages | ForEach-Object { $_.name.value } | Where-Object {
            $p = $_.Split('.', 3)
            $p.Count -ge 2 -and $p[1] -eq $Dep.Sku -and (($p.Count -eq 2 -and $Dep.Sku -match 'Provisioned') -or ($p.Count -eq 3 -and (Get-NormName $p[2]) -eq $nm))
        } | Sort-Object -Unique)
    if ($hits.Count -eq 1) { return @($hits[0], 'NameMatch') }
    if ($hits.Count -gt 1) { return @(($hits -join '|'), 'Ambiguous') }
    return @('', 'Unmapped')
}

# Refreshes every cached token that expires within 20 minutes or that a worker reported as rejected (401).
# Only the coordinator calls the auth provider; workers read the shared cache.
function Update-TokenCache {
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    foreach ($k in @($TokenCache.Keys)) {
        if (($TokenCache[$k].ExpiresOn - $now) -ge 1200 -and -not $TokenCache[$k].Rejected) { continue }
        $t, $a = $k.Split('|', 2)
        try { $null = Get-AccessToken $t $a -Force } catch { $TokenCache[$k].Rejected = $false; Write-Log "Token refresh failed for ${k}: $($_.Exception.Message)" 'WARN' }
    }
}

$WorkerBlock = {
    $U = $_
    $logs = [System.Collections.Generic.List[string]]::new()
    function Write-Log([string]$Message, [string]$Level = 'INFO') { $logs.Add("$Level`t$Message") }
    Set-StrictMode -Version 3.0
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    $Top = $using:Top; $MaxRetries = $using:MaxRetries; $TokenTimeoutSeconds = $using:TokenTimeoutSeconds
    $Calls = $using:Calls; $Interval = $using:Interval; $RequiredMetrics = $using:RequiredMetrics
    $TokenCache = $using:TokenCache; $Limiters = $using:Limiters; $BatchCap = $using:BatchCap; $BatchCapReason = $using:BatchCapReason
    $MetricsApiMode = $using:MetricsApiMode; $PartsDir = $using:PartsDir; $AzExe = $using:AzExe
    $AuthMode = $using:AuthMode; $ManagedIdentityClientId = $using:ManagedIdentityClientId
    $CanRefreshTokens = $false
    . $using:CommonFile
    try { $r = Invoke-WorkUnit $U }
    catch { $r = [pscustomobject]@{ Type = 'Error'; Unit = $U; Error = "$($_.Exception.Message) @ $($_.InvocationInfo.PositionMessage)" } }
    $r | Add-Member -NotePropertyName Logs -NotePropertyValue $logs -Force
    $r
}

# Runs work units in parallel; $OnResult runs on this thread for every result as it arrives and $OnError
# for every unit that failed as a whole. Tokens are refreshed on a timer, independently of unit completion.
function Invoke-Units([object[]]$Units, [scriptblock]$OnResult, [scriptblock]$OnError, [string]$Activity) {
    if (-not $Units.Count) { return }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $n = 0; $lastLog = 0; $lastRefresh = 0
    $job = $Units | ForEach-Object -Parallel $WorkerBlock -ThrottleLimit $Parallel -AsJob
    try {
        while ($true) {
            $finished = $job.State -notin 'NotStarted', 'Running'
            foreach ($r in @(Receive-Job $job -ErrorAction Continue -ErrorVariable jobErr)) {
                $n++
                foreach ($l in $r.Logs) { $lv, $msg = $l.Split("`t", 2); Write-Log $msg $lv }
                if ($r.Type -eq 'Error') {
                    Write-Log "$Activity unit failed: $($r.Error)" 'WARN'
                    if ($OnError) { & $OnError $r }
                }
                else { & $OnResult $r }
            }
            foreach ($e in @($jobErr)) { if ($e) { Write-Log "$Activity worker error: $e" 'WARN' } }
            $el = $sw.Elapsed.TotalSeconds
            if ($el - $lastRefresh -ge 60 -or @($TokenCache.Values | Where-Object Rejected).Count) { $lastRefresh = $el; Update-TokenCache }
            if ($n -and ($n -eq $Units.Count -or $el - $lastLog -ge 30)) {
                $lastLog = $el
                $eta = if ($n -lt $Units.Count) { [timespan]::FromSeconds($el / $n * ($Units.Count - $n)).ToString('hh\:mm\:ss') } else { '00:00:00' }
                $rate = ($Limiters.Snapshot() | Where-Object { $_.Key -in 'arm-global', 'batch-global' } | ForEach-Object { $_.Value.Calls } | Measure-Object -Sum).Sum
                Write-Log ("{0}: {1}/{2} units, {3:N0} s elapsed, ETA {4}, {5:N1} calls/s overall" -f $Activity, $n, $Units.Count, $el, $eta, ($rate / [math]::Max(1, ($script:RunClock.Elapsed.TotalSeconds))))
                if ($n -eq $Units.Count) { $lastLog = [double]::MaxValue }
            }
            if ($finished) { break }
            Start-Sleep -Milliseconds 250
        }
    }
    finally { Remove-Job $job -Force -ErrorAction SilentlyContinue }
    if ($n -lt $Units.Count) { Write-Log "$Activity`: only $n of $($Units.Count) unit(s) returned a result" 'WARN' }
}
#endregion

#region run setup -------------------------------------------------------------------------------
$AuthMode = Resolve-AuthMode $AuthMode
$AzExe = $null
if ($AuthMode -eq 'AzCli') {
    $azCmd = Get-Command az -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $azCmd) { throw 'Azure CLI (az) is required for -AuthMode AzCli. Install it and run az login.' }
    $AzExe = $azCmd.Source
}
if ($AuthMode -eq 'ManagedIdentity' -and -not $TenantId) { throw '-AuthMode ManagedIdentity requires -TenantId (a managed identity belongs to one tenant).' }

$RunStartedAt = [datetime]::UtcNow
$script:RunClock = [System.Diagnostics.Stopwatch]::StartNew()
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $OutputRoot ('{0}_{1:yyyyMMdd-HHmmss}Z' -f $RunName, $RunStartedAt) }
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$PartsDir = Join-Path $OutputDirectory 'parts'
New-Item -ItemType Directory -Force -Path $PartsDir | Out-Null
# Exclusive lock for the whole run: two processes resuming the same folder would corrupt the manifest.
try { $RunLock = [System.IO.File]::Open((Join-Path $OutputDirectory '.run.lock'), 'OpenOrCreate', 'ReadWrite', 'None') }
catch { throw "Another extraction is already using $OutputDirectory (cannot lock .run.lock)." }
$script:LogFile = Join-Path $OutputDirectory 'run.log'
$ManifestFile = Join-Path $OutputDirectory 'manifest.csv'
$RunFile = Join-Path $OutputDirectory 'run.json'
$ResourcesFile = Join-Path $OutputDirectory 'resources.csv'
$OutputCsv = Join-Path $OutputDirectory 'foundry-token-metrics.csv'

$prev = if (Test-Path $RunFile) { Get-Content $RunFile -Raw | ConvertFrom-Json -AsHashtable } else { $null }
$nowUtc = [datetime]::UtcNow
if ($prev -and -not ($PSBoundParameters.Keys | Where-Object { $_ -in 'StartTime', 'EndTime', 'Days' })) {
    # Resume: reuse the original window.
    $StartTime = ConvertTo-UtcDate $prev.Fingerprint.StartTime
    $EndTime = ConvertTo-UtcDate $prev.Fingerprint.EndTime
}
else {
    if (-not $PSBoundParameters.ContainsKey('EndTime')) {
        $EndTime = [datetime]::new($nowUtc.Year, $nowUtc.Month, $nowUtc.Day, $nowUtc.Hour, 0, 0, [DateTimeKind]::Utc)
    }
    $EndTime = ConvertTo-UtcDate $EndTime
    $StartTime = if ($PSBoundParameters.ContainsKey('StartTime')) { ConvertTo-UtcDate $StartTime } else { $EndTime.AddDays(-$Days) }
}
$StartTime = $StartTime.AddTicks(-($StartTime.Ticks % [TimeSpan]::TicksPerMinute))
$EndTime = $EndTime.AddTicks(-($EndTime.Ticks % [TimeSpan]::TicksPerMinute))
if ($EndTime -le $StartTime) { throw 'EndTime must be after StartTime.' }
if ($StartTime -lt $nowUtc.AddDays(-93)) { Write-Warning 'StartTime is older than the 93-day platform metrics retention; early chunks will be empty.' }

$fingerprint = [ordered]@{
    SchemaVersion = $SchemaVersion
    StartTime     = Format-Utc $StartTime
    EndTime       = Format-Utc $EndTime
    Interval      = $Interval
    ChunkHours    = $ChunkHours
    Calls         = @($Calls.Keys)
    # Discovery scope: resuming with a different scope would silently mix resource sets.
    Tenants       = if ($TenantId) { @($TenantId | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique) } else { @('*') }
    Subscriptions = if ($SubscriptionId) { @($SubscriptionId | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique) } else { @('*') }
    Kinds         = if ($IncludeAllKinds) { @('*') } else { @($Kind | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique) }
}
if ($prev) {
    foreach ($k in $fingerprint.Keys) {
        if (($prev.Fingerprint[$k] | ConvertTo-Json -Compress) -ne ($fingerprint[$k] | ConvertTo-Json -Compress)) {
            throw "Output folder belongs to a run with a different '$k' ($($prev.Fingerprint[$k] -join ',') vs $($fingerprint[$k] -join ',')). Use a new -OutputDirectory or the original parameters."
        }
    }
}
if (-not (Test-Path $ManifestFile)) { Set-Content -LiteralPath $ManifestFile -Value ($ManifestColumns -join ',') -Encoding utf8NoBOM }
Write-Log ("{0} run in {1}" -f ($(if ($prev) { 'Resuming' } else { 'Starting' })), $OutputDirectory)

# Latest manifest status and row count per (resource|chunk|call), and the part files on disk.
$done = @{}
$doneRows = @{}
foreach ($m in (Import-Csv $ManifestFile)) {
    $k = "$($m.ResourceId)|$($m.ChunkStart)|$($m.Call)"
    $done[$k] = $m.Status; $doneRows[$k] = [long]$m.Rows
}
$partNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($f in (Get-ChildItem -LiteralPath $PartsDir -File -Filter '*.csv')) { [void]$partNames.Add($f.Name) }

Write-Log "Window $(Format-Utc $StartTime) -> $(Format-Utc $EndTime) ($Interval, chunks of $ChunkHours h); metrics API $MetricsApiMode, $Parallel worker(s), idle pre-flight $(-not $NoSkipIdle)"
#endregion

#region discovery -------------------------------------------------------------------------------
if (-not $TenantId) {
    # Home tenant of the signed-in identity lists every tenant it belongs to.
    $TenantId = @(Get-ArmPaged '' '/tenants?api-version=2022-12-01' | ForEach-Object { $_.tenantId } | Where-Object { $_ } | Sort-Object -Unique)
}
Write-Log "Tenants to process: $($TenantId.Count) (auth: $AuthMode)"

$subsSeen = @{}
$tenantSubs = [ordered]@{}
$failedTenants = [System.Collections.Generic.List[string]]::new()
foreach ($t in $TenantId) {
    try {
        $subs = @(Get-ArmPaged $t '/subscriptions?api-version=2022-12-01' | Where-Object { $_.state -eq 'Enabled' })
    }
    catch {
        Write-Log "Tenant ${t}: $($_.Exception.Message)" 'WARN'
        $failedTenants.Add($t); continue
    }
    # A subscription visible from several tenants (e.g. Lighthouse) stays a candidate in each one;
    # resources are de-duplicated by id after Resource Graph so a failing tenant path is not fatal.
    $list = @(foreach ($s in $subs) {
            if ($SubscriptionId -and $s.subscriptionId -notin $SubscriptionId) { continue }
            if (-not $subsSeen.ContainsKey($s.subscriptionId)) { $subsSeen[$s.subscriptionId] = $s.displayName }
            [pscustomobject]@{ Id = $s.subscriptionId; Name = $s.displayName }
        })
    if ($list.Count) { $tenantSubs[$t] = $list }
    Write-Log "Tenant ${t}: $($list.Count) subscription(s)"
}

$kindFilter = if ($IncludeAllKinds) { '' } else { "| where kind in~ ($((($Kind | ForEach-Object { "'$_'" }) -join ',')))" }
$argQuery = "resources | where type =~ 'microsoft.cognitiveservices/accounts' $kindFilter | project id, name, resourceGroup, subscriptionId, kind, location, sku = tostring(sku.name) | order by id asc"

$resources = [System.Collections.Generic.List[object]]::new()
$resById = @{}
foreach ($t in $tenantSubs.Keys) {
    try { $rows = Invoke-ResourceGraph $t @($tenantSubs[$t].Id) $argQuery }
    catch { Write-Log "Resource Graph failed for tenant ${t}: $($_.Exception.Message)" 'WARN'; continue }
    foreach ($r in $rows) {
        $rk = $r.id.ToLowerInvariant()
        if ($resById.ContainsKey($rk)) { continue }
        $o = [pscustomobject]@{
            TenantId = $t; Id = $r.id; Name = $r.name; ResourceGroup = $r.resourceGroup
            SubscriptionId = $r.subscriptionId; SubscriptionName = $subsSeen[$r.subscriptionId]
            Kind = $r.kind; Location = $r.location; Sku = $r.sku
            Status = 'Pending'; Message = ''; Deployments = @{}; MetricNames = @()
        }
        $resById[$rk] = $o
        $resources.Add($o)
    }
}
Write-Log "Candidate Cognitive Services accounts: $($resources.Count)"

# Validate accounts and read deployments in parallel.
Invoke-Units @($resources | ForEach-Object { [pscustomobject]@{ Type = 'Discover'; Resource = $_ } }) -Activity 'Discovery' -OnResult {
    param($r)
    $o = $resById[$r.Id.ToLowerInvariant()]
    $o.Status = $r.Status; $o.Message = $r.Message; $o.Deployments = $r.Deployments; $o.MetricNames = @($r.MetricNames)
}
foreach ($o in $resources) { if ($o.Status -eq 'Pending') { $o.Status = 'Error'; $o.Message = 'Discovery failed (see run.log)' } }

$resources | ForEach-Object {
    [pscustomobject]@{
        TenantId = $_.TenantId; SubscriptionId = $_.SubscriptionId; SubscriptionName = $_.SubscriptionName
        ResourceGroup = $_.ResourceGroup; ResourceName = $_.Name; Kind = $_.Kind; Location = $_.Location; Sku = $_.Sku
        Status = $_.Status
        Deployments = (($_.Deployments.GetEnumerator() | Sort-Object Key | ForEach-Object {
                    "$($_.Key)=$($_.Value.Model):$($_.Value.Version):$($_.Value.Sku):$($_.Value.Capacity)" }) -join ';')
        Message = $_.Message; ResourceId = $_.Id
    }
} | Export-Csv -LiteralPath $ResourcesFile -NoTypeInformation -Encoding utf8NoBOM

$targets = @($resources | Where-Object Status -eq 'OpenAI')
$resources | Group-Object Status | ForEach-Object { Write-Log "Accounts $($_.Name): $($_.Count)" }

# Quota snapshot: quota is assigned per subscription + region (account location) + quota name
# (deployment type and model, or a shared PTU pool). Needs read access at subscription scope.
$capturedAt = Format-Utc ([datetime]::UtcNow)
$quotaRows = [System.Collections.Generic.List[object]]::new()
$depRows = [System.Collections.Generic.List[object]]::new()
$quotaGroups = @{}
foreach ($grp in ($targets | Group-Object SubscriptionId, Location)) { $quotaGroups[$grp.Name] = $grp.Group }
$quotaData = @{}
Invoke-Units @($quotaGroups.Keys | Sort-Object | ForEach-Object {
        $f = $quotaGroups[$_][0]
        [pscustomobject]@{ Type = 'Quota'; Key = $_; TenantId = $f.TenantId; SubscriptionId = $f.SubscriptionId; Location = $f.Location }
    }) -Activity 'Quota' -OnResult { param($r) $quotaData[$r.Key] = $r }

$quotaFailures = 0
foreach ($key in ($quotaGroups.Keys | Sort-Object)) {
    $grp = $quotaGroups[$key]
    $first = $grp[0]
    $qd = $quotaData[$key]
    $usages = @(); $models = @()
    if (-not $qd -or $qd.UsageError) {
        $quotaFailures++
        Write-Log "Quota usages not readable for $($first.SubscriptionName)/$($first.Location): $(if ($qd) { $qd.UsageError } else { 'unit failed' })" 'WARN'
    }
    else { $usages = @($qd.Usages) }
    if ($qd -and $qd.ModelError) { Write-Log "Model catalog not readable for $($first.SubscriptionName)/$($first.Location): $($qd.ModelError)" 'WARN' }
    elseif ($qd) { $models = @($qd.Models) }
    $usageByName = @{}
    foreach ($u in $usages) { $usageByName[$u.name.value] = $u }
    $mapped = @{}
    foreach ($res in $grp) {
        foreach ($kv in ($res.Deployments.GetEnumerator() | Sort-Object Key)) {
            $d = $kv.Value
            $qn, $qm = if ($usages.Count -or $models.Count) { Resolve-QuotaName $d $res.Kind $models $usages } else { @('', 'NoQuotaAccess') }
            $u = if ($qn) { $usageByName[$qn] } else { $null }
            if ($u) { $mapped[$qn] = $true }
            $depRows.Add([pscustomobject]@{
                    TenantId = $res.TenantId; SubscriptionId = $res.SubscriptionId; SubscriptionName = $res.SubscriptionName
                    Region = $res.Location; ResourceGroup = $res.ResourceGroup; ResourceName = $res.Name; Kind = $res.Kind
                    DeploymentName = $kv.Key; ModelFormat = $d.Format; ModelName = $d.Model; ModelVersion = $d.Version
                    DeploymentSku = $d.Sku; Capacity = $d.Capacity
                    TokensPerMinuteLimit = $d.TokensPerMinute; RequestsPerMinuteLimit = $d.RequestsPerMinute
                    QuotaName = $qn; QuotaMatch = $qm
                    QuotaLimit = if ($u) { $u.limit } else { $null }; QuotaCurrentValue = if ($u) { $u.currentValue } else { $null }
                    ResourceId = $res.Id; CapturedAt = $capturedAt
                })
        }
    }
    foreach ($u in $usages) {
        $parts = $u.name.value.Split('.', 3)
        $quotaRows.Add([pscustomobject]@{
                TenantId = $first.TenantId; SubscriptionId = $first.SubscriptionId; SubscriptionName = $first.SubscriptionName
                Region = $first.Location; QuotaName = $u.name.value
                DeploymentType = if ($parts.Count -ge 2) { $parts[1] } else { '' }
                QuotaModel = if ($parts.Count -ge 3) { $parts[2] } else { '' }
                CurrentValue = $u.currentValue; Limit = $u.limit; Unit = $u.unit
                AssignedPct = if ([double]$u.limit -gt 0) { [math]::Round(100.0 * [double]$u.currentValue / [double]$u.limit, 2) } else { $null }
                UsedByDeployments = [bool]$mapped[$u.name.value]; CapturedAt = $capturedAt
            })
    }
}
$depRows | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'deployments.csv') -NoTypeInformation -Encoding utf8NoBOM
$quotaRows | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'quota.csv') -NoTypeInformation -Encoding utf8NoBOM
Write-Log "Quota snapshot: $($quotaRows.Count) quota row(s), $($depRows.Count) deployment(s); unmapped: $(@($depRows | Where-Object { $_.QuotaMatch -notin 'ModelsApi', 'NameMatch' }).Count)"
if ($quotaFailures) { Write-Log "Quota not readable for $quotaFailures subscription/region pair(s): Reader at subscription scope is required." 'WARN' }

$runInfo = [ordered]@{
    Fingerprint     = $fingerprint
    RunName         = Split-Path $OutputDirectory -Leaf
    RunTimestamp    = if ($prev -and $prev['RunTimestamp']) { $prev.RunTimestamp } else { Format-Utc $RunStartedAt }
    LastRunAt       = Format-Utc ([datetime]::UtcNow)
    AuthMode        = $AuthMode
    Tenants         = $TenantId
    FailedTenants   = @($failedTenants)
    Subscriptions   = $subsSeen.Count
    TargetResources = @($targets | ForEach-Object Id)
}
$runInfo | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $RunFile -Encoding utf8NoBOM

if ($InventoryOnly) { Write-Log "Inventory written to $ResourcesFile"; return }
#endregion

#region extraction ------------------------------------------------------------------------------
$chunks = [System.Collections.Generic.List[object]]::new()
for ($c = $StartTime; $c -lt $EndTime; $c = $c.AddHours($ChunkHours)) {
    $e = $c.AddHours($ChunkHours); if ($e -gt $EndTime) { $e = $EndTime }
    $chunks.Add([pscustomobject]@{ Start = $c; End = $e })
}

# Only 'Success' and 'NotSupported' are terminal. 'Truncated' and 'Failed' chunks are retried; all calls of an
# incomplete chunk are re-run because its part file is rewritten atomically. A successful chunk with rows whose
# part file is missing is re-extracted too.
$TerminalStatus = @('Success', 'NotSupported')
function Get-PartName([string]$ResourceId, [datetime]$ChunkStart) { '{0}_{1:yyyyMMddHHmm}.csv' -f (Get-ResourceHash $ResourceId), $ChunkStart }
function Test-Pending($Res, $Chunk) {
    $cs = Format-Utc $Chunk.Start
    $rows = 0L
    foreach ($k in $Calls.Keys) {
        $key = "$($Res.Id)|$cs|$k"
        if ($done[$key] -notin $TerminalStatus) { return $true }
        $rows += [long]$doneRows[$key]
    }
    return ($rows -gt 0 -and -not $partNames.Contains((Get-PartName $Res.Id $Chunk.Start)))
}

if ($MetricsApiMode -eq 'Batch' -and $targets.Count) {
    foreach ($t in @($targets | ForEach-Object TenantId | Sort-Object -Unique)) {
        try { $null = Get-AccessToken $t $MetricsAudience }
        catch {
            Write-Log "No token for $MetricsAudience in tenant ${t}: $($_.Exception.Message). Using the ARM metrics API for it." 'WARN'
            foreach ($r in @($targets | Where-Object TenantId -eq $t)) { Set-BatchUnavailable "$($r.SubscriptionId)|$($r.Location)".ToLowerInvariant() 'No metrics token' }
        }
    }
}

# Batches: resources of the same tenant + subscription + region, sized so the most expensive call
# (Requests: deployments x status codes) stays within the points-per-response budget.
$seriesBudget = [math]::Max(1, [math]::Floor($MaxPointsPerResponse / ($ChunkHours * 60)))
$batches = [System.Collections.Generic.List[object]]::new()
foreach ($grp in ($targets | Sort-Object Id | Group-Object TenantId, SubscriptionId, Location)) {
    $cur = [System.Collections.Generic.List[object]]::new(); $est = 0
    foreach ($r in $grp.Group) {
        $e = 4 * [math]::Max(1, $r.Deployments.Count)
        $limit = if ($MetricsApiMode -eq 'Batch') { $MaxBatchResources } else { 1 }
        if ($cur.Count -and ($cur.Count -ge $limit -or $est + $e -gt $seriesBudget)) { $batches.Add($cur.ToArray()); $cur.Clear(); $est = 0 }
        $cur.Add($r); $est += $e
    }
    if ($cur.Count) { $batches.Add($cur.ToArray()) }
}

$idleIds = [System.Collections.Generic.HashSet[string]]::new()
$idleManifest = 0
if (-not $NoSkipIdle) {
    # Pre-flight in up-to-50-resource batches (no dimensions, so there is no series budget to respect).
    $pfUnits = [System.Collections.Generic.List[object]]::new()
    foreach ($grp in ($targets | Sort-Object Id | Group-Object TenantId, SubscriptionId, Location)) {
        $pend = @($grp.Group | Where-Object { $r = $_; @($chunks | Where-Object { Test-Pending $r $_ }).Count })
        for ($i = 0; $i -lt $pend.Count; $i += $MaxBatchResources) {
            $pfUnits.Add([pscustomobject]@{ Type = 'Preflight'; Resources = @($pend[$i..([math]::Min($i + $MaxBatchResources, $pend.Count) - 1)]); From = $StartTime; To = $EndTime })
        }
    }
    Invoke-Units $pfUnits.ToArray() -Activity 'Idle pre-flight' -OnResult { param($r) foreach ($id in $r.Idle) { [void]$idleIds.Add($id) } }
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $targets) {
        if (-not $idleIds.Contains($r.Id.ToLowerInvariant())) { continue }
        foreach ($ch in $chunks) {
            $cs = Format-Utc $ch.Start
            if (-not (Test-Pending $r $ch)) { continue }
            $stale = Join-Path $PartsDir (Get-PartName $r.Id $ch.Start)
            if (Test-Path -LiteralPath $stale) { Remove-Item -LiteralPath $stale -Force }
            [void]$partNames.Remove((Get-PartName $r.Id $ch.Start))
            foreach ($k in $Calls.Keys) {
                $lines.Add((ConvertTo-CsvLine @((Format-Utc ([datetime]::UtcNow)), $r.Id, $cs, (Format-Utc $ch.End), $k, 'Success', 0, 0, $false, 'Idle (pre-flight)')))
                $done["$($r.Id)|$cs|$k"] = 'Success'; $doneRows["$($r.Id)|$cs|$k"] = 0
            }
        }
    }
    if ($lines.Count) { Add-Content -LiteralPath $ManifestFile -Value $lines -Encoding utf8NoBOM }
    $idleManifest = $lines.Count
    Write-Log "Idle pre-flight: $($idleIds.Count) of $($targets.Count) account(s) had no activity in the window"
}

$units = [System.Collections.Generic.List[object]]::new()
foreach ($ch in $chunks) {
    foreach ($b in $batches) {
        $pend = @($b | Where-Object { Test-Pending $_ $ch })
        if ($pend.Count) { $units.Add([pscustomobject]@{ Type = 'Extract'; Resources = $pend; Chunk = $ch }) }
    }
}
Write-Log "Extracting $($targets.Count) resource(s) x $($chunks.Count) chunk(s) x $($Calls.Count) call(s): $($units.Count) work unit(s) in $($batches.Count) batch(es)"
$extractClock = [System.Diagnostics.Stopwatch]::StartNew()
$extractRows = 0L
Invoke-Units $units.ToArray() -Activity 'Extraction' -OnResult {
    param($r)
    if ($r.Manifest.Count) { Add-Content -LiteralPath $ManifestFile -Value $r.Manifest -Encoding utf8NoBOM }
    foreach ($k in $r.Status.Keys) { $done[$k] = $r.Status[$k]; $doneRows[$k] = [long]$r.RowsBy[$k] }
    foreach ($p in $r.Parts) { [void]$partNames.Add($p) }
    $script:extractRows += $r.Rows
} -OnError {
    # A unit that failed as a whole still leaves a 'Failed' trace for each of its calls, so the run is
    # reported incomplete and the next resume retries it.
    param($r)
    $now = Format-Utc ([datetime]::UtcNow); $cs = Format-Utc $r.Unit.Chunk.Start; $ce = Format-Utc $r.Unit.Chunk.End
    $msg = ($r.Error -split "`n")[0]
    $lines = foreach ($res in $r.Unit.Resources) {
        foreach ($k in $Calls.Keys) { ConvertTo-CsvLine @($now, $res.Id, $cs, $ce, $k, 'Failed', 0, 0, $false, $msg); $done["$($res.Id)|$cs|$k"] = 'Failed' }
    }
    Add-Content -LiteralPath $ManifestFile -Value @($lines) -Encoding utf8NoBOM
}
$extractClock.Stop()
#endregion

#region merge -----------------------------------------------------------------------------------
# Only chunks whose calls all completed ('Success', or 'Truncated' = flagged partial data) in the latest
# manifest are merged: a part left behind by an interrupted or failed unit never leaks into the final CSV.
$latest = @{}
foreach ($m in (Import-Csv $ManifestFile)) { $latest["$($m.ResourceId)|$($m.ChunkStart)|$($m.Call)"] = $m }
$chunkOk = @{}
foreach ($m in $latest.Values) {
    $ck = "$($m.ResourceId)|$($m.ChunkStart)"
    if (-not $chunkOk.ContainsKey($ck)) { $chunkOk[$ck] = $true }
    if ($m.Status -notin 'Success', 'Truncated', 'NotSupported') { $chunkOk[$ck] = $false }
}
$mergeable = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($ck in $chunkOk.Keys) {
    if (-not $chunkOk[$ck]) { continue }
    $rid, $cs = $ck.Split('|', 2)
    if (@($Calls.Keys | Where-Object { -not $latest.ContainsKey("$ck|$_") }).Count) { continue }
    [void]$mergeable.Add((Get-PartName $rid (ConvertTo-UtcDate $cs)))
}

# Completeness: every target x chunk x call must have a latest 'Success' entry, and every successful chunk
# with rows must have its part file.
$missing = 0; $bad = 0; $noPart = 0
foreach ($r in $targets) {
    foreach ($ch in $chunks) {
        $cs = Format-Utc $ch.Start
        $rows = 0L
        foreach ($k in $Calls.Keys) {
            $m = $latest["$($r.Id)|$cs|$k"]
            if (-not $m) { $missing++ } elseif ($m.Status -notin $TerminalStatus) { $bad++ } else { $rows += [long]$m.Rows }
        }
        if ($rows -gt 0 -and -not $partNames.Contains((Get-PartName $r.Id $ch.Start))) { $noPart++ }
    }
}

Write-Log 'Merging part files'
$tmpOut = "$OutputCsv.tmp"
$out = [System.IO.StreamWriter]::new($tmpOut, $false, $Utf8)
$total = 0L
$skippedParts = 0
try {
    $out.WriteLine(($CsvColumns -join ','))
    foreach ($f in (Get-ChildItem -LiteralPath $PartsDir -File | Where-Object Extension -eq '.csv' | Sort-Object Name)) {
        if (-not $mergeable.Contains($f.Name)) { $skippedParts++; continue }
        $reader = [System.IO.StreamReader]::new($f.FullName, $Utf8)
        try { while ($null -ne ($line = $reader.ReadLine())) { if ($line) { $out.WriteLine($line); $total++ } } }
        finally { $reader.Dispose() }
    }
}
finally { $out.Dispose() }
Move-Item -LiteralPath $tmpOut -Destination $OutputCsv -Force

$limiterStats = [ordered]@{}
foreach ($kv in ($Limiters.Snapshot() | Sort-Object Key)) {
    $limiterStats[$kv.Key] = [ordered]@{ Calls = $kv.Value.Calls; Throttles = $kv.Value.Throttles; Slowdowns = $kv.Value.Slowdowns
        FinalRate = [math]::Round($kv.Value.Rate, 2); PeakRate = [math]::Round($kv.Value.PeakRate, 2); LowestRate = [math]::Round($kv.Value.LowRate, 2) }
}
$runInfo['LastRunAt'] = Format-Utc ([datetime]::UtcNow)
$runInfo['Completeness'] = [ordered]@{
    Status          = if ($missing + $bad + $noPart -eq 0) { 'Complete' } else { 'Incomplete' }
    ExpectedCalls   = $targets.Count * $chunks.Count * $Calls.Count
    MissingCalls    = $missing
    NotSuccessful   = $bad
    MissingParts    = $noPart
    SkippedParts    = $skippedParts
}
$runInfo['Performance'] = [ordered]@{
    MetricsApi          = $MetricsApiMode
    Parallel            = $Parallel
    SkipIdle            = -not $NoSkipIdle
    DurationSeconds     = [math]::Round($script:RunClock.Elapsed.TotalSeconds, 1)
    ExtractionSeconds   = [math]::Round($extractClock.Elapsed.TotalSeconds, 1)
    WorkUnits           = $units.Count
    Batches             = $batches.Count
    MaxBatchSize        = ($batches | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    SeriesBudget        = $seriesBudget
    IdleAccounts        = $idleIds.Count
    IdleManifestRows    = $idleManifest
    BatchCapability     = [ordered]@{}
    Requests            = $limiterStats
}
foreach ($k in ($BatchCap.Keys | Sort-Object)) {
    $reason = $null; [void]$BatchCapReason.TryGetValue($k, [ref]$reason)
    $runInfo.Performance.BatchCapability[$k] = if ($reason) { "$($BatchCap[$k]) ($reason)" } else { $BatchCap[$k] }
}
$runInfo | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $RunFile -Encoding utf8NoBOM

$thr = ($Limiters.Snapshot() | Where-Object { $_.Key -in 'arm-global', 'batch-global' } | ForEach-Object { $_.Value.Throttles } | Measure-Object -Sum).Sum
Write-Log ("Done in {0:N0} s (extraction {1:N0} s, {2} unit(s), throttle events {3}). Rows: {4} -> {5}" -f $script:RunClock.Elapsed.TotalSeconds, $extractClock.Elapsed.TotalSeconds, $units.Count, $thr, $total, $OutputCsv)
if ($failedTenants.Count) { Write-Log "Tenants not accessible (run 'az login --tenant <id>' and re-run with the same -OutputDirectory): $($failedTenants -join ', ')" 'WARN' }
if ($bad -or $missing -or $noPart) { Write-Log "Run incomplete: $bad call(s) not successful, $missing call(s) never recorded, $noPart part file(s) missing (see manifest.csv). Re-run with the same -OutputDirectory to retry them." 'WARN' }
if ($skippedParts) { Write-Log "$skippedParts part file(s) of incomplete chunks were not merged" 'WARN' }
$RunLock.Dispose()
#endregion
