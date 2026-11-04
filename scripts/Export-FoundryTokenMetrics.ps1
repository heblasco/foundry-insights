#Requires -Version 7.2
<#
.SYNOPSIS
    Extracts per-minute Azure OpenAI / Foundry token and request metrics from every accessible
    subscription (across tenants) into a CSV file.

.DESCRIPTION
    1. Resolves tenants (default: every tenant known to Azure CLI) and gets an ARM token per tenant.
    2. Lists enabled subscriptions per tenant and finds Microsoft.CognitiveServices/accounts with
       Azure Resource Graph (queried per tenant, never mixing tenants in one call).
    3. Validates each account through its metric definitions (must expose the Azure OpenAI metrics).
    4. Queries Azure Monitor per resource and per time chunk at PT1M:
         - Tokens   : ProcessedPromptTokens, GeneratedTokens (Total) split by ModelDeploymentName,
                      ModelVersion, Region. These metrics have NO ModelName dimension.
         - Requests : AzureOpenAIRequests (Total) split by ModelDeploymentName, ModelName,
                      ModelVersion, Region. Used as the historical source of ModelName.
         - PTU      : (optional) AzureOpenAIProvisionedManagedUtilizationV2 (Maximum, Average).
    5. Writes one atomic part file per resource/chunk plus a manifest of every call, so a run can be
       resumed safely and statistics can tell real zeros apart from missing data.
    6. Merges the parts into <OutputDirectory>\foundry-token-metrics.csv.

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
    Tenants to process. Default: distinct tenants from 'az account list --all'.

.PARAMETER SubscriptionId
    Optional subscription allow-list.

.PARAMETER Kind
    Account kinds considered (default OpenAI, AIServices). Use -IncludeAllKinds to test every
    Cognitive Services account; capability is always confirmed through metric definitions.

.PARAMETER IncludePtuUtilization
    Also extract AzureOpenAIProvisionedManagedUtilizationV2 (Maximum and Average per minute).

.PARAMETER Top
    Max time series per metrics call. If a call returns this many series, it is re-queried per
    deployment to avoid silent truncation. If a single deployment still reaches the limit the
    chunk is marked 'Truncated': it is excluded from statistics coverage and retried on resume
    (you may raise -Top when resuming).

.PARAMETER InventoryOnly
    Only discover subscriptions/resources/deployments (resources.csv), no metrics.

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
    [ValidateRange(10, 10000)] [int] $Top = 1000,
    [ValidateRange(1, 10)] [int] $MaxRetries = 6,
    [ValidateRange(10, 600)] [int] $TokenTimeoutSeconds = 60,
    [switch]   $InventoryOnly
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$SchemaVersion = 4
$Arm = 'https://management.azure.com'
$MetricsApi = '2023-10-01'
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8 = [System.Text.UTF8Encoding]::new($false)

$CsvColumns = @(
    'Timestamp', 'SubscriptionId', 'SubscriptionName', 'ResourceGroup', 'ResourceName', 'Region',
    'DeploymentName', 'ModelName', 'ModelVersion', 'Metric', 'Total',
    'TenantId', 'Kind', 'Aggregation', 'DeploymentSku', 'DeploymentCapacity', 'ModelNameSource', 'ResourceId'
)
$ManifestColumns = @('RecordedAt', 'ResourceId', 'ChunkStart', 'ChunkEnd', 'Call', 'Status', 'SeriesCount', 'Rows', 'Truncated', 'Message')

#region helpers ---------------------------------------------------------------------------------
function Write-Log([string]$Message, [string]$Level = 'INFO') {
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level, $Message
    if ($Level -eq 'WARN') { Write-Warning $Message } else { Write-Host $line }
    Add-Content -LiteralPath $script:LogFile -Value $line -Encoding utf8NoBOM
}

function ConvertTo-CsvField($Value) {
    if ($null -eq $Value) { return '' }
    $s = if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) { ([double]$Value).ToString('R', $Inv) }
    elseif ($Value -is [datetime]) { $Value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $Inv) }
    else { [string]$Value }
    if ($s.IndexOfAny([char[]]@(',', '"', "`r", "`n")) -ge 0) { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

function ConvertTo-CsvLine([object[]]$Values) { ($Values | ForEach-Object { ConvertTo-CsvField $_ }) -join ',' }

function ConvertTo-UtcDate($Value) {
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        return $Value.ToUniversalTime()
    }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime }
    return [datetimeoffset]::Parse([string]$Value, $Inv).UtcDateTime
}

function Format-Utc([datetime]$d) { $d.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $Inv) }

function Get-ArmToken([string]$Tenant, [switch]$Force) {
    $cached = $script:Tokens[$Tenant]
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    if (-not $Force -and $cached -and ($cached.ExpiresOn - $now) -gt 300) { return $cached.Token }
    # az may block waiting for interactive login (MFA/CA) on some tenants: bound it with a timeout.
    $psi = [System.Diagnostics.ProcessStartInfo]::new($script:AzExe)
    foreach ($a in @('account', 'get-access-token', '--tenant', $Tenant, '--resource', "$Arm/", '-o', 'json', '--only-show-errors')) { $psi.ArgumentList.Add($a) }
    $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.RedirectStandardInput = $true; $psi.UseShellExecute = $false
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.StandardInput.Close()
    $outTask = $proc.StandardOutput.ReadToEndAsync(); $null = $proc.StandardError.ReadToEndAsync()
    $json = $null
    if ($proc.WaitForExit($TokenTimeoutSeconds * 1000)) {
        if ($proc.ExitCode -eq 0) { $json = $outTask.GetAwaiter().GetResult() }
    }
    else {
        try { $proc.Kill($true) } catch { }
        Write-Log "Tenant ${Tenant}: token request timed out after $TokenTimeoutSeconds s (interactive login required?)" 'WARN'
    }
    $proc.Dispose()
    if (-not $json) {
        throw "Cannot get an ARM token for tenant $Tenant. Run: az login --tenant $Tenant"
    }
    $t = $json | ConvertFrom-Json
    $exp = if ($t.PSObject.Properties['expires_on'] -and $t.expires_on) { [long]$t.expires_on } else { $now + 1800 }
    $script:Tokens[$Tenant] = [pscustomobject]@{ Token = $t.accessToken; ExpiresOn = $exp }
    return $t.accessToken
}

class ArmException : System.Exception {
    [int]$StatusCode
    [string]$Body
    ArmException([int]$code, [string]$body, [string]$msg) : base($msg) { $this.StatusCode = $code; $this.Body = $body }
}

function Invoke-Arm {
    param([string]$Tenant, [string]$Uri, [string]$Method = 'GET', $Body)
    if ($Uri -notmatch '^https://') { $Uri = $Arm + $Uri }
    $attempt = 0
    $refreshed = $false
    while ($true) {
        $attempt++
        $headers = @{ Authorization = "Bearer $(Get-ArmToken $Tenant)" }
        $params = @{ Uri = $Uri; Method = $Method; Headers = $headers; SkipHttpErrorCheck = $true; TimeoutSec = 300 }
        if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 20 -Compress); $params.ContentType = 'application/json' }
        $resp = $null
        try { $resp = Invoke-WebRequest @params }
        catch {
            if ($attempt -le $MaxRetries) {
                $wait = [math]::Min(120, [math]::Pow(2, $attempt))
                Write-Verbose "Transport error ($($_.Exception.Message)); retry in $wait s"
                Start-Sleep -Seconds $wait; continue
            }
            throw [ArmException]::new(0, '', "Transport error: $($_.Exception.Message)")
        }
        $code = [int]$resp.StatusCode
        $content = [string]$resp.Content
        if ($code -ge 200 -and $code -lt 300) {
            if ([string]::IsNullOrWhiteSpace($content)) { return $null }
            return $content | ConvertFrom-Json -Depth 100
        }
        if ($code -eq 401 -and -not $refreshed) { $null = Get-ArmToken $Tenant -Force; $refreshed = $true; continue }
        if (($code -eq 429 -or $code -ge 500) -and $attempt -le $MaxRetries) {
            $ra = 0
            $raHeader = $resp.Headers['Retry-After']
            if ($raHeader) { [void][int]::TryParse(([string]($raHeader | Select-Object -First 1)), [ref]$ra) }
            $wait = if ($ra -gt 0) { $ra } else { [math]::Min(120, [math]::Pow(2, $attempt)) }
            Write-Verbose "HTTP $code on $Uri; retry $attempt/$MaxRetries in $wait s"
            Start-Sleep -Seconds $wait; continue
        }
        $msg = $content
        try { $e = $content | ConvertFrom-Json; if ($e.error.message) { $msg = "$($e.error.code): $($e.error.message)" } } catch { }
        throw [ArmException]::new($code, $content, "HTTP $code - $msg")
    }
}

function Get-ArmPaged([string]$Tenant, [string]$Uri) {
    $items = [System.Collections.Generic.List[object]]::new()
    while ($Uri) {
        $r = Invoke-Arm -Tenant $Tenant -Uri $Uri
        if ($r -and $r.PSObject.Properties['value'] -and $r.value) { $items.AddRange([object[]]$r.value) }
        $Uri = if ($r -and $r.PSObject.Properties['nextLink']) { $r.nextLink } else { $null }
    }
    return , $items
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

function Get-Prop($Obj, [string]$Path) {
    foreach ($p in $Path.Split('.')) {
        if ($null -eq $Obj -or -not $Obj.PSObject.Properties[$p]) { return $null }
        $Obj = $Obj.$p
    }
    return $Obj
}
#endregion

#region metrics ---------------------------------------------------------------------------------
$Calls = [ordered]@{
    Tokens   = @{ Metrics = 'ProcessedPromptTokens,GeneratedTokens'; Aggregation = 'Total'; Dims = @('ModelDeploymentName', 'ModelVersion', 'Region') }
    Requests = @{ Metrics = 'AzureOpenAIRequests'; Aggregation = 'Total'; Dims = @('ModelDeploymentName', 'ModelName', 'ModelVersion', 'Region', 'StatusCode') }
}
# Requests split by StatusCode are re-aggregated: AzureOpenAIRequests = all codes, ThrottledMetric = 429 only.
$ThrottledMetric = 'ThrottledRequests(429)'
if ($IncludePtuUtilization) {
    $Calls['PtuUtilization'] = @{ Metrics = 'AzureOpenAIProvisionedManagedUtilizationV2'; Aggregation = 'Maximum,Average'; Dims = @('ModelDeploymentName', 'ModelName', 'ModelVersion', 'Region') }
}
$RequiredMetrics = @('ProcessedPromptTokens', 'GeneratedTokens', 'AzureOpenAIRequests')

function New-MetricsUri([string]$ResourceId, [hashtable]$Call, [datetime]$From, [datetime]$To, [string]$Deployment, [switch]$Metadata) {
    $clauses = foreach ($d in $Call.Dims) {
        if ($Deployment -and $d -eq 'ModelDeploymentName') { "ModelDeploymentName eq '$($Deployment.Replace("'", "''"))'" } else { "$d eq '*'" }
    }
    $q = [ordered]@{
        'api-version'       = $MetricsApi
        metricnamespace     = 'Microsoft.CognitiveServices/accounts'
        metricnames         = $Call.Metrics
        timespan            = "$(Format-Utc $From)/$(Format-Utc $To)"
        interval            = 'PT1M'
        aggregation         = $Call.Aggregation
        '$filter'           = ($clauses -join ' and ')
        top                 = $Top
        AutoAdjustTimegrain = 'false'
        ValidateDimensions  = 'false'
    }
    if ($Metadata) {
        $q['resultType'] = 'metadata'; $q.Remove('interval'); $q.Remove('AutoAdjustTimegrain')
        $q['$filter'] = "ModelDeploymentName eq '*'"; $q['top'] = 10000
    }
    $qs = ($q.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join '&'
    return "$Arm$ResourceId/providers/Microsoft.Insights/metrics?$qs"
}

function Test-TooLarge([ArmException]$Ex) {
    if ($Ex.StatusCode -in 0, 413, 504) { return $true }
    if ($Ex.StatusCode -eq 400 -and $Ex.Message -match '(?i)exceed|too (many|large|much)|limit|timeout|timed out') { return $true }
    return $false
}

function ConvertTo-Series($Metric, $Ts) {
    $dims = @{}
    if ($Ts.PSObject.Properties['metadatavalues']) {
        foreach ($mv in @($Ts.metadatavalues)) { if ($mv) { $dims[$mv.name.value.ToLowerInvariant()] = [string]$mv.value } }
    }
    $data = if ($Ts.PSObject.Properties['data']) { @($Ts.data) } else { @() }
    [pscustomobject]@{ Metric = $Metric.name.value; Dims = $dims; Data = $data }
}

function Assert-MetricResponse($Resp) {
    if ($Resp.PSObject.Properties['interval'] -and $Resp.interval -and $Resp.interval -ne 'PT1M') {
        throw "Service returned interval $($Resp.interval) instead of PT1M"
    }
    foreach ($m in @($Resp.value)) {
        if ($m.PSObject.Properties['errorCode'] -and $m.errorCode -and $m.errorCode -ne 'Success') {
            throw "Metric $($m.name.value) error: $($m.errorCode) $($m.errorMessage)"
        }
    }
}

# Returns @{ Series = list of @{Metric; Dims; Data}; Truncated; Notes }
function Get-MetricSeries([object]$Res, [string]$CallName, [datetime]$From, [datetime]$To) {
    $call = $Calls[$CallName]
    $result = [pscustomobject]@{ Series = [System.Collections.Generic.List[object]]::new(); Truncated = $false; Notes = [System.Collections.Generic.List[string]]::new() }
    try {
        $resp = Invoke-Arm -Tenant $Res.TenantId -Uri (New-MetricsUri $Res.Id $call $From $To)
    }
    catch [ArmException] {
        $ex = $_.Exception
        if ((Test-TooLarge $ex) -and ($To - $From).TotalMinutes -gt 60) {
            $mid = $From.AddMinutes([math]::Floor(($To - $From).TotalMinutes / 2))
            Write-Log "  $CallName window $(Format-Utc $From) too large ($($ex.StatusCode)); splitting" 'WARN'
            $a = Get-MetricSeries $Res $CallName $From $mid
            $b = Get-MetricSeries $Res $CallName $mid $To
            $a.Series.AddRange($b.Series); $a.Notes.AddRange($b.Notes); $a.Truncated = $a.Truncated -or $b.Truncated
            return $a
        }
        throw
    }
    Assert-MetricResponse $resp

    $truncated = $false
    foreach ($m in $resp.value) { if (@($m.timeseries).Count -ge $Top) { $truncated = $true } }

    if (-not $truncated) {
        foreach ($m in $resp.value) { foreach ($ts in $m.timeseries) { $result.Series.Add((ConvertTo-Series $m $ts)) } }
        return $result
    }

    # top reached: re-query deployment by deployment so no series is silently dropped.
    $meta = Invoke-Arm -Tenant $Res.TenantId -Uri (New-MetricsUri $Res.Id $call $From $To -Metadata)
    Assert-MetricResponse $meta
    $deps = @($meta.value | ForEach-Object { $_.timeseries } | ForEach-Object { $_.metadatavalues } |
        Where-Object { $_.name.value -eq 'modeldeploymentname' } | ForEach-Object { $_.value } | Sort-Object -Unique)
    $result.Notes.Add("top=$Top reached; re-queried $($deps.Count) deployments individually")
    foreach ($dep in $deps) {
        $r = Invoke-Arm -Tenant $Res.TenantId -Uri (New-MetricsUri $Res.Id $call $From $To -Deployment $dep)
        Assert-MetricResponse $r
        foreach ($m in $r.value) {
            if (@($m.timeseries).Count -ge $Top) { $result.Truncated = $true; $result.Notes.Add("deployment $dep still returns top=$Top series") }
            foreach ($ts in $m.timeseries) { $result.Series.Add((ConvertTo-Series $m $ts)) }
        }
    }
    return $result
}

# Collapses the StatusCode dimension: one AzureOpenAIRequests series (all codes) plus one
# ThrottledMetric series (429) per remaining dimension set, so later dedupe never drops codes.
function Merge-StatusCodeSeries($R) {
    $groups = [ordered]@{}
    foreach ($s in $R.Series) {
        $code = [string]$s.Dims['statuscode']
        $dims = @{}
        foreach ($k in $s.Dims.Keys) { if ($k -ne 'statuscode') { $dims[$k] = $s.Dims[$k] } }
        $base = (@($dims.Keys | Sort-Object) | ForEach-Object { "$_=$($dims[$_])" }) -join '|'
        $names = @($s.Metric); if ($code -eq '429') { $names += $ThrottledMetric }
        foreach ($n in $names) {
            $gk = "$n|$base"
            if (-not $groups.Contains($gk)) { $groups[$gk] = @{ Metric = $n; Dims = $dims; Sum = [ordered]@{} } }
            $g = $groups[$gk]
            foreach ($p in $s.Data) {
                if (-not $p.PSObject.Properties['total'] -or $null -eq $p.total) { continue }
                $tk = Format-Utc (ConvertTo-UtcDate $p.timeStamp)
                $g.Sum[$tk] = $(if ($g.Sum.Contains($tk)) { $g.Sum[$tk] } else { 0.0 }) + [double]$p.total
            }
        }
    }
    $R.Series.Clear()
    foreach ($g in $groups.Values) {
        $data = @(foreach ($e in $g.Sum.GetEnumerator()) { [pscustomobject]@{ timeStamp = $e.Key; total = $e.Value } })
        $R.Series.Add([pscustomobject]@{ Metric = $g.Metric; Dims = $g.Dims; Data = $data })
    }
    return $R
}

function ConvertTo-PerMinute($Limits, [string]$Key) {
    $l = @($Limits | Where-Object { $_ -and $_.key -eq $Key }) | Select-Object -First 1
    if (-not $l -or $null -eq $l.count) { return $null }
    $period = if ($l.PSObject.Properties['renewalPeriod'] -and $l.renewalPeriod) { [double]$l.renewalPeriod } else { 60.0 }
    return [double]$l.count * 60.0 / $period
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
#endregion

#region run setup -------------------------------------------------------------------------------
$azCmd = Get-Command az -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $azCmd) { throw 'Azure CLI (az) is required. Install it and run az login.' }
$script:AzExe = $azCmd.Source

$RunStartedAt = [datetime]::UtcNow
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $OutputRoot ('{0}_{1:yyyyMMdd-HHmmss}Z' -f $RunName, $RunStartedAt) }
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$PartsDir = Join-Path $OutputDirectory 'parts'
New-Item -ItemType Directory -Force -Path $PartsDir | Out-Null
$script:LogFile = Join-Path $OutputDirectory 'run.log'
$ManifestFile = Join-Path $OutputDirectory 'manifest.csv'
$RunFile = Join-Path $OutputDirectory 'run.json'
$ResourcesFile = Join-Path $OutputDirectory 'resources.csv'
$OutputCsv = Join-Path $OutputDirectory 'foundry-token-metrics.csv'
$script:Tokens = @{}

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
    Interval      = 'PT1M'
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

# Latest manifest status per (resource|chunk|call)
$done = @{}
foreach ($m in (Import-Csv $ManifestFile)) { $done["$($m.ResourceId)|$($m.ChunkStart)|$($m.Call)"] = $m.Status }

Write-Log "Window $(Format-Utc $StartTime) -> $(Format-Utc $EndTime) (PT1M, chunks of $ChunkHours h)"
#endregion

#region discovery -------------------------------------------------------------------------------
if (-not $TenantId) {
    $TenantId = @((& az account list --all -o json | ConvertFrom-Json) | ForEach-Object { $_.tenantId } | Where-Object { $_ } | Sort-Object -Unique)
}
Write-Log "Tenants to process: $($TenantId.Count)"

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
$resSeen = @{}
foreach ($t in $tenantSubs.Keys) {
    try { $rows = Invoke-ResourceGraph $t @($tenantSubs[$t].Id) $argQuery }
    catch { Write-Log "Resource Graph failed for tenant ${t}: $($_.Exception.Message)" 'WARN'; continue }
    foreach ($r in $rows) {
        $rk = $r.id.ToLowerInvariant()
        if ($resSeen.ContainsKey($rk)) { continue }
        $resSeen[$rk] = $t
        $resources.Add([pscustomobject]@{
                TenantId = $t; Id = $r.id; Name = $r.name; ResourceGroup = $r.resourceGroup
                SubscriptionId = $r.subscriptionId; SubscriptionName = $subsSeen[$r.subscriptionId]
                Kind = $r.kind; Location = $r.location; Sku = $r.sku
                Status = 'Pending'; Message = ''; Deployments = @{}
            })
    }
}
Write-Log "Candidate Cognitive Services accounts: $($resources.Count)"

foreach ($res in $resources) {
    try {
        $defs = Invoke-Arm -Tenant $res.TenantId -Uri "$($res.Id)/providers/Microsoft.Insights/metricDefinitions?api-version=$MetricsApi&metricnamespace=Microsoft.CognitiveServices/accounts"
        $names = @($defs.value | ForEach-Object { $_.name.value })
        $missing = @($RequiredMetrics | Where-Object { $_ -notin $names })
        if ($missing) { $res.Status = 'NotOpenAI'; $res.Message = "Missing metrics: $($missing -join ',')"; continue }
    }
    catch {
        $code = if ($_.Exception -is [ArmException]) { $_.Exception.StatusCode } else { -1 }
        $res.Status = if ($code -in 401, 403) { 'Forbidden' } else { 'Error' }
        $res.Message = $_.Exception.Message; continue
    }
    try {
        foreach ($d in (Get-ArmPaged $res.TenantId "$($res.Id)/deployments?api-version=2024-10-01")) {
            $res.Deployments[$d.name] = [pscustomobject]@{
                Model = Get-Prop $d 'properties.model.name'; Version = [string](Get-Prop $d 'properties.model.version')
                Format = Get-Prop $d 'properties.model.format'
                Sku = Get-Prop $d 'sku.name'; Capacity = Get-Prop $d 'sku.capacity'
                TokensPerMinute = ConvertTo-PerMinute (Get-Prop $d 'properties.rateLimits') 'token'
                RequestsPerMinute = ConvertTo-PerMinute (Get-Prop $d 'properties.rateLimits') 'request'
            }
        }
    }
    catch { $res.Message = "Deployments not readable: $($_.Exception.Message)" }
    $res.Status = 'OpenAI'
}

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
$quotaFailures = 0
foreach ($grp in ($targets | Group-Object SubscriptionId, Location)) {
    $first = $grp.Group[0]
    $base = "/subscriptions/$($first.SubscriptionId)/providers/Microsoft.CognitiveServices/locations/$($first.Location)"
    $usages = @(); $models = @()
    try { $usages = Get-ArmPaged $first.TenantId "$base/usages?api-version=2024-10-01" }
    catch { $quotaFailures++; Write-Log "Quota usages not readable for $($first.SubscriptionName)/$($first.Location): $($_.Exception.Message)" 'WARN' }
    try { $models = Get-ArmPaged $first.TenantId "$base/models?api-version=2024-10-01" }
    catch { Write-Log "Model catalog not readable for $($first.SubscriptionName)/$($first.Location): $($_.Exception.Message)" 'WARN' }
    $usageByName = @{}
    foreach ($u in $usages) { $usageByName[$u.name.value] = $u }
    $mapped = @{}
    foreach ($res in $grp.Group) {
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

[ordered]@{
    Fingerprint     = $fingerprint
    RunName         = Split-Path $OutputDirectory -Leaf
    RunTimestamp    = if ($prev -and $prev['RunTimestamp']) { $prev.RunTimestamp } else { Format-Utc $RunStartedAt }
    LastRunAt       = Format-Utc ([datetime]::UtcNow)
    Tenants         = $TenantId
    FailedTenants   = @($failedTenants)
    Subscriptions   = $subsSeen.Count
    TargetResources = @($targets | ForEach-Object Id)
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $RunFile -Encoding utf8NoBOM

if ($InventoryOnly) { Write-Log "Inventory written to $ResourcesFile"; return }
#endregion

#region extraction ------------------------------------------------------------------------------
$chunks = [System.Collections.Generic.List[object]]::new()
for ($c = $StartTime; $c -lt $EndTime; $c = $c.AddHours($ChunkHours)) {
    $e = $c.AddHours($ChunkHours); if ($e -gt $EndTime) { $e = $EndTime }
    $chunks.Add([pscustomobject]@{ Start = $c; End = $e })
}
$totalWork = $targets.Count * $chunks.Count
Write-Log "Extracting $($targets.Count) resource(s) x $($chunks.Count) chunk(s) x $($Calls.Count) call(s)"

function Resolve-Model($Res, [string]$Dep, [string]$Ver, [string]$Region, [string]$Ts, [hashtable]$ByMinute, [hashtable]$ByChunk) {
    # 1) model that served requests for this deployment/version/region in the same minute
    $set = $ByMinute["$Dep|$Ver|$Region|$Ts"]
    if ($set -and $set.Count -eq 1) { return @(@($set)[0], 'RequestsMetric') }
    # 2) the only model seen for this deployment/version in the chunk
    $chunkSet = $ByChunk["$Dep|$Ver"]
    if ($chunkSet -and $chunkSet.Count -eq 1) { return @(@($chunkSet)[0], 'RequestsChunk') }
    $cand = if ($set -and $set.Count) { $set } elseif ($chunkSet -and $chunkSet.Count) { $chunkSet } else { $null }
    # 3) current ARM deployment, if it matches the version (and any request-side candidates)
    $cur = $Res.Deployments[$Dep]
    if ($cur -and ($cur.Version -eq $Ver -or [string]::IsNullOrEmpty($Ver)) -and (-not $cand -or $cand.Contains([string]$cur.Model))) {
        return @($cur.Model, 'CurrentDeployment')
    }
    if ($cand) { return @((@($cand | Sort-Object) -join '|'), 'Ambiguous') }
    return @('Unknown', 'Unknown')
}

$sha = [System.Security.Cryptography.SHA256]::Create()
$work = 0
foreach ($res in $targets) {
    $resHash = [System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($res.Id.ToLowerInvariant()))).Replace('-', '').Substring(0, 16)
    foreach ($ch in $chunks) {
        $work++
        $cs = Format-Utc $ch.Start; $ce = Format-Utc $ch.End
        Write-Progress -Activity 'Extracting Azure OpenAI metrics' -Status "$($res.Name) $cs" -PercentComplete ([int](100 * $work / [math]::Max(1, $totalWork)))
        # Only 'Success' is terminal. 'Truncated' chunks are retried (e.g. after raising -Top) and are
        # never counted as observed coverage by the statistics.
        $pending = @($Calls.Keys | Where-Object { $done["$($res.Id)|$cs|$_"] -ne 'Success' })
        if (-not $pending) { continue }

        # All calls of an incomplete chunk are re-run: the part file is rewritten atomically.
        $results = [ordered]@{}
        foreach ($callName in $Calls.Keys) {
            try { $results[$callName] = Get-MetricSeries $res $callName $ch.Start $ch.End }
            catch { $results[$callName] = [string]$_.Exception.Message; Write-Log "  $($res.Name) $cs $callName FAILED: $($_.Exception.Message)" 'WARN' }
        }
        if ($results['Requests'] -isnot [string]) { $null = Merge-StatusCodeSeries $results['Requests'] }

        # Historical ModelName from AzureOpenAIRequests (tokens metrics lack that dimension):
        # per minute (exact) and per chunk (fallback), keyed by deployment/version[/region/minute].
        $byMinute = @{}; $byChunk = @{}
        if ($results['Requests'] -isnot [string]) {
            foreach ($s in $results['Requests'].Series) {
                $mn = $s.Dims['modelname']; if (-not $mn) { continue }
                $d = $s.Dims['modeldeploymentname']; $vv = $s.Dims['modelversion']
                $rg = if ($s.Dims['region']) { ($s.Dims['region'] -replace '\s', '').ToLowerInvariant() } else { $res.Location }
                foreach ($p in $s.Data) {
                    if (-not $p.PSObject.Properties['total'] -or $null -eq $p.total -or [double]$p.total -eq 0) { continue }
                    $pts = ConvertTo-UtcDate $p.timeStamp
                    if ($pts -lt $ch.Start -or $pts -ge $ch.End) { continue }
                    $mk = "$d|$vv|$rg|$(Format-Utc $pts)"
                    if (-not $byMinute.ContainsKey($mk)) { $byMinute[$mk] = [System.Collections.Generic.HashSet[string]]::new() }
                    [void]$byMinute[$mk].Add($mn)
                    if (-not $byChunk.ContainsKey("$d|$vv")) { $byChunk["$d|$vv"] = [System.Collections.Generic.HashSet[string]]::new() }
                    [void]$byChunk["$d|$vv"].Add($mn)
                }
            }
        }

        $partFile = Join-Path $PartsDir ('{0}_{1:yyyyMMddHHmm}.csv' -f $resHash, $ch.Start)
        $tmp = "$partFile.tmp"
        $sw = [System.IO.StreamWriter]::new($tmp, $false, $Utf8)
        $manifestLines = [System.Collections.Generic.List[string]]::new()
        try {
            foreach ($callName in $Calls.Keys) {
                $r = $results[$callName]
                if ($r -is [string]) {
                    $manifestLines.Add((ConvertTo-CsvLine @((Format-Utc ([datetime]::UtcNow)), $res.Id, $cs, $ce, $callName, 'Failed', 0, 0, $false, $r)))
                    continue
                }
                $rows = 0
                $seen = [System.Collections.Generic.HashSet[string]]::new()
                $aggs = if ($callName -eq 'PtuUtilization') { @('maximum', 'average') } else { @('total') }
                foreach ($s in $r.Series) {
                    $dep = [string]$s.Dims['modeldeploymentname']; $ver = [string]$s.Dims['modelversion']
                    $region = if ($s.Dims['region']) { ($s.Dims['region'] -replace '\s', '').ToLowerInvariant() } else { $res.Location }
                    $hasModelDim = $s.Dims.ContainsKey('modelname') -and $s.Dims['modelname']
                    $cur = $res.Deployments[$dep]
                    $depSku = if ($cur) { $cur.Sku } else { '' }
                    $depCap = if ($cur) { $cur.Capacity } else { '' }
                    foreach ($p in $s.Data) {
                        $ts = ConvertTo-UtcDate $p.timeStamp
                        if ($ts -lt $ch.Start -or $ts -ge $ch.End) { continue }   # half-open [start, end)
                        $tsS = Format-Utc $ts
                        $model = $null
                        foreach ($agg in $aggs) {
                            if (-not $p.PSObject.Properties[$agg]) { continue }
                            $v = $p.$agg
                            if ($null -eq $v -or [double]$v -eq 0) { continue }
                            if ($null -eq $model) {
                                if ($hasModelDim) { $model = $s.Dims['modelname']; $src = 'Metric' }
                                else { $model, $src = Resolve-Model $res $dep $ver $region $tsS $byMinute $byChunk }
                            }
                            $aggName = $agg.Substring(0, 1).ToUpperInvariant() + $agg.Substring(1)
                            if (-not $seen.Add("$tsS|$dep|$ver|$region|$($s.Metric)|$aggName|$model")) { continue }
                            $sw.WriteLine((ConvertTo-CsvLine @(
                                        $tsS, $res.SubscriptionId, $res.SubscriptionName, $res.ResourceGroup, $res.Name, $region,
                                        $dep, $model, $ver, $s.Metric, [double]$v,
                                        $res.TenantId, $res.Kind, $aggName, $depSku, $depCap, $src, $res.Id)))
                            $rows++
                        }
                    }
                }
                $status = if ($r.Truncated) { 'Truncated' } else { 'Success' }
                $manifestLines.Add((ConvertTo-CsvLine @((Format-Utc ([datetime]::UtcNow)), $res.Id, $cs, $ce, $callName, $status,
                            $r.Series.Count, $rows, $r.Truncated, ($r.Notes -join '; '))))
            }
        }
        finally { $sw.Dispose() }
        Move-Item -LiteralPath $tmp -Destination $partFile -Force
        Add-Content -LiteralPath $ManifestFile -Value $manifestLines -Encoding utf8NoBOM
        foreach ($callName in $Calls.Keys) {
            $done["$($res.Id)|$cs|$callName"] = if ($results[$callName] -is [string]) { 'Failed' } elseif ($results[$callName].Truncated) { 'Truncated' } else { 'Success' }
        }
    }
}
$sha.Dispose()
Write-Progress -Activity 'Extracting Azure OpenAI metrics' -Completed
#endregion

#region merge -----------------------------------------------------------------------------------
Write-Log 'Merging part files'
$tmpOut = "$OutputCsv.tmp"
$out = [System.IO.StreamWriter]::new($tmpOut, $false, $Utf8)
$total = 0L
try {
    $out.WriteLine(($CsvColumns -join ','))
    foreach ($f in (Get-ChildItem -LiteralPath $PartsDir -File | Where-Object Extension -eq '.csv' | Sort-Object Name)) {
        $reader = [System.IO.StreamReader]::new($f.FullName, $Utf8)
        try { while ($null -ne ($line = $reader.ReadLine())) { if ($line) { $out.WriteLine($line); $total++ } } }
        finally { $reader.Dispose() }
    }
}
finally { $out.Dispose() }
Move-Item -LiteralPath $tmpOut -Destination $OutputCsv -Force

$latest = @{}
foreach ($m in (Import-Csv $ManifestFile)) { $latest["$($m.ResourceId)|$($m.ChunkStart)|$($m.Call)"] = $m.Status }
$bad = @($latest.GetEnumerator() | Where-Object { $_.Value -ne 'Success' })
Write-Log "Done. Rows: $total -> $OutputCsv"
if ($failedTenants.Count) { Write-Log "Tenants not accessible (run 'az login --tenant <id>' and re-run with the same -OutputDirectory): $($failedTenants -join ', ')" 'WARN' }
if ($bad.Count) { Write-Log "$($bad.Count) call(s) not successful (see manifest.csv). Re-run with the same -OutputDirectory to retry failed ones." 'WARN' }
#endregion
