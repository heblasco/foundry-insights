# Shared code for Export-FoundryTokenMetrics.ps1: dot-sourced by the coordinator and by every parallel
# worker runspace. Functions read their settings ($Top, $MaxRetries, $Calls, $Interval, $TokenCache,
# $Limiters, $BatchCap, ...) from the caller's scope, which the coordinator and the workers both set.

$Arm = 'https://management.azure.com'
$MetricsAudience = 'https://metrics.monitor.azure.com'
$MetricsApiVersion = '2023-10-01'
$MetricsNamespace = 'Microsoft.CognitiveServices/accounts'
$ThrottledMetric = 'ThrottledRequests(429)'
$ClientErrorMetric = 'ClientErrors(4xx)'   # 4xx except 429
$ServerErrorMetric = 'ServerErrors(5xx)'
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8 = [System.Text.UTF8Encoding]::new($false)

# Compiled once per process; types are visible from every runspace.
if (-not ('FoundryInsights.LimiterSet' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace FoundryInsights {
    public class ArmException : Exception {
        public int StatusCode;
        public string Body;
        public ArmException(int code, string body, string msg) : base(msg) { StatusCode = code; Body = body; }
    }

    // Token bucket with AIMD: starts conservative, grows slowly after sustained success, halves and
    // pauses every caller sharing the bucket when the service throttles.
    public sealed class Limiter {
        readonly object _l = new object();
        public double Rate, Min, Max, Step, PeakRate, LowRate;
        double _tokens = 1;
        DateTime _last = DateTime.UtcNow, _pause = DateTime.MinValue;
        int _clean;
        public long Calls, Throttles, Slowdowns;
        public Limiter(double initial, double min, double max) {
            Rate = initial; Min = min; Max = max; Step = Math.Max(0.05, initial * 0.1); PeakRate = initial; LowRate = initial;
        }
        public TimeSpan TryTake() {
            lock (_l) {
                var now = DateTime.UtcNow;
                if (now < _pause) return _pause - now;
                _tokens = Math.Min(Math.Max(1.0, Rate), _tokens + (now - _last).TotalSeconds * Rate);
                _last = now;
                if (_tokens >= 1) { _tokens -= 1; Calls++; return TimeSpan.Zero; }
                return TimeSpan.FromSeconds((1 - _tokens) / Rate);
            }
        }
        public void Success() {
            lock (_l) { if (++_clean >= 20) { Rate = Math.Min(Max, Rate + Step); _clean = 0; if (Rate > PeakRate) PeakRate = Rate; } }
        }
        public void Throttle(double pauseSeconds) {
            lock (_l) {
                Throttles++; Rate = Math.Max(Min, Rate / 2); _clean = 0; _tokens = 0;
                if (Rate < LowRate) LowRate = Rate;
                var until = DateTime.UtcNow.AddSeconds(pauseSeconds);
                if (until > _pause) _pause = until;
            }
        }
        public void Slow() {
            lock (_l) { Slowdowns++; Rate = Math.Max(Min, Rate / 2); _clean = 0; if (Rate < LowRate) LowRate = Rate; }
        }
    }

    public sealed class LimiterSet {
        readonly ConcurrentDictionary<string, Limiter> _d = new ConcurrentDictionary<string, Limiter>();
        readonly ConcurrentDictionary<string, double[]> _cfg = new ConcurrentDictionary<string, double[]>();
        public void Configure(string prefix, double initial, double min, double max) { _cfg[prefix] = new[] { initial, min, max }; }
        public Limiter Get(string key) {
            return _d.GetOrAdd(key, k => {
                var p = k.Split('|')[0];
                double[] c;
                if (!_cfg.TryGetValue(p, out c)) c = new[] { 2.0, 0.2, 5.0 };
                return new Limiter(c[0], c[1], c[2]);
            });
        }
        public void Acquire(string[] keys) {
            foreach (var k in keys) {
                var l = Get(k);
                while (true) {
                    var w = l.TryTake();
                    if (w <= TimeSpan.Zero) break;
                    Thread.Sleep(w < TimeSpan.FromMilliseconds(5) ? TimeSpan.FromMilliseconds(5) : w);
                }
            }
        }
        public void Success(string[] keys) { foreach (var k in keys) Get(k).Success(); }
        public void Throttle(string[] keys, double pauseSeconds) { foreach (var k in keys) Get(k).Throttle(pauseSeconds); }
        public void Slow(string[] keys) { foreach (var k in keys) Get(k).Slow(); }
        public KeyValuePair<string, Limiter>[] Snapshot() { return _d.ToArray(); }
    }
}
'@
}

#region helpers ---------------------------------------------------------------------------------
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

function Get-Prop($Obj, [string]$Path) {
    foreach ($p in $Path.Split('.')) {
        if ($null -eq $Obj -or -not $Obj.PSObject.Properties[$p]) { return $null }
        $Obj = $Obj.$p
    }
    return $Obj
}

function ConvertTo-PerMinute($Limits, [string]$Key) {
    $l = @($Limits | Where-Object { $_ -and $_.key -eq $Key }) | Select-Object -First 1
    if (-not $l -or $null -eq $l.count) { return $null }
    $period = if ($l.PSObject.Properties['renewalPeriod'] -and $l.renewalPeriod) { [double]$l.renewalPeriod } else { 60.0 }
    return [double]$l.count * 60.0 / $period
}

function Get-IntervalMinutes([string]$Iso) {
    if ($Iso -match '^PT(\d+)M$') { return [int]$Matches[1] }
    if ($Iso -match '^PT(\d+)H$') { return 60 * [int]$Matches[1] }
    if ($Iso -eq 'P1D') { return 1440 }
    throw "Unsupported interval $Iso"
}
#endregion

#region auth and HTTP ---------------------------------------------------------------------------
# Tokens are cached per tenant + audience in a ConcurrentDictionary shared with the workers. Only the
# coordinator ($CanRefreshTokens) calls Azure CLI; workers read the cache and wait for its refresh.
function Get-AccessToken([string]$Tenant, [string]$Audience = "$Arm/", [switch]$Force) {
    $key = "$Tenant|$Audience"
    $cached = $null
    [void]$TokenCache.TryGetValue($key, [ref]$cached)
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    if (-not $Force -and $cached -and ($cached.ExpiresOn - $now) -gt 300) { return $cached.Token }
    if (-not $CanRefreshTokens) {
        if ($cached -and ($cached.ExpiresOn - $now) -gt 30) { return $cached.Token }
        throw [FoundryInsights.ArmException]::new(401, '', "No valid token for tenant $Tenant ($Audience) in worker")
    }
    $t = switch ($AuthMode) {
        'AzCli' { Get-AzCliToken $Tenant $Audience }
        'AzPowerShell' { Get-AzPowerShellToken $Tenant $Audience }
        'ManagedIdentity' { Get-ManagedIdentityToken $Audience }
        default { throw "Unsupported AuthMode '$AuthMode'" }
    }
    if ($t.ExpiresOn -le $now) { $t.ExpiresOn = $now + 1800 }
    # Workers set Rejected on a 401 so the coordinator force-refreshes the token on its next poll.
    $t | Add-Member -NotePropertyName Rejected -NotePropertyValue $false -Force
    $TokenCache[$key] = $t
    return $t.Token
}

# Picks the first AVAILABLE provider (installed / reachable). It never falls through to another provider after
# an authentication failure, so a run always uses one well-known identity.
function Resolve-AuthMode([string]$Mode) {
    if ($Mode -ne 'Auto') { return $Mode }
    if (Get-Command az -CommandType Application -ErrorAction SilentlyContinue) { return 'AzCli' }
    if (Get-Module -ListAvailable Az.Accounts) { return 'AzPowerShell' }
    if ($env:IDENTITY_ENDPOINT) { return 'ManagedIdentity' }
    throw 'No authentication provider found: install Azure CLI (az login) or Az.Accounts (Connect-AzAccount), or use -AuthMode ManagedIdentity.'
}

function Get-AzPowerShellToken([string]$Tenant, [string]$Audience) {
    Import-Module Az.Accounts -ErrorAction Stop -Verbose:$false
    $p = @{ ResourceUrl = $Audience; ErrorAction = 'Stop' }
    if ($Tenant) { $p.TenantId = $Tenant }
    $r = Get-AzAccessToken @p
    # Az.Accounts >= 5 returns a SecureString token.
    $tok = if ($r.Token -is [securestring]) { ConvertFrom-SecureString $r.Token -AsPlainText } else { $r.Token }
    [pscustomobject]@{ Token = $tok; ExpiresOn = [DateTimeOffset]::new($r.ExpiresOn.UtcDateTime).ToUnixTimeSeconds() }
}

# App Service / Container Apps / Functions (IDENTITY_ENDPOINT) or VM / VMSS / AKS (IMDS). Single tenant by design.
function Get-ManagedIdentityToken([string]$Audience) {
    $res = [uri]::EscapeDataString($Audience)
    $cid = if ($ManagedIdentityClientId) { "&client_id=$([uri]::EscapeDataString($ManagedIdentityClientId))" } else { '' }
    if ($env:IDENTITY_ENDPOINT -and $env:IDENTITY_HEADER) {
        $r = Invoke-RestMethod -Uri "$($env:IDENTITY_ENDPOINT)?api-version=2019-08-01&resource=$res$cid" -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER } -TimeoutSec $TokenTimeoutSeconds
    }
    else {
        $r = Invoke-RestMethod -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=$res$cid" -Headers @{ Metadata = 'true' } -NoProxy -TimeoutSec $TokenTimeoutSeconds
    }
    [pscustomobject]@{ Token = $r.access_token; ExpiresOn = [long]$r.expires_on }
}

function Get-AzCliToken([string]$Tenant, [string]$Audience) {
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $psi = [System.Diagnostics.ProcessStartInfo]::new($AzExe)
    $azArgs = @('account', 'get-access-token', '--resource', $Audience, '-o', 'json', '--only-show-errors')
    if ($Tenant) { $azArgs += @('--tenant', $Tenant) }
    foreach ($a in $azArgs) { $psi.ArgumentList.Add($a) }
    $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.RedirectStandardInput = $true; $psi.UseShellExecute = $false
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.StandardInput.Close()
    $outTask = $proc.StandardOutput.ReadToEndAsync(); $null = $proc.StandardError.ReadToEndAsync()
    $json = $null
    # az may block waiting for interactive login (MFA/CA) on some tenants: bound it with a timeout.
    if ($proc.WaitForExit($TokenTimeoutSeconds * 1000)) {
        if ($proc.ExitCode -eq 0) { $json = $outTask.GetAwaiter().GetResult() }
    }
    else {
        try { $proc.Kill($true) } catch { }
        Write-Log "Tenant ${Tenant}: token request timed out after $TokenTimeoutSeconds s (interactive login required?)" 'WARN'
    }
    $proc.Dispose()
    if (-not $json) { throw "Cannot get a token for tenant $Tenant ($Audience). Run: az login --tenant $Tenant" }
    $t = $json | ConvertFrom-Json
    $exp = if ($t.PSObject.Properties['expires_on'] -and $t.expires_on) { [long]$t.expires_on } else { $now + 1800 }
    [pscustomobject]@{ Token = $t.accessToken; ExpiresOn = $exp }
}

# Rate-limit buckets touched by a request:
#   batch-global, batch|<sub>|<region>             metrics:getBatch (regional data plane)
#   arm-global, arm|<sub>, armmetrics|<sub>         ARM (metrics API has a lower per-subscription budget)
#   arm-global, arg                                 Resource Graph
function Get-LimiterKeys([string]$Uri) {
    $u = [uri]$Uri
    $path = $u.AbsolutePath
    if ($u.Host -like '*.metrics.monitor.azure.com') {
        $sub = ($path.Split('/', [StringSplitOptions]::RemoveEmptyEntries))[1].ToLowerInvariant()
        return [string[]]@('batch-global', "batch|$sub|$($u.Host.Split('.')[0].ToLowerInvariant())")
    }
    $keys = [System.Collections.Generic.List[string]]::new()
    $keys.Add('arm-global')
    if ($path -match '(?i)^/providers/Microsoft\.ResourceGraph') { $keys.Add('arg') }
    elseif ($path -match '(?i)^/subscriptions/([^/]+)') {
        $sub = $Matches[1].ToLowerInvariant()
        $keys.Add("arm|$sub")
        if ($path -match '(?i)/providers/Microsoft\.Insights/metrics$') { $keys.Add("armmetrics|$sub") }
    }
    return $keys.ToArray()
}

function Invoke-Arm {
    param([string]$Tenant, [string]$Uri, [string]$Method = 'GET', $Body, [string]$Audience = "$Arm/")
    if ($Uri -notmatch '^https?://') { $Uri = $Arm + $Uri }
    $keys = Get-LimiterKeys $Uri
    $attempt = 0
    $authRetries = 0
    while ($true) {
        $attempt++
        $token = Get-AccessToken $Tenant $Audience
        $Limiters.Acquire($keys)
        $params = @{ Uri = $Uri; Method = $Method; Headers = @{ Authorization = "Bearer $token" }; SkipHttpErrorCheck = $true; TimeoutSec = 300 }
        if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 20 -Compress); $params.ContentType = 'application/json' }
        $resp = $null
        try { $resp = Invoke-WebRequest @params }
        catch {
            if ($attempt -le $MaxRetries) {
                Start-Sleep -Seconds (Get-Random -Minimum 1.0 -Maximum ([math]::Min(120, [math]::Pow(2, $attempt)) + 1))
                continue
            }
            throw [FoundryInsights.ArmException]::new(0, '', "Transport error: $($_.Exception.Message)")
        }
        $code = [int]$resp.StatusCode
        $content = if ($resp.Content -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($resp.Content) } else { [string]$resp.Content }
        if ($code -ge 200 -and $code -lt 300) {
            $Limiters.Success($keys)
            $remaining = $resp.Headers['x-ms-ratelimit-remaining-subscription-reads']
            if ($remaining) {
                $rem = 0
                if ([int]::TryParse([string]($remaining | Select-Object -First 1), [ref]$rem) -and $rem -lt 100) { $Limiters.Slow($keys) }
            }
            if ([string]::IsNullOrWhiteSpace($content)) { return $null }
            return $content | ConvertFrom-Json -Depth 100
        }
        if ($code -eq 401 -and $authRetries -lt 3) {
            $authRetries++
            if ($CanRefreshTokens) { $null = Get-AccessToken $Tenant $Audience -Force }
            else {
                # Ask the coordinator to force-refresh the shared token, then wait for it.
                $c = $null
                if ($TokenCache.TryGetValue("$Tenant|$Audience", [ref]$c) -and $c) { $c.Rejected = $true }
                Start-Sleep -Seconds 15
            }
            continue
        }
        if (($code -in 429, 503, 529 -or $code -ge 500) -and $attempt -le $MaxRetries) {
            $ra = 0
            $raHeader = $resp.Headers['Retry-After']
            if ($raHeader) { [void][int]::TryParse(([string]($raHeader | Select-Object -First 1)), [ref]$ra) }
            $wait = if ($ra -gt 0) { [double]$ra } else { Get-Random -Minimum 1.0 -Maximum ([math]::Min(120, [math]::Pow(2, $attempt)) + 1) }
            # Throttling pauses every worker sharing the bucket, not only this one.
            if ($code -in 429, 503, 529) { $Limiters.Throttle($keys, $wait) }
            Start-Sleep -Seconds $wait
            continue
        }
        $msg = $content
        try { $e = $content | ConvertFrom-Json; if ($e.error.message) { $msg = "$($e.error.code): $($e.error.message)" } } catch { }
        throw [FoundryInsights.ArmException]::new($code, $content, "HTTP $code - $msg")
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

function Get-StatusCode($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    if ($ex -is [FoundryInsights.ArmException]) { return $ex.StatusCode }
    return -1
}
#endregion

#region metrics ---------------------------------------------------------------------------------
function New-MetricsUri([string]$ResourceId, [hashtable]$Call, [datetime]$From, [datetime]$To, [string]$Deployment, [string[]]$Deployments, [switch]$Metadata) {
    $clauses = foreach ($d in $Call.Dims) {
        if ($Deployment -and $d -eq 'ModelDeploymentName') { "ModelDeploymentName eq '$($Deployment.Replace("'", "''"))'" }
        elseif ($Deployments -and $d -eq 'ModelDeploymentName') { Get-DeploymentClause $Deployments }
        else { "$d eq '*'" }
    }
    $q = [ordered]@{
        'api-version'       = $MetricsApiVersion
        metricnamespace     = $MetricsNamespace
        metricnames         = $Call.Metrics
        timespan            = "$(Format-Utc $From)/$(Format-Utc $To)"
        interval            = $Interval
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

function Get-DeploymentClause([string[]]$Names) {
    '(' + ((@($Names | Sort-Object -Unique) | ForEach-Object { "ModelDeploymentName eq '$($_.Replace("'", "''"))'" }) -join ' or ') + ')'
}

function New-BatchUri([string]$SubscriptionId, [string]$Region, [string]$Metrics, [string]$Aggregation, [datetime]$From, [datetime]$To,
    [string]$Grain, [string[]]$Dims, [string[]]$Deployments) {
    $q = [ordered]@{
        'api-version'   = $MetricsApiVersion
        metricnamespace = $MetricsNamespace
        metricnames     = $Metrics
        starttime       = Format-Utc $From
        endtime         = Format-Utc $To
        interval        = $Grain
        aggregation     = $Aggregation
    }
    if ($Dims) {
        $q['filter'] = (@($Dims | ForEach-Object { if ($Deployments -and $_ -eq 'ModelDeploymentName') { Get-DeploymentClause $Deployments } else { "$_ eq '*'" } }) -join ' and ')
        $q['top'] = $Top
    }
    $qs = ($q.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join '&'
    return "https://$($Region.ToLowerInvariant()).metrics.monitor.azure.com/subscriptions/$SubscriptionId/metrics:getBatch?$qs"
}

# Only unambiguous size signals split a request. Transport errors, 5xx and 529 have already been retried
# by Invoke-Arm; subdividing them would multiply load on a service that is already struggling.
$TooLargePattern = '(?i)exceed|too (many|large|much)|limit'
function Test-TooLarge([FoundryInsights.ArmException]$Ex) {
    if ($Ex.StatusCode -eq 413) { return $true }
    if ($Ex.StatusCode -eq 400 -and $Ex.Message -match $TooLargePattern) { return $true }
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

# Rejects responses with a different grain, a metric error, or a requested metric missing/duplicated
# (a partial answer must never be recorded as a successful, empty chunk).
function Assert-MetricResponse($Resp, [string[]]$Wanted, [string]$Expected = $Interval) {
    if ($Resp.PSObject.Properties['interval'] -and $Resp.interval -and $Resp.interval -ne $Expected) {
        throw "Service returned interval $($Resp.interval) instead of $Expected"
    }
    $count = @{}
    foreach ($m in @($Resp.value)) {
        if (-not $m) { continue }
        if ($m.PSObject.Properties['errorCode'] -and $m.errorCode -and $m.errorCode -ne 'Success') {
            throw "Metric $($m.name.value) error: $($m.errorCode) $($m.errorMessage)"
        }
        $n = [string]$m.name.value
        $count[$n] = 1 + [int]$count[$n]
    }
    foreach ($w in $Wanted) {
        if ([int]$count[$w] -ne 1) { throw "Metric $w returned $([int]$count[$w]) time(s) instead of once" }
    }
}

function New-SeriesResult { [pscustomobject]@{ Series = [System.Collections.Generic.List[object]]::new(); Truncated = $false; Notes = [System.Collections.Generic.List[string]]::new() } }

# Classic ARM metrics API for one resource. Splits the window when the service rejects it as too large and
# re-queries deployment by deployment when 'top' is reached. Returns @{ Series; Truncated; Notes }.
function Get-MetricSeries([object]$Res, [string]$CallName, [datetime]$From, [datetime]$To, [string[]]$Deployments) {
    $call = $Calls[$CallName]
    $wanted = $call.Metrics.Split(',')
    $result = New-SeriesResult
    try {
        $resp = Invoke-Arm -Tenant $Res.TenantId -Uri (New-MetricsUri $Res.Id $call $From $To -Deployments $Deployments)
    }
    catch [FoundryInsights.ArmException] {
        $ex = $_.Exception
        $minWindow = 60 * (Get-IntervalMinutes $Interval)
        if ((Test-TooLarge $ex) -and ($To - $From).TotalMinutes -gt $minWindow) {
            $step = Get-IntervalMinutes $Interval
            $mid = $From.AddMinutes($step * [math]::Floor(($To - $From).TotalMinutes / 2 / $step))
            Write-Log "  $($Res.Name) $CallName window $(Format-Utc $From) too large ($($ex.StatusCode)); splitting" 'WARN'
            $a = Get-MetricSeries $Res $CallName $From $mid $Deployments
            $b = Get-MetricSeries $Res $CallName $mid $To $Deployments
            $a.Series.AddRange($b.Series); $a.Notes.AddRange($b.Notes); $a.Truncated = $a.Truncated -or $b.Truncated
            return $a
        }
        throw
    }
    Assert-MetricResponse $resp $wanted

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
    if (-not $deps.Count) { $result.Truncated = $true; $result.Notes.Add('metadata returned no deployments'); return $result }
    foreach ($dep in $deps) {
        $r = Invoke-Arm -Tenant $Res.TenantId -Uri (New-MetricsUri $Res.Id $call $From $To -Deployment $dep)
        Assert-MetricResponse $r $wanted
        foreach ($m in $r.value) {
            if (@($m.timeseries).Count -ge $Top) { $result.Truncated = $true; $result.Notes.Add("deployment $dep still returns top=$Top series") }
            foreach ($ts in $m.timeseries) { $result.Series.Add((ConvertTo-Series $m $ts)) }
        }
    }
    return $result
}

function Set-BatchUnavailable([string]$Key, [string]$Reason) {
    if ($BatchCap.TryAdd($Key, 'No') -or $BatchCap[$Key] -ne 'No') {
        $BatchCap[$Key] = 'No'
        $BatchCapReason[$Key] = $Reason
        Write-Log "  metrics:getBatch unavailable for $Key ($Reason); using the ARM metrics API" 'WARN'
    }
}

# metrics:getBatch for resources of one subscription + region. Returns a hashtable lower(resourceId) ->
# @{ Series; Truncated; Notes } for every resource answered completely; resources missing from the result
# (error, duplicate, absent, truncated) must be queried through Get-MetricSeries.
function Get-BatchSeries([object[]]$Resources, [string]$CallName, [datetime]$From, [datetime]$To, [string[]]$Deployments) {
    $call = $Calls[$CallName]
    $first = $Resources[0]
    $capKey = "$($first.SubscriptionId)|$($first.Location)".ToLowerInvariant()
    $out = @{}
    $uri = New-BatchUri $first.SubscriptionId $first.Location $call.Metrics $call.Aggregation $From $To $Interval $call.Dims $Deployments
    try {
        $resp = Invoke-Arm -Tenant $first.TenantId -Method POST -Uri $uri -Audience $MetricsAudience -Body @{ resourceids = @($Resources | ForEach-Object Id) }
    }
    catch {
        $code = Get-StatusCode $_
        if ($code -in 401, 403, 404 -or ($code -eq 400 -and $_.Exception.Message -notmatch $TooLargePattern)) {
            Set-BatchUnavailable $capKey "HTTP $code"
            return $out
        }
        if ($_.Exception -is [FoundryInsights.ArmException] -and (Test-TooLarge $_.Exception) -and $Resources.Count -gt 1) {
            $half = [int][math]::Ceiling($Resources.Count / 2)
            Write-Log "  getBatch $CallName ($($Resources.Count) resources) too large ($code); splitting batch" 'WARN'
            $a = Get-BatchSeries $Resources[0..($half - 1)] $CallName $From $To $Deployments
            $b = Get-BatchSeries $Resources[$half..($Resources.Count - 1)] $CallName $From $To $Deployments
            foreach ($k in $b.Keys) { $a[$k] = $b[$k] }
            return $a
        }
        Write-Log "  getBatch $CallName failed ($($_.Exception.Message)); falling back to the ARM metrics API" 'WARN'
        return $out
    }
    $BatchCap[$capKey] = 'Yes'
    $wanted = @{}
    foreach ($r in $Resources) { $wanted[$r.Id.ToLowerInvariant()] = $true }
    $seen = @{}
    foreach ($item in @($resp.values)) {
        if (-not $item -or -not $item.PSObject.Properties['resourceid']) { continue }
        $rid = ([string]$item.resourceid).ToLowerInvariant()
        if (-not $wanted.ContainsKey($rid)) { continue }
        if ($seen.ContainsKey($rid)) { $out.Remove($rid); $seen[$rid] = 'dup'; continue }
        $seen[$rid] = 'ok'
        try { Assert-MetricResponse $item $call.Metrics.Split(',') } catch { Write-Log "  getBatch $CallName $rid : $($_.Exception.Message)" 'WARN'; continue }
        $truncated = $false
        foreach ($m in @($item.value)) { if (@($m.timeseries).Count -ge $Top) { $truncated = $true } }
        if ($truncated) { continue }   # per-deployment fallback through the ARM API
        $res = New-SeriesResult
        foreach ($m in @($item.value)) { foreach ($ts in @($m.timeseries)) { if ($ts) { $res.Series.Add((ConvertTo-Series $m $ts)) } } }
        $out[$rid] = $res
    }
    foreach ($k in @($seen.Keys)) { if ($seen[$k] -eq 'dup') { $out.Remove($k) } }
    return $out
}

# Collapses the StatusCode dimension: one AzureOpenAIRequests series (all codes) plus ThrottledMetric (429),
# ClientErrorMetric (other 4xx) and ServerErrorMetric (5xx) series per remaining dimension set, so later
# dedupe never drops codes.
function Merge-StatusCodeSeries($R) {
    $groups = [ordered]@{}
    foreach ($s in $R.Series) {
        $code = [string]$s.Dims['statuscode']
        $dims = @{}
        foreach ($k in $s.Dims.Keys) { if ($k -ne 'statuscode') { $dims[$k] = $s.Dims[$k] } }
        $base = (@($dims.Keys | Sort-Object) | ForEach-Object { "$_=$($dims[$_])" }) -join '|'
        $names = @($s.Metric)
        if ($code -eq '429') { $names += $ThrottledMetric }
        elseif ($code -match '^4\d\d$') { $names += $ClientErrorMetric }
        elseif ($code -match '^5\d\d$') { $names += $ServerErrorMetric }
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

function Get-ResourceHash([string]$Id) {
    [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::UTF8.GetBytes($Id.ToLowerInvariant()))).Replace('-', '').Substring(0, 16)
}

# Writes the atomic part file of one resource/chunk (rows sorted, so output does not depend on the API
# or on response order) and returns its manifest lines and per-call status.
function Write-ChunkPart($Res, $Chunk, [System.Collections.Specialized.OrderedDictionary]$Results) {
    $cs = Format-Utc $Chunk.Start; $ce = Format-Utc $Chunk.End
    if ($Results['Requests'] -and $Results['Requests'] -isnot [string]) { $null = Merge-StatusCodeSeries $Results['Requests'] }

    # Historical ModelName from AzureOpenAIRequests (tokens metrics lack that dimension):
    # per minute (exact) and per chunk (fallback), keyed by deployment/version[/region/minute].
    $byMinute = @{}; $byChunk = @{}
    $activeMin = [System.Collections.Generic.HashSet[string]]::new()
    if ($Results['Requests'] -and $Results['Requests'] -isnot [string] -and -not $Results['Requests'].PSObject.Properties['NotSupported']) {
        foreach ($s in $Results['Requests'].Series) {
            if ($s.Metric -ne 'AzureOpenAIRequests') { continue }
            $mn = $s.Dims['modelname']
            $d = $s.Dims['modeldeploymentname']; $vv = $s.Dims['modelversion']
            $rg = if ($s.Dims['region']) { ($s.Dims['region'] -replace '\s', '').ToLowerInvariant() } else { $Res.Location }
            foreach ($p in $s.Data) {
                if (-not $p.PSObject.Properties['total'] -or $null -eq $p.total -or [double]$p.total -eq 0) { continue }
                $pts = ConvertTo-UtcDate $p.timeStamp
                if ($pts -lt $Chunk.Start -or $pts -ge $Chunk.End) { continue }
                [void]$activeMin.Add("$d|$rg|$(Format-Utc $pts)")
                if (-not $mn) { continue }
                $mk = "$d|$vv|$rg|$(Format-Utc $pts)"
                if (-not $byMinute.ContainsKey($mk)) { $byMinute[$mk] = [System.Collections.Generic.HashSet[string]]::new() }
                [void]$byMinute[$mk].Add($mn)
                if (-not $byChunk.ContainsKey("$d|$vv")) { $byChunk["$d|$vv"] = [System.Collections.Generic.HashSet[string]]::new() }
                [void]$byChunk["$d|$vv"].Add($mn)
            }
        }
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    $manifest = [System.Collections.Generic.List[string]]::new()
    $status = @{}
    $rowsBy = @{}
    foreach ($callName in $Results.Keys) {
        $r = $Results[$callName]
        if ($r -is [string]) {
            $manifest.Add((ConvertTo-CsvLine @((Format-Utc ([datetime]::UtcNow)), $Res.Id, $cs, $ce, $callName, 'Failed', 0, 0, $false, $r)))
            $status[$callName] = 'Failed'; $rowsBy[$callName] = 0
            continue
        }
        if ($r.PSObject.Properties['NotSupported']) {
            $manifest.Add((ConvertTo-CsvLine @((Format-Utc ([datetime]::UtcNow)), $Res.Id, $cs, $ce, $callName, 'NotSupported', 0, 0, $false, $r.NotSupported)))
            $status[$callName] = 'NotSupported'; $rowsBy[$callName] = 0
            continue
        }
        $rows = 0
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        $call = $Calls[$callName]
        $aggs = if ($call.Contains('Aggs')) { $call.Aggs } else { @('total') }
        $activeOnly = $call.Contains('ActiveOnly') -and $call.ActiveOnly
        foreach ($s in $r.Series) {
            $dep = [string]$s.Dims['modeldeploymentname']; $ver = [string]$s.Dims['modelversion']
            $region = if ($s.Dims['region']) { ($s.Dims['region'] -replace '\s', '').ToLowerInvariant() } else { $Res.Location }
            $hasModelDim = $s.Dims.ContainsKey('modelname') -and $s.Dims['modelname']
            $cur = $Res.Deployments[$dep]
            $depSku = if ($cur) { $cur.Sku } else { '' }
            $depCap = if ($cur) { $cur.Capacity } else { '' }
            foreach ($p in $s.Data) {
                $ts = ConvertTo-UtcDate $p.timeStamp
                if ($ts -lt $Chunk.Start -or $ts -ge $Chunk.End) { continue }   # half-open [start, end)
                $tsS = Format-Utc $ts
                # Active-only calls keep zeros (0 % availability is a real outage) but only in minutes with requests.
                if ($activeOnly -and -not $activeMin.Contains("$dep|$region|$tsS")) { continue }
                $model = $null
                foreach ($agg in $aggs) {
                    if (-not $p.PSObject.Properties[$agg]) { continue }
                    $v = $p.$agg
                    if ($null -eq $v -or (-not $activeOnly -and [double]$v -eq 0)) { continue }
                    if ($null -eq $model) {
                        if ($hasModelDim) { $model = $s.Dims['modelname']; $src = 'Metric' }
                        else { $model, $src = Resolve-Model $Res $dep $ver $region $tsS $byMinute $byChunk }
                    }
                    $aggName = $agg.Substring(0, 1).ToUpperInvariant() + $agg.Substring(1)
                    if (-not $seen.Add("$tsS|$dep|$ver|$region|$($s.Metric)|$aggName|$model")) { continue }
                    $lines.Add((ConvertTo-CsvLine @(
                                $tsS, $Res.SubscriptionId, $Res.SubscriptionName, $Res.ResourceGroup, $Res.Name, $region,
                                $dep, $model, $ver, $s.Metric, [double]$v,
                                $Res.TenantId, $Res.Kind, $aggName, $depSku, $depCap, $src, $Res.Id)))
                    $rows++
                }
            }
        }
        $st = if ($r.Truncated) { 'Truncated' } else { 'Success' }
        $status[$callName] = $st; $rowsBy[$callName] = $rows
        $manifest.Add((ConvertTo-CsvLine @((Format-Utc ([datetime]::UtcNow)), $Res.Id, $cs, $ce, $callName, $st,
                    $r.Series.Count, $rows, $r.Truncated, ($r.Notes -join '; '))))
    }

    $partFile = Join-Path $PartsDir ('{0}_{1:yyyyMMddHHmm}.csv' -f (Get-ResourceHash $Res.Id), $Chunk.Start)
    $tmp = "$partFile.$([guid]::NewGuid().ToString('N')).tmp"
    $sorted = $lines.ToArray()
    [array]::Sort($sorted, [System.StringComparer]::Ordinal)
    [System.IO.File]::WriteAllLines($tmp, $sorted, $Utf8)
    Move-Item -LiteralPath $tmp -Destination $partFile -Force
    return [pscustomobject]@{ Manifest = $manifest; Status = $status; RowsBy = $rowsBy; Rows = $lines.Count; Part = (Split-Path $partFile -Leaf) }
}
#endregion

#region work units (executed by parallel workers) -------------------------------------------------
# Deployment names with at least one request inside the chunk (raw Requests result, before merging).
function Get-ActiveDeployments($R, $Chunk) {
    $set = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($s in $R.Series) {
        if ($s.Metric -ne 'AzureOpenAIRequests') { continue }
        foreach ($p in $s.Data) {
            if (-not $p.PSObject.Properties['total'] -or $null -eq $p.total -or [double]$p.total -eq 0) { continue }
            $ts = ConvertTo-UtcDate $p.timeStamp
            if ($ts -ge $Chunk.Start -and $ts -lt $Chunk.End) { [void]$set.Add([string]$s.Dims['modeldeploymentname']); break }
        }
    }
    return [string[]]@($set)
}
# Metrics of a call that the account does not define (from discovery). Unknown definitions = supported.
function Get-UnsupportedMetrics($Res, [string]$CallName) {
    # No PSObject.Properties here: resources are shared by parallel workers and the PSObject adapter is not
    # thread-safe (concurrent lookups intermittently fail). Every resource is created with MetricNames.
    if (-not @($Res.MetricNames).Count) { return @() }
    $defined = [System.Collections.Generic.HashSet[string]]::new([string[]]@($Res.MetricNames), [StringComparer]::OrdinalIgnoreCase)
    return @($Calls[$CallName].Metrics.Split(',') | Where-Object { -not $defined.Contains($_) })
}
function Invoke-DiscoverUnit($U) {
    $res = $U.Resource
    $out = [pscustomobject]@{ Type = 'Discover'; Id = $res.Id; Status = 'Pending'; Message = ''; Deployments = @{}; MetricNames = @() }
    try {
        $defs = Invoke-Arm -Tenant $res.TenantId -Uri "$($res.Id)/providers/Microsoft.Insights/metricDefinitions?api-version=$MetricsApiVersion&metricnamespace=$MetricsNamespace"
        $names = @($defs.value | ForEach-Object { $_.name.value })
        $out.MetricNames = $names
        $missing = @($RequiredMetrics | Where-Object { $_ -notin $names })
        if ($missing) { $out.Status = 'NotOpenAI'; $out.Message = "Missing metrics: $($missing -join ',')"; return $out }
    }
    catch {
        $code = Get-StatusCode $_
        $out.Status = if ($code -in 401, 403) { 'Forbidden' } else { 'Error' }
        $out.Message = $_.Exception.Message
        return $out
    }
    try {
        foreach ($d in (Get-ArmPaged $res.TenantId "$($res.Id)/deployments?api-version=2024-10-01")) {
            $out.Deployments[$d.name] = [pscustomobject]@{
                Model = Get-Prop $d 'properties.model.name'; Version = [string](Get-Prop $d 'properties.model.version')
                Format = Get-Prop $d 'properties.model.format'
                Sku = Get-Prop $d 'sku.name'; Capacity = Get-Prop $d 'sku.capacity'
                TokensPerMinute = ConvertTo-PerMinute (Get-Prop $d 'properties.rateLimits') 'token'
                RequestsPerMinute = ConvertTo-PerMinute (Get-Prop $d 'properties.rateLimits') 'request'
            }
        }
    }
    catch { $out.Message = "Deployments not readable: $($_.Exception.Message)" }
    $out.Status = 'OpenAI'
    return $out
}

function Invoke-QuotaUnit($U) {
    $base = "/subscriptions/$($U.SubscriptionId)/providers/Microsoft.CognitiveServices/locations/$($U.Location)"
    $out = [pscustomobject]@{ Type = 'Quota'; Key = $U.Key; Usages = @(); Models = @(); UsageError = ''; ModelError = '' }
    try { $out.Usages = (Get-ArmPaged $U.TenantId "$base/usages?api-version=2024-10-01").ToArray() } catch { $out.UsageError = $_.Exception.Message }
    try { $out.Models = (Get-ArmPaged $U.TenantId "$base/models?api-version=2024-10-01").ToArray() } catch { $out.ModelError = $_.Exception.Message }
    return $out
}

# True when every point of every requested metric is null or zero.
function Test-AllZero($Metrics, [string[]]$Wanted, [string]$Agg) {
    $byName = @{}
    foreach ($m in @($Metrics)) {
        if (-not $m) { continue }
        if ($byName.ContainsKey($m.name.value)) { return $false }   # ambiguous answer: never assume idle
        $byName[$m.name.value] = $m
    }
    foreach ($w in $Wanted) {
        $m = $byName[$w]
        if (-not $m) { return $false }   # unknown signal: never assume idle
        if ($m.PSObject.Properties['errorCode'] -and $m.errorCode -and $m.errorCode -ne 'Success') { return $false }
        foreach ($ts in @($m.timeseries)) {
            if (-not $ts) { continue }
            foreach ($p in @($ts.data)) {
                if ($p -and $p.PSObject.Properties[$Agg] -and $null -ne $p.$Agg -and [double]$p.$Agg -ne 0) { return $false }
            }
        }
    }
    return $true
}

# Resource-level idle check over the whole window at P1D, without dimensions (so no 'top' truncation).
# A resource is idle only if every signal that would be extracted is present and zero. The window is
# widened by one day on each side so day-aligned buckets can only over-count, never miss activity.
function Invoke-PreflightUnit($U) {
    $now = [datetime]::UtcNow
    $now = $now.AddTicks(-($now.Ticks % [TimeSpan]::TicksPerMinute))
    # The one-day padding is limited to retention (now-92d) and to now, but never narrows the requested window.
    $from = $U.From.AddDays(-1); if ($from -lt $now.AddDays(-92)) { $from = $now.AddDays(-92) }
    if ($from -gt $U.From) { $from = $U.From }
    $to = $U.To.AddDays(1); if ($to -gt $now) { $to = $now }
    if ($to -lt $U.To) { $to = $U.To }
    $checks = [System.Collections.Generic.List[object]]::new()
    $checks.Add(@{ Metrics = @('ProcessedPromptTokens', 'GeneratedTokens', 'AzureOpenAIRequests'); Agg = 'total' })
    if ($Calls.Contains('PtuUtilization')) { $checks.Add(@{ Metrics = @('AzureOpenAIProvisionedManagedUtilizationV2'); Agg = 'maximum' }) }
    $idle = @{}
    foreach ($r in $U.Resources) { $idle[$r.Id.ToLowerInvariant()] = $true }
    $first = $U.Resources[0]
    $capKey = "$($first.SubscriptionId)|$($first.Location)".ToLowerInvariant()
    foreach ($c in $checks) {
        $aggName = $c.Agg.Substring(0, 1).ToUpperInvariant() + $c.Agg.Substring(1)
        $answered = @{}
        if ($MetricsApiMode -eq 'Batch' -and $BatchCap[$capKey] -ne 'No') {
            $uri = New-BatchUri $first.SubscriptionId $first.Location ($c.Metrics -join ',') $aggName $from $to 'P1D' $null
            try {
                $resp = Invoke-Arm -Tenant $first.TenantId -Method POST -Uri $uri -Audience $MetricsAudience -Body @{ resourceids = @($U.Resources | ForEach-Object Id) }
                $BatchCap[$capKey] = 'Yes'
                foreach ($item in @($resp.values)) {
                    if (-not $item -or -not $item.PSObject.Properties['resourceid']) { continue }
                    $rid = ([string]$item.resourceid).ToLowerInvariant()
                    if (-not $idle.ContainsKey($rid)) { continue }
                    $answered[$rid] = if ($answered.ContainsKey($rid)) { $false } else { Test-AllZero $item.value $c.Metrics $c.Agg }
                }
            }
            catch {
                $code = Get-StatusCode $_
                if ($code -in 401, 403, 404) { Set-BatchUnavailable $capKey "HTTP $code" }
                else { Write-Log "  pre-flight getBatch failed ($($_.Exception.Message))" 'WARN' }
            }
        }
        foreach ($r in $U.Resources) {
            $rid = $r.Id.ToLowerInvariant()
            if (-not $idle[$rid]) { continue }
            if (-not $answered.ContainsKey($rid)) {
                $q = [ordered]@{
                    'api-version' = $MetricsApiVersion; metricnamespace = $MetricsNamespace; metricnames = ($c.Metrics -join ',')
                    timespan = "$(Format-Utc $from)/$(Format-Utc $to)"; interval = 'P1D'; aggregation = $aggName
                }
                $qs = ($q.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join '&'
                try {
                    $resp = Invoke-Arm -Tenant $r.TenantId -Uri "$Arm$($r.Id)/providers/Microsoft.Insights/metrics?$qs"
                    $answered[$rid] = Test-AllZero $resp.value $c.Metrics $c.Agg
                }
                catch { $answered[$rid] = $false }
            }
            if (-not $answered[$rid]) { $idle[$rid] = $false }
        }
    }
    return [pscustomobject]@{ Type = 'Preflight'; Idle = @($idle.Keys | Where-Object { $idle[$_] }) }
}

function Invoke-ExtractUnit($U) {
    $results = @{}
    foreach ($r in $U.Resources) { $results[$r.Id.ToLowerInvariant()] = [ordered]@{} }
    foreach ($callName in $Calls.Keys) {
        # Optional calls are only sent to accounts that define every metric of the call: an undefined metric
        # makes getBatch fail for the whole batch (and would disable batch for the subscription/region).
        $supported = [System.Collections.Generic.List[object]]::new()
        foreach ($r in $U.Resources) {
            $miss = @(Get-UnsupportedMetrics $r $callName)
            if ($miss.Count) { $results[$r.Id.ToLowerInvariant()][$callName] = [pscustomobject]@{ NotSupported = "Metrics not defined: $($miss -join ',')" } }
            else { $supported.Add($r) }
        }
        if (-not $supported.Count) { continue }
        $call = $Calls[$callName]
        $deps = $null
        if ($call.Contains('ActiveOnly') -and $call.ActiveOnly) {
            # Only deployments with requests in this chunk are queried: latency metrics report 0 ms / 100 %
            # for idle minutes, which would be millions of useless points on large estates.
            $active = [System.Collections.Generic.List[object]]::new()
            $names = [System.Collections.Generic.HashSet[string]]::new()
            $wildcard = $false
            foreach ($r in $supported) {
                $rid = $r.Id.ToLowerInvariant()
                $req = $results[$rid]['Requests']
                if ($null -eq $req -or $req -is [string] -or $req.PSObject.Properties['NotSupported']) {
                    $results[$rid][$callName] = 'Skipped: Requests call failed'; continue
                }
                $ad = @(Get-ActiveDeployments $req $U.Chunk)
                if (-not $ad.Count) { $e = New-SeriesResult; $e.Notes.Add('No requests in chunk'); $results[$rid][$callName] = $e; continue }
                foreach ($n in $ad) { if ($n) { [void]$names.Add($n) } else { $wildcard = $true } }
                $active.Add($r)
            }
            $supported = $active
            if (-not $wildcard -and $names.Count -and $names.Count -le 20) { $deps = [string[]]@($names) }
        }
        if (-not $supported.Count) { continue }
        $first = $supported[0]
        $capKey = "$($first.SubscriptionId)|$($first.Location)".ToLowerInvariant()
        $map = @{}
        if ($MetricsApiMode -eq 'Batch' -and $BatchCap[$capKey] -ne 'No') {
            $map = Get-BatchSeries $supported.ToArray() $callName $U.Chunk.Start $U.Chunk.End $deps
        }
        foreach ($r in $supported) {
            $rid = $r.Id.ToLowerInvariant()
            $v = $map[$rid]
            if ($null -eq $v) {
                try { $v = Get-MetricSeries $r $callName $U.Chunk.Start $U.Chunk.End $deps }
                catch { $v = [string]$_.Exception.Message; Write-Log "  $($r.Name) $(Format-Utc $U.Chunk.Start) $callName FAILED: $v" 'WARN' }
            }
            $results[$rid][$callName] = $v
        }
    }
    $manifest = [System.Collections.Generic.List[string]]::new()
    $status = @{}
    $rowsBy = @{}
    $parts = [System.Collections.Generic.List[string]]::new()
    $rows = 0
    foreach ($r in $U.Resources) {
        $w = Write-ChunkPart $r $U.Chunk $results[$r.Id.ToLowerInvariant()]
        $manifest.AddRange($w.Manifest)
        foreach ($k in $w.Status.Keys) {
            $key = "$($r.Id)|$(Format-Utc $U.Chunk.Start)|$k"
            $status[$key] = $w.Status[$k]; $rowsBy[$key] = $w.RowsBy[$k]
        }
        $parts.Add($w.Part)
        $rows += $w.Rows
    }
    return [pscustomobject]@{ Type = 'Extract'; Manifest = $manifest; Status = $status; RowsBy = $rowsBy; Parts = $parts; Rows = $rows }
}

function Invoke-WorkUnit($U) {
    switch ($U.Type) {
        'Discover' { return Invoke-DiscoverUnit $U }
        'Quota' { return Invoke-QuotaUnit $U }
        'Preflight' { return Invoke-PreflightUnit $U }
        'Extract' { return Invoke-ExtractUnit $U }
        default { throw "Unknown work unit type $($U.Type)" }
    }
}
#endregion
