<#
.SYNOPSIS
    Exports Omnissa Horizon 8 inventory to CSV, mirroring the Citrix
    Get-BrokerCatalog / Get-ProvScheme / Get-AcctIdentityPool exports.

.DESCRIPTION
    Uses the Horizon Connection Server REST API (no PowerCLI module needed).

      Citrix command          Horizon equivalent             Output file
      ----------------------  -----------------------------  -------------------------
      Get-BrokerCatalog       Desktop pools (+ RDS pools)    HorizonDesktopPools.csv
      Get-ProvScheme          Pool/farm provisioning,        HorizonProvisioning.csv
                              golden image VM + snapshot
      Get-AcctIdentityPool    Pool/farm naming patterns      HorizonNamingPatterns.csv

    Column notes for HorizonDesktopPools.csv (these differ from Citrix):
      AvailableCount   Machines whose Horizon state is AVAILABLE (idle and ready
                       for a user). This is NOT the Citrix catalog AvailableCount,
                       which counts machines not yet in a delivery group.
      UnassignedCount  Machines with no assigned user. Only filled in for
                       DEDICATED pools; blank for floating and RDS pools, where
                       machines are never assigned to a user.

    Works on Windows PowerShell 5.1 and PowerShell 7+.
    Requires a Horizon admin account with read access (the built-in
    "Administrators (Read only)" role is sufficient).

    On Windows PowerShell 5.1, TLS and certificate settings are process-wide.
    The script restores them when it finishes, even if it fails.

.PARAMETER TimeoutSec
    Per-request timeout. The base-VM lookup enumerates vCenter VMs and can be
    slow in large environments. Default 300.

.EXAMPLE
    .\Get-OmnissaHorizonInventory.ps1 -Server cs01.corp.local -Domain CORP -OutputDir C:\HorizonExport

.EXAMPLE
    $cred = Get-Credential CORP\svc_horizon
    .\Get-OmnissaHorizonInventory.ps1 -Server cs01.corp.local -Credential $cred -SkipCertificateCheck
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Server,
    [pscredential] $Credential,
    [string] $Domain,
    [string] $OutputDir,
    [ValidateRange(1, 3600)] [int] $TimeoutSec = 300,
    [switch] $SkipCertificateCheck   # lab / self-signed certificates only
)

$ErrorActionPreference = 'Stop'

# Highest version first; falls back automatically on older Connection Servers.
$PoolVersions    = 'v8','v7','v6','v5','v4','v3','v2','v1'
$FarmVersions    = 'v5','v4','v3','v2','v1'
$MachineVersions = 'v4','v3','v2','v1'
$RdsVersions     = 'v2','v1'
$VcVersions      = 'v2','v1'
$PageSize        = 1000
$BaseUri         = "https://$Server/rest"
$IsPS7           = $PSVersionTable.PSVersion.Major -ge 6
# Excel only detects UTF-8 reliably with a BOM. 'UTF8' writes one on 5.1; on 7+ it does not.
$CsvEncoding     = if ($IsPS7) { 'utf8BOM' } else { 'UTF8' }
# Status codes that mean "this API version doesn't exist here" -> try the next one.
$VersionMissingCodes = 400, 404, 405

if (-not $OutputDir) {
    $OutputDir = if ($IsPS7 -and -not $IsWindows) { Join-Path (Get-Location) 'HorizonExport' } else { 'C:\HorizonExport' }
}

$PoolCols = 'PoolName','DisplayName','PoolType','ProvisioningType','Source',
            'SessionSupport','UserAssignment','PersistUserChanges','Enabled',
            'FarmName','TotalMachines','AvailableCount','UsedCount',
            'UnassignedCount','PoolId'
$ProvCols = 'Name','ObjectType','ProvisioningType','Source','FullClone',
            'vCenter','GoldenImageVM','GoldenImagePath','Snapshot','SnapshotPath',
            'VMTemplate','NamingPattern','ParentVmId','BaseSnapshotId',
            'VmTemplateId','Id'
$NameCols = 'Name','ObjectType','NamingMethod','NamingPattern','MaxMachines',
            'MinMachines','SpareMachines','ProvisioningTime','Id'

# ---------------------------------------------------------------- credentials
if (-not $Credential) {
    $Credential = Get-Credential -Message "Horizon admin account (DOMAIN\user or user)"
}
$user = $Credential.UserName
if ($user -match '^(?<d>[^\\]+)\\(?<u>.+)$') {
    if ($Domain -and $Domain -ne $Matches.d) {
        Write-Warning "Using domain '$($Matches.d)' from the user name; -Domain '$Domain' is ignored."
    }
    $Domain = $Matches.d; $user = $Matches.u
}
elseif ($user -match '^(?<u>[^@]+)@(?<d>.+)$') {
    $user = $Matches.u
    if (-not $Domain) {
        $Domain = $Matches.d.Split('.')[0]
        Write-Warning "Domain guessed as '$Domain' from the UPN. If login fails, pass -Domain with the NetBIOS name."
    }
}
if (-not $Domain) { throw "Specify -Domain or enter the user as DOMAIN\user." }
Write-Verbose "Using domain '$Domain', user '$user'"

$script:Headers      = @{}
$script:RefreshToken = $null

# ---------------------------------------------------------------- HTTP helpers
function Invoke-Hzn {
    param([string]$Method = 'GET', [string]$Path, [hashtable]$Query, $Body, [switch]$NoAuthRetry)

    $uri = "$BaseUri/$Path"
    if ($Query -and $Query.Count) {
        $qs = ($Query.GetEnumerator() | ForEach-Object {
            '{0}={1}' -f $_.Key, [uri]::EscapeDataString([string]$_.Value) }) -join '&'
        $uri = "$uri`?$qs"
    }
    $p = @{ Method = $Method; Uri = $uri; Headers = $script:Headers; ContentType = 'application/json'; TimeoutSec = $TimeoutSec }
    if ($null -ne $Body) { $p.Body = ($Body | ConvertTo-Json -Depth 5) }
    if ($IsPS7 -and $SkipCertificateCheck) { $p.SkipCertificateCheck = $true }

    try {
        # Assign first, then return: on Windows PowerShell 5.1 a JSON array from
        # Invoke-RestMethod is emitted as ONE object unless it is re-enumerated.
        $res = Invoke-RestMethod @p
        return $res
    }
    catch {
        $code = Get-StatusCode $_
        if ($code -eq 401 -and -not $NoAuthRetry -and (Update-HznToken)) {
            return Invoke-Hzn -Method $Method -Path $Path -Query $Query -Body $Body -NoAuthRetry
        }
        throw
    }
}

function Get-StatusCode($err) {
    try { return [int]$err.Exception.Response.StatusCode } catch { return $null }
}

function Update-HznToken {
    if (-not $script:RefreshToken) { return $false }
    try {
        $r = Invoke-Hzn -Method POST -Path 'refresh' -Body @{ refresh_token = $script:RefreshToken } -NoAuthRetry
        $script:Headers['Authorization'] = "Bearer $($r.access_token)"
        return $true
    } catch { return $false }
}

function Get-HznPaged([string]$Path, [hashtable]$Query = @{}) {
    $all = New-Object System.Collections.Generic.List[object]
    $page = 1; $lastFirst = $null
    while ($true) {
        $q = $Query.Clone(); $q.page = $page; $q.size = $PageSize
        $batch = @(Invoke-Hzn -Path $Path -Query $q)
        if ($batch.Count -eq 0 -or ($lastFirst -and $batch[0].id -eq $lastFirst)) { break }
        $all.AddRange([object[]]$batch)
        if ($batch.Count -lt $PageSize) { break }
        $lastFirst = $batch[0].id; $page++
    }
    , $all.ToArray()
}

function Get-HznVersionedList([string]$Area, [string]$Resource, [string[]]$Versions, [switch]$NoPaging) {
    $lastErr = $null
    foreach ($v in $Versions) {
        $path = "$Area/$v/$Resource"
        try {
            $data = if ($NoPaging) { @(Invoke-Hzn -Path $path) } else { Get-HznPaged $path }
            return [pscustomobject]@{ Data = $data; Version = $v }
        }
        catch {
            # Only "version not supported" responses fall through to the next version.
            # Auth failures, 5xx, timeouts and network errors ($null code) are real errors.
            $code = Get-StatusCode $_
            if ($code -notin $VersionMissingCodes) { throw }
            Write-Verbose "$path returned $code; trying an older API version"
            $lastErr = $_
        }
    }
    throw "No supported API version for $Area/$Resource. Last error: $lastErr"
}

# ---------------------------------------------------------------- vCenter lookups (cached)
$script:VcNames   = @{}
$script:BaseVms   = @{}
$script:Templates = @{}
$script:Snaps     = @{}
$script:Warned    = @{}

function Write-WarnOnce($key, $msg) {
    if (-not $script:Warned.ContainsKey($key)) { $script:Warned[$key] = $true; Write-Warning $msg }
}

function Get-VmInfo($VcId, $VmId) {
    if (-not $VcId -or -not $VmId) { return $null }
    if (-not $script:BaseVms.ContainsKey($VcId)) {
        $map = @{}
        try { foreach ($v in @(Invoke-Hzn -Path 'external/v1/base-vms' -Query @{ vcenter_id = $VcId })) { $map[$v.id] = $v } }
        catch { Write-WarnOnce "vm$VcId" "Base VM lookup failed for vCenter $VcId : $($_.Exception.Message)" }
        $script:BaseVms[$VcId] = $map
    }
    $script:BaseVms[$VcId][$VmId]
}

function Get-TemplateInfo($VcId, $TplId) {
    if (-not $VcId -or -not $TplId) { return $null }
    if (-not $script:Templates.ContainsKey($VcId)) {
        $map = @{}
        try { foreach ($t in @(Invoke-Hzn -Path 'external/v1/vm-templates' -Query @{ vcenter_id = $VcId })) { $map[$t.id] = $t } }
        catch { Write-WarnOnce "tpl$VcId" "Template lookup failed for vCenter $VcId : $($_.Exception.Message)" }
        $script:Templates[$VcId] = $map
    }
    $script:Templates[$VcId][$TplId]
}

function Get-SnapInfo($VcId, $VmId, $SnapId) {
    if (-not $VcId -or -not $VmId -or -not $SnapId) { return $null }
    $key = "$VcId|$VmId"
    if (-not $script:Snaps.ContainsKey($key)) {
        $map = @{}
        try { foreach ($s in @(Invoke-Hzn -Path 'external/v1/base-snapshots' -Query @{ vcenter_id = $VcId; base_vm_id = $VmId })) { $map[$s.id] = $s } }
        catch { Write-WarnOnce "snap$key" "Snapshot lookup failed for $VmId : $($_.Exception.Message)" }
        $script:Snaps[$key] = $map
    }
    $script:Snaps[$key][$SnapId]
}

# ---------------------------------------------------------------- mapping helpers
function Get-ProvisioningLabel([string]$Source, [string]$PoolType) {
    $s = "$Source".ToUpper()
    if ($PoolType -eq 'RDS')               { return 'RDS Farm' }
    if ($s -like '*INSTANT_CLONE*')        { return 'Instant Clone' }
    if ($s -eq 'VIEW_COMPOSER')            { return 'Linked Clone (Composer)' }
    if ($s -eq 'VIRTUAL_CENTER')           { if ($PoolType -eq 'AUTOMATED') { return 'Full Clone' } else { return 'Manual (vCenter VMs)' } }
    if ($s -eq 'UNMANAGED')                { return 'Manual (Physical/Unmanaged)' }
    if ($s) { return $s } else { return $PoolType }
}

function Get-MachineUsers($m) {
    if ($m.user_ids) { return @($m.user_ids) }
    if ($m.assigned_user_ids) { return @($m.assigned_user_ids) }
    @()
}

function Coalesce { foreach ($a in $args) { if ($null -ne $a -and "$a" -ne '') { return $a } }; '' }

$UsedStates = 'CONNECTED','DISCONNECTED','UNASSIGNED_USER_CONNECTED',
              'UNASSIGNED_USER_DISCONNECTED','ALREADY_USED','IN_PROGRESS'

function New-ProvRow($ObjType, $Obj, $Source, $Prov, $Naming, $VcId) {
    $parent = $Prov.parent_vm_id; $snapId = $Prov.base_snapshot_id; $tplId = $Prov.vm_template_id
    $vm   = Get-VmInfo $VcId $parent
    $snap = Get-SnapInfo $VcId $parent $snapId
    $tpl  = Get-TemplateInfo $VcId $tplId
    [pscustomobject]@{
        Name             = $Obj.name
        ObjectType       = $ObjType
        ProvisioningType = Get-ProvisioningLabel $Source 'AUTOMATED'
        Source           = $Source
        FullClone        = ("$Source".ToUpper() -eq 'VIRTUAL_CENTER')
        vCenter          = if ($VcId -and $script:VcNames.ContainsKey($VcId)) { $script:VcNames[$VcId] } else { Coalesce $VcId }
        GoldenImageVM    = Coalesce $vm.name $parent
        GoldenImagePath  = Coalesce $vm.path
        Snapshot         = Coalesce $snap.name $snapId
        SnapshotPath     = Coalesce $snap.path
        VMTemplate       = Coalesce $tpl.name $tplId
        NamingPattern    = Coalesce $Naming.naming_pattern
        ParentVmId       = Coalesce $parent
        BaseSnapshotId   = Coalesce $snapId
        VmTemplateId     = Coalesce $tplId
        Id               = $Obj.id
    }
}

function New-NamingRow($ObjType, $Obj, $Method, $Pattern, $Specified) {
    if (-not $Method) { $Method = if ($Pattern) { 'PATTERN' } elseif ($Specified) { 'SPECIFIED' } else { '' } }
    [pscustomobject]@{
        Name             = $Obj.name
        ObjectType       = $ObjType
        NamingMethod     = $Method
        NamingPattern    = Coalesce $Pattern.naming_pattern
        MaxMachines      = Coalesce $Pattern.max_number_of_machines $Pattern.max_number_of_rds_servers
        MinMachines      = Coalesce $Pattern.min_number_of_machines
        SpareMachines    = Coalesce $Pattern.number_of_spare_machines $Specified.number_of_spare_machines
        ProvisioningTime = Coalesce $Pattern.provisioning_time
        Id               = $Obj.id
    }
}

function Export-HznCsv($Rows, [string[]]$Columns, [string]$Path) {
    if ($Rows.Count) {
        $Rows | Select-Object $Columns | Export-Csv -Path $Path -NoTypeInformation -Encoding $CsvEncoding
    } else {
        # Export-Csv writes no header for empty input; keep the columns stable for consumers.
        ('"' + ($Columns -join '","') + '"') | Set-Content -Path $Path -Encoding $CsvEncoding
    }
    Write-Host ("  Wrote {0,4} rows -> {1}" -f $Rows.Count, $Path)
}

# ---------------------------------------------------------------- TLS setup
# Windows PowerShell 5.1 has no per-request TLS options; these settings are
# process-wide, so they are saved here and restored in the finally block.
$script:SavedProtocol   = $null
$script:SavedCertPolicy = $null
$script:CertPolicyChanged = $false
if (-not $IsPS7) {
    $script:SavedProtocol = [Net.ServicePointManager]::SecurityProtocol
    [Net.ServicePointManager]::SecurityProtocol = $script:SavedProtocol -bor [Net.SecurityProtocolType]::Tls12
    if ($SkipCertificateCheck) {
        if (-not ('TrustAllCertsPolicy' -as [type])) {
            Add-Type @"
using System.Net; using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }
}
"@
        }
        $script:SavedCertPolicy = [Net.ServicePointManager]::CertificatePolicy
        [Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
        $script:CertPolicyChanged = $true
    }
}

try {
    # ------------------------------------------------------------ login
    Write-Host "Logging in to $Server as $Domain\$user..."
    $login = Invoke-Hzn -Method POST -Path 'login' -NoAuthRetry -Body @{
        username = $user
        password = $Credential.GetNetworkCredential().Password
        domain   = $Domain
    }
    $script:Headers['Authorization'] = "Bearer $($login.access_token)"
    $script:RefreshToken = $login.refresh_token

    # ------------------------------------------------------------ collect
    Write-Host 'Collecting desktop pools...'
    $r = Get-HznVersionedList 'inventory' 'desktop-pools' $PoolVersions
    $pools = @($r.Data); Write-Host "  $($pools.Count) pools (API $($r.Version))"

    Write-Host 'Collecting farms...'
    $farms = @()
    try { $r = Get-HznVersionedList 'inventory' 'farms' $FarmVersions; $farms = @($r.Data); Write-Host "  $($farms.Count) farms (API $($r.Version))" }
    catch { Write-Warning "Could not list farms: $($_.Exception.Message)" }
    $farmById = @{}; foreach ($f in $farms) { $farmById[$f.id] = $f }

    Write-Host 'Collecting machines (for pool counts)...'
    $r = Get-HznVersionedList 'inventory' 'machines' $MachineVersions
    $machines = @($r.Data); Write-Host "  $($machines.Count) machines (API $($r.Version))"
    $byPool = @{}
    foreach ($m in $machines) {
        $k = "$($m.desktop_pool_id)"
        if (-not $byPool.ContainsKey($k)) { $byPool[$k] = New-Object System.Collections.Generic.List[object] }
        $byPool[$k].Add($m)
    }

    $rdsByFarm = @{}
    if ($farms.Count) {
        try {
            $r = Get-HznVersionedList 'inventory' 'rds-servers' $RdsVersions
            foreach ($s in @($r.Data)) {
                $k = "$($s.farm_id)"
                if (-not $rdsByFarm.ContainsKey($k)) { $rdsByFarm[$k] = New-Object System.Collections.Generic.List[object] }
                $rdsByFarm[$k].Add($s)
            }
        } catch { Write-Warning "Could not list RDS servers: $($_.Exception.Message)" }
    }

    try {
        $r = Get-HznVersionedList 'config' 'virtual-centers' $VcVersions -NoPaging
        foreach ($vc in @($r.Data)) { $script:VcNames[$vc.id] = Coalesce $vc.server_name $vc.name $vc.id }
    } catch { Write-Warning "Could not list vCenters; IDs will be shown instead. $($_.Exception.Message)" }

    $poolRows = New-Object System.Collections.Generic.List[object]
    $provRows = New-Object System.Collections.Generic.List[object]
    $nameRows = New-Object System.Collections.Generic.List[object]

    # ------------------------------------------------------------ desktop pools
    foreach ($p in $pools) {
        $ptype = "$($p.type)"; $source = "$($p.source)"
        $farm  = if ($p.farm_id) { $farmById[$p.farm_id] } else { $null }
        $ms    = if ($byPool.ContainsKey("$($p.id)")) { @($byPool["$($p.id)"]) } else { @() }
        $assign = "$($p.user_assignment)"

        if ($ptype -eq 'RDS') {
            $servers = if ($rdsByFarm.ContainsKey("$($p.farm_id)")) { @($rdsByFarm["$($p.farm_id)"]) } else { @() }
            $total = $servers.Count; $avail = ''; $used = ''; $unassigned = ''
        } else {
            $total      = $ms.Count
            $avail      = @($ms | Where-Object { $_.state -eq 'AVAILABLE' }).Count
            $used       = @($ms | Where-Object { $_.state -in $UsedStates }).Count
            # Floating machines are never assigned to a user, so the count is only meaningful for DEDICATED pools.
            $unassigned = if ($assign -eq 'DEDICATED') { @($ms | Where-Object { (Get-MachineUsers $_).Count -eq 0 }).Count } else { '' }
        }

        $persist = switch ($assign) { 'DEDICATED' { 'Persistent (dedicated)' } 'FLOATING' { 'Non-persistent' } default { '' } }

        $poolRows.Add([pscustomobject]@{
            PoolName           = $p.name
            DisplayName        = $p.display_name
            PoolType           = $ptype
            ProvisioningType   = Get-ProvisioningLabel $source $ptype
            Source             = $source
            SessionSupport     = if ($ptype -eq 'RDS') { 'MultiSession' } else { 'SingleSession' }
            UserAssignment     = if ($assign) { $assign } elseif ($ptype -eq 'RDS') { 'SHARED (RDS)' } else { '' }
            PersistUserChanges = $persist
            Enabled            = $p.enabled
            FarmName           = if ($farm) { $farm.name } else { '' }
            TotalMachines      = $total
            AvailableCount     = $avail
            UsedCount          = $used
            UnassignedCount    = $unassigned
            PoolId             = $p.id
        })

        if ($ptype -eq 'AUTOMATED') {
            $provRows.Add((New-ProvRow 'DesktopPool' $p $source $p.provisioning_settings $p.pattern_naming_settings $p.vcenter_id))
            $spec = if ($p.specific_naming_settings) { $p.specific_naming_settings } else { $p.specified_naming_settings }
            $nameRows.Add((New-NamingRow 'DesktopPool' $p $p.naming_method $p.pattern_naming_settings $spec))
        }
    }

    # ------------------------------------------------------------ automated RDS farms
    foreach ($f in @($farms | Where-Object { $_.type -eq 'AUTOMATED' })) {
        $auto = $f.automated_farm_settings
        $vcId = Coalesce $auto.vcenter_id $f.vcenter_id
        $provRows.Add((New-ProvRow 'Farm' $f "$($f.source)" $auto.provisioning_settings $auto.pattern_naming_settings $vcId))
        $nameRows.Add((New-NamingRow 'Farm' $f 'PATTERN' $auto.pattern_naming_settings $null))
    }

    # ------------------------------------------------------------ export
    if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
    Write-Host 'Writing CSV files...'
    Export-HznCsv $poolRows $PoolCols (Join-Path $OutputDir 'HorizonDesktopPools.csv')
    Export-HznCsv $provRows $ProvCols (Join-Path $OutputDir 'HorizonProvisioning.csv')
    Export-HznCsv $nameRows $NameCols (Join-Path $OutputDir 'HorizonNamingPatterns.csv')
}
finally {
    if ($script:RefreshToken) {
        # Refresh first so logout still works if the access token expired during collection.
        Update-HznToken | Out-Null
        try { Invoke-Hzn -Method POST -Path 'logout' -NoAuthRetry -Body @{ refresh_token = $script:RefreshToken } | Out-Null } catch { }
    }
    if (-not $IsPS7) {
        if ($null -ne $script:SavedProtocol)   { [Net.ServicePointManager]::SecurityProtocol = $script:SavedProtocol }
        if ($script:CertPolicyChanged)         { [Net.ServicePointManager]::CertificatePolicy = $script:SavedCertPolicy }
    }
}
Write-Host 'Done.'
