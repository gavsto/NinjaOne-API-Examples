# --------------------------------------------------
# Author: Gavin Stone (NinjaOne)
# Attribution: Luke Whitelock (NinjaOne) for his work on the Authentication Functions
# Date: 2026-03-10
# Description: Find all devices managed by NinjaOne that have Google Chrome installed. Return the device name, OS, and the Chrome version found.

# Version: 1.0
# --------------------------------------------------

# User editable variables:
$NinjaOneInstance = 'eu.ninjarmm.com' # Please replace with the region instance you login to (app.ninjarmm.com, us2.ninjarmm.com, eu.ninjarmm.com, ca.ninjarmm.com, oc.ninjarmm.com)
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''

# Regex pattern to match software names against (e.g. 'Google Chrome', 'Firefox')
$SoftwareNameFilter = 'Google Chrome'

# Functions for Authentication
function Get-NinjaOneToken {
    [CmdletBinding()]
    param()

    if ($Script:NinjaOneInstance -and $Script:NinjaOneClientID -and $Script:NinjaOneClientSecret ) {
        if ($Script:NinjaTokenExpiry -and (Get-Date) -lt $Script:NinjaTokenExpiry) {
            return $Script:NinjaToken
        }
        else {

            if ($Script:NinjaOneRefreshToken) {
                $Body = @{
                    'grant_type'    = 'refresh_token'
                    'client_id'     = $Script:NinjaOneClientID
                    'client_secret' = $Script:NinjaOneClientSecret
                    'refresh_token' = $Script:NinjaOneRefreshToken
                }
            }
            else {

                $body = @{
                    grant_type    = 'client_credentials'
                    client_id     = $Script:NinjaOneClientID
                    client_secret = $Script:NinjaOneClientSecret
                    scope         = 'monitoring management'
                }
            }

            $token = Invoke-RestMethod -Uri "https://$($Script:NinjaOneInstance -replace '/ws','')/ws/oauth/token" -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing

            $Script:NinjaTokenExpiry = (Get-Date).AddSeconds($Token.expires_in)
            $Script:NinjaToken = $token

            Write-Host 'Fetched New Token'
            return $token
        }
    }
    else {
        Throw 'Please run Connect-NinjaOne first'
    }

}

function Connect-NinjaOne {
    [CmdletBinding()]
    param (
        [Parameter(mandatory = $true)]
        $NinjaOneInstance,
        [Parameter(mandatory = $true)]
        $NinjaOneClientID,
        [Parameter(mandatory = $true)]
        $NinjaOneClientSecret,
        $NinjaOneRefreshToken
    )

    $Script:NinjaOneInstance = $NinjaOneInstance
    $Script:NinjaOneClientID = $NinjaOneClientID
    $Script:NinjaOneClientSecret = $NinjaOneClientSecret
    $Script:NinjaOneRefreshToken = $NinjaOneRefreshToken


    try {
        $Null = Get-NinjaOneToken -ea Stop
    }
    catch {
        Throw "Failed to Connect to NinjaOne: $_"
    }

}

function Invoke-NinjaOneRequest {
    param(
        $Method,
        $Body,
        $InputObject,
        $Path,
        $QueryParams,
        [Switch]$Paginate,
        [Switch]$AsArray
    )

    $Token = Get-NinjaOneToken

    if ($InputObject) {
        if ($AsArray) {
            $Body = $InputObject | ConvertTo-Json -depth 100
            if (($InputObject | Measure-Object).count -eq 1 ) {
                $Body = '[' + $Body + ']'
            }
        }
        else {
            $Body = $InputObject | ConvertTo-Json -depth 100
        }
    }

    try {
        if ($Method -in @('GET', 'DELETE')) {
            if ($Paginate) {

                $After = 0
                $PageSize = 1000
                $NinjaResult = do {
                    $Result = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)?pageSize=$PageSize&after=$After$(if ($QueryParams){"&$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json' -UseBasicParsing
                    $Result
                    $ResultCount = ($Result.id | Measure-Object -Maximum)
                    $After = $ResultCount.maximum

                } while ($ResultCount.count -eq $PageSize)
            }
            else {
                $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json; charset=utf-8' -UseBasicParsing
            }

        }
        elseif ($Method -in @('PATCH', 'PUT', 'POST')) {
            $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -Body $Body -ContentType 'application/json; charset=utf-8' -UseBasicParsing
        }
        else {
            Throw 'Unknown Method'
        }
    }
    catch {
        Throw "Error Occured: $_"
    }

    try {
        return $NinjaResult.content | ConvertFrom-Json -ea stop
    }
    catch {
        return $NinjaResult.content
    }

}

# Connect to NinjaOne API
try {
    Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
}
catch {
    Write-Output "Failed to connect to NinjaOne API: $_"
    exit 1
}

# ─── Script Logic ─────────────────────────────────────────────────────────────

# Query all software inventory in bulk using cursor pagination
Write-Host "Fetching software inventory..."
$AllSoftware = [System.Collections.Generic.List[PSCustomObject]]::new()
$CursorName = $null
$PageSize = 1000

do {
    $QueryParams = "pageSize=$PageSize"
    if ($CursorName) {
        $QueryParams += "&cursor=$CursorName"
    }

    $Response = Invoke-NinjaOneRequest -Method GET -Path 'queries/software' -QueryParams $QueryParams

    if ($Response.results) {
        foreach ($Item in $Response.results) {
            $AllSoftware.Add($Item)
        }
    }

    $CursorName = $Response.cursor.name
    $PageCount = if ($Response.results) { $Response.results.Count } else { 0 }
    Write-Host "  Retrieved $($AllSoftware.Count) records so far..."
} while ($CursorName -and $PageCount -eq $PageSize)

Write-Host "Retrieved $($AllSoftware.Count) total software record(s)"

# Filter to matching software
$MatchingSoftware = $AllSoftware | Where-Object { $_.name -match $SoftwareNameFilter }

if ($MatchingSoftware.Count -eq 0) {
    Write-Host "No installations found matching '$SoftwareNameFilter'."
    exit 0
}

# Get unique device IDs that have the software
$DeviceIds = $MatchingSoftware | Select-Object -ExpandProperty deviceId -Unique
Write-Host "Found $($MatchingSoftware.Count) installation(s) across $($DeviceIds.Count) device(s). Fetching device details..."

# Fetch all devices to resolve names (single paginated call)
$AllDevices = Invoke-NinjaOneRequest -Method GET -Path 'devices' -Paginate
$DeviceLookup = @{}
foreach ($Device in $AllDevices) {
    $DeviceLookup[$Device.id] = $Device
}

# Build results
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($App in $MatchingSoftware) {
    $Device = $DeviceLookup[$App.deviceId]
    $Results.Add([PSCustomObject]@{
        DeviceName   = if ($Device) { $Device.systemName } else { "Unknown (ID: $($App.deviceId))" }
        DeviceId     = $App.deviceId
        DeviceClass  = if ($Device) { $Device.nodeClass } else { 'N/A' }
        SoftwareName = $App.name
        Version      = $App.version
        Publisher    = $App.publisher
        InstallDate  = $App.installDate
    })
}

Write-Host "`nFound $($Results.Count) installation(s) across $($DeviceIds.Count) device(s)."
$Results | Format-Table -AutoSize

# Uncomment the next line to export to a CSV
$Results | Export-Csv -Path ('C:\Temp\NinjaOneSoftwareReport_' + (Get-Date -f "yyyyMMdd_HHmm") + '.csv') -NoTypeInformation
