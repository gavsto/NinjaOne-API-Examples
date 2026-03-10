# --------------------------------------------------
# Author: Gavin Stone (NinjaOne)
# Attribution: Luke Whitelock (NinjaOne) for his work on the Authentication Functions
# Date: 2026-03-10
# Description: Update the assigned policy on a specific device by its device ID. Only the policyId field should be changed — no other device properties should be modified.

# Version: 1.0
# --------------------------------------------------

# User editable variables:
$NinjaOneInstance = 'eu.ninjarmm.com' # Please replace with the region instance you login to (app.ninjarmm.com, us2.ninjarmm.com, eu.ninjarmm.com, ca.ninjarmm.com, oc.ninjarmm.com)
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''

# The NinjaOne device ID to update
$DeviceId = 1

# The policy ID to assign to the device (set to $null to revert to the organisation-level policy)
$NewPolicyId = 42

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

# Validate inputs
if (-not $DeviceId) {
    Write-Host "Error: Please set the `$DeviceId variable to the target device's ID."
    exit 1
}

if (-not $NewPolicyId -and $NewPolicyId -ne 0) {
    Write-Host "Error: Please set the `$NewPolicyId variable to the policy ID to assign (or `$null to revert to the organisation policy)."
    exit 1
}

# Fetch the device first to confirm it exists and show current state
Write-Host "Fetching device $DeviceId..."
try {
    $Device = Invoke-NinjaOneRequest -Method GET -Path "device/$DeviceId"
} catch {
    Write-Host "Error: Could not find device with ID $DeviceId. Please check the ID and try again."
    exit 1
}

Write-Host "Device:         $($Device.systemName) (ID: $($Device.id))"
Write-Host "Current Policy: $(if ($Device.policyId) { $Device.policyId } else { '(organisation default)' })"
Write-Host "New Policy:     $(if ($NewPolicyId) { $NewPolicyId } else { '(revert to organisation default)' })"

if ($Device.policyId -eq $NewPolicyId) {
    Write-Host "`nDevice is already assigned to policy $NewPolicyId. No changes made."
    exit 0
}

# Build the update body — only include policyId to avoid changing anything else
$UpdateBody = @{
    policyId = $NewPolicyId
}

Write-Host "`nUpdating policy..."
try {
    Invoke-NinjaOneRequest -Method PATCH -Path "device/$DeviceId" -InputObject $UpdateBody
    Write-Host "Successfully updated device '$($Device.systemName)' (ID: $DeviceId) to policy ID $NewPolicyId."
} catch {
    Write-Host "Error: Failed to update device policy: $_"
    exit 1
}
