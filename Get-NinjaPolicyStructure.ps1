# --------------------------------------------------
# Author: Gavin Stone (NinjaOne)
# Status: Work in Progress / Draft / Testing
# Attribution: Luke Whitelock (NinjaOne) for his work on the Authentication Functions
# Date: 20th June 2024
# Description: Tries to be clever and get the policy structure from NinjaOne. Lists out an overview of the policy structure,
#              including the number of devices and snowflake devices for each policy. Also lists out the snowflake devices for each policy.
#              A snowflake device is a device that has a policy override applied to it.
#              The output is grouped by node class and the policies are listed in a hierarchical structure.            
#              The output is written to a file in C:\Temp\NinjaOnePolicyOutput_<currentDateTime>.txt
#              The output is also printed to the console.
# Version: 1.0
# --------------------------------------------------

# User editable variables:
$NinjaOneInstance = '' # Please replace with the region instance you login to (app.ninjarmm.com, us2.ninjarmm.com, eu.ninjarmm.com, ca.ninjarmm.com, oc.ninjarmm.com)
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''

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
        else {
            Throw 'Please run Connect-NinjaOne first'
        }
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

# Get the current date and time for the file name
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFilePath = "C:\Temp\NinjaOnePolicyOutput_$currentDateTime.txt"

# Create a new file and ensure it's empty if it already exists (clean slate)
New-Item -Path $outputFilePath -ItemType File -Force | Out-Null

# Function to write output to both console and file
function Write-OutputAndFile {
    param (
        [string]$message
    )
    Write-Output $message
    Add-Content -Path $outputFilePath -Value $message
}

# Get all policies and devices in NinjaOne (like a boss)
$Policies = Invoke-NinjaOneRequest -Method GET -Path 'policies' -Paginate
$Devices = Invoke-NinjaOneRequest -Method GET -Path 'devices' -Paginate

# Create a hashtable to map policy IDs to policy names (because we love hashtables)
$PolicyMap = @{}
foreach ($Policy in $Policies) {
    $PolicyMap[$Policy.id] = $Policy.name
}

# Initialize hashtables to store various counts and mappings
$policyDepth = @{}
$policyDeviceCount = @{}
$policySnowFlakeCount = @{}
$policySnowFlakeDevices = @{}
$nodeClassPolicies = @{}

# Recursive function to calculate the depth of a policy
function Get-PolicyDepth {
    param (
        [int]$policyId
    )

    if ($policyDepth.ContainsKey($policyId)) {
        return $policyDepth[$policyId]
    }

    $policy = $Policies | Where-Object { $_.id -eq $policyId }
    if (-not $policy.parentPolicyId) {
        $policyDepth[$policyId] = 0
        return 0
    }

    $parentDepth = Get-PolicyDepth -policyId $policy.parentPolicyId
    $policyDepth[$policyId] = $parentDepth + 1
    return $policyDepth[$policyId]
}

# Calculate depth for each policy and add parent policy name if applicable
foreach ($Policy in $Policies) {
    $depth = Get-PolicyDepth -policyId $Policy.id
    $Policy | Add-Member -MemberType NoteProperty -Name Depth -Value $depth -Force

    $deviceCount = ($Devices | Where-Object { $_.rolePolicyId -eq $Policy.id -and [string]::IsNullOrEmpty($_.policyId) }).Count
    $snowFlakeDevices = $Devices | Where-Object { $_.policyId -eq $Policy.id }
    $snowFlakeCount = $snowFlakeDevices.Count
    
    $policyDeviceCount[$Policy.id] = $deviceCount
    $policySnowFlakeCount[$Policy.id] = $snowFlakeCount
    $policySnowFlakeDevices[$Policy.id] = $snowFlakeDevices
    
    $Policy | Add-Member -MemberType NoteProperty -Name DeviceCount -Value $deviceCount -Force
    $Policy | Add-Member -MemberType NoteProperty -Name SnowFlakeCount -Value $snowFlakeCount -Force

    if (-not $nodeClassPolicies.ContainsKey($Policy.nodeclass)) {
        $nodeClassPolicies[$Policy.nodeclass] = @()
    }
    $nodeClassPolicies[$Policy.nodeclass] = $nodeClassPolicies[$Policy.nodeclass] + $Policy
}

# Function to print child policies
function Show-ChildPolicies {
    param (
        [int]$policyId,
        [int]$depth
    )

    $childPolicies = $Policies | Where-Object { $_.parentPolicyId -eq $policyId }
    foreach ($childPolicy in $childPolicies) {
        $indentation = " " * ($depth * 4)
        $message = "$indentation$($childPolicy.name) | Devices: $($childPolicy.DeviceCount) | SnowFlakeDevices: $($childPolicy.SnowFlakeCount)"
        Write-OutputAndFile -message $message
        
        $halfTabIndentation = $indentation + "  "
        foreach ($device in $policySnowFlakeDevices[$childPolicy.id]) {
            $snowFlakeMessage = "$halfTabIndentation Device Override: $($device.systemName)"
            Write-OutputAndFile -message $snowFlakeMessage
        }
        
        Show-ChildPolicies -policyId $childPolicy.id -depth ($depth + 1)
    }
}

# Loop through each nodeclass
foreach ($nodeclass in $nodeClassPolicies.Keys) {
    Write-OutputAndFile -message "[NODE CLASS = '$nodeclass']"

    foreach ($Policy in $nodeClassPolicies[$nodeclass] | Where-Object { $_.Depth -eq 0 }) {
        $message = "$($Policy.name) | Devices: $($Policy.DeviceCount) | SnowFlakeDevices: $($Policy.SnowFlakeCount)"
        Write-OutputAndFile -message $message
        
        $halfTabIndentation = "  "
        foreach ($device in $policySnowFlakeDevices[$Policy.id]) {
            $snowFlakeMessage = "$halfTabIndentation Device Override: $($device.systemName)"
            Write-OutputAndFile -message $snowFlakeMessage
        }
        
        Show-ChildPolicies -policyId $Policy.id -depth 1
    }

    Write-OutputAndFile -message ""
}
