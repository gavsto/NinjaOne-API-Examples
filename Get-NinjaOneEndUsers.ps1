# --------------------------------------------------
# Author: Gavin Stone (NinjaOne)
# Attribution: Luke Whitelock (NinjaOne) for his work on the Authentication Functions
# Date: 19th June 2024
# Description: Gets a list of all End Users in NinjaOne and exports them to a CSV file
# Version: 1.0
# --------------------------------------------------

# PRE REQUISITES --------------------------------------------------
# For API access you will need to generate a Client ID and Client Secret in NinjaOne
# Go to Administration > Apps > API and the Client App IDs tab. 
# Click the 'Add' button in the top right
# For application platform, select API Services (machine-to-machine)
# Name your token something you will recognize (e.g. Recurring Maintenance Mode Script API Token)
# Set Redirect URI to http://localhost
# Set the scopes to monitoring and management
# For allowed grant types, select client credentials only
# Click save in the top right, enter 2FA prompt if required.
# You will be presented with the client secret credential only once. Use the copy icon to copy this into the clipboard and store it somewhere secure. Enter this into the $NinjaOneClientSecret variable below
# Close this window, which will take you back to the Client App IDs tab. Click copy on the Client ID and store this somewhere secure. Enter this into the $NinjaOneClientId variable below

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

# Get all users in NinjaOne
$Users = Invoke-NinjaOneRequest -Method GET -Path 'users' -Paginate

if ($Users.Count -eq 0) {
    Write-Output 'No Users Found'
    exit
}

# Filters for the users where userType = END_USER
$EndUsers = $Users | Where-Object { $_.userType -eq 'END_USER' }

# Get a list of all organizations in NinjaOne
$Organizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations' -Paginate

# Create a hashtable to map organizationId to organizationName for faster lookup
$OrgHashTable = @{}
$Organizations | ForEach-Object {
    $OrgHashTable[$_.id] = $_.name
}

# Process each End User: make deviceIds a comma-delimited string and add OrganizationName
$EndUsers | ForEach-Object {
    # Make deviceIds a comma-delimited string
    $DeviceIds = $_.deviceIds -join ','
    $_.deviceIds = $DeviceIds

    # Add OrganizationName property
    $OrganizationName = $OrgHashTable[$_.organizationId]
    $_ | Add-Member -MemberType NoteProperty -Name OrganizationName -Value $OrganizationName
}

# Export all End Users to a CSV
$PathToSaveCSV = "C:\Temp\NinjaOneEndUsers_" + (Get-Date -f "yyyyMMdd_HHmm") + ".csv"
$EndUsers | Export-Csv -Path $PathToSaveCSV -NoTypeInformation
