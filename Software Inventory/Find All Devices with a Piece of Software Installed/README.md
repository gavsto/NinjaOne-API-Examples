# Get-DevicesWithSoftware

## Overview

This script finds all devices managed by NinjaOne that have a specific piece of software installed (defaults to Google Chrome). It uses the bulk software inventory query endpoint for efficient retrieval across the entire estate, then resolves device names from the devices endpoint.

## Attribution

- **Author:** Gavin Stone (NinjaOne)

## Requirements / Prerequisites

- **NinjaOne API Credentials:**
  - `NinjaOneClientId`
  - `NinjaOneClientSecret`
- **NinjaOne Instance URL:** e.g., `eu.ninjarmm.com`

## How It Works

1. **Authentication:** The script authenticates with the NinjaOne API using OAuth2 Client Credentials, obtaining an access token via `Connect-NinjaOne`.
2. **Bulk Software Query:** Retrieves the entire software inventory via `queries/software` using cursor-based pagination, collecting all records in a single bulk operation rather than querying each device individually.
3. **Filter Matches:** Filters the software list using the `$SoftwareNameFilter` regex pattern (default: `Google Chrome`).
4. **Resolve Device Names:** Fetches all devices via `devices` to build a lookup table, mapping device IDs from the software results to device names and classes.
5. **Output Results:** Displays a summary count and a formatted table of all matching installations with device name, version, publisher, and install date.

## Usage

1. **Set Variables:**
   - Open `Get-DevicesWithSoftware.ps1` in an editor (e.g., VS Code, PowerShell ISE).
   - Fill in your NinjaOne credentials and any script-specific variables at the top:

     ```powershell
     $NinjaOneInstance     = "eu.ninjarmm.com"
     $NinjaOneClientId     = "your_client_id"
     $NinjaOneClientSecret = "your_client_secret"
     $SoftwareNameFilter = 'Google Chrome'
     ```

2. **Run the Script:**

   ```powershell
   .\Get-DevicesWithSoftware.ps1
   ```

### Expected Output

```
Fetching software inventory...
  Retrieved 1000 records so far...
  Retrieved 2000 records so far...
  Retrieved 2420 records so far...
Retrieved 2420 total software record(s)
Found 12 installation(s) across 12 device(s). Fetching device details...

Found 12 installation(s) across 12 device(s).

DeviceName     DeviceId DeviceClass          SoftwareName          Version        Publisher      InstallDate
----------     -------- -----------          ------------          -------        ---------      -----------
DESK-001            142 WINDOWS_WORKSTATION  Google Chrome         133.0.6943.142 Google LLC     2025-01-15
DESK-002            287 WINDOWS_WORKSTATION  Google Chrome         133.0.6943.142 Google LLC     2025-02-01
SRV-WEB-01          403 WINDOWS_SERVER       Google Chrome         133.0.6943.142 Google LLC     2024-11-20
LAPTOP-JS           891 WINDOWS_WORKSTATION  Google Chrome          132.0.6834.83 Google LLC     2025-01-10
```

### Troubleshooting

- **Issue:** Authentication fails with an error message.
  - **Solution:** Verify that `$NinjaOneClientId` and `$NinjaOneClientSecret` are correct and that the API client has the required scopes (`monitoring`, `management`).

- **Issue:** The script cannot connect to the NinjaOne API.
  - **Solution:** Ensure that `$NinjaOneInstance` is correct (e.g., `eu.ninjarmm.com`, `app.ninjarmm.com`, `oc.ninjarmm.com`) and accessible from your network.

- **Issue:** No results returned but the software is known to be installed.
  - **Solution:** Check that `$SoftwareNameFilter` matches the display name shown in NinjaOne. The filter uses regex — special characters like `+` or `.` may need escaping (e.g. `\.NET` instead of `.NET`).

## Notes

- Ensure that your NinjaOne API credentials are kept secure and not shared.
- This script uses the bulk `queries/software` endpoint which retrieves the entire software inventory in just 2 paginated API calls (software + devices) — much faster than the per-device approach which requires N+1 calls.
- The `queries/software` endpoint uses named cursor pagination (different from the `after`-based pagination used by `devices`), which is handled directly in the script.
- Change `$SoftwareNameFilter` to search for any software — e.g. `'Firefox'`, `'7-Zip'`, `'Microsoft Teams'`.
- Uncomment the last line of the script to export results to a CSV file.
