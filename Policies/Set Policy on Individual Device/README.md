# Set-DevicePolicy

## Overview

This script updates the assigned policy on a specific NinjaOne device. It sends only the `policyId` field in the PATCH request, ensuring no other device properties are modified.

## Attribution

- **Author:** Gavin Stone (NinjaOne)

## Requirements / Prerequisites

- **NinjaOne API Credentials:**
  - `NinjaOneClientId`
  - `NinjaOneClientSecret`
- **NinjaOne Instance URL:** e.g., `eu.ninjarmm.com`
- **Policy ID to change too:** You can find this by loading a policy, and it's the number that appears at the top, IE for editor/policy/106 it would be 106

## How It Works

1. **Authentication:** The script authenticates with the NinjaOne API using OAuth2 Client Credentials, obtaining an access token via `Connect-NinjaOne`.
2. **Fetch Device:** Retrieves the target device via `device/{id}` to confirm it exists and display its current policy assignment.
3. **Compare Policies:** Checks whether the device is already assigned to the requested policy and exits early if no change is needed.
4. **Update Policy:** Sends a PATCH request to `device/{id}` with only `{ policyId: <value> }` in the body, ensuring no other device properties are touched.
5. **Confirm Result:** Displays a success message with the device name and new policy ID.

## Usage

1. **Set Variables:**
   - Open `Set-DevicePolicy.ps1` in an editor (e.g., VS Code, PowerShell ISE).
   - Fill in your NinjaOne credentials and any script-specific variables at the top:

     ```powershell
     $NinjaOneInstance     = "eu.ninjarmm.com"
     $NinjaOneClientId     = "your_client_id"
     $NinjaOneClientSecret = "your_client_secret"
     $DeviceId = 1
     $NewPolicyId = 42 # You can find this by loading a policy, and it's the number that appears at the top
     ```

2. **Run the Script:**

   ```powershell
   .\Set-DevicePolicy.ps1
   ```

### Expected Output

```
Fetching device 142...
Device:         DESK-001 (ID: 142)
Current Policy: 15
New Policy:     42

Updating policy...
Successfully updated device 'DESK-001' (ID: 142) to policy ID 42.
```

### Troubleshooting

- **Issue:** Authentication fails with an error message.
  - **Solution:** Verify that `$NinjaOneClientId` and `$NinjaOneClientSecret` are correct and that the API client has the required scopes (`monitoring`, `management`).

- **Issue:** The script cannot connect to the NinjaOne API.
  - **Solution:** Ensure that `$NinjaOneInstance` is correct (e.g., `eu.ninjarmm.com`, `app.ninjarmm.com`, `oc.ninjarmm.com`) and accessible from your network.

- **Issue:** Error: Could not find device with ID.
  - **Solution:** Verify the device ID exists in your NinjaOne instance. You can find device IDs in the NinjaOne dashboard or by querying the `/v2/devices` endpoint.

- **Issue:** 403 Forbidden when updating the device.
  - **Solution:** Ensure your API client credentials have the 'management' scope and write permissions for the target device's organisation.

## Notes

- Ensure that your NinjaOne API credentials are kept secure and not shared.
- The PATCH request sends only `policyId` — no other device fields (displayName, location, userData, etc.) are included, so they remain unchanged.
- To update multiple devices, you can wrap the script in a loop or modify it to accept an array of device IDs.
