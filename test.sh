#!/usr/bin/env bash
# Smoke test for --broker-cmd on Dean's CB-managed Windows host (WSL2 → azureauth.exe → WAM).
# Exercises the WSL bridge designed in deasmi/ms-365-mcp-server issue #1
# and the upstream Conditional Access failure in Softeria/ms-365-mcp-server#443.
set -euo pipefail

cd "$(dirname "$0")"

AZUREAUTH='/mnt/c/Users/dean.smith/AppData/Local/Microsoft/WinGet/Packages/Microsoft.Azure.Auth_Microsoft.Winget.Source_8wekyb3d8bbwe/azureauth.exe'
# Microsoft Graph PowerShell — first-party multi-tenant client with the WAM
# broker redirect URI registered AND broad Graph delegated permissions.
# Requires one-time admin consent in the CB tenant (in progress) and a
# one-time interactive seed from PowerShell on Windows side to populate
# WAM's account cache. After that, the WSL bridge silent-acquires.
CLIENT_ID='14d82eec-204b-4c2f-b7e8-296a70dab67e'
TENANT_ID='a51576e9-76df-4a71-a887-b066ed4e5c27'   # Credit Benchmark tenant
RESOURCE='https://graph.microsoft.com'

BROKER_CMD="$AZUREAUTH aad --client $CLIENT_ID --tenant $TENANT_ID --resource $RESOURCE --mode broker --output json"

echo "==> azureauth: $AZUREAUTH"
echo "==> broker cmd: $BROKER_CMD"
echo

node dist/index.js --broker-cmd "$BROKER_CMD" --verify-login
