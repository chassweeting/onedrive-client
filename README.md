# OneDrive Integration

Python client for OneDrive/SharePoint file operations via Microsoft Graph API.

## Quick Start

### 1. Prerequisites

- Python 3.11+
- [Azure CLI](https://docs.microsoft.com/en-us/cli/azure/install-azure-cli)
- [Poetry](https://python-poetry.org/docs/#installation)

### 2. Setup

```bash
# Install dependencies
make install

# Create Azure AD app registration and generate .env
make setup
```

The setup script will:
- Log you into Azure (if needed)
- Create an Azure AD app registration
- Add Microsoft Graph permissions (User.Read, Files.Read.All, Sites.Read.All)
- Generate a `.env` file with your credentials

### 3. Use the Client

```python
import os
from azure.identity import DeviceCodeCredential
from src.onedrive import OneDriveClient

credential = DeviceCodeCredential(
    client_id=os.environ["AZURE_CLIENT_ID"],
    tenant_id=os.environ["AZURE_TENANT_ID"],
)

client = OneDriveClient(credential=credential)

# This prints a URL + code - open the URL and enter the code to authenticate
# You'll see a consent prompt - click Accept
name = await client.get_user_display_name()
print(name)

# List your OneDrive files
drive_id = await client.get_my_drive_id()
items = await client.list_items(drive_id)
for item in items:
    print(f"{'📁' if item.is_folder else '📄'} {item.name}")
```

## Make Commands

| Command | Description |
|---------|-------------|
| `make install` | Install dependencies with Poetry |
| `make setup` | Create Azure AD app and generate `.env` |
| `make lint` | Check code with ruff |
| `make fix` | Auto-fix lint issues |
| `make format` | Format code with ruff |

---

## Reference

### Manual Setup Commands

If you prefer to run the setup steps manually:

```bash
# Login to Azure
az login

# Get your tenant ID
az account show --query tenantId -o tsv

# Create the app registration
az ad app create \
  --display-name "OneDrive Integration Client" \
  --sign-in-audience "AzureADMyOrg" \
  --enable-id-token-issuance true \
  --enable-access-token-issuance true \
  --public-client-redirect-uris "http://localhost"

# Get the App (Client) ID
APP_ID=$(az ad app list --display-name "OneDrive Integration Client" --query "[0].appId" -o tsv)

# Add Microsoft Graph permissions
GRAPH_API="00000003-0000-0000-c000-000000000000"

# User.Read (delegated)
az ad app permission add --id $APP_ID \
  --api $GRAPH_API \
  --api-permissions e1fe6dd8-ba31-4d61-89e7-88639da4683d=Scope

# Files.Read.All (delegated)
az ad app permission add --id $APP_ID \
  --api $GRAPH_API \
  --api-permissions df85f4d6-205c-4ac5-a5ea-6bf408dba283=Scope

# Sites.Read.All (delegated)
az ad app permission add --id $APP_ID \
  --api $GRAPH_API \
  --api-permissions 205e70e5-aba6-4c52-a976-6d2d46c48043=Scope
```

### Permission GUIDs

The `--api` flag uses Microsoft Graph's well-known App ID:
- `00000003-0000-0000-c000-000000000000` = Microsoft Graph API

The `--api-permissions` flag uses permission-specific GUIDs:

| GUID | Permission | Type |
|------|------------|------|
| `e1fe6dd8-ba31-4d61-89e7-88639da4683d` | User.Read | Delegated |
| `df85f4d6-205c-4ac5-a5ea-6bf408dba283` | Files.Read.All | Delegated |
| `205e70e5-aba6-4c52-a976-6d2d46c48043` | Sites.Read.All | Delegated |

### Delegated vs Application Permissions

- `=Scope` = **Delegated** - runs as the signed-in user, user consent required
- `=Role` = **Application** - runs as the app itself, admin consent required

For interactive use (IPython, desktop apps), use **delegated** permissions.

### Discovering Permission GUIDs

```bash
# List all delegated permissions
az ad sp show --id 00000003-0000-0000-c000-000000000000 \
  --query "oauth2PermissionScopes[].{name:value, id:id}" -o table

# List all application permissions
az ad sp show --id 00000003-0000-0000-c000-000000000000 \
  --query "appRoles[].{name:value, id:id}" -o table

# Find a specific permission
az ad sp show --id 00000003-0000-0000-c000-000000000000 \
  --query "oauth2PermissionScopes[?value=='Files.Read.All']"
```

### Official Documentation

- [Microsoft Graph Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Well-known Microsoft App IDs](https://learn.microsoft.com/en-us/troubleshoot/entra/entra-id/governance/verify-first-party-apps-sign-in#application-ids-of-commonly-used-microsoft-applications)

---

## Troubleshooting

### "accessDenied - This operation is not supported with the provided scopes"

Your credential doesn't have Graph API permissions. Make sure you:
1. Ran `make setup` to create the app registration
2. Are using `DeviceCodeCredential` with the correct `client_id` and `tenant_id` from `.env`

### "Authorization_RequestDenied - This operation can only be performed by an administrator"

You tried to run `az ad app permission admin-consent` without admin rights. This is fine for delegated permissions - the consent prompt will appear when you sign in via the device code flow.

### "You do not have access to create this personal site"

Your account doesn't have OneDrive provisioned. Try:
1. Visit https://portal.office.com and click OneDrive to trigger provisioning
2. Or use SharePoint sites instead: `await client.list_followed_sites()`
