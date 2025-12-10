# SharePoint Permission Management Scripts

This directory contains PowerShell scripts and batch file wrappers for managing SharePoint site permissions for applications via Microsoft Graph API.

## ⚠️ Important Understanding

**Microsoft Graph API only supports SITE-LEVEL permissions for applications.**

- Permissions granted apply to **ALL libraries** in the SharePoint site
- There is no library-specific permission granularity for applications
- Library access control must be implemented in your application logic

See [../docs/SHAREPOINT_PERMISSIONS.md](../docs/SHAREPOINT_PERMISSIONS.md) for detailed explanation.

---

## Scripts Overview

| Script | Purpose |
|--------|---------|
| `grant-library-access.ps1` | Grant site-level permissions to an application |
| `grant-access.bat` | Batch wrapper for grant script |
| `check-library-access.ps1` | Check existing site permissions |
| `check-access.bat` | Batch wrapper for check script |
| `revoke-site-access.ps1` | Revoke (delete) site permissions for an application |
| `revoke-access.bat` | Batch wrapper for revoke script |

---

## Prerequisites

### 1. Microsoft Graph PowerShell Module

The scripts will auto-install if needed, but you can install manually:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Sites -Scope CurrentUser -Force
```

### 2. Azure AD Permissions

Your user account needs:
- **Sites.FullControl.All** (for grant/revoke operations)
- **Sites.Read.All** (for check operations)

### 3. Application Registration

You need:
- **Tenant ID** - Your Azure AD tenant ID
- **Site URL** - SharePoint site URL (e.g., `https://contoso.sharepoint.com/sites/demo`)
- **App ID** - Application (client) ID from Azure AD app registration

---

## How to Run the Scripts

### Option 1: Using Batch Files (Easiest)

#### Step 1: Edit Configuration

Open the `.bat` file in a text editor and update these variables:

```batch
SET TenantId=YOUR-TENANT-ID
SET SiteUrl=https://YOUR-SITE-URL
SET AppId=YOUR-APP-ID
```

#### Step 2: Run the Batch File

Double-click the `.bat` file or run from command prompt:

```cmd
cd scripts
grant-access.bat     REM Grant permissions
check-access.bat     REM Check permissions
revoke-access.bat    REM Revoke permissions
```

### Option 2: Using PowerShell Directly

#### Grant Site Permission

```powershell
cd scripts
.\grant-library-access.ps1 `
    -TenantId "a3e8f00c-1024-4869-a9e1-96ca13af6290" `
    -SiteUrl "https://ledang1001.sharepoint.com/sites/demo" `
    -AppId "4ca0dc20-704b-45ae-871d-a53fae722bf3" `
    -Permission "write"
```

**Parameters:**
- `TenantId` (required) - Azure AD tenant ID
- `SiteUrl` (required) - SharePoint site URL
- `AppId` (required) - Application client ID
- `Permission` (optional) - `read` or `write` (default: `write`)

#### Check Site Permissions

```powershell
.\check-library-access.ps1 `
    -TenantId "a3e8f00c-1024-4869-a9e1-96ca13af6290" `
    -SiteUrl "https://ledang1001.sharepoint.com/sites/demo" `
    -AppId "4ca0dc20-704b-45ae-871d-a53fae722bf3"
```

**Parameters:**
- `TenantId` (required) - Azure AD tenant ID
- `SiteUrl` (required) - SharePoint site URL
- `AppId` (optional) - Filter to specific application

#### Revoke Site Permission

```powershell
.\revoke-site-access.ps1 `
    -TenantId "a3e8f00c-1024-4869-a9e1-96ca13af6290" `
    -SiteUrl "https://ledang1001.sharepoint.com/sites/demo" `
    -AppId "4ca0dc20-704b-45ae-871d-a53fae722bf3"
```

**Parameters:**
- `TenantId` (required) - Azure AD tenant ID
- `SiteUrl` (required) - SharePoint site URL
- `AppId` (required) - Application client ID to revoke

⚠️ **Warning:** Requires typing "YES" to confirm deletion.

---

## Typical Workflow

### 1. Grant Permission to Your Application

```cmd
REM Edit grant-access.bat first
grant-access.bat
```

**What happens:**
- Connects to Microsoft Graph with your credentials
- Resolves the SharePoint site ID from the URL
- Grants site-level permissions to the application
- Permission applies to ALL libraries in the site

### 2. Verify Permissions Were Granted

```cmd
check-access.bat
```

**Expected output:**
```
Site Permissions Found: 1
  Permission ID: aTowaS50...
  Roles:         write
  App ID:        4ca0dc20-704b-45ae-871d-a53fae722bf3
  App Name:      YourAppName
```

### 3. Revoke Permission (When No Longer Needed)

```cmd
revoke-access.bat
```

**What happens:**
- Finds all permissions for the specified App ID
- Shows which permissions will be deleted
- Asks for "YES" confirmation
- Deletes the permissions via Graph API

---

## Authentication Flow

When you run any script, you'll see:

```
Connecting to Microsoft Graph...
```

A browser window will open asking you to:
1. Sign in with your Azure AD account
2. Grant consent for the requested permissions
3. Close the browser when prompted

The connection is cached, so subsequent runs won't require re-authentication (unless the session expires).

---

## Troubleshooting

### "Access Denied" or "Insufficient Privileges"

**Solution:** Your account needs admin permissions or delegated access to manage site permissions.

Contact your SharePoint admin to grant you:
- SharePoint Administrator role, or
- Site Collection Administrator for the specific site

### "Module Not Found"

**Solution:** Install Microsoft Graph modules:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Sites -Scope CurrentUser -Force
```

### "Site Not Found"

**Solution:** Verify the Site URL format:
- ✅ Correct: `https://contoso.sharepoint.com/sites/demo`
- ❌ Wrong: `https://contoso.sharepoint.com/sites/demo/`
- ❌ Wrong: `https://contoso.sharepoint.com/sites/demo/Shared Documents`

### Permission Granted But App Can't Access Files

**Cause:** Azure AD app registration needs SharePoint API permissions.

**Solution:** In Azure Portal → App Registration → API Permissions, add:
- `Sites.Read.All` or `Sites.ReadWrite.All` (Application permission)
- Grant admin consent

---

## API References

- [Grant Site Permission](https://learn.microsoft.com/en-us/graph/api/site-post-permissions)
- [List Site Permissions](https://learn.microsoft.com/en-us/graph/api/site-list-permissions)
- [Delete Site Permission](https://learn.microsoft.com/en-us/graph/api/site-delete-permission)
- [SharePoint Sites Permissions](https://learn.microsoft.com/en-us/graph/permissions-reference#sitesreadall)

---

## Security Best Practices

1. **Use least privilege permissions**
   - Grant `read` if the app only needs to read files
   - Only use `write` if the app needs to create/modify files

2. **Audit permissions regularly**
   - Run `check-access.bat` periodically to review what apps have access
   - Revoke permissions for unused applications

3. **Don't commit credentials to Git**
   - The `.bat` files contain your tenant/site/app IDs
   - Consider creating `.bat.sample` templates and gitignore the actual `.bat` files

4. **Use managed identities in production**
   - For Azure-hosted apps, use Managed Identity instead of App ID/Secret
   - This eliminates the need for credential management

---

## Next Steps

After granting permissions:

1. **Update your application code** to use the SharePoint service
2. **Implement library access control** in your application logic (see [SHAREPOINT_PERMISSIONS.md](../docs/SHAREPOINT_PERMISSIONS.md))
3. **Test file operations** (read, write, search)
4. **Monitor access logs** in Azure AD and SharePoint

---

## Questions?

See the detailed documentation:
- [SharePoint Permissions Guide](../docs/SHAREPOINT_PERMISSIONS.md) - Understanding permission limitations
- [Deployment Guide](../DEPLOYMENT.md) - Full application deployment steps
