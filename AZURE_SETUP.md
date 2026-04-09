# Azure AD App Registration — MHL Services PWA

Complete these steps once to get your `AZURE_CLIENT_ID` and `AZURE_TENANT_ID`, then paste them into `report.html` and `crew.html`.

---

## Step 1 — Sign in to Azure Portal

1. Go to [https://portal.azure.com](https://portal.azure.com)
2. Sign in with your **MHL Services Microsoft 365 admin account** (e.g. oliver@mhlservices.com)

---

## Step 2 — Register a new app

1. In the search bar at the top, search for **"App registrations"** and click it
2. Click **+ New registration**
3. Fill in:
   - **Name:** `MHL Services PWA`
   - **Supported account types:** select **"Accounts in this organizational directory only (MHL Services only — Single tenant)"**
   - **Redirect URI:**
     - Platform: **Single-page application (SPA)**
     - URL: `https://oliverbid.github.io/mhl-pwa/report.html`
4. Click **Register**

---

## Step 3 — Add the second redirect URI (for crew form)

After registering, you need to add the crew page redirect URI:

1. In your new app, go to **Authentication** (left sidebar)
2. Under **Single-page application**, click **Add URI**
3. Add: `https://oliverbid.github.io/mhl-pwa/crew.html`
4. Click **Save**

> **Why two URIs?** MSAL redirects back to the exact page the user was on. Both `report.html` and `crew.html` authenticate independently, so both need to be registered.

---

## Step 4 — Note your IDs

On the app's **Overview** page, copy:

| Value | Where to find it |
|-------|-----------------|
| **Application (client) ID** | Overview page, top section |
| **Directory (tenant) ID** | Overview page, top section |

These are your `AZURE_CLIENT_ID` and `AZURE_TENANT_ID`.

---

## Step 5 — Set API permissions

1. Go to **API permissions** (left sidebar)
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Delegated permissions**
5. Search for and check: **`Files.ReadWrite`**
6. Click **Add permissions**
7. Click **Grant admin consent for MHL Services** → confirm **Yes**

> **Important:** Without admin consent, each user will see a consent prompt the first time they sign in. Granting admin consent here skips that prompt for everyone in your organisation.

---

## Step 6 — Paste your IDs into the PWA

Open both `report.html` and `crew.html`. Near the top of the `<script>` block, replace:

```js
const AZURE_CLIENT_ID  = 'AZURE_CLIENT_ID';
const AZURE_TENANT_ID  = 'AZURE_TENANT_ID';
```

with your actual values, e.g.:

```js
const AZURE_CLIENT_ID  = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890';
const AZURE_TENANT_ID  = '11223344-5566-7788-99aa-bbccddeeff00';
```

Commit and push both files to GitHub.

---

## Step 7 — Set up the Excel file in OneDrive

The PWA writes to a fixed path in the signed-in user's OneDrive:

```
/MHL Forms Data/MHL Job Reports.xlsx
```

### Create the file and folder:

1. Open **OneDrive** (https://onedrive.live.com or via Microsoft 365)
2. Sign in as **Oliver** (the account that will own the data)
3. Create a folder called exactly: `MHL Forms Data`
4. Inside that folder, create a new Excel workbook called: `MHL Job Reports.xlsx`

### Create the Excel table:

1. Open `MHL Job Reports.xlsx`
2. In Sheet1, add these headers in row 1 (one per column, A through Q):

| Col | Header |
|-----|--------|
| A | SubmittedBy |
| B | SubmittedAt |
| C | Name |
| D | InvoiceNumber |
| E | Vessel |
| F | Port |
| G | Client |
| H | OperationType |
| I | JobStart |
| J | JobEnd |
| K | WorkHours |
| L | StandbyDays |
| M | TravelDays |
| N | PerDiemPeople |
| O | PerDiemDays |
| P | LogisticsHours |
| Q | CashExpenses |

3. Select cells **A1 through Q1**, then go to **Insert → Table**
4. Check "My table has headers" → click OK
5. Name the table **`JobReports`**:
   - Click anywhere in the table
   - Go to **Table Design** tab (appears at top when table is selected)
   - Change the **Table Name** field (top-left) from `Table1` to `JobReports`
6. Save the file

> **The table name must be exactly `JobReports`** (capital J and R, no spaces). The Graph API call targets this table by name.

---

## Who signs in — important note

The Graph API endpoint used is:

```
/me/drive/root:/MHL Forms Data/MHL Job Reports.xlsx:/workbook/...
```

The `/me/drive` path refers to **the signed-in user's own OneDrive**. This means:

- **Oliver** signs in → data writes to Oliver's OneDrive → ✅ works perfectly
- **Crew members** sign in with their own M365 accounts → the app looks in *their* OneDrive, not Oliver's → the file won't be found

### Recommended approaches for crew:

**Option A — Crew members all sign in as Oliver (simplest)**
Share Oliver's login credentials with crew, or have Oliver's account available on shared devices. All submissions land in one place. Not ideal for security but fine for a small team.

**Option B — Share the file with crew members (recommended)**
1. In OneDrive, right-click `MHL Job Reports.xlsx` → Share → share with each crew member's M365 account (Edit permissions)
2. Update the `EXCEL_ENDPOINT` in `crew.html` to use the file's **drive item ID** instead of the `/me/drive` path:
   - Open the file in OneDrive, look at the URL — it contains a `resid=` parameter with the item ID
   - Or use Graph Explorer (https://developer.microsoft.com/en-us/graph/graph-explorer) to find the item ID
   - Change the endpoint to: `https://graph.microsoft.com/v1.0/me/drive/items/{item-id}/workbook/tables/JobReports/rows`
   - A shared file accessible via "Shared with me" can be reached this way

**Option C — Use SharePoint instead of OneDrive (most robust)**
Put the Excel file in a SharePoint site or Teams channel that everyone has access to. Update the endpoint to:
```
https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/MHL Job Reports.xlsx:/workbook/tables/JobReports/rows
```

---

## Testing the connection

Use **Microsoft Graph Explorer** (no coding required) to verify your setup before going live:

1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
2. Sign in with Oliver's account
3. Run this GET request to confirm the file is reachable:
   ```
   GET https://graph.microsoft.com/v1.0/me/drive/root:/MHL Forms Data/MHL Job Reports.xlsx:/workbook/tables/JobReports/rows
   ```
4. If you see `{"value": []}` or an empty rows array — setup is correct
5. If you see a 404 — check the folder/file name for typos (spaces matter) and confirm the table is named `JobReports`

---

## Summary checklist

- [ ] App registered in Azure AD as type **Single-page application (SPA)**
- [ ] Redirect URI added: `https://oliverbid.github.io/mhl-pwa/report.html`
- [ ] Redirect URI added: `https://oliverbid.github.io/mhl-pwa/crew.html`
- [ ] `Files.ReadWrite` delegated permission added
- [ ] Admin consent granted
- [ ] `AZURE_CLIENT_ID` and `AZURE_TENANT_ID` pasted into `report.html` and `crew.html`
- [ ] OneDrive folder `MHL Forms Data` created under Oliver's account
- [ ] Excel file `MHL Job Reports.xlsx` created with 17-column table named `JobReports`
- [ ] Tested via Graph Explorer — table endpoint returns 200
- [ ] Changes pushed to GitHub Pages
