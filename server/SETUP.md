# Google Sheets + Apps Script Setup

## Step 1: Create a Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet
2. Name it something like **"Order Checklist Manager"**

## Step 2: Open Apps Script Editor

1. In your Google Sheet, go to **Extensions → Apps Script**
2. This opens the Apps Script editor in a new tab
3. Delete any existing code in the editor

## Step 3: Paste the Code

1. Open `server/google-apps-script.js` from this project
2. Copy the entire contents
3. Paste it into the Apps Script editor (replacing everything)
4. Click the **Save** icon (or Ctrl+S)

## Step 4: Run Seed Data

1. In the Apps Script editor, select **`seedData`** from the function dropdown at the top
2. Click the **Run** button (▶)
3. You'll be prompted to authorize — click **Review Permissions**, choose your Google account, then click **Allow**
4. After it runs, go back to your Google Sheet — you should see 7 tabs:
   - **OrderTypes** (5 rows of data)
   - **Customers** (1 row)
   - **Checklists** (6 rows)
   - **AssignmentRules** (5 rows)
   - **Orders** (empty, headers only)
   - **OrderChecklists** (empty, headers only)
   - **ChecklistResponses** (empty, headers only)

## Step 5: Deploy as Web App

1. In the Apps Script editor, click **Deploy → New deployment**
2. Click the gear icon next to "Select type" and choose **Web app**
3. Set:
   - **Description**: (optional) e.g., "Checklist API v1"
   - **Execute as**: **Me**
   - **Who has access**: **Anyone**
4. Click **Deploy**
5. Copy the **Web app URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

## Step 6: Configure the Frontend

1. Open `order-checklist-manager.jsx`
2. Find this line near the top:
   ```javascript
   const APPS_SCRIPT_URL = "YOUR_APPS_SCRIPT_URL_HERE";
   ```
3. Replace `YOUR_APPS_SCRIPT_URL_HERE` with your Web app URL from Step 5

## Step 7: Run the App

```bash
npm install
npm run dev
```

Open the URL shown by Vite (usually http://localhost:5173).

## Updating the Deployment

If you modify the Apps Script code:

1. In the Apps Script editor, click **Deploy → Manage deployments**
2. Click the edit (pencil) icon on your deployment
3. Under **Version**, select **New version**
4. Click **Deploy**

The URL stays the same — no need to update the frontend config.

## Troubleshooting

- **"Authorization required"**: Re-run `seedData` and accept permissions
- **CORS errors**: Make sure "Who has access" is set to "Anyone" in the deployment
- **Stale data**: Apps Script caches aggressively — wait a few seconds between rapid writes/reads
- **Slow first request**: Apps Script cold-starts can take 2-5 seconds; subsequent requests are faster
