# HH Warehouse Register — public frontend, test MVP v0.1

Independent warehouse web application with the visual approach of HH Project Register. Does NOT connect to or modify the active Project Register.

## Run locally
Node.js 20+:
```
npm install
npm run dev
```
The front page offers **Explore demo version**. All demo names, quantities and transactions are synthetic; changes are temporary in browser memory and reset on reload. The demo does not require Entra or PostgreSQL.

## Test functionality
Dashboard; inventory catalogue; create/edit item; critical stock based on **total**; receive stock; issue consumable; rent equipment; full return/loss; month-end rental billing; project and employee directories; search/sort; resize table columns; +/− and Ctrl+wheel zoom; fullscreen; audit display; PWA shell; HTML launcher template for OneDrive.

The real backend is the private `IceHummmer/warehouse-be` repository. Before connected testing, configure `public/config.js` with public Microsoft Entra client/tenant IDs, Warehouse API scope and API URL. These IDs are public identifiers, NOT secrets. The separate Entra API must expose an access scope, and the backend must validate its bearer tokens.

Build: `npm run build`; Render static-site Publish Directory: `dist`.

## Known limitations
This is a prototype, not production: no historical Excel import, no tested Render deployment, no backups/restore, incomplete admin/system message UI, no PDF/documents/SharePoint integration yet, no explicit edit-lease dialog (stale server updates are rejected with revision conflicts), no partial rental returns, no equipment transfers. Real data is never included in public repo.

When real backend is configured, all mutations are authorized and validated by the API, not by UI visibility. Do not release to staff or connect production stock until acceptance and reconciliation.

The `public/opener.html` launcher requires its URL to be replaced before distributing in OneDrive.
