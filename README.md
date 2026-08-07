# OPB Check-in — QR scanner web app

A mobile-first PWA that replaces the old Power App + Power Automate flow. A volunteer
scans a guest's invoice QR code; the app applies the original business rules and
records the check-in (`Status` + `DateTime`) in the OPB Excel Online workbook.

## Architecture

| Part | What | Where |
| ---- | ---- | ----- |
| Front end | Self-contained PWA scanner, vanilla JS, no build, no secrets | `webapp/index.html` (+ `sw.js`, `manifest.webmanifest`, icons) — GitHub Pages |
| Back end | Express proxy: business rules + Microsoft Graph writes | `webapp/api/` — Azure App Service (Free F1) |

The workbook lives on a **personal OneDrive** (`osloprobaseebangali@outlook.com`), which
has **no service principals**. So the backend reaches it with a **delegated** Graph
refresh token minted once by the owner (`authorize.js`). No password is ever stored;
the only server-side secrets are the app's client secret and that refresh token.

```
QR "3499;" → PWA (order 3499) → POST /api/register → backend
   → Graph read Table → apply rules → PATCH Status/DateTime → SUCCESS/ALREADY/INVALID
```

## Business rules (ported from `OPB_Excel_QRCodeScannerFlow`)

1. `order` = number before the first `;` in the QR.
2. Candidates = rows where `Order Number`/`RegistrationID` == order.
3. Session = Oslo local time after **14:00 → Dinner**, else **Lunch**.
4. Valid = candidates whose pass **Date == today (Oslo)** and **PassType == session**.
   - Discrete-column workbooks match on `Date`/`PassType`.
   - Collapsed-key workbooks match `AppKey`/`UniqueKey` starting with `order+dateSerial+session`.
5. No valid row → `ERROR!!! This Pass is NOT valid at this moment!`
6. First valid row already `REGISTERED` → `ERROR!!! {name} is already registered!!!`
7. Otherwise set `Status=REGISTERED`, `DateTime=now(Oslo)` on **all** valid rows →
   `SUCCESS!! {name} has been registered successfully!`

Both workbook variants are supported automatically (see `rules.js` column detection):
- **2025 style** (`App_Source` Table1): `RegistrationID, UniqueKey, First/Last, Item, Date, PassType, FoodOption, Quantity, Status, DateTime`.
- **2026 style**: `Order Number, UniqueKey, First/Last, Quantity, AppKey, Status, DateTime`.

## One-time setup

### 1. Register the app (free, no tenant admin)
At <https://entra.microsoft.com> → **App registrations** → **New registration**:
- Supported accounts: **Personal Microsoft accounts only** (or "…and personal").
- Platform **Web**, redirect URI `http://localhost:53682/callback`.
- **Certificates & secrets** → new client secret → copy the value.
- **API permissions** → Microsoft Graph → **Delegated** → `Files.ReadWrite`, `offline_access`.

### 2. Mint the refresh token (sign in once as the workbook owner)
```powershell
cd webapp/api
npm install
$env:OPB_CLIENT_ID="<app id>"; $env:OPB_CLIENT_SECRET="<secret>"
npm run authorize          # opens sign-in; MFA/passkey handled here; prints the refresh token
```

### 3. Find the workbook locator
Either a path (simplest) or the driveItem id:
- Path form: `GRAPH_WORKBOOK="path:/OPB Boishakhi Adda 2026.xlsx"`
- Id form: `GRAPH_WORKBOOK="id:<driveItemId>"`
Set `TABLE_NAME` to the table you update (default `Table1`).

### 4. Provision + deploy the backend (bare minimum, tenant `ad340c84…`)
```powershell
az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
cd webapp/api
./provision.ps1 -Subscription "<sub>" -PagesOrigin "https://<user>.github.io"
az webapp config appsettings set -n opb-checkin-api -g rg-opb-checkin --settings `
  OPB_CLIENT_ID="<app id>" OPB_CLIENT_SECRET="<secret>" OPB_REFRESH_TOKEN="<from step 2>" `
  GRAPH_WORKBOOK="path:/OPB Boishakhi Adda 2026.xlsx" TABLE_NAME="Table1"
./push.ps1 -Subscription "<sub>"
# verify: GET https://opb-checkin-api.azurewebsites.net/health -> {"ok":true}
```

### 5. Publish the front end
Set `BACKEND_URL` at the top of `webapp/index.html` to the App Service URL, then serve
`webapp/` from GitHub Pages (or any static host). Open on a phone and "Add to Home Screen".

## Run / test locally
```powershell
cd webapp/api
npm test                    # business-rule tests for both workbook variants
node server.js              # needs the env vars above; GET /health, POST /api/register
# front end:
cd ..; python -m http.server 8765   # http://localhost:8765/index.html
```

## Security notes
- No secrets in the browser or repo; the refresh token/client secret live only in App Service settings.
- Backend fails CORS closed to `ALLOWED_ORIGINS` and rate-limits per IP.
- **Deferred (next phase):** volunteer/admin login (identity provider + roles). Until that
  is added, keep `ALLOWED_ORIGINS` tight and treat the endpoint as semi-public.
- Consider signing the QR payload (HMAC) so a fabricated order number can't be walked in.
