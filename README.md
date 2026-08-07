# OPB Check-in — QR scanner web app

A mobile-first PWA that replaces the old Power App + Power Automate flow. A volunteer
scans a guest's invoice QR code; the app applies the original business rules and
records the check-in (`Status` + `DateTime`) in the OPB Excel Online workbook.

## Architecture

| Part | What | Where |
| ---- | ---- | ----- |
| Front end | Self-contained PWA scanner, vanilla JS, no build, no secrets | `webapp/index.html` (+ `sw.js`, `manifest.webmanifest`, icons) — GitHub Pages |
| Back end | Express proxy: business rules + Microsoft Graph writes | `webapp/api/` — Azure App Service (Free F1) in tenant `ad340c84…`, RG `rg-opb-checkin` |

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
3. Session = CET/Oslo local time strictly after the **cut-off → Dinner**, else **Lunch**.
   The cut-off is **16:00** (was 14:00 in the original flow) and is configurable via the
   `SESSION_CUTOFF` app setting (HHmm, e.g. `1600`). Exactly at the cut-off stays Lunch.
4. Valid = candidates whose pass **Date == today (Oslo)** and **PassType == session** — so a
   pass is honoured only for its own date and the current meal, never for all days/both meals.
   - Discrete-column workbooks match on `Date`/`PassType`.
   - Collapsed-key workbooks match `AppKey`/`UniqueKey` starting with `order+dateSerial+session`.
5. No valid row → `ERROR!!! This Pass is NOT valid at this moment!`
6. First valid row already `REGISTERED` → `ERROR!!! {name} is already registered!!!`
7. Otherwise set `Status=REGISTERED`, `DateTime=now(Oslo)` on **all** valid rows →
   `SUCCESS!! {name} has been registered successfully!`

Both workbook variants are supported automatically (see `rules.js` column detection):
- **2025 style** (`App_Source` Table1): `RegistrationID, UniqueKey, First/Last, Item, Date, PassType, FoodOption, Quantity, Status, DateTime`.
- **2026 style**: `Order Number, UniqueKey, First/Last, Quantity, AppKey, Status, DateTime`.

## Sign-in (Google / Microsoft social login)

Volunteers sign in with **Google** or **Microsoft (personal/Live)**; the backend
verifies the provider ID token (signature, issuer, audience, expiry, `email_verified`)
and checks the email against an **allowlist** before issuing a short-lived session JWT.
The scan and guest-list endpoints require that session; `ADMIN_EMAILS` get the admin role.

Set up (all free, no admin):
1. **Google** — Google Cloud Console → *APIs & Services → Credentials → OAuth client ID*
   (type *Web*); add your Pages URL as an authorized JavaScript origin. Copy the client ID.
2. **Microsoft** — an app registration (personal accounts) with a **SPA** redirect URI =
   your Pages URL. Copy the client ID. (Can be a separate registration from the Graph one.)
3. Put both IDs in `webapp/index.html` → `AUTH_CFG`. Set the allowlist + IDs on the backend:
   ```powershell
   az webapp config appsettings set -n opb-checkin-api -g rg-opb-checkin --settings `
     GOOGLE_CLIENT_ID="<google id>" MS_CLIENT_ID="<ms id>" `
     ADMIN_EMAILS="you@gmail.com" USER_EMAILS="vol1@gmail.com,vol2@live.com"
   ```
`SESSION_SECRET` is generated automatically by `provision.ps1`.


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

### 4. Provision + deploy the backend (dedicated F1 resources, tenant `ad340c84…`)
```powershell
az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
cd webapp/api
./provision.ps1        # creates rg-opb-checkin + Free F1 plan + web app opb-checkin-api
az webapp config appsettings set -n opb-checkin-api -g rg-opb-checkin --settings `
  OPB_CLIENT_ID="<app id>" OPB_CLIENT_SECRET="<secret>" OPB_REFRESH_TOKEN="<from step 2>"
./push.ps1             # zip-deploy
# verify: GET https://opb-checkin-api.azurewebsites.net/health -> {"ok":true}
```
(`GRAPH_WORKBOOK="path:/Oslo Durgotsav 2026_Test.xlsx"` and `TABLE_NAME="Table1"` are preset by provision.ps1. Default subscription: `ME-MngEnvMCAP218279-didharch-2`.)

### 5. Publish the front end
Set `BACKEND_URL` at the top of `webapp/index.html` to the App Service URL, then serve
`webapp/` from GitHub Pages (or any static host). Open on a phone and "Add to Home Screen".

### Versioning
A build-version badge (bottom-left) shows the running build and prompts a reload when a
newer one is deployed. On each release, bump the same `vNN` in **three** places:
`index.html` → `APP_VERSION`, `sw.js` → `VERSION`, and `version.json`.

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
- Volunteer/admin access requires **Google/Microsoft sign-in** verified server-side against an
  email allowlist; scan and guest-list endpoints require a session (see "Sign-in" above).
- Consider signing the QR payload (HMAC) so a fabricated order number can't be walked in.
