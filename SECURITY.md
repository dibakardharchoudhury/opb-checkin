# Security model

## Trust boundaries

The GitHub Pages frontend is public and must be treated as untrusted. It contains no Graph
credential, workbook content, application signing secret, or allowlist. Google and Microsoft
ID tokens are verified by the Azure API before an application session is issued. Every request
that reads or changes workbook or configuration data is independently authorized by the API.

The three `Transport & Logistics` views are deliberately public. Anonymous navigation to them:

- loads Star Schedule and Food Pickup from static frontend content;
- calls only `GET /api/public/transport` for validated Pickup / Drop Plan card overrides;
- does not create, restore, or receive an application session;
- does not load workbook data; and
- does not bypass the login gate on any other module.

Adding a name to `PUBLIC_VIEWS` is appropriate only for content embedded in the public frontend.
Any dynamic public content needs a dedicated response allowlist; `PUBLIC_VIEWS` must never be used
as authorization for an operational API-backed view.

## Backend controls

- Provider tokens are checked for signature, issuer, audience, expiry, and the expected RS256
  algorithm. Google accounts with `email_verified=false` are rejected.
- A verified identity must also appear in the server-side user/admin allowlist.
- Application sessions are issuer-bound, expire after 12 hours, and accept HS256 only.
- All workbook, guest, configuration, food, parking, and walk-in routes use `requireAuth()`;
  administrative routes additionally require the `admin` role.
- `GET /api/public/transport` returns only public card fields: time/meal, topic, location/note,
  assigned people, and backup people. Its corresponding write
  route requires the `admin` role and validates IDs, lengths, types, and item counts.
- CORS permits only configured exact origins and permits no origin when configuration is absent.
  CORS is browser isolation and defense in depth, not authentication.
- API responses use `Cache-Control: no-store`, HSTS, frame denial, MIME sniffing protection,
  a no-referrer policy, bounded JSON bodies, and per-IP rate limits.
- Internal Graph/token errors are logged server-side and are not returned to clients.
- The service worker never intercepts or caches API requests.
- App Service enforces HTTPS and TLS 1.2 or newer. Provisioning disables FTP/FTPS and remote
  debugging remains disabled.

The API must remain internet reachable because requests originate from volunteers' devices.
IP restrictions are useful only when all users have known VPN or corporate egress addresses.

## Public transport analysis

Publishing the transport pages does not weaken operational authorization because two are static and
the only anonymous backend path returns transport labels intended for public display. They cannot
access the workbook or general configuration. Knowledge of an API URL or order number does not
satisfy the signed session and allowlist checks, and an attacker cannot use CORS as a credential substitute.

The pages are not risk-free. Exact pickup times, flight numbers, hotel or Airbnb addresses, vehicle
labels, and named travellers can reveal attendance and movement patterns. This is an operational privacy
and physical-security consideration even though it does not expose workbook data or credentials.
Only publish details approved for unrestricted distribution. Prefer meeting points over private
addresses, roles or initials over full names, and remove stale movements promptly after the event.

## Verification checklist

Run before each backend release:

```powershell
cd api
npm test
npm audit --omit=dev
```

The automated suite covers role resolution, token signature/tamper/expiry/algorithm rejection,
malformed authorization, admin enforcement, fail-closed CORS, security headers, sanitized errors,
business rules, configuration, Graph helpers, and user storage.

Live release checks should verify:

1. `/health` returns 200, and the only anonymous data endpoint is the content-limited
  `GET /api/public/transport` route.
2. Every `/api/*` data route returns 401 without a session.
3. Invalid provider credentials and malformed/tampered sessions return 401.
4. The configured Pages origin receives CORS headers; another origin does not.
5. HTTP redirects to HTTPS and API responses include the documented security/cache headers.
6. Anonymous navigation is limited to the three Transport & Logistics pages; only Pickup / Drop Plan
  calls the public transport read endpoint. Protected modules open the sign-in gate, and transport
  writes reject non-admin users.

## Residual risks

- A compromised allowlisted social account can operate with that account's assigned role until
  it is removed; provider MFA/passkeys and prompt offboarding reduce this risk.
- Session tokens are stored in `sessionStorage`, limiting persistence but remaining accessible to
  JavaScript running in the page. Third-party scripts are pinned to versions where possible;
  dependency updates and frontend script sources require review.
- The delegated Graph refresh token is a high-value secret. Rotate it and the client secret after
  suspected exposure, and keep App Service setting access narrowly assigned.
- QR payloads are order numbers rather than signed claims. Server-side authentication prevents
  anonymous use, but a signed QR format would better resist guessing by an approved operator.

Report a suspected exposure to the application owner. Revoke affected sessions by rotating
`SESSION_SECRET`; rotate Graph credentials separately if workbook access may be affected.
