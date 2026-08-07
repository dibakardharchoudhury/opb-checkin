// One-time consent helper — run locally to mint the delegated refresh token that the
// backend uses to reach the personal-OneDrive workbook. You sign in ONCE as the
// workbook owner (osloprobaseebangali@outlook.com); MFA/passkey is handled by
// Microsoft in the browser at that moment. Nothing here stores a password.
//
// Prereqs (all free, no tenant admin):
//   1. Register an app at https://entra.microsoft.com -> App registrations -> New:
//        - Supported account types: "Personal Microsoft accounts only" (or
//          "…and personal Microsoft accounts").
//        - Platform: Web. Redirect URI: http://localhost:53682/callback
//        - Certificates & secrets -> new client secret (copy the VALUE).
//        - API permissions -> Microsoft Graph -> Delegated -> Files.ReadWrite,
//          offline_access (openid, profile). No admin consent needed for personal files.
//   2. Set env before running:
//        $env:OPB_CLIENT_ID="...";  $env:OPB_CLIENT_SECRET="..."
//   3. node authorize.js   -> opens the sign-in URL, then prints the refresh token.
//
// Paste the printed refresh token into the App Service setting OPB_REFRESH_TOKEN.

import http from "node:http";
import crypto from "node:crypto";
import { exec } from "node:child_process";

const CLIENT_ID = process.env.OPB_CLIENT_ID;
const CLIENT_SECRET = process.env.OPB_CLIENT_SECRET || "";
const REDIRECT = "http://localhost:53682/callback";
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const SCOPE = "Files.ReadWrite offline_access openid profile";

if (!CLIENT_ID) { console.error("Set OPB_CLIENT_ID (and OPB_CLIENT_SECRET) first."); process.exit(1); }

const verifier = crypto.randomBytes(32).toString("base64url");
const challenge = crypto.createHash("sha256").update(verifier).digest("base64url");
const state = crypto.randomBytes(8).toString("hex");

const authUrl = `${AUTH}/authorize?` + new URLSearchParams({
  client_id: CLIENT_ID, response_type: "code", redirect_uri: REDIRECT,
  scope: SCOPE, state, code_challenge: challenge, code_challenge_method: "S256",
  prompt: "select_account",
}).toString();

const server = http.createServer(async (req, res) => {
  if (!req.url.startsWith("/callback")) { res.writeHead(404).end(); return; }
  const url = new URL(req.url, REDIRECT);
  if (url.searchParams.get("state") !== state) { res.writeHead(400).end("state mismatch"); return; }
  const code = url.searchParams.get("code");
  if (!code) { res.writeHead(400).end("no code: " + (url.searchParams.get("error_description") || "")); return; }

  try {
    const body = new URLSearchParams({
      client_id: CLIENT_ID, client_secret: CLIENT_SECRET, grant_type: "authorization_code",
      code, redirect_uri: REDIRECT, scope: SCOPE, code_verifier: verifier,
    });
    const r = await fetch(`${AUTH}/token`, {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body,
    });
    const j = await r.json();
    if (!r.ok) throw new Error(j.error_description || j.error || r.status);
    res.writeHead(200, { "Content-Type": "text/plain" }).end("Done — you can close this tab and return to the terminal.");
    console.log("\n==================  REFRESH TOKEN  ==================\n");
    console.log(j.refresh_token);
    console.log("\n====================================================");
    console.log("Set it as the App Service setting OPB_REFRESH_TOKEN.\n");
  } catch (e) {
    res.writeHead(500).end("token exchange failed: " + (e?.message || e));
    console.error(e);
  } finally {
    setTimeout(() => server.close(() => process.exit(0)), 500);
  }
});

server.listen(53682, () => {
  console.log("Opening sign-in in your browser…\nIf it doesn't open, paste this URL:\n" + authUrl + "\n");
  const cmd = process.platform === "win32" ? `start "" "${authUrl}"`
    : process.platform === "darwin" ? `open "${authUrl}"` : `xdg-open "${authUrl}"`;
  exec(cmd);
});
