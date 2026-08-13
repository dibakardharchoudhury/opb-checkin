// One-time consent helper — run locally to mint the delegated refresh token that the
// backend uses to reach the personal-OneDrive workbook. You sign in ONCE as the
// workbook owner account; MFA/passkey is handled by
// Microsoft in the browser at that moment. Nothing here stores a password.
//
// The app registration is a PUBLIC client (no client secret). Just run:
//     cd webapp/api
//     npm run authorize
// A browser opens; sign in as the workbook owner. The refresh token is then stored
// DIRECTLY into the App Service setting OPB_REFRESH_TOKEN via `az` (you must be
// `az login`-ed) — it is never printed. If az storing fails it falls back to printing.
//
// Overridable via env: OPB_CLIENT_ID, OPB_APP, OPB_RG.

import http from "node:http";
import crypto from "node:crypto";
import { exec, execFile } from "node:child_process";

const CLIENT_ID = process.env.OPB_CLIENT_ID || "b56289aa-27f5-4380-a2cf-58a829e7c638";
const CLIENT_SECRET = process.env.OPB_CLIENT_SECRET || ""; // public client: none
const APP = process.env.OPB_APP || "opb-checkin-api";
const RG = process.env.OPB_RG || "rg-opb-checkin";
const REDIRECT = "http://localhost:53682/callback";
const AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const SCOPE = "Files.ReadWrite offline_access openid profile";

const verifier = crypto.randomBytes(32).toString("base64url");
const challenge = crypto.createHash("sha256").update(verifier).digest("base64url");
const state = crypto.randomBytes(8).toString("hex");

const authUrl = `${AUTH}/authorize?` + new URLSearchParams({
  client_id: CLIENT_ID, response_type: "code", redirect_uri: REDIRECT,
  scope: SCOPE, state, code_challenge: challenge, code_challenge_method: "S256",
  prompt: "select_account",
}).toString();

// Store the token straight into App Service so the secret never touches the console.
function storeToken(token) {
  return new Promise((resolve) => {
    execFile("az", ["webapp", "config", "appsettings", "set", "-n", APP, "-g", RG,
      "--settings", `OPB_REFRESH_TOKEN=${token}`], { shell: true }, (err) => resolve(!err));
  });
}

const server = http.createServer(async (req, res) => {
  if (!req.url.startsWith("/callback")) { res.writeHead(404).end(); return; }
  const url = new URL(req.url, REDIRECT);
  if (url.searchParams.get("state") !== state) { res.writeHead(400).end("state mismatch"); return; }
  const code = url.searchParams.get("code");
  if (!code) { res.writeHead(400).end("no code: " + (url.searchParams.get("error_description") || "")); return; }

  try {
    const params = { client_id: CLIENT_ID, grant_type: "authorization_code",
      code, redirect_uri: REDIRECT, scope: SCOPE, code_verifier: verifier };
    if (CLIENT_SECRET) params.client_secret = CLIENT_SECRET;
    const r = await fetch(`${AUTH}/token`, {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(params),
    });
    const j = await r.json();
    if (!r.ok) throw new Error(j.error_description || j.error || r.status);
    res.writeHead(200, { "Content-Type": "text/plain" }).end("Done — you can close this tab and return to the terminal.");

    const stored = await storeToken(j.refresh_token);
    if (stored) {
      console.log("\nRefresh token stored in App Service (OPB_REFRESH_TOKEN). Scanning + Guest List are now live.\n");
    } else {
      console.log("\nCould not store via az. Set this value as OPB_REFRESH_TOKEN manually:\n");
      console.log(j.refresh_token + "\n");
    }
  } catch (e) {
    res.writeHead(500).end("token exchange failed: " + (e?.message || e));
    console.error(e);
  } finally {
    setTimeout(() => server.close(() => process.exit(0)), 500);
  }
});

server.listen(53682, () => {
  console.log("Opening sign-in in your browser - sign in as the workbook OWNER account.");
  console.log("If it doesn't open, paste this URL:\n" + authUrl + "\n");
  const cmd = process.platform === "win32" ? `start "" "${authUrl}"`
    : process.platform === "darwin" ? `open "${authUrl}"` : `xdg-open "${authUrl}"`;
  exec(cmd);
});
