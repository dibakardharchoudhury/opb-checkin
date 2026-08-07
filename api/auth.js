// Social sign-in (Google + Microsoft personal/Live) for the OPB check-in app.
//
// Security model:
//  - The browser signs in with Google or Microsoft (which enforce MFA/passkeys) and
//    gets a signed ID token (JWT).
//  - We VERIFY that token here server-side: provider signature (JWKS), issuer, audience
//    (== our OAuth client id, so a token minted for another app can't be replayed),
//    expiry, and email_verified.
//  - Social identity alone is not permission: the verified email must be on an ALLOWLIST
//    (ADMIN_EMAILS / USER_EMAILS in app settings). Everyone else is rejected.
//  - We then issue our own short-lived session JWT (HS256, SESSION_SECRET) carrying the
//    role; the SPA sends it as a Bearer token on each API call.

import { createRemoteJWKSet, jwtVerify, SignJWT } from "jose";
import { listStoredUsers } from "./userstore.js";

// Provider config. jwks getters are lazy (no network until first verify). Injectable
// for tests via the `deps` argument to verifyProviderToken.
export const PROVIDERS = {
  google: {
    issuers: ["https://accounts.google.com", "accounts.google.com"],
    jwks: createRemoteJWKSet(new URL("https://www.googleapis.com/oauth2/v3/certs")),
    audEnv: "GOOGLE_CLIENT_ID",
  },
  microsoft: {
    // personal Microsoft accounts (consumers) tenant
    issuers: ["https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0"],
    jwks: createRemoteJWKSet(new URL("https://login.microsoftonline.com/consumers/discovery/v2.0/keys")),
    audEnv: "MS_CLIENT_ID",
  },
};

export async function verifyProviderToken(provider, token, deps = PROVIDERS) {
  const p = deps[provider];
  if (!p) throw new Error("unknown provider");
  const aud = process.env[p.audEnv];
  if (!aud) throw new Error(`${p.audEnv} not configured`);
  const { payload } = await jwtVerify(token, p.jwks, { issuer: p.issuers, audience: aud });
  const email = String(payload.email || payload.preferred_username || "").toLowerCase();
  if (!email) throw new Error("token has no email");
  // Google returns a reliable email_verified boolean; reject if explicitly false.
  if (provider === "google" && payload.email_verified === false) throw new Error("email not verified");
  return { email, name: payload.name || email };
}

const listEmails = (s) => String(s || "").split(/[,;\s]+/).map((x) => x.trim().toLowerCase()).filter(Boolean);

// admin > user; unlisted => null (rejected).
export function resolveRole(email, env = process.env) {
  const e = String(email || "").toLowerCase();
  if (listEmails(env.ADMIN_EMAILS).includes(e)) return "admin";
  if (listEmails(env.USER_EMAILS).includes(e)) return "user";
  return null;
}

// Role resolution used at login: bootstrap admins from config, then the in-app
// user store, then any USER_EMAILS fallback. Config admins always win (no lockout).
export async function resolveRoleMerged(email) {
  const e = String(email || "").toLowerCase();
  if (listEmails(process.env.ADMIN_EMAILS).includes(e)) return "admin";
  const stored = await listStoredUsers();
  const hit = stored.find((u) => u.email === e);
  if (hit) return hit.role === "admin" ? "admin" : "user";
  if (listEmails(process.env.USER_EMAILS).includes(e)) return "user";
  return null;
}

const SESSION_TTL_SECONDS = 12 * 3600;
const secretKey = () => new TextEncoder().encode(process.env.SESSION_SECRET || "dev-only-insecure-secret");

export async function issueSession({ email, name, role }) {
  return new SignJWT({ name, role })
    .setProtectedHeader({ alg: "HS256" })
    .setSubject(email).setIssuer("opb-checkin").setIssuedAt()
    .setExpirationTime(`${SESSION_TTL_SECONDS}s`)
    .sign(secretKey());
}

export async function verifySession(token) {
  const { payload } = await jwtVerify(token, secretKey(), { issuer: "opb-checkin" });
  return { email: payload.sub, name: payload.name, role: payload.role };
}

// Express middleware. requireAuth() = any signed-in allowlisted user; requireAuth("admin") = admin only.
export function requireAuth(role) {
  return async (req, res, next) => {
    try {
      const m = (req.headers.authorization || "").match(/^Bearer (.+)$/i);
      if (!m) return res.status(401).json({ error: "Please sign in." });
      const user = await verifySession(m[1]);
      if (role === "admin" && user.role !== "admin") return res.status(403).json({ error: "Admin access only." });
      req.user = user;
      next();
    } catch {
      return res.status(401).json({ error: "Session expired — sign in again." });
    }
  };
}
