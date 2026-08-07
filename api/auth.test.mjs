import { test } from "node:test";
import assert from "node:assert/strict";
import { generateKeyPair, SignJWT, exportJWK } from "jose";
import { verifyProviderToken, resolveRole, issueSession, verifySession } from "./auth.js";

test("resolveRole: admin > user > rejected", () => {
  const env = { ADMIN_EMAILS: "Boss@opb.no", USER_EMAILS: "vol1@gmail.com, vol2@live.com" };
  assert.equal(resolveRole("boss@opb.no", env), "admin");
  assert.equal(resolveRole("VOL1@gmail.com", env), "user");
  assert.equal(resolveRole("stranger@gmail.com", env), null);
  assert.equal(resolveRole("", env), null);
});

test("session issue + verify round-trip", async () => {
  process.env.SESSION_SECRET = "test-secret-please-rotate";
  const tok = await issueSession({ email: "a@b.com", name: "Aisha", role: "admin" });
  const s = await verifySession(tok);
  assert.equal(s.email, "a@b.com");
  assert.equal(s.role, "admin");
  assert.equal(s.name, "Aisha");
});

test("verifySession rejects a tampered token", async () => {
  process.env.SESSION_SECRET = "test-secret-please-rotate";
  const tok = await issueSession({ email: "a@b.com", name: "A", role: "user" });
  await assert.rejects(() => verifySession(tok.slice(0, -3) + "xyz"));
});

// Build a fake provider (local keypair as the JWKS) to exercise the ID-token verifier.
async function fakeProvider(issuer) {
  const { publicKey, privateKey } = await generateKeyPair("RS256");
  const deps = { google: { issuers: [issuer], jwks: publicKey, audEnv: "GOOGLE_CLIENT_ID" } };
  const mint = (claims) =>
    new SignJWT(claims).setProtectedHeader({ alg: "RS256" }).setIssuer(issuer)
      .setAudience(process.env.GOOGLE_CLIENT_ID).setIssuedAt().setExpirationTime("5m").sign(privateKey);
  return { deps, mint };
}

test("verifyProviderToken accepts a valid Google token", async () => {
  process.env.GOOGLE_CLIENT_ID = "client-123.apps.googleusercontent.com";
  const { deps, mint } = await fakeProvider("https://accounts.google.com");
  const token = await mint({ email: "Vol1@Gmail.com", email_verified: true, name: "Vol One" });
  const r = await verifyProviderToken("google", token, deps);
  assert.equal(r.email, "vol1@gmail.com");
  assert.equal(r.name, "Vol One");
});

test("verifyProviderToken rejects wrong audience (token minted for another app)", async () => {
  process.env.GOOGLE_CLIENT_ID = "client-123.apps.googleusercontent.com";
  const { deps } = await fakeProvider("https://accounts.google.com");
  const { privateKey } = await generateKeyPair("RS256");
  const badAudDeps = { google: { issuers: ["https://accounts.google.com"], jwks: (await generateKeyPair("RS256")).publicKey, audEnv: "GOOGLE_CLIENT_ID" } };
  const token = await new SignJWT({ email: "x@gmail.com", email_verified: true })
    .setProtectedHeader({ alg: "RS256" }).setIssuer("https://accounts.google.com")
    .setAudience("some-other-app").setIssuedAt().setExpirationTime("5m").sign(privateKey);
  await assert.rejects(() => verifyProviderToken("google", token, badAudDeps));
});

test("verifyProviderToken rejects email_verified=false", async () => {
  process.env.GOOGLE_CLIENT_ID = "client-123.apps.googleusercontent.com";
  const { deps, mint } = await fakeProvider("https://accounts.google.com");
  const token = await mint({ email: "x@gmail.com", email_verified: false });
  await assert.rejects(() => verifyProviderToken("google", token, deps));
});
