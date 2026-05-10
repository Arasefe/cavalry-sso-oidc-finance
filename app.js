require("dotenv").config();
const express = require("express");
const session = require("express-session");
const msal = require("@azure/msal-node");

const app = express();
const PORT = process.env.PORT || 4001;

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
  system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false, logLevel: msal.LogLevel.Warning } },
};

const pca = new msal.ConfidentialClientApplication(msalConfig);
const cryptoProvider = new msal.CryptoProvider();

app.use(session({
  secret: process.env.SESSION_SECRET || "finance-secret",
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, maxAge: 8 * 60 * 60 * 1000 },
}));

function isAuthenticated(req, res, next) {
  if (req.session.account) return next();
  res.redirect("/");
}

function decodeToken(token) {
  try { return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8")); }
  catch { return {}; }
}

const C = { primary: "#1b3a6b", accent: "#2563eb", red: "#dc2626" };

const head = `
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet"/>
  <style>
    body { background:#0a0f1e; color:#e2e8f0; font-family:'Segoe UI',sans-serif; }
    .card { background:#0f1629; border:1px solid #1e3a5f; border-radius:12px; }
    .claim-row:hover { background:#131d35; }
    pre { white-space:pre-wrap; word-break:break-all; }
    .badge { display:inline-block; padding:2px 10px; border-radius:9999px; font-size:11px; font-weight:600; }
  </style>`;

const nav = (user) => `
  <nav style="background:#060d1a;border-bottom:1px solid #1e3a5f;" class="px-8 py-4 flex items-center justify-between">
    <div class="flex items-center gap-3">
      <div style="background:${C.accent};border-radius:8px;" class="p-2">
        <span style="font-size:18px;">💰</span>
      </div>
      <div>
        <div class="font-bold text-white text-sm">Finance Portal</div>
        <div style="color:#4b7fb5;font-size:11px;">SSO · OIDC · Microsoft Entra ID</div>
      </div>
    </div>
    <div class="flex items-center gap-4">
      ${user
        ? `<span style="color:#9ca3af;font-size:13px;">Signed in as <strong style="color:#60a5fa;">${user.name || user.username}</strong></span>
           <a href="/dashboard" style="background:#1e3a5f;" class="px-4 py-2 rounded-lg text-sm text-white">Dashboard</a>
           <a href="/auth/logout" style="background:${C.red};" class="px-4 py-2 rounded-lg text-sm text-white font-semibold">Sign Out</a>`
        : `<a href="/auth/login" style="background:${C.accent};" class="px-5 py-2 rounded-lg text-sm font-semibold text-white">Sign In with Microsoft</a>`
      }
    </div>
  </nav>`;

// ── Home ──────────────────────────────────────────────────────────────────────
app.get("/", (req, res) => {
  const user = req.session.account;
  if (user) return res.redirect("/dashboard");
  res.send(`<!DOCTYPE html><html><head>${head}<title>Finance — SSO OIDC</title></head><body>
    ${nav(null)}
    <div class="max-w-3xl mx-auto px-6 py-16 text-center space-y-8">
      <div>
        <div style="font-size:64px;">💰</div>
        <h1 class="text-3xl font-bold text-white mt-4">Finance Portal</h1>
        <p style="color:#60a5fa;" class="text-lg font-semibold mt-1">cavalry-sso-oidc-finance</p>
        <p style="color:#9ca3af;" class="mt-3 text-sm max-w-xl mx-auto">
          Demonstrates OIDC Single Sign-On via Microsoft Entra ID. Sign in once — your identity
          token carries your claims into this app without a separate password.
        </p>
      </div>
      <div class="card p-6 text-left space-y-3">
        <h3 class="font-bold text-white text-sm">This app teaches:</h3>
        ${[
          ["🔑", "ID Token claims", "What Entra ID knows about you and sends to this app"],
          ["🔐", "OIDC Authorization Code + PKCE flow", "How the secure redirect-based login works"],
          ["📋", "JWT structure", "Header · Payload · Signature — decoded and explained"],
          ["✅", "SSO across apps", "One Entra ID session, multiple independent apps"],
        ].map(([icon, title, desc]) => `
          <div class="flex gap-3 items-start">
            <span>${icon}</span>
            <div><div class="text-sm font-semibold text-white">${title}</div>
            <div style="color:#9ca3af;font-size:12px;">${desc}</div></div>
          </div>`).join("")}
      </div>
      <a href="/auth/login" style="background:linear-gradient(135deg,${C.accent},#1d4ed8);"
         class="inline-flex items-center gap-3 px-8 py-4 rounded-xl text-white font-semibold text-base hover:opacity-90 transition shadow-lg">
        <svg width="20" height="20" viewBox="0 0 21 21" fill="none">
          <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
          <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
          <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
          <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
        </svg>
        Sign In with Microsoft Entra ID
      </a>
      <p style="color:#374151;font-size:12px;">Port 4001 · OIDC · cavalry-sso-oidc-finance</p>
    </div>
  </body></html>`);
});

// ── Auth ──────────────────────────────────────────────────────────────────────
app.get("/auth/login", async (req, res) => {
  try {
    const { verifier, challenge } = await cryptoProvider.generatePkceCodes();
    req.session.pkceCodes = { verifier, challenge };
    req.session.signedInAt = new Date().toISOString();
    const authUrl = await pca.getAuthCodeUrl({
      scopes: ["openid", "profile", "email", "User.Read"],
      redirectUri: process.env.REDIRECT_URI,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
    });
    res.redirect(authUrl);
  } catch (err) {
    res.send(errorPage("Login Failed", err.message));
  }
});

app.get("/auth/callback", async (req, res) => {
  try {
    const tokenResponse = await pca.acquireTokenByCode({
      code: req.query.code,
      scopes: ["openid", "profile", "email", "User.Read"],
      redirectUri: process.env.REDIRECT_URI,
      codeVerifier: req.session.pkceCodes?.verifier,
    });
    req.session.account = tokenResponse.account;
    req.session.idToken = tokenResponse.idToken;
    req.session.accessToken = tokenResponse.accessToken;
    res.redirect("/dashboard");
  } catch (err) {
    res.send(errorPage("Authentication Failed", err.message));
  }
});

app.get("/auth/logout", (req, res) => {
  req.session.destroy();
  res.redirect(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(`http://localhost:${PORT}`)}`);
});

// ── Dashboard ─────────────────────────────────────────────────────────────────
app.get("/dashboard", isAuthenticated, (req, res) => {
  const account = req.session.account;
  const idClaims = req.session.idToken ? decodeToken(req.session.idToken) : {};
  const signedInAt = req.session.signedInAt
    ? new Date(req.session.signedInAt).toLocaleString("en-GB", { dateStyle: "medium", timeStyle: "medium" })
    : "—";

  const claimDesc = {
    iss: "Issuer — Entra ID endpoint that issued this token",
    sub: "Subject — unique stable user identifier",
    aud: "Audience — Client ID of the app this token is for",
    exp: "Expiry — token is invalid after this Unix timestamp",
    iat: "Issued At — when this token was created",
    nbf: "Not Before — token is invalid before this Unix timestamp",
    name: "Display name of the signed-in user",
    preferred_username: "UPN / email of the signed-in user",
    oid: "Object ID — unique user ID in the Entra ID directory",
    tid: "Tenant ID — identifies your Entra ID organisation",
    sid: "Session ID — ties this token to an Entra ID browser session",
    nonce: "Replay-attack prevention nonce",
    ver: "Token version (1.0 or 2.0)",
    email: "User's email address",
  };

  const fmtVal = (k, v) => {
    if (["exp", "iat", "nbf"].includes(k))
      return `${v} <span style="color:#64748b;">(${new Date(v * 1000).toLocaleString()})</span>`;
    if (k === "aud") return `${v} <span style="background:#052e16;color:#4ade80;padding:1px 6px;border-radius:4px;font-size:10px;">✓ your Client ID</span>`;
    if (k === "tid") return `${v} <span style="background:#052e16;color:#4ade80;padding:1px 6px;border-radius:4px;font-size:10px;">✓ your Tenant ID</span>`;
    return String(v);
  };

  res.send(`<!DOCTYPE html><html><head>${head}<title>Finance — Dashboard</title></head><body>
    ${nav(account)}

    <!-- Success Banner -->
    <div style="background:linear-gradient(90deg,#0f2044,#1b3a6b);border-bottom:1px solid #1e3a5f;" class="px-8 py-6 flex items-center gap-5">
      <div style="font-size:40px;">✅</div>
      <div>
        <h2 class="text-xl font-bold text-white">Sign In Successful</h2>
        <p style="color:#93c5fd;" class="text-sm">Welcome, <strong>${account.name || account.username}</strong> — authenticated via Microsoft Entra ID at ${signedInAt}</p>
        <p style="color:#4b7fb5;font-size:12px;" class="mt-1">OIDC Authorization Code + PKCE flow complete · ID Token issued and verified</p>
      </div>
      <div class="ml-auto">
        <a href="/auth/logout" style="background:${C.red};" class="px-5 py-2 rounded-lg text-sm text-white font-semibold">Sign Out</a>
      </div>
    </div>

    <div class="max-w-4xl mx-auto px-6 py-8 space-y-8">

      <!-- Info cards -->
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;">
        ${[
          ["User", account.name || "—"],
          ["Email", account.username || "—"],
          ["Signed In", signedInAt],
          ["Protocol", "OIDC / OAuth 2.0"],
        ].map(([l, v]) => `
          <div class="card p-4" style="border-top:3px solid ${C.accent};">
            <div style="color:#64748b;font-size:11px;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px;">${l}</div>
            <div class="text-sm font-semibold text-white" style="word-break:break-all;">${v}</div>
          </div>`).join("")}
      </div>

      <!-- ID Token Claims -->
      <div class="card overflow-hidden">
        <div style="border-bottom:1px solid #1e3a5f;" class="px-6 py-4 flex justify-between items-center">
          <h3 class="font-bold text-white">ID Token Claims</h3>
          <span style="color:#64748b;font-size:12px;">${Object.keys(idClaims).length} claims</span>
        </div>
        <div class="divide-y" style="border-color:#1e3a5f;">
          ${Object.entries(idClaims).map(([k, v], i) => `
            <div class="claim-row px-6 py-3" style="display:grid;grid-template-columns:150px 1fr 1fr;gap:12px;">
              <code style="color:#f59e0b;font-size:13px;">${k}</code>
              <span style="font-size:12px;font-family:monospace;color:#e2e8f0;word-break:break-all;">${fmtVal(k, v)}</span>
              <span style="color:#64748b;font-size:11px;">${claimDesc[k] || ""}</span>
            </div>`).join("")}
        </div>
      </div>

      <!-- Raw JWT -->
      <div class="card p-5 space-y-3">
        <h3 class="font-bold text-white text-sm">Raw ID Token (JWT)</h3>
        <p style="color:#9ca3af;font-size:12px;">Three base64-encoded parts separated by dots. The signature proves Entra ID issued this token.</p>
        <div style="background:#060d1a;border-radius:8px;padding:12px;font-size:11px;font-family:monospace;line-height:1.6;overflow-x:auto;">
          ${req.session.idToken
            ? req.session.idToken.split(".").map((p, i) => `<span style="color:${["#60a5fa","#4ade80","#f87171"][i]}">${p}</span>`).join('<span style="color:#374151;">.</span>')
            : '<span style="color:#374151;">Token not available</span>'}
        </div>
        <div class="flex gap-6 text-xs" style="color:#64748b;">
          <span><span style="color:#60a5fa;">■</span> Header (algorithm)</span>
          <span><span style="color:#4ade80;">■</span> Payload (claims)</span>
          <span><span style="color:#f87171;">■</span> Signature (Entra ID's seal)</span>
        </div>
      </div>

      <div class="text-center pb-8">
        <a href="/auth/logout" style="background:#1f2937;" class="inline-flex items-center gap-2 px-6 py-3 rounded-lg text-sm text-white hover:opacity-90 transition">
          Sign Out &amp; Clear Session
        </a>
      </div>
    </div>
  </body></html>`);
});

function errorPage(title, msg) {
  return `<!DOCTYPE html><html><head>${head}<title>Error</title></head><body>
    <div class="max-w-xl mx-auto px-6 py-20 text-center space-y-4">
      <div style="font-size:48px;">⚠️</div>
      <h1 class="text-2xl font-bold text-white">${title}</h1>
      <div style="background:#1a1f2e;border:1px solid #7f1d1d;border-radius:8px;" class="p-4 text-left">
        <code style="color:#fca5a5;font-size:12px;">${msg}</code>
      </div>
      <a href="/" style="color:#60a5fa;" class="text-sm underline">← Back to Home</a>
    </div></body></html>`;
}

app.listen(PORT, () => {
  console.log(`cavalry-sso-oidc-finance  →  http://localhost:${PORT}`);
  console.log(`Redirect URI:                http://localhost:${PORT}/auth/callback`);
});
