# cavalry-sso-oidc-finance

A Microsoft Entra ID OIDC SSO demo app themed around a Finance Portal. Part of the Cavalry SSO demo suite.

**Port:** `4001` · **Protocol:** OIDC (OpenID Connect) · **Library:** `@azure/msal-node`

---

## Teaching Focus

This app teaches the **OIDC Authorization Code + PKCE flow** and **ID Token claims**.

Students will see:
- The full SP-initiated OIDC login redirect to `login.microsoftonline.com`
- The decoded ID Token claims returned by Entra ID after authentication
- The raw JWT structure — header · payload · signature — colour-coded and explained
- How the `aud` claim matches the app's Client ID and the `tid` claim matches the Tenant ID

---

## Prerequisites

- Node.js 18+
- A Microsoft Entra ID tenant
- An App Registration in Entra ID (see setup below)

---

## Entra ID App Registration Setup

1. Go to **Entra ID → App Registrations → New registration**
2. Name: `cavalry-sso-oidc-finance`
3. Supported account types: **Accounts in this organizational directory only**
4. Platform: **Web** — Redirect URI: `http://localhost:4001/auth/callback`
5. Click **Register**
6. Note the **Application (Client) ID** and **Directory (Tenant) ID** from Overview
7. Go to **Certificates & secrets → New client secret** — copy the secret value immediately

---

## Setup

```bash
git clone https://github.com/Arasefe/cavalry-sso-oidc-finance.git
cd cavalry-sso-oidc-finance
npm install
cp .env.example .env
```

Edit `.env`:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
REDIRECT_URI=http://localhost:4001/auth/callback
SESSION_SECRET=any-long-random-string
PORT=4001
```

---

## Run

```bash
npm start
```

Navigate to `http://localhost:4001` and click **Sign in with Microsoft Entra ID**.

---

## Endpoints

| Endpoint | Description |
|---|---|
| `GET /` | Landing page with sign-in button |
| `GET /auth/login` | Initiates OIDC login — redirects to Entra ID |
| `GET /auth/callback` | Receives authorization code, exchanges for tokens |
| `GET /dashboard` | Authenticated view — shows ID token claims and raw JWT |
| `GET /auth/logout` | Destroys local session and redirects to Entra ID logout |

---

## SSO Demo Sequence

Run all three apps together to demonstrate SSO session reuse:

```bash
# Tab 1 — Finance (sign in here first)
cd ~/Desktop/cavalry-sso-oidc-finance && npm start

# Tab 2 — HR (authenticates silently after Finance)
cd ~/Desktop/cavalry-sso-oidc-hr && npm start

# Tab 3 — Marketing (authenticates silently after Finance)
cd ~/Desktop/cavalry-sso-oidc-marketing && npm start
```

Sign into Finance → open HR in a new tab → open Marketing in a new tab.
Compare the `sid` claim across all three dashboards — it will be identical, proving one Entra ID session served all three apps without re-entering credentials.

---

## Part of the Cavalry Demo Suite

| App | Port | Protocol | Focus |
|---|---|---|---|
| `cavalry-sso-oidc-finance` | 4001 | OIDC | ID Token claims & JWT structure |
| `cavalry-sso-oidc-hr` | 4002 | OIDC | SSO session reuse & `sid` claim |
| `cavalry-sso-oidc-marketing` | 4003 | OIDC | Scopes, access token vs ID token |
| `cavalry-sso-saml-hr` | 3001 | SAML | SAML claims table |
| `cavalry-sso-saml-devtools` | 3002 | SAML | Raw SAML assertion payload |
| `cavalry-sso-saml-finance` | 3003 | SAML | User assignment access control |
| `cavalry-sso-saml-portal` | 3004 | SAML | SP-initiated SAML flow walkthrough |
| `cavalry-sso-saml-operations` | 3005 | SAML | Session metadata & continuity |
