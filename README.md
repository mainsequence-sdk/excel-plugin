# Main Sequence — Excel Add-in Skeleton (Spec)

**Objective:** Build the skeleton and base of a project that will become the Excel add-in for **Main Sequence**. Main Sequence is a finance and data app that allows clients to unify and consume diverse sources of data and perform asset management operations.

---

## Project Details

- For this task, use the `excel-addin` folder in this repository.
- The project should be properly documented, commented, and follow coding standards so it can be edited and maintained by the Main Sequence team if necessary.
- Keep an ordered folder structure that is **DRY** and **extendable**.

---

## Tasks

### Task 1: Implement Sign-In and Data Retrieval (Authentication)

- Implement sign-in to the platform.
- The Main Sequence platform requires users to sign in with a **username** and **password**.
- For this, you can check the schematics of the CLI below.
- When the user signs-in keep its info so can be displayed somewhere
- Build some debug/console button when if something fails user can see the full error of the response

# Auth workflow (email/password → JWT)

1) **Login (obtain tokens)**

- **Endpoint:** `POST /auth/jwt-token/token/`
- **Body (JSON):**
  ```json
  { "email": "<user@domain>", "password": "<secret>" }
  ```
  *(Server expects the field name `email`, not `username`.)*
- **Success (JSON):** returns **access** and **refresh** JWTs. The server may use different keys; normalize like:
  - Access token in: `access | token | jwt | access_token`
  - Refresh token in: `refresh | refresh_token`
- **Client action:** store `{ username: email, access, refresh }` in your token storage (sessionStorage/localStorage/memory).

2) **Use token for protected calls**

- **Header:** `Authorization: Bearer <access>`
- **Allowed paths:** only `/api/*`, `/auth/*`, `/pods/*`, `/orm/*`, `/user/*`  
  *(Client should reject anything else—mirrors `_normalize_api_path`.)*

3) **Handle expiration (single automatic refresh)**

- If a request returns **401**, the client:
  - Calls `POST /auth/jwt-token/token/refresh/` with `{ "refresh": "<refresh>" }`.
  - On success, server returns a new access token (key = `access`).
  - Client updates stored access token and **retries the original request once**.
- If the retry still returns **401**, treat the session as logged out and prompt re-login.

4) **Verify current user** *(optional, like the CLI)*

- **Ping:** `GET /auth/rest-auth/user/` → JSON with one of `id | pk | user.id | user_id`.
- **Full profile:** `GET /user/api/user/{id}/` → includes `username` and `organization.name` (or `organization_name`).

5) **Logout**

- Clear stored tokens. User must login again to continue.

**CLI schematic (reference):**
 **Only implement Login for the moment ** 
```typescript
// auth.ts
// Minimal, dependency-free client that follows mainsequence/cli/api.py semantics.

type HttpMethod = 'GET' | 'HEAD' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';

export class ApiError extends Error {}
export class NotLoggedIn extends ApiError {}

export interface LoginRequest {
  /** Server expects the field name 'email' (not 'username') */
  email: string;
  password: string;
}

export interface TokenResponseNormalized {
  access: string;
  refresh: string;
}

export interface TokensPersisted extends TokenResponseNormalized {
  username: string; // email used to log in
}

export interface AuthResult {
  username: string;
  backend: string;
}

export interface TokenStorage {
  get(): TokensPersisted | null;
  set(t: TokensPersisted): void;
  clear(): void;
}

const STORAGE_KEY = 'MAIN_SEQUENCE_TOKENS';

class MemoryTokenStorage implements TokenStorage {
  private v: TokensPersisted | null = null;
  get() { return this.v; }
  set(t: TokensPersisted) { this.v = t; }
  clear() { this.v = null; }
}

class WebStorage implements TokenStorage {
  constructor(private area: Storage) {}
  get(): TokensPersisted | null {
    try {
      const raw = this.area.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) as TokensPersisted : null;
    } catch { return null; }
  }
  set(t: TokensPersisted): void {
    this.area.setItem(STORAGE_KEY, JSON.stringify(t));
  }
  clear(): void { this.area.removeItem(STORAGE_KEY); }
}

function defaultStorage(): TokenStorage {
  // Prefer sessionStorage, then localStorage; fall back to memory (Node).
  try {
    if (typeof window !== 'undefined' && window.sessionStorage) {
      return new WebStorage(window.sessionStorage);
    }
  } catch {}
  try {
    if (typeof window !== 'undefined' && window.localStorage) {
      return new WebStorage(window.localStorage);
    }
  } catch {}
  return new MemoryTokenStorage();
}

function normalizePath(p: string): string {
  const path = '/' + (p || '').replace(/^\/+/, '');
  if (!/^\/(api|auth|pods|orm|user)(\/|$)/.test(path)) {
    throw new ApiError('Only /api/*, /auth/*, /pods/*, /orm/*, /user/* allowed');
  }
  return path;
}

function joinUrl(base: string, path: string): string {
  const b = base.replace(/\/+$/, '');
  const p = path.replace(/^\/+/, '');
  return `${b}/${p}`;
}

function isJson(res: Response): boolean {
  const ct = res.headers.get('content-type') || '';
  return ct.toLowerCase().startsWith('application/json');
}

async function readErrText(res: Response): Promise<string> {
  try {
    if (isJson(res)) {
      const j = await res.json();
      return j?.detail || j?.message || JSON.stringify(j);
    }
    return await res.text();
  } catch {
    return res.statusText || `HTTP ${res.status}`;
  }
}

function normalizeTokenResponse(raw: any): TokenResponseNormalized {
  const access = raw?.access ?? raw?.token ?? raw?.jwt ?? raw?.access_token;
  const refresh = raw?.refresh ?? raw?.refresh_token;
  if (!access || !refresh) {
    throw new ApiError('Server did not return expected tokens.');
  }
  return { access, refresh };
}

export interface AuthClientOptions {
  backend: string;            // e.g. https://api.example.com
  storage?: TokenStorage;     // optional custom storage
  fetchImpl?: typeof fetch;   // optional custom fetch
  defaultHeaders?: Record<string, string>;
}

export class AuthClient {
  private backend: string;
  private storage: TokenStorage;
  private fetchFn: typeof fetch;
  private defaultHeaders: Record<string, string>;
  private refreshInFlight: Promise<string> | null = null;

  constructor(opts: AuthClientOptions) {
    if (!opts?.backend) throw new Error('backend is required');
    this.backend = opts.backend;
    this.storage = opts.storage ?? defaultStorage();
    this.fetchFn = opts.fetchImpl ?? fetch;
    this.defaultHeaders = { 'Content-Type': 'application/json', ...(opts.defaultHeaders || {}) };
  }

  /** POST /auth/jwt-token/token/ with { email, password } */
  async login(email: string, password: string): Promise<AuthResult> {
    const url = joinUrl(this.backend, '/auth/jwt-token/token/');
    const res = await this.fetchFn(url, {
      method: 'POST',
      headers: this.defaultHeaders,
      body: JSON.stringify({ email: (email || '').trim(), password: (password || '').replace(/\r?\n$/, '') }),
    });

    if (!res.ok) {
      throw new ApiError(await readErrText(res));
    }

    const json = isJson(res) ? await res.json() : {};
    const { access, refresh } = normalizeTokenResponse(json);
    this.storage.set({ username: email, access, refresh });
    return { username: email, backend: this.backend };
  }

  /** POST /auth/jwt-token/token/refresh/ with { refresh } */
  async refreshAccess(): Promise<string> {
    const tokens = this.storage.get();
    if (!tokens?.refresh) {
      throw new NotLoggedIn('Not logged in. Please sign in.');
    }
    // Collapse concurrent refreshes into a single flight.
    if (this.refreshInFlight) return this.refreshInFlight;

    this.refreshInFlight = (async () => {
      const res = await this.fetchFn(joinUrl(this.backend, '/auth/jwt-token/token/refresh/'), {
        method: 'POST',
        headers: this.defaultHeaders,
        body: JSON.stringify({ refresh: tokens.refresh }),
      });

      if (!res.ok) {
        this.refreshInFlight = null;
        throw new NotLoggedIn(isJson(res) ? (await res.json())?.detail || 'Token refresh failed.' : 'Token refresh failed.');
      }

      const j = isJson(res) ? await res.json() : {};
      const access: string | undefined = j?.access;
      if (!access) {
        this.refreshInFlight = null;
        throw new NotLoggedIn('Refresh succeeded but no access token returned.');
      }
      this.storage.set({ username: tokens.username, access, refresh: tokens.refresh });
      this.refreshInFlight = null;
      return access;
    })();

    return this.refreshInFlight;
  }

  /** Low-level request with Bearer token + single automatic refresh on 401. */
  async authed(method: HttpMethod, apiPath: string, body?: unknown): Promise<Response> {
    const path = normalizePath(apiPath);
    const t = this.storage.get();
    let access = t?.access;
    if (!access) {
      // Try a single refresh if we have a refresh token.
      access = await this.refreshAccess();
    }

    const doReq = () => this.fetchFn(joinUrl(this.backend, path), {
      method,
      headers: { ...this.defaultHeaders, Authorization: `Bearer ${access}` },
      body: method === 'GET' || method === 'HEAD' ? undefined : JSON.stringify(body ?? {}),
    });

    let res = await doReq();
    if (res.status === 401) {
      access = await this.refreshAccess();
      res = await this.fetchFn(joinUrl(this.backend, path), {
        method,
        headers: { ...this.defaultHeaders, Authorization: `Bearer ${access}` },
        body: method === 'GET' || method === 'HEAD' ? undefined : JSON.stringify(body ?? {}),
      });
    }
    if (res.status === 401) {
      throw new NotLoggedIn('Not logged in.');
    }
    return res;
  }

  /** Mirrors get_current_user_profile() from the Python CLI. */
  async getCurrentUserProfile(): Promise<{ username: string; organization: string }> {
    const who = await this.authed('GET', '/auth/rest-auth/user/');
    const d = who.ok && isJson(who) ? await who.json() : {};
    const uid = d?.id ?? d?.pk ?? d?.user?.id ?? d?.user_id;
    if (!uid) return { username: '', organization: '' };
    const full = await this.authed('GET', `/user/api/user/${uid}/`);
    const u = full.ok && isJson(full) ? await full.json() : {};
    const orgName: string =
      u?.organization?.name || u?.organization_name || '';
    return { username: u?.username || '', organization: orgName };
  }

  /** Convenience helpers if you want the raw tokens or to log out. */
  getTokens(): TokensPersisted | null { return this.storage.get(); }
  logout() { this.storage.clear(); }
}

/* =======================
   Usage (browser or Node)
   =======================

import { AuthClient } from './auth';

const auth = new AuthClient({ backend: 'https://your-backend.example.com' });

async function signIn() {
  try {
    await auth.login('user@example.com', 'correct horse battery staple');
    // Optional: verify by pinging the user endpoint
    const profile = await auth.getCurrentUserProfile();
    console.log(profile);
  } catch (e) {
    if (e instanceof NotLoggedIn || e instanceof ApiError) {
      console.error(e.message);
    } else {
      console.error(e);
    }
  }
}

async function callApi() {
  const res = await auth.authed('POST', '/api/some/endpoint/', { foo: 'bar' });
  const data = await res.json();
  console.log(data);
}

*/

```


---

### Task 2: Data Retrieval (Post-Authentication)

- Once the user is authenticated, build a function that accepts `date_start` ...
