# Main Sequence — Excel Add-in Skeleton (Spec)

the main endpoint is
https://dev-tsorm.ngrok.app

**Objective:** Build the skeleton and base of a project that will become the Excel add-in for **Main Sequence**.
Main Sequence is a finance and data app that allows clients to unify and consume diverse 
sources of data and perform asset management operations.

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

- The request should be done to :
- 'http://ROOT/orm/api/ts_manager/dynamic_table/714/get_data_between_dates_from_remote/'
with a json payload

```
"json": {
                        "start_date": start_date.timestamp() if start_date else None,
                        "end_date": end_date.timestamp() if end_date else None,
                        "great_or_equal": great_or_equal,
                        "less_or_equal": less_or_equal,
                        "unique_identifier_list": unique_identifier_list,
                        "columns": columns,
                        "offset": offset,  # pagination offset
                       
                    }
```

the return is

```typescript
type ApiResponse<T = unknown> = {
  results: T[];
  limit: number;
  offset: number;
  returned_count: number;
  next_offset: number | null;
};
```
example return 
```json
{
  "results": [
    { "col_a": "value", "col_b": 123 },
    { "col_a": "value2", "col_b": 456 }
  ],
  "limit": 100,
  "offset": 0,
  "returned_count": 2,
  "next_offset": 100
}
```

This is the LLM translation  of our python client to type script for you to take as a guideline

```typescript
// Types you can tweak to your domain
export type TimestampSeconds = number;

export interface DateRangeDescriptor {
  start_date?: Date | TimestampSeconds;
  end_date?: Date | TimestampSeconds;
  // Allow arbitrary extra fields the server might accept
  [key: string]: unknown;
}

export type UniqueIdentifierRangeMap = Record<string, DateRangeDescriptor>;

export interface GetDataBetweenDatesParams {
  start_date?: Date | TimestampSeconds | null;
  end_date?: Date | TimestampSeconds | null;
  great_or_equal?: boolean | null;
  less_or_equal?: boolean | null;
  unique_identifier_list?: Array<string | number> | null;
  columns?: string[] | null;
  unique_identifier_range_map?: UniqueIdentifierRangeMap | null;
  // Present in the Python signature but not used in the payload there either:
  column_range_descriptor?: UniqueIdentifierRangeMap | null;
}

export interface ApiResponse<T = unknown> {
  results?: T[];
  next_offset?: number | null;
  // passthrough for any other fields
  [key: string]: unknown;
}

export interface RequestOptions {
  /** Base object URL, e.g. https://api.example.com/my-object */
  objectBaseUrl: string;
  /** The object id used in the Python f"/{self.id}/..." */
  id: string | number;
  /** Optional fetch to inject (Node <18 or custom). Defaults to global fetch. */
  fetchImpl?: typeof fetch;
  /** Extra headers to include on every POST. */
  headers?: Record<string, string>;
  /** Size of chunks for unique_identifier_range_map keys. Default: 100 */
  chunkSize?: number;
}

/**
 * Helper that mirrors the Python behavior:
 * - POSTs JSON including an "offset" for pagination
 * - Follows "next_offset" until it's null/undefined
 * - If unique_identifier_range_map is present, splits into 100-key chunks and fetches per chunk
 * - Converts Date objects to UNIX seconds
 */
export async function getDataBetweenDatesFromApi<T = unknown>(
  params: GetDataBetweenDatesParams,
  opts: RequestOptions
): Promise<T[]> {
  const {
    objectBaseUrl,
    id,
    fetchImpl = fetch,
    headers = {},
    chunkSize = 100,
  } = opts;

  const url = `${objectBaseUrl}/${id}/get_data_between_dates_from_remote/`;

  const toUnixSeconds = (d?: Date | number | null): number | null | undefined => {
    if (d == null) return d as null | undefined;
    if (d instanceof Date) return Math.floor(d.getTime() / 1000);
    // assume it's already seconds if a number
    return d;
  };

  // Shallow clone that preserves Date instances on values (good enough for the known fields).
  const cloneRangeMap = (
    map?: UniqueIdentifierRangeMap | null
  ): UniqueIdentifierRangeMap | undefined => {
    if (!map) return undefined;
    const out: UniqueIdentifierRangeMap = {};
    for (const [k, v] of Object.entries(map)) out[k] = { ...v };
    return out;
  };

  // Normalize top-level dates to seconds (Python: .timestamp())
  const startSeconds = toUnixSeconds(params.start_date);
  const endSeconds = toUnixSeconds(params.end_date);

  // Prepare (cloned) range map and convert any Date fields inside it to seconds
  const rangeMap = cloneRangeMap(params.unique_identifier_range_map);
  if (rangeMap) {
    for (const dateInfo of Object.values(rangeMap)) {
      if ("start_date" in dateInfo && dateInfo.start_date instanceof Date) {
        dateInfo.start_date = Math.floor(dateInfo.start_date.getTime() / 1000);
      }
      if ("end_date" in dateInfo && dateInfo.end_date instanceof Date) {
        dateInfo.end_date = Math.floor(dateInfo.end_date.getTime() / 1000);
      }
    }
  }

  // Inner helper: fetches one (possibly multi-offset) batch for a given chunked range map
  const fetchOneBatch = async (
    chunkRangeMap: UniqueIdentifierRangeMap | null | undefined
  ): Promise<T[]> => {
    const allResults: T[] = [];
    let offset = 0;

    // Follow next_offset just like the Python while True loop
    // Break when next_offset is null/undefined
    for (;;) {
      const body = {
        start_date: startSeconds ?? null,
        end_date: endSeconds ?? null,
        great_or_equal: params.great_or_equal ?? null,
        less_or_equal: params.less_or_equal ?? null,
        unique_identifier_list: params.unique_identifier_list ?? null,
        columns: params.columns ?? null,
        offset, // pagination offset (important!)
        unique_identifier_range_map: chunkRangeMap ?? null,
        // NOTE: The Python version does NOT send column_range_descriptor in the payload.
        // To stay faithful, we keep it omitted here as well.
      };

      const resp = await fetchImpl(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...headers,
        },
        body: JSON.stringify(body),
      });

      if (!resp.ok) {
        // Mirror the Python "warning then return []" behavior
        const text = await resp.text().catch(() => "");
        console.warn(`Error in request: ${text || resp.statusText}`);
        return [];
      }

      const data = (await resp.json()) as ApiResponse<T>;
      const chunk = Array.isArray(data.results) ? data.results : [];
      allResults.push(...chunk);

      const next = (data as ApiResponse<T>).next_offset;
      if (next === null || next === undefined) break;

      offset = next;
    }

    return allResults;
  };

  const allResults: T[] = [];

  // Python truthiness: only chunk if map exists AND has keys
  const hasRangeMap =
    !!rangeMap && Object.keys(rangeMap).length > 0;

  if (hasRangeMap) {
    const keys = Object.keys(rangeMap!);
    for (let i = 0; i < keys.length; i += chunkSize) {
      const slice = keys.slice(i, i + chunkSize);
      const chunkMap: UniqueIdentifierRangeMap = {};
      for (const k of slice) chunkMap[k] = rangeMap![k];

      const chunkResults = await fetchOneBatch(chunkMap);
      allResults.push(...chunkResults);
    }
  } else {
    // Single batch with offset-based pagination only
    const chunkResults = await fetchOneBatch(null);
    allResults.push(...chunkResults);
  }

  return allResults;
}

```

The final function should be call =GET_DATA_FROM_NODE(date_start,date_end,unique_identifier_list,great_or_equal,
less_or_equal) where each argument is a cell 