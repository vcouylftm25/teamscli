import { CliError } from '../errors.ts';

/**
 * Teams web services the CLI calls. Each has its own token (see docs/authentication.md).
 * `media` is the store of inline images; `sharepoint` the SharePoint or OneDrive site of an attached file.
 */
export type Service = 'chatService' | 'csa' | 'media' | 'sharepoint' | 'urlp';

/** Base URLs for the account's region, discovered at login (`authz`). */
export interface Endpoints {
  /** e.g. `https://<region>.ng.msg.teams.microsoft.com`; chats, channels, messages. */
  chatService: string;
  /** e.g. `https://teams.microsoft.com/api/csa/<region>`; teams and channels list. */
  csa: string;
  /**
   * URL preview service from authz `regionGtms.urlpV2`
   * (e.g. `https://us-prod.asyncgw.teams.microsoft.com/urlp`), else `regionGtms.urlp`.
   * Null when authz omits both.
   */
  urlp: string | null;
}

/** Implemented by auth/: valid credentials for each service, refreshed as needed. */
export interface Session {
  /** Headers that authenticate a request to `url`: `Authentication: skypetoken=…`, `Authorization: skype_token …` or `Authorization: Bearer …`. */
  headers(service: Service, url?: string): Promise<Record<string, string>>;
  endpoints(): Promise<Endpoints>;
}

export interface ClientDeps {
  fetch: typeof fetch;
  session: Session;
  sleep?: (ms: number) => Promise<void>;
  /** Retries after HTTP 429 before giving up with `rate_limited`. */
  maxRetries?: number;
  /** Maximum time for one request, including reading its body. Defaults to 60 seconds. */
  requestTimeoutMs?: number;
}

export interface RequestOptions {
  method?: 'GET' | 'POST' | 'PUT' | 'DELETE';
  /** Serialized as JSON. */
  body?: unknown;
  /** Sent as is, as `application/octet-stream`; takes the place of `body`. */
  data?: Uint8Array<ArrayBuffer>;
}

/** One page of a paged collection, with the raw links Teams returned for it. */
export interface Page<T> {
  items: T[];
  backwardLink: string | null;
  syncState: string | null;
}

export type TeamsClient = ReturnType<typeof createTeamsClient>;

export function createTeamsClient({ fetch, session, sleep = defaultSleep, maxRetries = 3, requestTimeoutMs = 60_000 }: ClientDeps) {
  /** Authenticated JSON request. Maps failures to CliError: 401 auth_required, 429 rate_limited (after retries), other remote. */
  async function json<T>(service: Service, url: string, options: RequestOptions = {}): Promise<T> {
    const text = new TextDecoder().decode((await request(service, url, options, 'application/json')).data);
    try {
      return (text ? JSON.parse(text) : null) as T;
    } catch (error) {
      throw new CliError('remote', `Teams returned invalid JSON (${options.method ?? 'GET'} ${new URL(url).pathname}).`, null, { cause: error });
    }
  }

  /**
   * Authenticated GET of a binary body, with the same failure mapping as `json`.
   * shortcut: buffers the whole body within the request timeout; stream to disk if large files time out.
   */
  async function bytes(service: Service, url: string): Promise<{ contentType: string | null; data: Uint8Array }> {
    return request(service, url, {}, '*/*');
  }

  async function request(service: Service, url: string, { method = 'GET', body, data }: RequestOptions, accept: string): Promise<{ contentType: string | null; data: Uint8Array }> {
    const where = `${method} ${new URL(url).pathname}`;
    for (let attempt = 0; ; attempt++) {
      const headers: Record<string, string> = { accept, ...(await session.headers(service, url)) };
      if (body !== undefined) headers['content-type'] = 'application/json';
      if (data !== undefined) headers['content-type'] = 'application/octet-stream';
      let response: Response;
      const signal = AbortSignal.timeout(requestTimeoutMs);
      try {
        response = await fetch(url, { method, headers, body: data ?? (body === undefined ? null : JSON.stringify(body)), signal });
      } catch (error) {
        if (signal.aborted) throw requestTimeoutError(where);
        throw new CliError('remote', `Could not reach Teams (${where}).`, null, { cause: error });
      }

      if (response.status === 429) {
        const delay = retryDelay(response.headers.get('retry-after'), attempt);
        if (attempt < maxRetries) {
          await sleep(delay);
          continue;
        }
        throw new CliError('rate_limited', 'Teams is throttling requests; try again later.', { retryAfterSeconds: Math.ceil(delay / 1000) });
      }
      if (response.status === 401) {
        throw new CliError('auth_required', 'The Teams session expired or was revoked. Run "teamscli auth".');
      }
      if (!response.ok) {
        throw new CliError('remote', `Teams request failed: ${response.status} ${where}.`, { status: response.status });
      }
      try {
        return { contentType: response.headers.get('content-type'), data: new Uint8Array(await response.arrayBuffer()) };
      } catch (error) {
        if (signal.aborted) throw requestTimeoutError(where);
        throw new CliError('remote', `Could not reach Teams (${where}).`, null, { cause: error });
      }
    }
  }

  /**
   * Follows `_metadata.backwardLink` from `url` (first page or a saved link) and yields every page,
   * including the last, empty one: its `syncState` is the checkpoint. `key` names the items array
   * (`conversations`, `messages`). Page order is not chronological; upsert by id.
   */
  async function* pages<T>(service: 'chatService' | 'csa', url: string, key: string): AsyncGenerator<Page<T>> {
    const origin = new URL((await session.endpoints())[service]).origin;
    let next: string | null = url;
    while (next) {
      assertTrustedUrl(next, origin);
      const body: Record<string, unknown> | null = await json(service, next);
      const metadata = (body?._metadata ?? {}) as { backwardLink?: string; syncState?: string };
      const items = (body?.[key] ?? []) as T[];
      const page: Page<T> = { items, backwardLink: metadata.backwardLink ?? null, syncState: metadata.syncState ?? null };
      if (page.backwardLink !== null) assertTrustedUrl(page.backwardLink, origin);
      if (page.syncState !== null) assertTrustedUrl(page.syncState, origin);
      yield page;
      next = items.length > 0 && page.backwardLink !== next ? page.backwardLink : null;
    }
  }

  return { json, bytes, pages };
}

function assertTrustedUrl(value: unknown, origin: string): URL {
  if (typeof value === 'string') {
    try {
      const url = new URL(value);
      if (url.protocol === 'https:' && url.origin === origin) return url;
    } catch {
      // Convert malformed remote links to the same safe error as unexpected origins.
    }
  }
  throw new CliError('remote', 'Teams returned a link to an unexpected host; refusing to follow it.');
}

function requestTimeoutError(where: string): CliError {
  return new CliError('remote', `Teams did not answer in time (${where}).`);
}

/** `Retry-After` in seconds or as an HTTP date; exponential backoff from 1 s when absent. */
function retryDelay(header: string | null, attempt: number): number {
  if (header) {
    const seconds = Number(header);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const date = Date.parse(header);
    if (!Number.isNaN(date)) return Math.max(0, date - Date.now());
  }
  return 1000 * 2 ** attempt;
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
