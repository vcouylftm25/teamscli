import type { Endpoints, Session } from '../../src/teams/client.ts';

export interface FakeRoute {
  method?: string;
  /** Exact URL or pattern. */
  url: string | RegExp;
  status?: number;
  /** Response body, serialized as JSON. Omit for an empty body. */
  json?: unknown;
  text?: string;
  data?: Uint8Array;
  headers?: Record<string, string>;
  /** How many requests this route answers; unlimited when omitted. Routes are tried in order. */
  times?: number;
}

export interface RecordedRequest {
  method: string;
  url: string;
  headers: Record<string, string>;
  body: unknown;
}

/**
 * A `fetch` that answers from a route table and records every request.
 * An unmatched request throws, so the test fails instead of reaching the network.
 */
export function fakeFetch(routes: FakeRoute[]) {
  const calls: RecordedRequest[] = [];
  const remaining = routes.map((route) => route.times ?? Infinity);

  const fetch = async (input: string | URL | Request, init: RequestInit = {}): Promise<Response> => {
    const url = String(input instanceof Request ? input.url : input);
    const method = (init.method ?? 'GET').toUpperCase();
    const body = typeof init.body === 'string' ? parseBody(init.body) : init.body instanceof Uint8Array ? new Uint8Array(init.body) : null;
    calls.push({ method, url, headers: Object.fromEntries(new Headers(init.headers).entries()), body });

    const index = routes.findIndex(
      (route, i) =>
        (remaining[i] ?? 0) > 0 &&
        (route.method ?? 'GET').toUpperCase() === method &&
        (typeof route.url === 'string' ? route.url === url : route.url.test(url)),
    );
    const route = routes[index];
    if (!route) throw new Error(`fakeFetch: unexpected request ${method} ${url}`);
    remaining[index] = (remaining[index] ?? 0) - 1;
    const payload = route.data ?? (route.text ?? (route.json === undefined ? null : JSON.stringify(route.json)));
    return new Response(payload as BodyInit | null, { status: route.status ?? 200, headers: route.headers ?? {} });
  };

  return { fetch: fetch as typeof globalThis.fetch, calls };
}

/** JSON bodies are parsed; anything else (e.g. a urlencoded form) is kept as text. */
function parseBody(body: string): unknown {
  try {
    return JSON.parse(body);
  } catch {
    return body;
  }
}

export const TEST_ENDPOINTS: Endpoints = {
  chatService: 'https://chat.example.test',
  csa: 'https://csa.example.test/api/csa/test',
};

/** Session with fixed, obviously fake credentials. */
export function fakeSession(endpoints: Endpoints = TEST_ENDPOINTS): Session {
  return {
    headers: async (service): Promise<Record<string, string>> =>
      service === 'chatService' ? { authentication: 'skypetoken=test-skypetoken' }
        : service === 'media' ? { authorization: 'skype_token test-skypetoken' }
        : { authorization: 'Bearer test-token' },
    endpoints: async () => endpoints,
  };
}
