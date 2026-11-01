import type { SecretStore } from '../../src/auth/secrets.ts';
import type { FakeRoute } from './fake-fetch.ts';

/** In-memory SecretStore; `secrets` exposes what was saved. */
export function memorySecretStore(initial: Record<string, string> = {}) {
  const secrets = new Map(Object.entries(initial));
  const store: SecretStore = {
    kind: 'file',
    get: async (account) => secrets.get(account) ?? null,
    set: async (account, secret) => void secrets.set(account, secret),
    delete: async (account) => void secrets.delete(account),
  };
  return { store, secrets };
}

/** Unsigned JWT with the given claims; enough for code that only reads the payload. */
export function fakeJwt(claims: Record<string, unknown>): string {
  return `e30.${Buffer.from(JSON.stringify(claims)).toString('base64url')}.signature`;
}

export const TEST_OID = '00000000-0000-0000-0000-0000000000aa';
export const TEST_TID = '00000000-0000-0000-0000-0000000000bb';

export const AAD_TOKEN_URL = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token';
export const AAD_DEVICECODE_URL = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode';
export const AUTHZ_URL = 'https://teams.microsoft.com/api/authsvc/v1.0/authz';

/** A successful AAD token response. */
export function tokenResponse(accessToken: string, refreshToken: string | null, expiresIn = 3600) {
  return { token_type: 'Bearer', access_token: accessToken, expires_in: expiresIn, ...(refreshToken ? { refresh_token: refreshToken } : {}) };
}

/** Synthetic authz answer for region `test`. */
export const AUTHZ_ROUTE: FakeRoute = {
  method: 'POST',
  url: AUTHZ_URL,
  json: {
    tokens: { skypeToken: 'test-skypetoken', expiresIn: 86400 },
    region: 'test',
    regionGtms: { chatService: 'https://chat.example.test/', urlpV2: 'https://urlp.example.test/urlp/' },
  },
};
