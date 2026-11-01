import { test } from 'node:test';
import assert from 'node:assert/strict';
import { SCOPES } from '../../src/auth/devicecode.ts';
import { createSession } from '../../src/auth/session.ts';
import { AAD_TOKEN_URL, AUTHZ_ROUTE, fakeJwt, memorySecretStore, TEST_OID, TEST_TID, tokenResponse } from '../helpers/fake-auth.ts';
import { fakeFetch, type RecordedRequest } from '../helpers/fake-fetch.ts';

const SPACES_TOKEN = fakeJwt({ oid: TEST_OID, tid: TEST_TID });

/** AAD that rotates the refresh token on every call: refresh-1 -> refresh-2 -> … */
function rotatingAad(expiresIn = 3600) {
  let n = 1;
  const { fetch, calls } = fakeFetch([AUTHZ_ROUTE, { method: 'POST', url: AAD_TOKEN_URL }]);
  const wrapped: typeof globalThis.fetch = async (input, init) => {
    if (String(input) !== AAD_TOKEN_URL) return fetch(input, init);
    await fetch(input, init);
    const scope = new URLSearchParams(String(init?.body)).get('scope') ?? '';
    const access = scope.startsWith(SCOPES.spaces) ? SPACES_TOKEN : 'test-chatsvcagg-token';
    return Response.json(tokenResponse(access, `refresh-${++n}`, expiresIn));
  };
  return { fetch: wrapped, calls };
}

const form = (call: RecordedRequest | undefined) => new URLSearchParams(String(call?.body));
const tokenCalls = (calls: RecordedRequest[]) => calls.filter((call) => call.url === AAD_TOKEN_URL);

test('authenticates each service, caching tokens and authz within the process', async () => {
  const { store } = memorySecretStore({ work: 'refresh-1' });
  const { fetch, calls } = rotatingAad();
  const session = createSession({ account: 'work', store, fetch, now: () => new Date('2026-01-01T00:00:00.000Z') });

  assert.deepEqual(await session.headers('chatService'), { authentication: 'skypetoken=test-skypetoken' });
  assert.deepEqual(await session.headers('csa'), { authorization: 'Bearer test-chatsvcagg-token' });
  assert.deepEqual(await session.headers('urlp', 'https://us-prod.asyncgw.teams.microsoft.com'), { authorization: 'Bearer test-chatsvcagg-token' });
  assert.deepEqual(await session.endpoints(), { chatService: 'https://chat.example.test', csa: 'https://teams.microsoft.com/api/csa/test' });
  await session.headers('chatService');
  await session.headers('csa');

  assert.equal(tokenCalls(calls).length, 2);
  assert.equal(calls.filter((call) => call.url === AUTHZ_ROUTE.url).length, 1);
});

test('every refresh uses and saves the rotated refresh token', async () => {
  const { store, secrets } = memorySecretStore({ work: 'refresh-1', other: 'other-refresh' });
  const { fetch, calls } = rotatingAad();
  const session = createSession({ account: 'work', store, fetch });

  await session.headers('chatService');
  await session.headers('csa');

  assert.deepEqual(tokenCalls(calls).map((call) => form(call).get('refresh_token')), ['refresh-1', 'refresh-2']);
  assert.equal(secrets.get('work'), 'refresh-3');
  assert.equal(secrets.get('other'), 'other-refresh');
});

test('persistRotation false keeps the rotated refresh token in memory only', async () => {
  const { store, secrets } = memorySecretStore({ work: 'refresh-1' });
  const { fetch, calls } = rotatingAad();
  const session = createSession({ account: 'work', store, fetch, persistRotation: false });

  await session.headers('chatService');
  await session.headers('csa');

  assert.deepEqual(tokenCalls(calls).map((call) => form(call).get('refresh_token')), ['refresh-1', 'refresh-2']);
  assert.equal(secrets.get('work'), 'refresh-1');
});

test('renews a token close to expiry', async () => {
  const { store } = memorySecretStore({ work: 'refresh-1' });
  const { fetch, calls } = rotatingAad(600);
  let clock = new Date('2026-01-01T00:00:00.000Z');
  const session = createSession({ account: 'work', store, fetch, now: () => clock });

  await session.headers('csa');
  clock = new Date('2026-01-01T00:05:00.000Z');
  await session.headers('csa');
  clock = new Date('2026-01-01T00:09:00.000Z');
  await session.headers('csa');
  assert.equal(tokenCalls(calls).length, 2);
});

test('missing or rejected login is auth_required, and a later call retries', async () => {
  const { store, secrets } = memorySecretStore();
  const { fetch } = rotatingAad();
  const session = createSession({ account: 'work', store, fetch });
  await assert.rejects(session.headers('csa'), { code: 'auth_required', message: /not signed in/ });

  secrets.set('work', 'refresh-1');
  assert.deepEqual(await session.headers('csa'), { authorization: 'Bearer test-chatsvcagg-token' });

  const rejected = fakeFetch([{ method: 'POST', url: AAD_TOKEN_URL, status: 400, json: { error: 'invalid_grant' } }]);
  const expired = createSession({ account: 'work', store, fetch: rejected.fetch });
  await assert.rejects(expired.headers('chatService'), { code: 'auth_required' });
});

test('a fresh login is used without reading the store; account() reports who and where', async () => {
  const { store, secrets } = memorySecretStore();
  const { fetch, calls } = rotatingAad();
  const login = { accessToken: SPACES_TOKEN, refreshToken: 'login-refresh', expiresAt: new Date(Date.now() + 3600_000) };
  const session = createSession({ account: 'work', store, fetch, login });

  assert.deepEqual(await session.account(), { name: 'work', userId: `8:orgid:${TEST_OID}`, tenantId: TEST_TID, region: 'test', displayName: null });
  assert.equal(tokenCalls(calls).length, 0);
  assert.equal(secrets.size, 0);
});
