import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createTeamsClient } from '../../src/teams/client.ts';
import { deleteMessage, editText, sendText } from '../../src/teams/send.ts';
import { fakeFetch, fakeSession, TEST_ENDPOINTS } from '../helpers/fake-fetch.ts';
import { aMessage } from '../helpers/builders.ts';

const chatId = 'chat-example';
const channelId = '19:channel@example.test';
const messageUrl = (conversationId: string) => {
  const marker = ';messageid=';
  const separator = conversationId.lastIndexOf(marker);
  const path = separator < 0
    ? encodeURIComponent(conversationId)
    : `${encodeURIComponent(conversationId.slice(0, separator))}${marker}${encodeURIComponent(conversationId.slice(separator + marker.length))}`;
  return `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${path}/messages?view=msnp24Equivalent&pageSize=200`;
};

function rawMessage(overrides: Record<string, unknown> = {}) {
  return {
    id: '1767261600000',
    messagetype: 'RichText/Html',
    content: '<p>Updated</p>',
    conversationid: chatId,
    from: 'https://chat.example.test/v1/users/ME/contacts/8:orgid:synthetic',
    imdisplayname: 'Alex Example',
    composetime: '2026-01-01T10:00:00.000Z',
    version: '1767261700000',
    clientmessageid: '123456789012345678',
    properties: {},
    ...overrides,
  };
}

test('sends a rich text reply to a thread and finds it by clientmessageid', async () => {
  const threadId = `${channelId};messageid=1767261600000`;
  const pageUrl = messageUrl(threadId);
  const fake = fakeFetch([
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)};messageid=1767261600000/messages`, status: 201, json: {} },
    { url: pageUrl, json: { messages: [], _metadata: {} } },
  ]);
  const fetch: typeof globalThis.fetch = async (input, init) => {
    const response = await fake.fetch(input, init);
    if ((init?.method ?? 'GET') !== 'GET') return response;
    const clientMessageId = (fake.calls[0]?.body as Record<string, unknown>).clientmessageid;
    return new Response(JSON.stringify({ messages: [rawMessage({ conversationid: channelId, clientmessageid: clientMessageId })], _metadata: {} }));
  };
  const client = createTeamsClient({ fetch, session: fakeSession() });
  const result = await sendText(client, TEST_ENDPOINTS, {
    conversationId: channelId,
    replyTo: '1767261600000',
    text: '<script>alert("x")</script>',
  });

  const body = fake.calls[0]?.body as Record<string, unknown>;
  assert.equal(fake.calls[0]?.method, 'POST');
  assert.equal(body.messagetype, 'RichText/Html');
  assert.equal(body.contenttype, 'text');
  assert.equal('postType' in body, false);
  assert.match(String(body.clientmessageid), /^\d+$/);
  assert.equal(body.content, '<p>&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;</p>');
  assert.equal(result.message?.id, '1767261600000');
  assert.equal(result.message?.clientMessageId, result.clientMessageId);
  assert.equal(fake.calls[1]?.url, pageUrl);
});

test('edits and soft-deletes a message at its conversation path', async () => {
  const id = '1767261600000';
  const endpoint = `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(chatId)}/messages/${id}`;
  const fake = fakeFetch([
    { method: 'PUT', url: endpoint, status: 200, json: {} },
    { url: messageUrl(chatId), times: 1, json: { messages: [rawMessage()], _metadata: {} } },
    { method: 'DELETE', url: `${endpoint}?behavior=softDelete`, status: 200, json: {} },
    { url: messageUrl(chatId), json: { messages: [rawMessage({ content: '', version: '1767261800000', properties: { deletetime: 1767261800000 } })], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  const message = aMessage({ id, conversationId: chatId });

  const edited = await editText(client, TEST_ENDPOINTS, message, '<updated>');
  assert.ok(edited);
  assert.equal((fake.calls[0]?.body as Record<string, unknown>).content, '<p>&lt;updated&gt;</p>');
  const editedBody = fake.calls[0]?.body as Record<string, unknown>;
  assert.match(String(editedBody.clientmessageid), /^\d+$/);
  assert.equal(editedBody.messagetype, 'RichText/Html');
  assert.equal((await deleteMessage(client, TEST_ENDPOINTS, message))?.deletedAt, '2026-01-01T10:03:20.000Z');
  assert.deepEqual(fake.calls.map(({ method }) => method), ['PUT', 'GET', 'DELETE', 'GET']);
});

test('uploads one SVG favicon per domain and places its AMS image before each link', async () => {
  const channel = '19:icons@example.test';
  const page = 'https://icons.example.test/one';
  const fake = fakeFetch([
    { url: page, text: '<html><head><link rel="apple-touch-icon" href="/touch.svg"></head></html>', headers: { 'content-type': 'text/html' } },
    { url: 'https://icons.example.test/touch.svg', text: '<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"><circle cx="16" cy="16" r="14" fill="#123456"/></svg>', headers: { 'content-type': 'image/svg+xml' } },
    { method: 'POST', url: 'https://us-api.asm.skype.com/v1/objects', status: 201, json: { id: '0-synthetic-icon' } },
    { method: 'PUT', url: 'https://us-api.asm.skype.com/v1/objects/0-synthetic-icon/content/imgpsh', json: {} },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, {
    conversationId: channel,
    channelPost: true,
    text: `See ${page} and https://icons.example.test/two`,
    linkIcons: true,
    fetch: fake.fetch,
  });

  const upload = fake.calls.find((call) => call.method === 'PUT' && call.url.includes('/content/imgpsh'));
  assert.ok(upload);
  assert.ok(upload.body instanceof Uint8Array && upload.body[0] === 0x89 && upload.body[1] === 0x50);
  assert.equal(fake.calls.filter((call) => call.url === page).length, 1);
  assert.equal(fake.calls.filter((call) => call.url === 'https://icons.example.test/touch.svg').length, 1);
  assert.equal(fake.calls.find((call) => call.url === page)?.headers.authorization, undefined);
  const body = fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>;
  assert.equal((body.content as string).match(/itemtype="http:\/\/schema\.skype\.com\/AMSImage"/g)?.length, 2, String(body.content));
  assert.match(String(body.content), /vertical-align:middle"><a href="https:\/\/icons\.example\.test\/one">/);
  assert.deepEqual(body.amsreferences, ['0-synthetic-icon']);
});

test('URLP supplies authenticated preview metadata and only the selected link gets a card', async () => {
  const channel = '19:preview@example.test';
  const selected = 'https://second.example.test/page';
  const urlp = `https://chat-prod.asyncgw.teams.microsoft.com/urlp/v1/url/info?url=${encodeURIComponent(selected)}`;
  const fake = fakeFetch([
    { url: urlp, json: { title: 'URLP title', description: 'URLP summary', image: 'https://cdn.example.test/card.png', token: 'synthetic-image-token' } },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, {
    conversationId: channel,
    channelPost: true,
    text: 'First https://first.example.test/ and second https://second.example.test/page',
    linkPreviews: true,
    previewLink: selected,
    fetch: fake.fetch,
  });

  assert.equal(fake.calls.find((call) => call.url === urlp)?.headers.authorization, 'Bearer test-token');
  const body = fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>;
  const links = JSON.parse(String((body.properties as Record<string, unknown>).links)) as Array<Record<string, any>>;
  assert.equal(links.length, 1);
  assert.equal(links[0]?.url, selected);
  assert.equal(links[0]?.preview.title, 'URLP title');
  assert.equal(links[0]?.preview.description, 'URLP summary');
  assert.equal(links[0]?.preview.previewurl, `https://chat-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=${encodeURIComponent('https://cdn.example.test/card.png')}&token=synthetic-image-token`);
});

test('falls back to public page metadata when URLP is unavailable and uploads repeated inline images', async () => {
  const channel = '19:fallback@example.test';
  const page = 'https://fallback.example.test/page';
  const fake = fakeFetch([
    { url: `https://chat-prod.asyncgw.teams.microsoft.com/urlp/v1/url/info?url=${encodeURIComponent(page)}`, status: 503 },
    { url: page, text: '<meta property="og:title" content="Fallback title"><meta property="og:image" content="/card.png">', headers: { 'content-type': 'text/html' } },
    { method: 'POST', url: 'https://us-api.asm.skype.com/v1/objects', times: 1, status: 201, json: { id: '0-inline-one' } },
    { method: 'PUT', url: 'https://us-api.asm.skype.com/v1/objects/0-inline-one/content/imgpsh', json: {} },
    { method: 'POST', url: 'https://us-api.asm.skype.com/v1/objects', status: 201, json: { id: '0-inline-two' } },
    { method: 'PUT', url: 'https://us-api.asm.skype.com/v1/objects/0-inline-two/content/imgpsh', json: {} },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, {
    conversationId: channel,
    text: `Link ${page}`,
    linkPreviews: true,
    images: [
      { name: 'banner.png', format: 'png', data: new Uint8Array([1, 2, 3]) },
      { name: 'icon.png', format: 'png', data: new Uint8Array([4, 5, 6]) },
    ],
    fetch: fake.fetch,
  });

  const body = fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>;
  const links = JSON.parse(String((body.properties as Record<string, unknown>).links)) as Array<Record<string, any>>;
  assert.equal(links.length, 1);
  assert.equal(links[0]?.preview.title, 'Fallback title');
  assert.equal(links[0]?.preview.previewurl, 'https://chat-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=https%3A%2F%2Ffallback.example.test%2Fcard.png');
  assert.deepEqual(body.amsreferences, ['0-inline-one', '0-inline-two']);
  assert.equal((body.content as string).match(/itemtype="http:\/\/schema\.skype\.com\/AMSImage"/g)?.length, 2);
});
