import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createTeamsClient } from '../../src/teams/client.ts';
import { deleteMessage, editText, sendText, sourceLabel, withLinkSources } from '../../src/teams/send.ts';
import { pngSize } from '../../src/teams/link-icons.ts';
import { solidPng } from '../helpers/png.ts';
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
  assert.match(String(body.content), /<img [^>]*width="16" height="16" style="width:16px;height:16px;vertical-align:text-bottom">&nbsp;<a href="https:\/\/icons\.example\.test\/one">/);
  assert.deepEqual(body.amsreferences, ['0-synthetic-icon']);
});

test('URLP supplies authenticated preview metadata and only the selected link gets a card', async () => {
  const channel = '19:preview@example.test';
  const selected = 'https://second.example.test/page';
  const urlp = `${TEST_ENDPOINTS.urlp}/v1/url/info?url=${encodeURIComponent(selected)}`;
  const thumbnail = `https://us-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail?url=${encodeURIComponent('https://cdn.example.test/card.png?w=1&h=2')}`;
  const fake = fakeFetch([
    { url: urlp, json: { title: 'URLP title', description: 'URLP summary', thumbnail, image: { url: 'https://cdn.example.test/card.png?w=1&h=2' } } },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, {
    conversationId: channel,
    channelPost: true,
    subject: 'Weekly',
    text: 'First https://first.example.test/ and second https://second.example.test/page',
    linkPreviews: true,
    previewLink: selected,
    fetch: fake.fetch,
  });

  assert.equal(fake.calls.find((call) => call.url === urlp)?.headers.authorization, 'Bearer test-token');
  const body = fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>;
  const links = JSON.parse(String((body.properties as Record<string, unknown>).links)) as Array<Record<string, any>>;
  assert.equal(links.length, 1);
  assert.equal(links[0]?.['@type'], 'http://schema.skype.com/HyperLink');
  assert.equal(links[0]?.itemid, '0');
  assert.equal(links[0]?.url, selected);
  assert.equal(links[0]?.previewenabled, true);
  assert.equal(links[0]?.preview.title, 'URLP title');
  assert.equal(links[0]?.preview.description, 'URLP summary');
  assert.equal(links[0]?.preview.previewurl, thumbnail);
  assert.match(String(body.content), new RegExp(`<a href="${selected}" itemtype="http://schema.skype.com/HyperLink" rel="noreferrer noopener" target="_blank" itemid="0">`));
  assert.doesNotMatch(String(body.content), /first\.example\.test\/" itemtype=/);
  assert.equal((body.properties as Record<string, unknown>).subject, 'Weekly');
  assert.equal(body.postType, 'Standard');
});

test('builds an encoded thumbnail when URLP returns only the image, and refuses another host', async () => {
  const channel = '19:encoded@example.test';
  const page = 'https://encoded.example.test/a?b=1&c=2';
  const image = 'https://cdn.example.test/card.png?w=1&h=2';
  const fake = fakeFetch([
    { url: `${TEST_ENDPOINTS.urlp}/v1/url/info?url=${encodeURIComponent(page)}`, json: { title: 'Encoded', description: 'Query', thumbnail: 'https://evil.example.test/pixel.png', image } },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, {
    conversationId: channel,
    text: `See ${page} and https://other.example.test/`,
    linkPreviews: true,
    fetch: fake.fetch,
  });
  const body = fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>;
  const links = JSON.parse(String((body.properties as Record<string, unknown>).links)) as Array<Record<string, any>>;
  const previewurl = String(links[0]?.preview.previewurl);
  assert.equal(links[0]?.url, page);
  assert.equal(previewurl, `${TEST_ENDPOINTS.urlp}/v1/url/image/Thumbnail?url=${encodeURIComponent(image)}`);
  assert.equal(new URL(previewurl).searchParams.get('url'), image);
  assert.equal(previewurl.includes('evil.example.test'), false);
  assert.match(String(body.content), /itemid="0">https:\/\/encoded\.example\.test\/a\?b=1&amp;c=2</);
  assert.doesNotMatch(String(body.content), /other\.example\.test\/" itemtype=/);
});

test('--preview-link must match a link and a subject is ignored on a reply', async () => {
  const fake = fakeFetch([]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await assert.rejects(() => sendText(client, TEST_ENDPOINTS, {
    conversationId: channelId,
    text: 'See https://first.example.test/',
    linkPreviews: true,
    previewLink: 'https://missing.example.test/',
    fetch: fake.fetch,
  }), { code: 'usage' });
  assert.equal(fake.calls.length, 0);

  const posted = fakeFetch([
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channelId)};messageid=1767261600000/messages`, status: 201, json: {} },
    { url: messageUrl(`${channelId};messageid=1767261600000`), json: { messages: [], _metadata: {} } },
  ]);
  await sendText(createTeamsClient({ fetch: posted.fetch, session: fakeSession() }), TEST_ENDPOINTS, {
    conversationId: channelId,
    replyTo: '1767261600000',
    channelPost: true,
    subject: 'Not a root',
    text: 'Reply',
  });
  const body = posted.calls[0]?.body as Record<string, unknown>;
  assert.equal(body.properties, undefined);
  assert.equal('postType' in body, false);
});

test('falls back to public page metadata when URLP is unavailable and uploads repeated inline images', async () => {
  const channel = '19:fallback@example.test';
  const page = 'https://fallback.example.test/page';
  const fake = fakeFetch([
    { url: `${TEST_ENDPOINTS.urlp}/v1/url/info?url=${encodeURIComponent(page)}`, status: 503 },
    { url: page, text: '<meta content="Fallback title" property="og:title"><meta content="/card.png" property="og:image">', headers: { 'content-type': 'text/html' } },
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
  assert.equal(links[0]?.preview.previewurl, `${TEST_ENDPOINTS.urlp}/v1/url/image/Thumbnail?url=${encodeURIComponent('https://fallback.example.test/card.png')}`);
  assert.equal(fake.calls.find((call) => call.url === page)?.headers.authorization, undefined);
  assert.deepEqual(body.amsreferences, ['0-inline-one', '0-inline-two']);
  assert.equal((body.content as string).match(/itemtype="http:\/\/schema\.skype\.com\/AMSImage"/g)?.length, 2);
});

test('uploads markdown images through AMS and leaves the external address out of the message', async () => {
  const channel = '19:markdown-image@example.test';
  const remote = 'https://cdn.example.test/photos/chart.png';
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 9]);
  const fake = fakeFetch([
    { url: remote, data: png, headers: { 'content-type': 'image/png' } },
    { method: 'POST', url: 'https://us-api.asm.skype.com/v1/objects', times: 1, status: 201, json: { id: '0-remote-image' } },
    { method: 'PUT', url: 'https://us-api.asm.skype.com/v1/objects/0-remote-image/content/imgpsh', json: {} },
    { method: 'POST', url: 'https://us-api.asm.skype.com/v1/objects', status: 201, json: { id: '0-local-image' } },
    { method: 'PUT', url: 'https://us-api.asm.skype.com/v1/objects/0-local-image/content/imgpsh', json: {} },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, {
    conversationId: channel,
    text: `Before ![Chart](${remote}) after ![local](images/banner.jpg)`,
    fetch: fake.fetch,
    readFile: (path) => {
      assert.equal(path, 'images/banner.jpg');
      return new Uint8Array([0xff, 0xd8, 0xff, 1]);
    },
  });
  assert.equal(fake.calls.find((call) => call.url === remote)?.headers.authorization, undefined);
  const body = fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>;
  assert.equal(String(body.content).includes(remote), false);
  assert.match(String(body.content), /Before <img [^>]*alt="Chart"[^>]*itemid="0-remote-image"[^>]*> after <img [^>]*itemid="0-local-image"/);
  assert.equal(String(body.content).includes('images/banner.jpg'), false);
  assert.deepEqual(body.amsreferences, ['0-remote-image', '0-local-image']);
});

test('sends the link unchanged when its favicon cannot be fetched', async () => {
  const channel = '19:no-icon@example.test';
  const page = 'https://plain.example.test/post';
  const fake = fakeFetch([
    { url: page, status: 500 },
    { url: 'https://plain.example.test/favicon.ico', status: 404 },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, {
    conversationId: channel,
    text: `Read ${page}`,
    linkIcons: true,
    fetch: fake.fetch,
  });
  const body = fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>;
  assert.equal(String(body.content).includes('AMSImage'), false);
  assert.equal(body.amsreferences, undefined);
  assert.match(String(body.content), /<a href="https:\/\/plain\.example\.test\/post">/);
});

test('link sources add the site after each named link and skip links that already show it', () => {
  assert.equal(sourceLabel('access.redhat.com'), 'redhat.com');
  assert.equal(sourceLabel('www.cisa.gov'), 'cisa.gov');
  assert.equal(sourceLabel('www.gov.br'), 'gov.br');
  assert.equal(sourceLabel('noticias.uol.com.br'), 'uol.com.br');
  assert.equal(sourceLabel('github.com'), 'github.com');
  assert.equal(
    withLinkSources('<p><a href="https://access.redhat.com/x?a=1&amp;b=2">Red Hat: boletim</a> and <a href="https://github.com/o/r">https://github.com/o/r</a></p>'),
    '<p><a href="https://access.redhat.com/x?a=1&amp;b=2">Red Hat: boletim</a> <span style="font-size:12px;color:#616161">· redhat.com</span>'
      + ' and <a href="https://github.com/o/r">https://github.com/o/r</a></p>',
  );
});

test('compact radar post keeps every word, drops empty sections and puts icon and source on the link line', async () => {
  const channel = '19:radar@example.test';
  const text = [
    '## 🤖 IA',
    '**Item um**',
    'O que aconteceu, explicado por inteiro.',
    '> 💡 **Por que importa pra LFTM:** motivo completo.',
    '',
    '[Cursor Blog: aviso](https://cursor.com/blog/a)',
    '',
    '---',
    '',
    '## 📈 Mercado',
    '',
    '---',
  ].join('\n');
  const fake = fakeFetch([
    { url: 'https://cursor.com/blog/a', text: '<link rel="icon" href="/i.png">', headers: { 'content-type': 'text/html' } },
    { url: 'https://cursor.com/i.png', data: solidPng(64, 64), headers: { 'content-type': 'image/png' } },
    { method: 'POST', url: 'https://us-api.asm.skype.com/v1/objects', status: 201, json: { id: '0-cursor-icon' } },
    { method: 'PUT', url: 'https://us-api.asm.skype.com/v1/objects/0-cursor-icon/content/imgpsh', json: {} },
    { method: 'POST', url: `${TEST_ENDPOINTS.chatService}/v1/users/ME/conversations/${encodeURIComponent(channel)}/messages`, status: 201, json: {} },
    { url: messageUrl(channel), json: { messages: [], _metadata: {} } },
  ]);
  const client = createTeamsClient({ fetch: fake.fetch, session: fakeSession() });
  await sendText(client, TEST_ENDPOINTS, { conversationId: channel, channelPost: true, text, compact: true, linkIcons: true, linkSources: true, fetch: fake.fetch });
  const upload = fake.calls.find((call) => call.method === 'PUT' && call.url.includes('/content/imgpsh'));
  assert.deepEqual(pngSize(upload!.body as Uint8Array), { width: 16, height: 16 });
  const content = String((fake.calls.find((call) => call.method === 'POST' && call.url.endsWith('/messages'))?.body as Record<string, unknown>).content);
  assert.equal(content.includes('Mercado'), false);
  assert.equal(content.includes('<hr>'), false);
  assert.equal(content.includes('&nbsp;</p>'), false);
  for (const words of ['O que aconteceu, explicado por inteiro.', 'motivo completo.', 'Cursor Blog: aviso']) assert.ok(content.includes(words), words);
  assert.match(content, /^<h3>🤖 IA<\/h3><p><strong>Item um<\/strong><br>O que aconteceu/);
  assert.match(content, /<\/blockquote><p><img [^>]*width="16" height="16"[^>]*>&nbsp;<a href="https:\/\/cursor\.com\/blog\/a">Cursor Blog: aviso<\/a> <span style="font-size:12px;color:#616161">· cursor\.com<\/span><\/p>$/);
});
