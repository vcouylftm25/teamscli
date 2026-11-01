import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fetchLinkIcon } from '../../src/teams/link-icons.ts';
import { fakeFetch } from '../helpers/fake-fetch.ts';

test('tries apple-touch-icon before rel=icon and falls back to /favicon.ico', async () => {
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1]);
  const fake = fakeFetch([
    { url: 'https://site.example.test/page', text: '<link rel="icon" href="/icon.png"><link rel="apple-touch-icon" href="/touch.png">', headers: { 'content-type': 'text/html' } },
    { url: 'https://site.example.test/touch.png', status: 404 },
    { url: 'https://site.example.test/icon.png', status: 404 },
    { url: 'https://site.example.test/favicon.ico', data: png, headers: { 'content-type': 'image/png' } },
  ]);

  const icon = await fetchLinkIcon(fake.fetch, 'https://site.example.test/page');

  assert.equal(icon?.format, 'png');
  assert.deepEqual([...icon!.data], [...png]);
  assert.deepEqual(fake.calls.map(({ url }) => url), [
    'https://site.example.test/page',
    'https://site.example.test/touch.png',
    'https://site.example.test/icon.png',
    'https://site.example.test/favicon.ico',
  ]);
});

test('returns no icon when page and fallback icon are unavailable', async () => {
  const fake = fakeFetch([
    { url: 'https://site.example.test/page', text: '<html></html>', headers: { 'content-type': 'text/html' } },
    { url: 'https://site.example.test/favicon.ico', status: 404 },
  ]);
  assert.equal(await fetchLinkIcon(fake.fetch, 'https://site.example.test/page'), null);
});

test('converts a bitmap ICO favicon to PNG', async () => {
  const ico = new Uint8Array(70);
  const view = new DataView(ico.buffer);
  view.setUint16(2, 1, true);
  view.setUint16(4, 1, true);
  ico[6] = 1;
  ico[7] = 1;
  view.setUint16(10, 1, true);
  view.setUint16(12, 32, true);
  view.setUint32(14, 48, true);
  view.setUint32(18, 22, true);
  view.setUint32(22, 40, true);
  view.setInt32(26, 1, true);
  view.setInt32(30, 2, true);
  view.setUint16(34, 1, true);
  view.setUint16(36, 32, true);
  ico.set([0x33, 0x22, 0x11, 0xff], 62);
  const fake = fakeFetch([
    { url: 'https://site.example.test/page', text: '<link rel="icon" href="/favicon.ico">', headers: { 'content-type': 'text/html' } },
    { url: 'https://site.example.test/favicon.ico', data: ico, headers: { 'content-type': 'image/x-icon' } },
  ]);

  const icon = await fetchLinkIcon(fake.fetch, 'https://site.example.test/page');

  assert.equal(icon?.format, 'png');
  assert.deepEqual([...icon!.data.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  assert.equal(new DataView(icon!.data.buffer, icon!.data.byteOffset).getUint32(16), 1);
  assert.equal(new DataView(icon!.data.buffer, icon!.data.byteOffset).getUint32(20), 1);
});
