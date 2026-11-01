import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fetchLinkIcon, LINK_ICON_SIZE, pngSize } from '../../src/teams/link-icons.ts';
import { solidPng } from '../helpers/png.ts';
import { fakeFetch } from '../helpers/fake-fetch.ts';

test('tries apple-touch-icon before rel=icon and falls back to /favicon.ico', async () => {
  const png = solidPng(180, 180);
  const fake = fakeFetch([
    { url: 'https://site.example.test/page', text: '<link rel="icon" href="/icon.png"><link rel="apple-touch-icon" href="/touch.png">', headers: { 'content-type': 'text/html' } },
    { url: 'https://site.example.test/touch.png', status: 404 },
    { url: 'https://site.example.test/icon.png', status: 404 },
    { url: 'https://site.example.test/favicon.ico', data: png, headers: { 'content-type': 'image/png' } },
  ]);

  const icon = await fetchLinkIcon(fake.fetch, 'https://site.example.test/page');

  assert.equal(icon?.format, 'png');
  assert.deepEqual(pngSize(icon!.data), { width: 16, height: 16 });
  assert.deepEqual(fake.calls.map(({ url }) => url), [
    'https://site.example.test/page',
    'https://site.example.test/touch.png',
    'https://site.example.test/icon.png',
    'https://site.example.test/favicon.ico',
  ]);
});

test('tries /favicon.ico when the page itself cannot be fetched', async () => {
  const png = solidPng(32, 32);
  const fake = fakeFetch([
    { url: 'https://site.example.test/missing', status: 500 },
    { url: 'https://site.example.test/favicon.ico', data: png, headers: { 'content-type': 'image/png' } },
  ]);
  const icon = await fetchLinkIcon(fake.fetch, 'https://site.example.test/missing');
  assert.equal(icon?.format, 'png');
  assert.deepEqual(pngSize(icon!.data), { width: 16, height: 16 });
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
  assert.deepEqual(pngSize(icon!.data), { width: LINK_ICON_SIZE, height: LINK_ICON_SIZE });
});

test('scales every favicon to a 16px square, including wide SVGs and JPEGs it cannot decode', async () => {
  assert.equal(LINK_ICON_SIZE, 16);
  const wide = fakeFetch([
    { url: 'https://wide.example.test/favicon.ico', text: '<svg xmlns="http://www.w3.org/2000/svg" width="400" height="100"><rect width="400" height="100" fill="#f00"/></svg>', headers: { 'content-type': 'image/svg+xml' } },
  ]);
  const icon = await fetchLinkIcon(wide.fetch, 'https://wide.example.test/favicon.ico');
  assert.deepEqual(pngSize(icon!.data), { width: 16, height: 16 });

  const broken = fakeFetch([
    { url: 'https://broken.example.test/favicon.ico', data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1]), headers: { 'content-type': 'image/png' } },
  ]);
  assert.equal(await fetchLinkIcon(broken.fetch, 'https://broken.example.test/favicon.ico'), null, 'an undecodable icon is skipped, never sent at full size');
});
