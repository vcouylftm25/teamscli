import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { deflateSync } from 'node:zlib';
import { initWasm, Resvg } from '@resvg/resvg-wasm';

export interface LinkIcon {
  data: Uint8Array<ArrayBuffer>;
  format: 'png' | 'jpg' | 'gif' | 'webp';
}

const require = createRequire(import.meta.url);
const USER_AGENT = 'Mozilla/5.0 (compatible; teamscli link icon)';
// resvg throws if initWasm runs twice; this promise is the one shared initialization.
let wasm: Promise<void> | null = null;

function wasmReady(): Promise<void> {
  if (!wasm) {
    const bytes = readFileSync(require.resolve('@resvg/resvg-wasm/index_bg.wasm'));
    wasm = initWasm(bytes).catch((error: unknown) => {
      wasm = null;
      throw error;
    });
  }
  return wasm;
}

/** Finds and downloads one public favicon for a link's site. */
export async function fetchLinkIcon(fetcher: typeof fetch, pageUrl: string): Promise<LinkIcon | null> {
  let pageBase: URL;
  try {
    pageBase = new URL(pageUrl);
  } catch {
    return null;
  }
  if (!['https:', 'http:'].includes(pageBase.protocol) || pageBase.username || pageBase.password) return null;
  const candidates: string[] = [];
  try {
    const page = await fetcher(pageUrl, { headers: { 'user-agent': USER_AGENT }, redirect: 'follow', signal: AbortSignal.timeout(8000) });
    if (page.ok && page.headers.get('content-type')?.toLowerCase().includes('text/html')) {
      const html = (await safeText(page)).slice(0, 300_000);
      pageBase = safeUrl(page.url || pageUrl, pageUrl);
      for (const tag of html.matchAll(/<link\b[^>]*>/gi)) {
        const attrs = attributes(tag[0]);
        if (!attrs.href || /^data:/i.test(attrs.href)) continue;
        const rel = (attrs.rel ?? '').toLowerCase().split(/\s+/);
        const apple = rel.includes('apple-touch-icon') || rel.includes('apple-touch-icon-precomposed');
        const icon = rel.includes('icon') || rel.includes('shortcut');
        if (apple) candidates.unshift(attrs.href);
        else if (icon) candidates.push(attrs.href);
      }
    }
  } catch {
    // The site's /favicon.ico can still exist when the page itself does not answer.
  }
  candidates.push('/favicon.ico');
  const unique = new Set<string>();
  for (const candidate of candidates) {
    let iconUrl: URL;
    try {
      iconUrl = new URL(candidate, pageBase);
      if (!['https:', 'http:'].includes(iconUrl.protocol) || iconUrl.username || iconUrl.password) continue;
      if (unique.has(iconUrl.href)) continue;
      unique.add(iconUrl.href);
    } catch {
      continue;
    }
    try {
      const response = await fetcher(iconUrl, { headers: { 'user-agent': USER_AGENT }, redirect: 'follow', signal: AbortSignal.timeout(8000) });
      if (!response.ok) continue;
      const type = response.headers.get('content-type')?.split(';', 1)[0]?.trim().toLowerCase() ?? '';
      const bytes = new Uint8Array(await response.arrayBuffer());
      if (bytes.length === 0 || bytes.length > 2_000_000) continue;
      const converted = await toInlineIcon(bytes, type || iconUrl.pathname.toLowerCase());
      if (converted) return converted;
    } catch {
      continue;
    }
  }
  return null;
}

async function toInlineIcon(data: Uint8Array, contentTypeOrPath: string): Promise<LinkIcon | null> {
  if (contentTypeOrPath.includes('svg') || isSvg(data)) {
    try {
      await wasmReady();
      const svg = new Resvg(new TextDecoder().decode(data), { fitTo: { mode: 'width', value: 18 }, font: { loadSystemFonts: false } });
      const png = svg.render().asPng();
      svg.free();
      return { data: new Uint8Array(png), format: 'png' };
    } catch {
      return null;
    }
  }
  if (contentTypeOrPath.includes('icon') || isIco(data)) {
    const png = icoPngFrame(data) ?? icoBitmapFrame(data);
    return png ? { data: new Uint8Array(png), format: 'png' } : null;
  }
  if (contentTypeOrPath.includes('png') || isPng(data)) return { data: new Uint8Array(data), format: 'png' };
  if (contentTypeOrPath.includes('jpeg') || contentTypeOrPath.includes('jpg') || isJpeg(data)) return { data: new Uint8Array(data), format: 'jpg' };
  if (contentTypeOrPath.includes('gif') || isGif(data)) return { data: new Uint8Array(data), format: 'gif' };
  if (contentTypeOrPath.includes('webp') || isWebp(data)) return { data: new Uint8Array(data), format: 'webp' };
  return null;
}

function attributes(tag: string): Record<string, string> {
  const result: Record<string, string> = {};
  for (const match of tag.matchAll(/([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g)) {
    result[match[1]!.toLowerCase()] = decode(match[2] ?? match[3] ?? match[4] ?? '');
  }
  return result;
}

async function safeText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch {
    return '';
  }
}

function safeUrl(value: string, fallback: string): URL {
  try {
    const url = new URL(value);
    if (url.protocol === 'http:' || url.protocol === 'https:') return url;
  } catch {
    // The caller supplied URL is already validated by its link matcher.
  }
  return new URL(fallback);
}

function decode(value: string): string {
  return value.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>');
}

function isSvg(data: Uint8Array): boolean {
  return new TextDecoder().decode(data.subarray(0, 512)).trimStart().startsWith('<svg') || /<svg\b/i.test(new TextDecoder().decode(data.subarray(0, 512)));
}

function isIco(data: Uint8Array): boolean { return data[0] === 0 && data[1] === 0 && data[2] === 1 && data[3] === 0; }
function isPng(data: Uint8Array): boolean { return data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47; }
function isJpeg(data: Uint8Array): boolean { return data[0] === 0xff && data[1] === 0xd8; }
function isGif(data: Uint8Array): boolean { return data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46; }
function isWebp(data: Uint8Array): boolean { return data[0] === 0x52 && data[1] === 0x49 && data[2] === 0x46 && data[8] === 0x57; }

function icoPngFrame(data: Uint8Array): Uint8Array | null {
  if (!isIco(data) || data.length < 22) return null;
  const count = data[4]! | (data[5]! << 8);
  let selected: { data: Uint8Array; size: number } | null = null;
  for (let index = 0; index < count; index++) {
    const entry = 6 + index * 16;
    if (entry + 16 > data.length) break;
    const size = data[entry] || 256;
    const length = (data[entry + 8]! | (data[entry + 9]! << 8) | (data[entry + 10]! << 16) | (data[entry + 11]! << 24)) >>> 0;
    const offset = (data[entry + 12]! | (data[entry + 13]! << 8) | (data[entry + 14]! << 16) | (data[entry + 15]! << 24)) >>> 0;
    const frame = data.subarray(offset, offset + length);
    if (isPng(frame) && (!selected || Math.abs(size - 18) < Math.abs(selected.size - 18))) selected = { data: frame, size };
  }
  return selected?.data ?? null;
}

function icoBitmapFrame(data: Uint8Array): Uint8Array | null {
  if (!isIco(data) || data.length < 22) return null;
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const count = view.getUint16(4, true);
  let selected: { png: Uint8Array; size: number } | null = null;
  for (let index = 0; index < count; index++) {
    const entry = 6 + index * 16;
    if (entry + 16 > data.length) break;
    const size = data[entry] || 256;
    const length = view.getUint32(entry + 8, true);
    const offset = view.getUint32(entry + 12, true);
    if (offset + length > data.length || length < 40) continue;
    const headerSize = view.getUint32(offset, true);
    const width = view.getInt32(offset + 4, true);
    const doubleHeight = view.getInt32(offset + 8, true);
    const bits = view.getUint16(offset + 14, true);
    const compression = view.getUint32(offset + 16, true);
    const height = Math.abs(doubleHeight) / 2;
    if (headerSize < 40 || width <= 0 || !Number.isInteger(height) || height <= 0 || ![24, 32].includes(bits) || compression !== 0) continue;
    const rowBytes = Math.ceil(width * bits / 32) * 4;
    const pixelsOffset = offset + headerSize;
    if (pixelsOffset + rowBytes * height > offset + length) continue;
    const rgba = new Uint8Array(width * height * 4);
    let hasAlpha = false;
    for (let y = 0; y < height; y++) {
      const sourceY = doubleHeight > 0 ? height - y - 1 : y;
      for (let x = 0; x < width; x++) {
        const source = pixelsOffset + sourceY * rowBytes + x * (bits / 8);
        const target = (y * width + x) * 4;
        rgba[target] = data[source + 2]!;
        rgba[target + 1] = data[source + 1]!;
        rgba[target + 2] = data[source]!;
        rgba[target + 3] = bits === 32 ? data[source + 3]! : 255;
        hasAlpha ||= rgba[target + 3] !== 0;
      }
    }
    if (bits === 32 && !hasAlpha) for (let i = 3; i < rgba.length; i += 4) rgba[i] = 255;
    const png = encodePng(rgba, width, height);
    if (!selected || Math.abs(size - 18) < Math.abs(selected.size - 18)) selected = { png, size };
  }
  return selected?.png ?? null;
}

function encodePng(rgba: Uint8Array, width: number, height: number): Uint8Array<ArrayBuffer> {
  const scanlines = new Uint8Array(height * (width * 4 + 1));
  for (let y = 0; y < height; y++) scanlines.set(rgba.subarray(y * width * 4, (y + 1) * width * 4), y * (width * 4 + 1) + 1);
  const chunk = (type: string, bytes: Uint8Array): Uint8Array => {
    const typeBytes = new TextEncoder().encode(type);
    const output = new Uint8Array(12 + bytes.length);
    const view = new DataView(output.buffer);
    view.setUint32(0, bytes.length);
    output.set(typeBytes, 4);
    output.set(bytes, 8);
    view.setUint32(8 + bytes.length, crc32(output.subarray(4, 8 + bytes.length)));
    return output;
  };
  const header = new Uint8Array(13);
  const headerView = new DataView(header.buffer);
  headerView.setUint32(0, width);
  headerView.setUint32(4, height);
  header[8] = 8;
  header[9] = 6;
  const parts = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', header), chunk('IDAT', new Uint8Array(deflateSync(scanlines))), chunk('IEND', new Uint8Array())];
  const png = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) { png.set(part, offset); offset += part.length; }
  return png;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
