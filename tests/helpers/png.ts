import { deflateSync } from 'node:zlib';

/** A valid solid-colour RGBA PNG of the given size, for favicon tests. */
export function solidPng(width: number, height: number, rgba: [number, number, number, number] = [0x12, 0x34, 0x56, 0xff]): Uint8Array<ArrayBuffer> {
  const raw = new Uint8Array(height * (width * 4 + 1));
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) raw.set(rgba, y * (width * 4 + 1) + 1 + x * 4);
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  header[8] = 8;
  header[9] = 6;
  const parts = [new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', deflateSync(raw)), chunk('IEND', new Uint8Array())];
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) { out.set(part, offset); offset += part.length; }
  return out;
}

function chunk(type: string, bytes: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + bytes.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, bytes.length);
  out.set(new TextEncoder().encode(type), 4);
  out.set(bytes, 8);
  view.setUint32(8 + bytes.length, crc32(out.subarray(4, 8 + bytes.length)));
  return out;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}
