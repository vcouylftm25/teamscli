import { randomInt } from 'node:crypto';
import { CliError } from '../errors.ts';
import type { Message } from '../models.ts';
import { listMessages } from './messages.ts';
import type { Endpoints, TeamsClient } from './client.ts';
import { escapeHtml, maskMarkdownImages, quoteHtml, toTeamsHtml, toTeamsHtmlWithMentions, type MarkdownImageRef, type PersonMention, type QuotedMessage } from './compose.ts';
import { uploadImage, type UploadedImage } from './files.ts';
import { fetchLinkIcon, LINK_ICON_SIZE } from './link-icons.ts';

export interface SendTextOptions {
  conversationId: string;
  replyTo?: string;
  channelPost?: boolean;
  text: string;
  quote?: QuotedMessage;
  mentions?: PersonMention[];
  subject?: string;
  linkPreviews?: boolean;
  previewLink?: string;
  linkIcons?: boolean;
  /** Adds the link's site (`github.com`) in small grey text after each link. */
  linkSources?: boolean;
  /** Tighter layout; see `ComposeOptions.compact`. */
  compact?: boolean;
  fetch?: typeof fetch;
  /** Reads a local markdown image. The CLI passes this; tests pass a fake. */
  readFile?: (path: string) => Uint8Array<ArrayBuffer>;
  images?: Array<{ name: string; format: string; data: Uint8Array<ArrayBuffer> }>;
}

export async function sendText(client: TeamsClient, endpoints: Endpoints, { conversationId, replyTo, channelPost, text, quote, mentions, subject, linkPreviews, previewLink, linkIcons, linkSources, compact, fetch: fetcher, readFile, images }: SendTextOptions): Promise<{ clientMessageId: string; message: Message | null }> {
  const clientMessageId = newClientMessageId();
  const targetId = replyTo === undefined ? conversationId : `${conversationId};messageid=${replyTo}`;
  const masked = maskMarkdownImages(text);
  const body: Record<string, any> = messageBody(masked.text, clientMessageId, quote, mentions, compact === true);
  const properties: Record<string, unknown> = { ...(body.properties ?? {}) };
  const amsreferences: string[] = [...(body.amsreferences ?? [])];
  if (subject && channelPost && replyTo === undefined) properties.subject = subject;
  for (const image of masked.images) {
    const loaded = await loadMarkdownImage(image, fetcher, readFile);
    const uploaded = await uploadImage(client, conversationId, loaded.name, loaded.data);
    body.content = body.content.replaceAll(image.marker, amsImageTag({ ...uploaded, format: loaded.format }, image.alt));
    amsreferences.push(uploaded.id);
  }
  if (images?.length) {
    const uploaded = await Promise.all(images.map((image) => uploadImage(client, conversationId, image.name, image.data).then((item) => ({ ...item, format: image.format }))));
    body.content += uploaded.map((image) => `<p>${amsImageTag(image, 'image')}</p>`).join('');
    amsreferences.push(...uploaded.map((image) => image.id));
  }
  if (linkSources) body.content = withLinkSources(body.content);
  if (linkIcons && fetcher) {
    const icons = await withLinkIcons(client, body.content, conversationId, fetcher);
    body.content = icons.content;
    amsreferences.push(...icons.imageIds);
  }
  if (linkPreviews) {
    const previews = await withLinkPreviews(client, body.content, endpoints, previewLink, fetcher);
    body.content = previews.content;
    if (previews.links.length > 0) properties.links = JSON.stringify(previews.links);
  }
  if (amsreferences.length > 0) body.amsreferences = amsreferences;
  if (Object.keys(properties).length > 0) body.properties = properties;
  await client.json('chatService', `${conversationUrl(endpoints, targetId)}/messages`, {
    method: 'POST',
    body: {
      ...body,
      ...(channelPost && replyTo === undefined ? { postType: 'Standard' } : {}),
    },
  });
  const message = await firstMessage(client, endpoints, targetId, (item) => item.clientMessageId === clientMessageId);
  return { clientMessageId, message };
}

function amsImageTag(image: UploadedImage & { format: string }, alt: string, icon = false): string {
  // The icon file is already LINK_ICON_SIZE pixels; the attributes and the style say so again for
  // clients that do honour them.
  const px = LINK_ICON_SIZE;
  const size = icon ? ` width="${px}" height="${px}" style="width:${px}px;height:${px}px;vertical-align:text-bottom"` : '';
  return `<img src="${escapeHtml(image.url)}" itemtype="http://schema.skype.com/AMSImage" itemscope="${escapeHtml(image.format)}" alt="${escapeHtml(alt)}" id="x_${escapeHtml(image.id)}" itemid="${escapeHtml(image.id)}"${size}>`;
}

async function withLinkIcons(client: TeamsClient, content: string, conversationId: string, fetcher: typeof fetch): Promise<{ content: string; imageIds: string[] }> {
  const anchors = [...content.matchAll(/<a href="([^"]+)">/g)];
  const domains = new Map<string, Promise<(UploadedImage & { format: string }) | null>>();
  const byUrl = new Map<string, UploadedImage & { format: string }>();
  const imageIds: string[] = [];
  for (const [, encoded] of anchors) {
    const url = decodeHref(encoded!);
    let parsed: URL;
    try { parsed = new URL(url); } catch { continue; }
    if (!['http:', 'https:'].includes(parsed.protocol)) continue;
    const domain = parsed.hostname.toLowerCase();
    let imagePromise = domains.get(domain);
    if (!imagePromise) {
      imagePromise = (async () => {
        const icon = await fetchLinkIcon(fetcher, url);
        if (!icon) return null;
        try {
          const uploaded = await uploadImage(client, conversationId, `favicon-${domain}.${icon.format}`, icon.data);
          return { ...uploaded, format: icon.format };
        } catch {
          return null;
        }
      })();
      domains.set(domain, imagePromise);
    }
    const image = await imagePromise;
    if (image) {
      byUrl.set(url, image);
      if (!imageIds.includes(image.id)) imageIds.push(image.id);
    }
  }
  const next = content.replace(/<a href="([^"]+)">/g, (whole, encoded: string) => {
    const image = byUrl.get(decodeHref(encoded));
    if (!image) return whole;
    return `${amsImageTag(image, '', true)}&nbsp;${whole}`;
  });
  return { content: next, imageIds };
}

/** Appends ` · site.com` in small grey text after each link whose text does not already show its address. */
export function withLinkSources(content: string): string {
  return content.replace(/<a href="([^"]+)">([\s\S]*?)<\/a>/g, (whole, encoded: string, label: string) => {
    let url: URL;
    try { url = new URL(decodeHref(encoded)); } catch { return whole; }
    const source = sourceLabel(url.hostname);
    const text = label.replace(/<[^>]*>/g, '').toLowerCase();
    if (!source || text.includes(source) || text.includes(url.hostname.toLowerCase())) return whole;
    return `${whole} <span style="font-size:12px;color:#616161">· ${escapeHtml(source)}</span>`;
  });
}

// Second-level labels under which the registrable name has three parts (gov.br, co.uk, com.br ...).
const SHARED_SUFFIXES = new Set(['ac', 'co', 'com', 'edu', 'gov', 'net', 'org']);

/** `access.redhat.com` -> `redhat.com`, `www.gov.br` -> `gov.br`; IP addresses stay as they are. */
export function sourceLabel(hostname: string): string {
  const host = hostname.toLowerCase().replace(/\.$/, '').replace(/^www\./, '');
  if (!host || /^[\d.]+$/.test(host) || host.includes(':')) return host;
  const parts = host.split('.');
  if (parts.length <= 2) return host;
  const keep = SHARED_SUFFIXES.has(parts.at(-2)!) && parts.at(-1)!.length === 2 ? 3 : 2;
  return parts.slice(-keep).join('.');
}

export interface SendAttachmentOptions {
  conversationId: string;
  /** Optional markdown text shown above the attachment. */
  text?: string;
  image?: UploadedImage & { format: string };
  /** An entry from `uploadFile`. */
  file?: Record<string, unknown>;
  /** Sender name; Teams leaves it empty on attachment messages that do not carry it. */
  displayName: string | null;
}

/** Sends an image or file uploaded with `uploadImage` or `uploadFile`. */
export async function sendAttachment(client: TeamsClient, endpoints: Endpoints, { conversationId, text, image, file, displayName }: SendAttachmentOptions): Promise<{ clientMessageId: string; message: Message | null }> {
  const clientMessageId = newClientMessageId();
  const imageHtml = image
    ? `<p><img src="${escapeHtml(image.url)}" itemtype="http://schema.skype.com/AMSImage" itemscope="${escapeHtml(image.format)}" alt="image" id="x_${escapeHtml(image.id)}" itemid="${escapeHtml(image.id)}"></p>`
    : '';
  await client.json('chatService', `${conversationUrl(endpoints, conversationId)}/messages`, {
    method: 'POST',
    body: {
      content: `${text ? toTeamsHtml(text) : ''}${imageHtml}`,
      messagetype: 'RichText/Html',
      contenttype: 'text',
      clientmessageid: clientMessageId,
      ...(displayName ? { imdisplayname: displayName } : {}),
      ...(image ? { amsreferences: [image.id] } : {}),
      ...(file ? { properties: { files: JSON.stringify([file]) } } : {}),
    },
  });
  const message = await firstMessage(client, endpoints, conversationId, (item) => item.clientMessageId === clientMessageId);
  return { clientMessageId, message };
}

export async function editText(client: TeamsClient, endpoints: Endpoints, message: Message, text: string): Promise<Message | null> {
  await client.json('chatService', `${conversationUrl(endpoints, message.conversationId)}/messages/${encodeURIComponent(message.id)}`, {
    method: 'PUT',
    body: messageBody(text, newClientMessageId()),
  });
  const threadId = message.parentId === null ? message.conversationId : `${message.conversationId};messageid=${message.parentId}`;
  return firstMessage(client, endpoints, threadId, (item) => item.id === message.id);
}

// shortcut: channel reply mutations use stored conversationId; verify the path with D6.
export async function deleteMessage(client: TeamsClient, endpoints: Endpoints, message: Message): Promise<Message | null> {
  await client.json('chatService', `${conversationUrl(endpoints, message.conversationId)}/messages/${encodeURIComponent(message.id)}?behavior=softDelete`, {
    method: 'DELETE',
  });
  const threadId = message.parentId === null ? message.conversationId : `${message.conversationId};messageid=${message.parentId}`;
  return firstMessage(client, endpoints, threadId, (item) => item.id === message.id);
}

async function firstMessage(
  client: TeamsClient,
  endpoints: Endpoints,
  conversationId: string,
  matches: (message: Message) => boolean,
): Promise<Message | null> {
  const page = await listMessages(client, endpoints, conversationId).next();
  return page.value?.messages.find(matches) ?? null;
}

function messageBody(text: string, clientMessageId: string, quote?: QuotedMessage, mentions?: PersonMention[], compact = false) {
  const options = compact ? { compact } : {};
  const composed = mentions?.length ? toTeamsHtmlWithMentions(text, mentions, options) : null;
  return {
    content: `${quote ? quoteHtml(quote) : ''}${composed?.content ?? toTeamsHtml(text, options)}`,
    messagetype: 'RichText/Html',
    contenttype: 'text',
    clientmessageid: clientMessageId,
    ...(composed ? { properties: { mentions: JSON.stringify(composed.mentions) } } : {}),
  };
}

function conversationUrl(endpoints: Endpoints, conversationId: string): string {
  const marker = ';messageid=';
  const separator = conversationId.lastIndexOf(marker);
  const path = separator < 0
    ? encodeURIComponent(conversationId)
    : `${encodeURIComponent(conversationId.slice(0, separator))}${marker}${encodeURIComponent(conversationId.slice(separator + marker.length))}`;
  return `${endpoints.chatService}/v1/users/ME/conversations/${path}`;
}

function newClientMessageId(): string {
  return `${randomInt(1_000_000_000).toString().padStart(9, '0')}${randomInt(1_000_000_000).toString().padStart(9, '0')}`;
}

const HYPERLINK = 'http://schema.skype.com/HyperLink';
const IMAGE_FORMATS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp']);

async function withLinkPreviews(client: TeamsClient, content: string, endpoints: Endpoints, previewLink: string | undefined, fetcher: typeof fetch | undefined): Promise<{ content: string; links: Record<string, unknown>[] }> {
  const urls = httpUrls(content);
  const selected = previewLink === undefined ? urls[0] : urls.find((url) => url === previewLink);
  if (previewLink !== undefined && selected === undefined) throw new CliError('usage', '--preview-link must exactly match an http(s) URL in the message.');
  if (!selected) return { content, links: [] };
  const meta = await linkMeta(client, selected, endpoints, fetcher);
  const links = [{
    '@type': HYPERLINK,
    itemid: '0',
    url: selected,
    previewenabled: true,
    preview: {
      ...(meta.previewurl ? { previewurl: meta.previewurl } : {}),
      title: meta.title ?? selected,
      description: meta.description ?? '',
    },
  }];
  const next = content.replace(/<a href="([^"]+)">/g, (whole, href: string) => decodeHref(href) === selected
    ? `<a href="${href}" itemtype="${HYPERLINK}" rel="noreferrer noopener" target="_blank" itemid="0">`
    : whole);
  return { content: next, links };
}

function httpUrls(content: string): string[] {
  const urls: string[] = [];
  for (const match of content.matchAll(/<a href="([^"]+)">/g)) {
    const url = decodeHref(match[1]!);
    if (/^https?:\/\//i.test(url) && !urls.includes(url)) urls.push(url);
  }
  return urls;
}

interface LinkMetadata { title?: string; description?: string; image?: string; previewurl?: string }

async function linkMeta(client: TeamsClient, page: string, endpoints: Endpoints, fetcher: typeof fetch | undefined): Promise<LinkMetadata> {
  // shortcut: URLP's JSON fields and chatsvcagg auth are inferred from the service shape; a live card still needs a Teams post.
  if (endpoints.urlp) {
    try {
      const raw = await client.json<Record<string, unknown> | null>('urlp', `${endpoints.urlp.replace(/\/+$/, '')}/v1/url/info?url=${encodeURIComponent(page)}`);
      if (raw && typeof raw === 'object') {
        const parsed = parseInfo(raw, page);
        if (parsed.title || parsed.description || parsed.image || parsed.previewurl) return withThumbnail(parsed, endpoints.urlp);
      }
    } catch {
      // Public Open Graph is the fallback when URLP does not answer.
    }
  }
  if (!fetcher) return {};
  return withThumbnail(await pageMeta(fetcher, page), endpoints.urlp);
}

function parseInfo(raw: Record<string, unknown>, page: string): LinkMetadata {
  const thumbnail = resourceUrl(raw, 'thumbnail', 'previewurl', 'previewUrl', 'thumbnailUrl');
  const image = absoluteHttp(resourceUrl(raw, 'image', 'imageUrl', 'imageurl'), page);
  if (thumbnail && teamsThumbnail(thumbnail)) {
    return {
      title: plain(raw, 'title'),
      description: plain(raw, 'description'),
      image: image ?? absoluteHttp(new URL(thumbnail).searchParams.get('url') ?? undefined, page),
      previewurl: thumbnail,
    };
  }
  return { title: plain(raw, 'title'), description: plain(raw, 'description'), image: image ?? absoluteHttp(thumbnail, page) };
}

/** Teams loads the card image from its own thumbnail service. A preview URL on any other host is not used. */
function withThumbnail(meta: LinkMetadata, base: string | null): LinkMetadata {
  if (meta.previewurl || !meta.image || !base) return meta;
  return { ...meta, previewurl: `${base.replace(/\/+$/, '')}/v1/url/image/Thumbnail?url=${encodeURIComponent(meta.image)}` };
}

function teamsThumbnail(value: string): boolean {
  try {
    const url = new URL(value);
    const hostOk = url.protocol === 'https:' && (url.hostname === 'urlp.asm.skype.com' || /^[a-z0-9-]+\.asyncgw\.teams\.microsoft\.com$/.test(url.hostname));
    return hostOk && /\/v1\/url\/image\/Thumbnail$/i.test(url.pathname);
  } catch {
    return false;
  }
}

function plain(source: Record<string, unknown>, ...names: string[]): string | undefined {
  for (const name of names) {
    const value = source[name];
    if (typeof value === 'string' && value.trim()) return value.trim();
  }
  return undefined;
}

function resourceUrl(source: Record<string, unknown>, ...names: string[]): string | undefined {
  for (const name of names) {
    const value = source[name];
    if (typeof value === 'string' && value.trim()) return value.trim();
    if (value && typeof value === 'object' && 'url' in value && typeof value.url === 'string' && value.url.trim()) return value.url.trim();
  }
  return undefined;
}

function absoluteHttp(value: string | undefined, base: string): string | undefined {
  if (!value) return undefined;
  try {
    const url = new URL(value, base);
    if ((url.protocol !== 'http:' && url.protocol !== 'https:') || url.username || url.password) return undefined;
    return url.toString();
  } catch {
    return undefined;
  }
}

async function pageMeta(fetcher: typeof fetch, url: string): Promise<LinkMetadata> {
  try {
    const response = await fetcher(url, { headers: { 'user-agent': 'Mozilla/5.0 (compatible; teamscli link preview)' }, signal: AbortSignal.timeout(8000), redirect: 'follow' });
    const html = (await response.text()).slice(0, 300_000);
    const pick = (...names: string[]) => {
      for (const tag of html.matchAll(/<meta\b[^>]*>/gi)) {
        const attrs = tagAttributes(tag[0]);
        const key = (attrs.property ?? attrs.name ?? '').toLowerCase();
        if (names.includes(key) && attrs.content) return attrs.content;
      }
      return undefined;
    };
    const image = absoluteHttp(pick('og:image', 'twitter:image'), response.url || url);
    const title = pick('og:title', 'twitter:title') ?? (decode(html.match(/<title[^>]*>([^<]*)<\/title>/i)?.[1] ?? '').trim() || undefined);
    return { title, description: pick('og:description', 'description', 'twitter:description'), ...(image ? { image } : {}) };
  } catch {
    return {};
  }
}

function tagAttributes(tag: string): Record<string, string> {
  const result: Record<string, string> = {};
  for (const match of tag.matchAll(/([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g)) {
    result[match[1]!.toLowerCase()] = decode(match[2] ?? match[3] ?? match[4] ?? '');
  }
  return result;
}

async function loadMarkdownImage(image: MarkdownImageRef, fetcher: typeof fetch | undefined, readFile: SendTextOptions['readFile']): Promise<{ name: string; format: string; data: Uint8Array<ArrayBuffer> }> {
  if (image.kind === 'file') {
    if (!readFile) throw new CliError('usage', 'Could not read a local image from the message.');
    let data: Uint8Array<ArrayBuffer>;
    try {
      data = readFile(image.target);
    } catch (error) {
      if (error instanceof CliError) throw error;
      throw new CliError('usage', 'Could not read a local image from the message.');
    }
    const format = imageFormat(image.target, '', data);
    if (!format) throw new CliError('usage', 'Markdown images must be png, jpg, jpeg, gif or webp files.');
    return { name: imageName(image.target, format), format, data: new Uint8Array(data) };
  }
  if (!fetcher) throw new CliError('remote', 'Could not download a markdown image.');
  let response: Response;
  try {
    response = await fetcher(image.target, { headers: { 'user-agent': 'Mozilla/5.0 (compatible; teamscli)' }, redirect: 'follow', signal: AbortSignal.timeout(8000) });
  } catch {
    throw new CliError('remote', 'Could not download a markdown image.');
  }
  if (!response.ok || (response.url && !/^https?:/i.test(response.url))) throw new CliError('remote', 'Could not download a markdown image.');
  const data = new Uint8Array(await response.arrayBuffer());
  if (data.byteLength === 0 || data.byteLength > 8_000_000) throw new CliError('usage', 'A markdown image must be a png, jpg, jpeg, gif or webp file under 8 MB.');
  const format = imageFormat(image.target, response.headers.get('content-type') ?? '', data);
  if (!format) throw new CliError('usage', 'A markdown image must be a png, jpg, jpeg, gif or webp file.');
  return { name: imageName(image.target, format), format, data };
}

function imageFormat(name: string, contentType: string, data: Uint8Array): string | null {
  const type = contentType.split(';', 1)[0]?.trim().toLowerCase() ?? '';
  const fromType = type === 'image/png' ? 'png' : type === 'image/gif' ? 'gif' : type === 'image/webp' ? 'webp' : type === 'image/jpeg' || type === 'image/jpg' ? 'jpeg' : null;
  if (fromType === 'jpeg') return /\.jpg(?:$|[?#])/i.test(name) ? 'jpg' : 'jpeg';
  if (fromType) return fromType;
  if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return 'png';
  if (data[0] === 0xff && data[1] === 0xd8) return 'jpg';
  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) return 'gif';
  if (data.length > 11 && data[0] === 0x52 && data[1] === 0x49 && data[2] === 0x46 && data[8] === 0x57) return 'webp';
  const ext = (name.split(/[?#]/)[0] ?? '').split('.').at(-1)?.toLowerCase() ?? '';
  return IMAGE_FORMATS.has(ext) ? ext : null;
}

function imageName(target: string, format: string): string {
  let base = 'image';
  try {
    base = /^https?:/i.test(target) ? new URL(target).pathname.split('/').at(-1) || 'image' : target.split(/[\\/]/).at(-1) || 'image';
  } catch {
    base = 'image';
  }
  const clean = base.replace(/[^A-Za-z0-9._-]/g, '').slice(0, 80);
  return IMAGE_FORMATS.has(clean.split('.').at(-1)?.toLowerCase() ?? '') ? clean : `${clean || 'image'}.${format}`;
}

function decodeHref(value: string): string {
  return decode(value);
}

function decode(value: string): string {
  return value.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>');
}
