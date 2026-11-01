import { randomInt } from 'node:crypto';
import { CliError } from '../errors.ts';
import type { Message } from '../models.ts';
import { listMessages } from './messages.ts';
import type { Endpoints, TeamsClient } from './client.ts';
import { escapeHtml, quoteHtml, toTeamsHtml, toTeamsHtmlWithMentions, type PersonMention, type QuotedMessage } from './compose.ts';
import { uploadImage, type UploadedImage } from './files.ts';
import { fetchLinkIcon } from './link-icons.ts';

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
  fetch?: typeof fetch;
  images?: Array<{ name: string; format: string; data: Uint8Array<ArrayBuffer> }>;
}

export async function sendText(client: TeamsClient, endpoints: Endpoints, { conversationId, replyTo, channelPost, text, quote, mentions, subject, linkPreviews, previewLink, linkIcons, fetch: fetcher, images }: SendTextOptions): Promise<{ clientMessageId: string; message: Message | null }> {
  const clientMessageId = newClientMessageId();
  const targetId = replyTo === undefined ? conversationId : `${conversationId};messageid=${replyTo}`;
  const body: Record<string, any> = messageBody(text, clientMessageId, quote, mentions);
  const properties: Record<string, unknown> = { ...(body.properties ?? {}) };
  if (subject && channelPost && replyTo === undefined) properties.subject = subject;
  if (images?.length) {
    const uploaded = await Promise.all(images.map((image) => uploadImage(client, conversationId, image.name, image.data).then((item) => ({ ...item, format: image.format }))));
    body.content += uploaded.map((image) => inlineImageHtml(image)).join('');
    body.amsreferences = [...(body.amsreferences ?? []), ...uploaded.map((image) => image.id)];
  }
  if (linkIcons && fetcher) {
    const { content, imageIds } = await withLinkIcons(client, body.content, conversationId, fetcher);
    body.content = content;
    if (imageIds.length > 0) body.amsreferences = [...(body.amsreferences ?? []), ...imageIds];
  }
  if (linkPreviews) {
    const { content, links } = await withLinkPreviews(client, body.content, endpoints, previewLink, fetcher ?? fetch);
    body.content = content;
    if (links.length > 0) properties.links = JSON.stringify(links);
  }
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

function inlineImageHtml(image: UploadedImage & { format: string }): string {
  return `<p><img src="${escapeHtml(image.url)}" itemtype="http://schema.skype.com/AMSImage" itemscope="${escapeHtml(image.format)}" alt="image" id="x_${escapeHtml(image.id)}" itemid="${escapeHtml(image.id)}"></p>`;
}

async function withLinkIcons(client: TeamsClient, content: string, conversationId: string, fetcher: typeof fetch): Promise<{ content: string; imageIds: string[] }> {
  const anchors = [...content.matchAll(/<a href="([^"]+)">/g)];
  const domains = new Map<string, Promise<(UploadedImage & { format: string }) | null>>();
  const byUrl = new Map<string, UploadedImage & { format: string }>();
  const imageIds: string[] = [];
  for (const [, encoded] of anchors) {
    const url = encoded!.replace(/&amp;/g, '&');
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
    const image = byUrl.get(encoded.replace(/&amp;/g, '&'));
    if (!image) return whole;
    return `<img src="${escapeHtml(image.url)}" itemtype="http://schema.skype.com/AMSImage" itemscope="${image.format}" alt="" id="x_${escapeHtml(image.id)}" itemid="${escapeHtml(image.id)}" width="18" height="18" style="vertical-align:middle">${whole}`;
  });
  return { content: next, imageIds };
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

function messageBody(text: string, clientMessageId: string, quote?: QuotedMessage, mentions?: PersonMention[]) {
  const composed = mentions?.length ? toTeamsHtmlWithMentions(text, mentions) : null;
  return {
    content: `${quote ? quoteHtml(quote) : ''}${composed?.content ?? toTeamsHtml(text)}`,
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

async function withLinkPreviews(client: TeamsClient, content: string, endpoints: Endpoints, previewLink: string | undefined, fetcher: typeof fetch): Promise<{ content: string; links: Record<string, unknown>[] }> {
  const anchors = [...content.matchAll(/<a href="([^"]+)">/g)];
  const urls = [...new Set(anchors.map(([, href]) => href!.replace(/&amp;/g, '&')).filter((url) => /^https?:/i.test(url)))];
  const selected = previewLink === undefined ? urls[0] : urls.find((url) => url === previewLink);
  if (previewLink !== undefined && selected === undefined) throw new CliError('usage', '--preview-link must exactly match an http(s) URL in the message.');
  if (!selected) return { content, links: [] };
  const meta = await linkMeta(client, selected, endpoints, fetcher);
  const previewurl = meta.previewurl ?? (meta.image
    ? `${thumbnailBase(endpoints)}?url=${encodeURIComponent(meta.image)}${meta.token ? `&token=${encodeURIComponent(meta.token)}` : ''}`
    : undefined);
  const itemid = '0';
  const links = [{
    '@type': 'http://schema.skype.com/HyperLink',
    itemid,
    url: selected,
    previewenabled: true,
    preview: { ...(previewurl ? { previewurl } : {}), title: meta.title ?? selected, description: meta.description ?? '' },
  }];
  const escapedSelected = selected.replace(/&/g, '&amp;');
  const next = content.replace(/<a href="([^"]+)">/g, (whole, href: string) => href === escapedSelected
    ? `<a href="${href}" itemtype="http://schema.skype.com/HyperLink" rel="noreferrer noopener" target="_blank" itemid="${itemid}">`
    : whole);
  return { content: next, links };
}

function thumbnailBase(endpoints: Endpoints): string {
  const host = new URL(endpoints.chatService).hostname.split('.')[0]!.replace(/-.*$/, '') || 'us';
  return `https://${host}-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail`;
}

interface LinkMetadata { title?: string; description?: string; image?: string; previewurl?: string; token?: string }

async function linkMeta(client: TeamsClient, url: string, endpoints: Endpoints, fetcher: typeof fetch): Promise<LinkMetadata> {
  try {
    const region = new URL(endpoints.chatService).hostname.split('.')[0]!.replace(/-.*$/, '') || 'us';
    const infoUrl = `https://${region}-prod.asyncgw.teams.microsoft.com/urlp/v1/url/info?url=${encodeURIComponent(url)}`;
    const raw = await client.json<Record<string, unknown> | null>('urlp', infoUrl);
    if (raw) {
      const get = (...names: string[]) => names.map((name) => raw[name]).find((value) => typeof value === 'string') as string | undefined;
      const result: LinkMetadata = {
        title: get('title'),
        description: get('description'),
        image: get('image', 'imageUrl', 'imageurl', 'thumbnail'),
        previewurl: get('previewurl', 'previewUrl', 'thumbnailUrl', 'thumbnailurl'),
        token: get('imageToken', 'imageAuthToken', 'authToken', 'token'),
      };
      if (result.title || result.description || result.image || result.previewurl) return result;
    }
  } catch {
  }
  return pageMeta(fetcher, url);
}

async function pageMeta(fetcher: typeof fetch, url: string): Promise<LinkMetadata> {
  try {
    const response = await fetcher(url, { headers: { 'user-agent': 'Mozilla/5.0 (compatible; teamscli link preview)' }, signal: AbortSignal.timeout(8000), redirect: 'follow' });
    const html = (await response.text()).slice(0, 300_000);
    const pick = (...names: string[]) => {
      for (const name of names) {
        const tag = html.match(new RegExp(`<meta[^>]+(?:property|name)=["']${name}["'][^>]*>`, 'i'))?.[0];
        const value = tag?.match(/content=["']([^"']*)["']/i)?.[1];
        if (value) return decode(value);
      }
      return undefined;
    };
    const image = pick('og:image', 'twitter:image');
    return {
      title: pick('og:title', 'twitter:title') ?? (decode(html.match(/<title[^>]*>([^<]*)<\/title>/i)?.[1] ?? '').trim() || undefined),
      description: pick('og:description', 'description', 'twitter:description'),
      image: image ? new URL(image, response.url || url).toString() : undefined,
    };
  } catch {
    return {};
  }
}

function decode(value: string): string {
  return value.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>');
}
