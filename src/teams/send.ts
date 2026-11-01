import { randomInt } from 'node:crypto';
import type { Message } from '../models.ts';
import { listMessages } from './messages.ts';
import type { Endpoints, TeamsClient } from './client.ts';
import { escapeHtml, quoteHtml, toTeamsHtml, toTeamsHtmlWithMentions, type PersonMention, type QuotedMessage } from './compose.ts';
import type { UploadedImage } from './files.ts';

export interface SendTextOptions {
  conversationId: string;
  replyTo?: string;
  channelPost?: boolean;
  text: string;
  quote?: QuotedMessage;
  mentions?: PersonMention[];
  subject?: string;
  linkPreviews?: boolean;
}

export async function sendText(client: TeamsClient, endpoints: Endpoints, { conversationId, replyTo, channelPost, text, quote, mentions, subject, linkPreviews }: SendTextOptions): Promise<{ clientMessageId: string; message: Message | null }> {
  const clientMessageId = newClientMessageId();
  const targetId = replyTo === undefined ? conversationId : `${conversationId};messageid=${replyTo}`;
  const body: Record<string, any> = messageBody(text, clientMessageId, quote, mentions);
  const properties: Record<string, unknown> = { ...(body.properties ?? {}) };
  if (subject && channelPost && replyTo === undefined) properties.subject = subject;
  if (linkPreviews) {
    const { content, links } = await withLinkPreviews(body.content, endpoints);
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

async function withLinkPreviews(content: string, endpoints: Endpoints): Promise<{ content: string; links: Record<string, unknown>[] }> {
  const links: Record<string, unknown>[] = [];
  const seen = new Map<string, string>();
  const anchors = [...content.matchAll(/<a href="([^"]+)">/g)];
  for (const [, href] of anchors) {
    const url = href!.replace(/&amp;/g, '&');
    if (seen.has(url) || !/^https?:/i.test(url)) continue;
    const itemid = String(links.length);
    seen.set(url, itemid);
    const meta = await pageMeta(url);
    links.push({
      '@type': 'http://schema.skype.com/HyperLink',
      itemid,
      url,
      previewenabled: true,
      preview: {
        ...(meta.image ? { previewurl: `${thumbnailBase(endpoints)}?url=${encodeURIComponent(meta.image)}` } : {}),
        title: meta.title ?? url,
        description: meta.description ?? '',
      },
    });
  }
  const next = content.replace(/<a href="([^"]+)">/g, (whole, href: string) => {
    const itemid = seen.get(href.replace(/&amp;/g, '&'));
    return itemid === undefined ? whole : `<a href="${href}" itemtype="http://schema.skype.com/HyperLink" rel="noreferrer noopener" target="_blank" itemid="${itemid}">`;
  });
  return { content: next, links };
}

function thumbnailBase(endpoints: Endpoints): string {
  const host = new URL(endpoints.chatService).hostname.split('.')[0]!.replace(/-.*$/, '') || 'us';
  return `https://${host}-prod.asyncgw.teams.microsoft.com/urlp/v1/url/image/Thumbnail`;
}

async function pageMeta(url: string): Promise<{ title?: string; description?: string; image?: string }> {
  try {
    const response = await fetch(url, { headers: { 'user-agent': 'Mozilla/5.0 (compatible; teamscli link preview)' }, signal: AbortSignal.timeout(8000), redirect: 'follow' });
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
