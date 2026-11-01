import { readFileSync, statSync } from 'node:fs';
import { basename, extname } from 'node:path';
import { CliError } from '../../errors.ts';
import type { Message } from '../../models.ts';
import { readIdentity } from '../../auth/identity.ts';
import { databasePath } from '../../runtime/paths.ts';
import { resolveConversation, resolvePerson } from '../../store/resolve.ts';
import { openDatabase } from '../../store/database.ts';
import { getMessage, upsertMessages } from '../../store/messages.ts';
import { getPersonName } from '../../store/people.ts';
import { createTeamsClient } from '../../teams/client.ts';
import { sendText } from '../../teams/send.ts';
import type { Command } from '../command.ts';
import { renderHuman, renderMessageDetail } from '../output.ts';
import { teamsSession } from '../session.ts';

type SendResult = Message | { sent: true; clientMessageId: string };

const command: Command<SendResult> = {
  name: 'send text',
  summary: 'Send a text message to a chat or reply in a channel thread',
  usage: '(--to TARGET | --chat ID | --channel ID) [TEXT | --message TEXT] [--mention PERSON] [--reply-to ID] [--quote ID] [--subject TEXT] [--link-previews] [--preview-link URL] [--image PATH] [--link-icons] [--link-sources] [--compact] [--radar]',
  arguments: '[TEXT]  Optional message text after the destination options',
  options: {
    to: { type: 'string' },
    chat: { type: 'string' },
    channel: { type: 'string' },
    message: { type: 'string' },
    'reply-to': { type: 'string' },
    quote: { type: 'string' },
    mention: { type: 'string', multiple: true },
    subject: { type: 'string' },
    'link-previews': { type: 'boolean' },
    'preview-link': { type: 'string' },
    image: { type: 'string', multiple: true },
    'link-icons': { type: 'boolean' },
    'link-sources': { type: 'boolean' },
    compact: { type: 'boolean' },
    radar: { type: 'boolean' },
  },
  optionHelp: {
    to: { valueName: 'TARGET', description: 'Chat or person to message by id or name.' },
    chat: { valueName: 'ID', description: 'Chat to message (id, short id or name).' },
    channel: { valueName: 'ID', description: 'Channel to post or reply in (id, short id or name).' },
    message: { valueName: 'TEXT', description: 'Message text; alternatively pass it after the options.' },
    'reply-to': { valueName: 'ID', description: 'Reply to this channel message.' },
    quote: { valueName: 'ID', description: 'Quote this locally stored chat message.' },
    mention: { valueName: 'PERSON', description: 'Mention a person from the local data; may be repeated.' },
    subject: { valueName: 'TEXT', description: 'Subject (large title) of a new channel post.' },
    'link-previews': { description: 'Attach one link preview card. Teams shows a single card; --preview-link chooses which.' },
    'preview-link': { valueName: 'URL', description: 'Choose which link gets the single preview card; must match a link in the text.' },
    image: { valueName: 'PATH', description: 'Upload an inline image; may be repeated.' },
    'link-icons': { description: 'Add each site favicon as a small inline image before its links.' },
    'link-sources': { description: 'Show the site (e.g. github.com) in small grey text after each link.' },
    compact: { description: 'Tighter spacing: no rules, smaller headings, empty sections hidden; text is never cut.' },
    radar: { description: 'Newsletter preset: --compact, --link-icons and --link-sources together.' },
  },
  maxPositionals: 1,
  mutates: true,
  async run(ctx, { positionals, values }) {
    if (ctx.options.events) throw new CliError('usage', 'send text does not support --events.');
    const to = option(values, 'to');
    const chat = option(values, 'chat');
    const channel = option(values, 'channel');
    const positionalText = positionals[0];
    const optionText = option(values, 'message');
    const replyTo = option(values, 'reply-to');
    const quoteId = option(values, 'quote');
    const previewLink = option(values, 'preview-link');
    const imagePaths = stringOptions(values, 'image');
    const targets = [to, chat, channel].filter((value) => value !== undefined);
    if (targets.length === 0) throw new CliError('usage', 'Missing the destination: pass --to TARGET, --chat ID or --channel ID (id, short id or name).');
    if (targets.length > 1) throw new CliError('usage', 'Choose one destination: pass exactly one of --to, --chat or --channel.');
    if (positionalText !== undefined && optionText !== undefined) {
      throw new CliError('usage', 'Pass the message text either after the options or with --message, not both.');
    }
    const messageText = optionText ?? positionalText;
    if (!messageText?.trim()) throw new CliError('usage', 'Missing the message text: pass it after the options or with --message.');
    if (replyTo === '' || quoteId === '') throw new CliError('usage', '--reply-to and --quote must not be empty.');
    if (previewLink === '') throw new CliError('usage', '--preview-link must not be empty.');
    if (previewLink !== undefined && values['link-previews'] !== true) throw new CliError('usage', '--preview-link requires --link-previews.');
    const images = imagePaths.map((path) => readInlineImage(path));

    const db = openDatabase(databasePath(ctx.home, ctx.account));
    try {
      const session = teamsSession(ctx);
      const identity = readIdentity(ctx.home, ctx.account);
      const query = to ?? chat ?? channel ?? '';
      let selfId = identity?.userId ?? null;
      if (to !== undefined && selfId === null) selfId = (await session.account()).userId;
      const conversation = resolveConversation(db, query, selfId);
      const mentions = stringOptions(values, 'mention').map((mention) => {
        const person = resolvePerson(db, mention);
        if (!person.displayName?.trim()) throw new CliError('usage', 'The person has no display name in local data; sync the chat or choose another person.');
        return { query: mention, mri: person.id, displayName: person.displayName };
      });
      if (channel !== undefined && conversation.kind !== 'channel') throw new CliError('usage', '--channel must name a channel in the local data.');
      if (chat !== undefined && conversation.kind === 'channel') throw new CliError('usage', '--chat must name a chat in the local data.');

      if (conversation.kind !== 'channel' && replyTo !== undefined) {
        throw new CliError('usage', '--reply-to is only supported for channel messages.');
      }
      if (conversation.kind === 'channel' && quoteId !== undefined) {
        throw new CliError('usage', '--quote is only supported in chats.');
      }

      const reply = replyTo === undefined ? null : getMessage(db, conversation.id, replyTo);
      if (replyTo !== undefined && !reply) throw new CliError('not_found', 'Reply target not found in the local channel.');
      const quoted = quoteId === undefined ? null : getMessage(db, conversation.id, quoteId);
      if (quoteId !== undefined && !quoted) throw new CliError('not_found', 'Quoted message not found in the local chat.');
      if (quoted && (!quoted.senderId || !quoted.senderName)) {
        throw new CliError('usage', 'The quoted message has no sender details in local data.');
      }

      if (quoted) ctx.log('The quoted message is not verified by Teams.');
      const client = createTeamsClient({ fetch: ctx.fetch, session });
      const result = await sendText(client, await session.endpoints(), {
        conversationId: conversation.id,
        ...(reply ? { replyTo: reply.parentId ?? reply.id } : {}),
        channelPost: conversation.kind === 'channel' && replyTo === undefined,
        text: messageText,
        ...(mentions.length > 0 ? { mentions } : {}),
        ...(option(values, 'subject') ? { subject: option(values, 'subject') } : {}),
        ...(values['link-previews'] === true ? { linkPreviews: true } : {}),
        ...(previewLink ? { previewLink } : {}),
        ...(values['link-icons'] === true || values.radar === true ? { linkIcons: true } : {}),
        ...(values['link-sources'] === true || values.radar === true ? { linkSources: true } : {}),
        ...(values.compact === true || values.radar === true ? { compact: true } : {}),
        fetch: ctx.fetch,
        readFile: (path) => new Uint8Array(readFileSync(path)),
        ...(images.length > 0 ? { images } : {}),
        ...(quoted ? { quote: { id: quoted.id, senderId: quoted.senderId!, senderName: quoted.senderName!, text: quoted.text } } : {}),
      });
      if (!result.message) {
        ctx.log('Teams accepted the message, but it was not found in the first page; the local store was not updated.');
        return { sent: true, clientMessageId: result.clientMessageId };
      }
      let message = result.message;
      if (message.senderName === null) {
        let senderName = identity?.displayName || null;
        if (senderName === null) {
          const account = await session.account();
          senderName = getPersonName(db, account.userId) ?? account.displayName;
        }
        if (senderName !== null) message = { ...message, senderName };
      }
      upsertMessages(db, [message]);
      return getMessage(db, conversation.id, message.id) ?? message;
    } finally {
      db.close();
    }
  },
  render: (result) => 'sent' in result ? renderHuman(result) : renderMessageDetail(result),
};

function readInlineImage(path: string): { name: string; format: string; data: Uint8Array<ArrayBuffer> } {
  try {
    if (!statSync(path).isFile()) throw new CliError('usage', '--image must point to a regular file.');
    const name = basename(path);
    const format = extname(name).slice(1).toLowerCase();
    if (!new Set(['png', 'jpg', 'jpeg', 'gif', 'webp']).has(format)) throw new CliError('usage', '--image supports png, jpg, jpeg, gif or webp files.');
    return { name, format, data: new Uint8Array(readFileSync(path)) };
  } catch (error) {
    if (error instanceof CliError) throw error;
    throw new CliError('usage', `Could not read --image: ${(error as NodeJS.ErrnoException).code ?? 'unknown error'}.`);
  }
}

function option(values: Record<string, string | boolean | string[] | boolean[] | undefined>, name: string): string | undefined {
  const value = values[name];
  return typeof value === 'string' ? value : undefined;
}

function stringOptions(values: Record<string, string | boolean | string[] | boolean[] | undefined>, name: string): string[] {
  const value = values[name];
  if (typeof value === 'string') return [value];
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === 'string') : [];
}

export default command;
