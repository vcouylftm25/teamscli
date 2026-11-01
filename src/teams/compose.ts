// Builds the HTML body Teams expects (`messagetype: RichText/Html`) from user text. Nothing the user
// types reaches the output as markup: every segment is escaped, and only the markdown below is rendered.

/**
 * Plain text or markdown to Teams HTML, in the shape the Teams client itself produces. Blocks: blank
 * lines separate paragraphs and single newlines become `<br>`; `#` headings (`#`/`##` -> h2, deeper
 * -> h3), `>` quotes, `-`/`*`/`+` and `1.` lists (nested by indentation), `|` tables with a `---`
 * separator row and fenced code blocks with an optional language. A line of `---`, `***` or `___` is a
 * horizontal rule. Inline: `**bold**`, `*italic*`, `` `code` ``, `[text](url)` and bare URLs; links only with
 * http or https (anything else stays literal text). Markdown images are not left as external `<img>` tags:
 * `maskMarkdownImages` lifts them out so the sender can upload them.
 */
export function toTeamsHtml(text: string, options: ComposeOptions = {}): string {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const compact = options.compact === true;
  return blocks(compact ? dropEmptySections(lines) : lines, { codeBlocks: 0, compact });
}

export interface ComposeOptions {
  /**
   * Tighter layout for long posts that Teams would fold behind "see more": no `---` rules, `#`/`##`
   * headings as h3, headings with nothing under them dropped, and a blank paragraph only before a
   * paragraph that opens with bold text (the title of the next item), never around quotes or links.
   */
  compact?: boolean;
}

interface BlockState {
  codeBlocks: number;
  compact?: boolean;
}

/** Removes headings whose section (up to the next heading of the same or a higher level) holds only blank lines or rules. */
export function dropEmptySections(lines: string[]): string[] {
  let current = lines;
  for (;;) {
    const fenced = fencedLines(current);
    const level = (index: number) => (fenced[index] ? 0 : HEADING.exec(current[index]!)?.[1]!.length ?? 0);
    const drop = new Set<number>();
    for (let index = 0; index < current.length; index++) {
      const own = level(index);
      if (own === 0) continue;
      let end = index + 1;
      while (end < current.length && !(level(end) > 0 && level(end) <= own)) end++;
      const body = current.slice(index + 1, end);
      if (body.every((line, offset) => !fenced[index + 1 + offset] && (!line.trim() || RULE.test(line)))) {
        for (let next = index; next < end; next++) drop.add(next);
      }
    }
    if (drop.size === 0) return current;
    current = current.filter((_, index) => !drop.has(index));
  }
}

function fencedLines(lines: string[]): boolean[] {
  let open = false;
  return lines.map((line) => {
    if (FENCE.test(line)) {
      if (!open) { open = true; return true; }
      if (FENCE.exec(line)![1] === '') { open = false; return true; }
    }
    return open;
  });
}

const FENCE = /^\s*```\s*([\w+#.-]*)\s*$/;
const HEADING = /^\s*(#{1,6})\s+(.*?)\s*#*\s*$/;
const QUOTE = /^\s*>\s?(.*)$/;
const ITEM = /^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$/;
const RULE = /^\s*(?:-{3,}|\*{3,}|_{3,})\s*$/;
const TABLE_ROW = /^\s*\|.*\|\s*$/;
const TABLE_SEPARATOR = /^\s*\|(?:\s*:?-{1,}:?\s*\|)+\s*$/;

function blocks(lines: string[], state: BlockState): string {
  let html = '';
  let index = 0;
  let gap = false;
  const take = (matches: (line: string) => boolean) => {
    const taken: string[] = [];
    while (index < lines.length && matches(lines[index]!)) taken.push(lines[index++]!);
    return taken;
  };

  while (index < lines.length) {
    const line = lines[index]!;
    const fence = FENCE.exec(line);
    const heading = HEADING.exec(line);
    const item = ITEM.exec(line);
    const rule = RULE.test(line);
    if (!line.trim()) {
      index++;
      continue;
    }
    // Teams renders consecutive paragraphs with no gap and writes a blank line as this empty paragraph;
    // lists, code blocks and the bottom of headings bring their own margin.
    const spaced = !fence && !item && !rule;
    if (state.compact) {
      // Only the start of a new item gets air: a paragraph opening with bold text after other content.
      if (gap && spaced && !heading && /^\s*\*\*/.test(line)) html += '<p>&nbsp;</p>';
      // Section headings after other content get air too, or Teams glues them to the previous link line.
      if (heading && html) html += '<p>&nbsp;</p>';
      gap = !heading && !rule;
    } else {
      if (gap && spaced) html += '<p>&nbsp;</p>';
      gap = spaced && !heading;
    }
    if (rule) {
      index++;
      if (!state.compact) html += '<hr>';
      gap = false;
    } else if (fence) {
      index++;
      const code = take((next) => !FENCE.test(next) || FENCE.exec(next)![1] !== '');
      index++; // closing fence, if any
      const id = ++state.codeBlocks;
      html += `<p itemtype="http://schema.skype.com/CodeBlockEditor" id="${id}">&nbsp;</p>`
        + `<pre class="language-${(fence[1] || 'plaintext').toLowerCase()}" itemid="${id}"><code>${code.map(escapeHtml).join('<br>')}</code></pre>`;
    } else if (heading) {
      index++;
      const tag = heading[1]!.length <= 2 && !state.compact ? 'h2' : 'h3';
      html += `<${tag}>${inline(heading[2]!)}</${tag}>`;
    } else if (QUOTE.test(line)) {
      const quoted = take((next) => QUOTE.test(next)).map((next) => QUOTE.exec(next)![1]!);
      html += `<blockquote>${blocks(quoted, state)}</blockquote>`;
    } else if (item) {
      html += list(take((next) => next.trim() !== '' && (ITEM.test(next) || indent(next) > indent(line))), state);
    } else if (TABLE_ROW.test(line) && TABLE_SEPARATOR.test(lines[index + 1] ?? '')) {
      const [header, , ...rows] = take((next) => TABLE_ROW.test(next));
      html += `<table><tbody>${row(header!, true)}${rows.map((next) => row(next, false)).join('')}</tbody></table>`;
    } else {
      const paragraph = take((next) => next.trim() !== '' && !startsBlock(next));
      if (paragraph.length === 0) paragraph.push(lines[index++]!); // a table row without separator
      html += `<p>${paragraph.map((next) => inline(next.trim())).join('<br>')}</p>`;
    }
  }
  return html;
}

function startsBlock(line: string): boolean {
  return FENCE.test(line) || HEADING.test(line) || QUOTE.test(line) || ITEM.test(line) || RULE.test(line) || TABLE_ROW.test(line);
}

/** Consecutive items at the first item's indentation; deeper lines belong to the item above them. */
function list(lines: string[], state: BlockState): string {
  const first = ITEM.exec(lines[0]!)!;
  const depth = first[1]!.length;
  const tag = /\d/.test(first[2]!) ? 'ol' : 'ul';
  const items: Array<{ text: string; children: string[] }> = [];
  for (const line of lines) {
    const item = ITEM.exec(line);
    if (item && item[1]!.length <= depth) items.push({ text: item[3]!, children: [] });
    else items.at(-1)!.children.push(line);
  }
  return `<${tag}>${items.map(({ text, children }) => `<li>${inline(text)}${children.length ? blocks(children, state) : ''}</li>`).join('')}</${tag}>`;
}

function row(line: string, header: boolean): string {
  const cells = line.trim().replace(/^\|/, '').replace(/\|$/, '').split('|');
  return `<tr>${cells.map((value) => {
    const content = inline(value.trim());
    const body = header && content ? `<strong>${content}</strong>` : content;
    // Teams shows cell text that sits in a paragraph; a bare text node in <td> does not.
    return `<td><p>${body || '&nbsp;'}</p></td>`;
  }).join('')}</tr>`;
}

export interface MarkdownImageRef {
  alt: string;
  /** http(s) URL or a local path with an image extension. */
  target: string;
  kind: 'remote' | 'file';
  /** Private-use placeholder left in the text; the sender replaces it after uploading. */
  marker: string;
}

const MARKDOWN_IMAGE = /!\[([^\]]*)\]\(([^)\s]+)\)/g;
const LOCAL_IMAGE = /\.(?:png|jpe?g|gif|webp)$/i;

/**
 * Replaces sendable markdown images with placeholders and leaves every other `![...](...)` as text.
 * Code fences and inline code are not images. With no images, the text is returned unchanged.
 */
export function maskMarkdownImages(text: string): { text: string; images: MarkdownImageRef[] } {
  const normalized = text.replace(/\r\n?/g, '\n');
  const images: MarkdownImageRef[] = [];
  let serial = 0;
  const markerFor = (index: number) => `\uE000teamscli-image-${serial}-${index}\uE001`;
  while (normalized.includes(markerFor(0))) serial++;

  let fenced = false;
  const lines = normalized.split('\n').map((line) => {
    if (FENCE.test(line)) {
      fenced = !fenced;
      return line;
    }
    if (fenced) return line;
    return line.split(/(`[^`]*`)/g).map((part, index) => {
      if (index % 2 === 1) return part;
      return part.replace(MARKDOWN_IMAGE, (whole, alt: string, target: string) => {
        const kind = imageKind(target);
        if (!kind) return whole;
        const marker = markerFor(images.length);
        images.push({ alt, target, kind, marker });
        return marker;
      });
    }).join('');
  });
  return images.length === 0 ? { text, images } : { text: lines.join('\n'), images };
}

function imageKind(target: string): MarkdownImageRef['kind'] | null {
  if (/^https?:\/\//i.test(target)) {
    try {
      const url = new URL(target);
      if ((url.protocol === 'http:' || url.protocol === 'https:') && !url.username && !url.password) return 'remote';
    } catch {
      return null;
    }
    return null;
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return null;
  const path = target.split(/[?#]/)[0] ?? target;
  return LOCAL_IMAGE.test(path) ? 'file' : null;
}

function indent(line: string): number {
  return line.length - line.trimStart().length;
}

export interface PersonMention {
  query: string;
  mri: string;
  displayName: string;
}

export interface MentionReference {
  '@type': 'http://schema.skype.com/Mention';
  itemid: string;
  mri: string;
  mentionType: 'person';
  displayName: string;
}

/** Renders mentions as Teams HTML and returns the references in content order. */
export function toTeamsHtmlWithMentions(text: string, people: PersonMention[], options: ComposeOptions = {}): { content: string; mentions: MentionReference[] } {
  const body = text.replace(/\r\n?/g, '\n');
  const normalized = normalizeWithOffsets(body);
  const ranges: Array<{ start: number; end: number; person: number }> = [];
  people.forEach((person, index) => {
    const candidates = [person.query, person.displayName]
      .flatMap((name) => mentionRanges(normalized, `@${name}`))
      .filter((range) => !ranges.some((used) => range.start < used.end && range.end > used.start))
      .sort((a, b) => a.start - b.start || b.end - a.end);
    const range = candidates[0];
    if (range) ranges.push({ ...range, person: index });
  });

  let markerPrefix = '\uE000teamscli-mention-0-';
  for (let nonce = 1; people.some((_, index) => body.includes(`${markerPrefix}${index}\uE001`)); nonce++) {
    markerPrefix = `\uE000teamscli-mention-${nonce}-`;
  }
  const marker = (index: number) => `${markerPrefix}${index}\uE001`;
  const missing = people.map((_, index) => index).filter((index) => !ranges.some((range) => range.person === index));
  const ordered = [...missing, ...ranges.toSorted((a, b) => a.start - b.start).map(({ person }) => person)];
  const itemids = new Map(ordered.map((person, index) => [person, String(index)]));
  let marked = body;
  for (const range of ranges.toSorted((a, b) => b.start - a.start)) {
    marked = `${marked.slice(0, range.start)}${marker(range.person)}${marked.slice(range.end)}`;
  }
  if (missing.length > 0) marked = `${missing.map(marker).join(' ')} ${marked}`;

  const mentions = ordered.map((person): MentionReference => ({
    '@type': 'http://schema.skype.com/Mention',
    itemid: itemids.get(person)!,
    mri: people[person]!.mri,
    mentionType: 'person',
    displayName: people[person]!.displayName,
  }));
  // shortcut: multiple mention itemid sequencing is not live-verified; validate when Teams is available.
  let content = toTeamsHtml(marked, options);
  for (let index = 0; index < people.length; index++) {
    const person = people[index]!;
    content = content.replaceAll(marker(index), `<readonly class="skipProofing" itemtype="http://schema.skype.com/Mention" spellcheck="false">`
      + `<span itemtype="http://schema.skype.com/Mention" itemscope itemid="${itemids.get(index)}">${escapeHtml(person.displayName)}</span></readonly>`);
  }
  return { content, mentions };
}

export interface QuotedMessage {
  id: string;
  senderId: string;
  senderName: string;
  /** Plain text preview of the quoted message. */
  text: string;
}

/** Reply-quote block (docs/authentication.md); goes right before the reply's own HTML. */
export function quoteHtml({ id, senderId, senderName, text }: QuotedMessage): string {
  return `<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="${escapeHtml(id)}">`
    + `<strong itemprop="mri" itemid="${escapeHtml(senderId)}">${escapeHtml(senderName)}</strong>`
    + `<span itemprop="time" itemid="${escapeHtml(id)}"></span>`
    + `<p itemprop="preview">${escapeHtml(text)}</p>`
    + '</blockquote>';
}

const ENTITIES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

export function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ENTITIES[char] ?? char);
}

// Code span | [text](http url) | bare http url. Matched on raw text; each piece is escaped on its own.
const INLINE = /`([^`]+)`|\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)|(https?:\/\/[^\s<>"'`]+)/gi;

function inline(line: string): string {
  let html = '';
  let last = 0;
  for (const match of line.matchAll(INLINE)) {
    html += emphasis(escapeHtml(line.slice(last, match.index)));
    const [whole, code, label, url, bare] = match;
    if (code !== undefined) html += `<code>${escapeHtml(code)}</code>`;
    else if (label !== undefined && url !== undefined) {
      html += url.includes('\uE000teamscli-mention-') ? escapeHtml(whole) : link(url, emphasis(escapeHtml(label)));
    }
    else if (bare !== undefined) {
      if (bare.includes('\uE000teamscli-mention-')) {
        html += escapeHtml(bare);
        last = match.index + whole.length;
        continue;
      }
      // Sentence punctuation after a bare URL is not part of it.
      const trimmed = bare.replace(/[.,;:!?)]+$/, '');
      html += link(trimmed, escapeHtml(trimmed)) + escapeHtml(bare.slice(trimmed.length));
    }
    last = match.index + whole.length;
  }
  return html + emphasis(escapeHtml(line.slice(last)));
}

function link(url: string, label: string): string {
  return `<a href="${escapeHtml(url)}">${label}</a>`;
}

function mentionRanges(text: { value: string; starts: number[]; ends: number[] }, name: string): Array<{ start: number; end: number }> {
  const wanted = normalizeMention(name);
  const ranges: Array<{ start: number; end: number }> = [];
  let offset = 0;
  while (wanted !== '' && offset <= text.value.length - wanted.length) {
    const index = text.value.indexOf(wanted, offset);
    if (index < 0) break;
    ranges.push({ start: text.starts[index]!, end: text.ends[index + wanted.length - 1]! });
    offset = index + 1;
  }
  return ranges;
}

function normalizeWithOffsets(value: string): { value: string; starts: number[]; ends: number[] } {
  let normalized = '';
  const starts: number[] = [];
  const ends: number[] = [];
  for (let index = 0; index < value.length;) {
    const char = String.fromCodePoint(value.codePointAt(index)!);
    const end = index + char.length;
    const folded = normalizeMention(char);
    if (folded === '' && ends.length > 0) ends[ends.length - 1] = end;
    normalized += folded;
    for (let unit = 0; unit < folded.length; unit++) {
      starts.push(index);
      ends.push(end);
    }
    index = end;
  }
  return { value: normalized, starts, ends };
}

function normalizeMention(value: string): string {
  return value.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase();
}

// shortcut: emphasis does not span code spans or links, and markers cannot nest; enough for chat text.
function emphasis(escaped: string): string {
  return escaped
    .replace(/\*\*(?!\s)([^*]+?)(?<!\s)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(?!\s)([^*]+?)(?<!\s)\*/g, '<em>$1</em>');
}
