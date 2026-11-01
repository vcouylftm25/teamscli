const namedEntities: Record<string, string> = {
  amp: '&',
  apos: "'",
  bull: '•',
  copy: '©',
  euro: '€',
  gt: '>',
  hellip: '…',
  ldquo: '“',
  lsquo: '‘',
  lt: '<',
  mdash: '—',
  nbsp: ' ',
  ndash: '–',
  quot: '"',
  raquo: '»',
  reg: '®',
  rdquo: '”',
  rsquo: '’',
  trade: '™',
};

/** Convert Teams HTML message content to the plain text used by search. */
export function htmlToText(html: string): string {
  return decodeEntities(html.replace(/<\s*\/?\s*(?:br|hr|p|div|li|h[1-6]|blockquote|pre|t[dhr])\b[^>]*>/gi, ' ').replace(/<[^>]*>/g, ''))
    .replace(/\s+/g, ' ')
    .trim();
}

function decodeEntities(value: string): string {
  return value.replace(/&(#x[\da-f]+|#\d+|[a-z][a-z\d]+);/gi, (entity, name: string) => {
    if (name[0] !== '#') {
      const decoded = namedEntities[name.toLowerCase()];
      return typeof decoded === 'string' ? decoded : entity;
    }

    const codePoint = name[1]?.toLowerCase() === 'x' ? Number.parseInt(name.slice(2), 16) : Number.parseInt(name.slice(1), 10);
    if (!Number.isInteger(codePoint) || codePoint <= 0 || codePoint > 0x10ffff || (codePoint >= 0xd800 && codePoint <= 0xdfff)) return '\uFFFD';
    return String.fromCodePoint(codePoint);
  });
}
