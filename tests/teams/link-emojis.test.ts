import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_LINK_EMOJI, linkEmoji, withLinkEmojis } from '../../src/teams/link-emojis.ts';

test('picks an emoji per kind of site and falls back to 🔗', () => {
  const cases: Array<[string, string]> = [
    ['https://www.cisa.gov/known-exploited-vulnerabilities-catalog', '🛡️'],
    ['https://nvd.nist.gov/vuln/detail/CVE-2026-1234', '🛡️'],
    ['https://github.com/acme/tool/security/advisories/GHSA-xxxx', '🛡️'],
    ['https://access.redhat.com/security/cve/cve-2026-1234', '🛡️'],
    ['https://msrc.microsoft.com/update-guide', '🛡️'],
    ['https://github.com/acme/tool', '💻'],
    ['https://huggingface.co/org/model', '💻'],
    ['https://arxiv.org/abs/2601.00001', '📄'],
    ['https://techcrunch.com/2026/11/01/x', '📰'],
    ['https://openai.com/news/item', '📰'],
    ['https://blog.example.com/post', '📰'],
    ['https://cursor.com/blog/a', '📰'],
    ['https://example.com/page', DEFAULT_LINK_EMOJI],
  ];
  for (const [url, emoji] of cases) assert.equal(linkEmoji(new URL(url)), emoji, url);
});

test('withLinkEmojis only touches http(s) links and leaves other markup alone', () => {
  const decode = (value: string) => value.replace(/&amp;/g, '&');
  const html = '<p><a href="https://example.com/?a=1&amp;b=2">x</a> <a href="mailto:a@example.com">m</a></p>';
  assert.equal(withLinkEmojis(html, decode), '<p>🔗&nbsp;<a href="https://example.com/?a=1&amp;b=2">x</a> <a href="mailto:a@example.com">m</a></p>');
});
