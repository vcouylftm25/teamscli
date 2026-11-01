import { escapeHtml } from './compose.ts';

/** Shown before a link when no rule below matches. */
export const DEFAULT_LINK_EMOJI = '🔗';

const SECURITY_HOSTS = ['cisa.gov', 'nvd.nist.gov', 'cve.org', 'cve.mitre.org', 'msrc.microsoft.com', 'cert.br', 'kb.cert.org', 'exploit-db.com', 'osv.dev', 'snyk.io', 'security.snyk.io'];
const CODE_HOSTS = ['github.com', 'gitlab.com', 'bitbucket.org', 'huggingface.co', 'npmjs.com', 'pypi.org'];
const PAPER_HOSTS = ['arxiv.org', 'openreview.net', 'paperswithcode.com'];
const NEWS_HOSTS = [
  'techcrunch.com', 'theverge.com', 'reuters.com', 'bloomberg.com', 'arstechnica.com', 'wired.com', 'theregister.com',
  'venturebeat.com', 'zdnet.com', 'cnbc.com', 'bbc.com', 'bbc.co.uk', 'nytimes.com', 'ft.com', 'wsj.com', 'theinformation.com',
  'bleepingcomputer.com', 'thehackernews.com', 'securityweek.com', 'darkreading.com', 'news.ycombinator.com', 'medium.com',
  'substack.com', 'g1.globo.com', 'globo.com', 'folha.uol.com.br', 'uol.com.br', 'valor.globo.com', 'estadao.com.br',
  'exame.com', 'infomoney.com.br', 'tecmundo.com.br', 'olhardigital.com.br', 'canaltech.com.br', 'convergenciadigital.com.br',
];

function onHost(host: string, list: string[]): boolean {
  return list.some((entry) => host === entry || host.endsWith(`.${entry}`));
}

/**
 * Picks the small emoji shown before a link: 🛡️ security advisories, 💻 code hosts,
 * 📄 papers, 📰 news and blogs, 🔗 anything else. Teams draws inline images inside a
 * large placeholder box, so a text emoji is the only icon that stays small.
 */
export function linkEmoji(url: URL): string {
  const host = url.hostname.toLowerCase().replace(/\.$/, '').replace(/^www\./, '');
  const path = url.pathname.toLowerCase();
  const security = onHost(host, SECURITY_HOSTS)
    || /^(security|advisories|psirt|cert)\./.test(host)
    || /\/(security|advisories|advisory|security-advisories|cve)(\/|$)|cve-\d{4}-\d+/.test(path);
  if (security) return '🛡️';
  if (onHost(host, CODE_HOSTS)) return '💻';
  if (onHost(host, PAPER_HOSTS)) return '📄';
  if (onHost(host, NEWS_HOSTS) || /^(blog|news|noticias)\./.test(host) || /\/(blog|blogs|news|noticias|press|newsroom)(\/|$)/.test(path)) return '📰';
  return DEFAULT_LINK_EMOJI;
}

/** Puts `linkEmoji(url)` and a non-breaking space before every HTTP(S) link in Teams HTML. */
export function withLinkEmojis(content: string, decodeHref: (value: string) => string): string {
  return content.replace(/<a href="([^"]+)">/g, (whole, encoded: string) => {
    let url: URL;
    try { url = new URL(decodeHref(encoded)); } catch { return whole; }
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return whole;
    return `${escapeHtml(linkEmoji(url))}&nbsp;${whole}`;
  });
}
