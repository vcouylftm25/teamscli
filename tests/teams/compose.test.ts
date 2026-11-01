import assert from 'node:assert/strict';
import test from 'node:test';
import { quoteHtml, toTeamsHtml } from '../../src/teams/compose.ts';
import { htmlToText } from '../../src/teams/html.ts';

test('paragraphs and line breaks', () => {
  assert.equal(toTeamsHtml('first\nline two\n\n\nsecond'), '<p>first<br>line two</p><p>&nbsp;</p><p>second</p>');
  // No spacer next to lists, code blocks or below headings: they carry their own margin in Teams.
  assert.equal(toTeamsHtml('a\n\n- x\n\nb'), '<p>a</p><ul><li>x</li></ul><p>b</p>');
  assert.equal(toTeamsHtml('a\n\n# T\n\nb'), '<p>a</p><p>&nbsp;</p><h2>T</h2><p>b</p>');
  assert.equal(toTeamsHtml('> q\n\n```\nc\n```').startsWith('<blockquote><p>q</p></blockquote><p itemtype'), true);
  assert.equal(toTeamsHtml('a\r\nb'), '<p>a<br>b</p>');
  assert.equal(toTeamsHtml('  \n\n '), '');
});

test('simple markdown', () => {
  assert.equal(toTeamsHtml('**bold** and *it* and `x < y`'), '<p><strong>bold</strong> and <em>it</em> and <code>x &lt; y</code></p>');
  assert.equal(toTeamsHtml('2 * 3 * 4'), '<p>2 * 3 * 4</p>');
  assert.equal(toTeamsHtml('`**not bold**`'), '<p><code>**not bold**</code></p>');
});

test('user HTML is escaped, never passed through', () => {
  assert.equal(toTeamsHtml('<script>alert("x")</script> & \'q\''),
    '<p>&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt; &amp; &#39;q&#39;</p>');
  assert.equal(toTeamsHtml('<b onclick="x">hi</b>'), '<p>&lt;b onclick=&quot;x&quot;&gt;hi&lt;/b&gt;</p>');
});

test('http and https links; anything else stays text', () => {
  assert.equal(toTeamsHtml('[docs](https://example.com/a?b=1&c="2")'), '<p><a href="https://example.com/a?b=1&amp;c=&quot;2&quot;">docs</a></p>');
  assert.equal(toTeamsHtml('see http://example.com/x.'), '<p>see <a href="http://example.com/x">http://example.com/x</a>.</p>');
  assert.equal(toTeamsHtml('[x](javascript:alert(1))'), '<p>[x](javascript:alert(1))</p>');
  assert.equal(toTeamsHtml('[x](data:text/html,<b>)'), '<p>[x](data:text/html,&lt;b&gt;)</p>');
  assert.equal(toTeamsHtml('https://example.com/"><script>'), '<p><a href="https://example.com/">https://example.com/</a>&quot;&gt;&lt;script&gt;</p>');
});

test('quote block escapes name and preview and round-trips to text', () => {
  const html = quoteHtml({ id: '1767261600000', senderId: '8:orgid:00000000-0000-0000-0000-000000000001', senderName: 'Alice <Ex> & "Co"', text: '<script>x</script>' });
  assert.equal(html,
    '<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="1767261600000">'
    + '<strong itemprop="mri" itemid="8:orgid:00000000-0000-0000-0000-000000000001">Alice &lt;Ex&gt; &amp; &quot;Co&quot;</strong>'
    + '<span itemprop="time" itemid="1767261600000"></span>'
    + '<p itemprop="preview">&lt;script&gt;x&lt;/script&gt;</p></blockquote>');
  assert.equal(htmlToText(html + toTeamsHtml('ok')), 'Alice <Ex> & "Co" <script>x</script> ok');
});

test('markdown blocks become the HTML Teams itself produces', () => {
  assert.equal(toTeamsHtml('# Title\n### Small *one*'), '<h2>Title</h2><h3>Small <em>one</em></h3>');
  assert.equal(toTeamsHtml('Steps:\n- one\n- **two**\n  - nested\n1. first'),
    '<p>Steps:</p><ul><li>one</li><li><strong>two</strong><ul><li>nested</li></ul></li><li>first</li></ul>');
  assert.equal(toTeamsHtml('1. a\n2. b'), '<ol><li>a</li><li>b</li></ol>');
  assert.equal(toTeamsHtml('> quoted\n> - item'), '<blockquote><p>quoted</p><ul><li>item</li></ul></blockquote>');
  assert.equal(toTeamsHtml('| A | B |\n|---|:-:|\n| `x` | <y> |'), '<table><tbody><tr><th>A</th><th>B</th></tr><tr><td><code>x</code></td><td>&lt;y&gt;</td></tr></tbody></table>');
  assert.equal(toTeamsHtml('| not | a table |'), '<p>| not | a table |</p>');
});

test('horizontal rules render as hr blocks with paragraph spacing around surrounding text', () => {
  assert.equal(toTeamsHtml('Before\n\n---\n\nAfter'), '<p>Before</p><hr><p>After</p>');
  assert.equal(toTeamsHtml('---'), '<hr>');
});

test('fenced code keeps its text escaped, with a language and an id per block', () => {
  assert.equal(toTeamsHtml('```TypeScript\nif (a < b) {\n  **x**;\n}\n```\ntext\n```\nplain'),
    '<p itemtype="http://schema.skype.com/CodeBlockEditor" id="1">&nbsp;</p><pre class="language-typescript" itemid="1"><code>if (a &lt; b) {<br>  **x**;<br>}</code></pre>'
    + '<p>text</p>'
    + '<p itemtype="http://schema.skype.com/CodeBlockEditor" id="2">&nbsp;</p><pre class="language-plaintext" itemid="2"><code>plain</code></pre>');
  assert.equal(toTeamsHtml('```"><script>\nx\n```'), '<p>```&quot;&gt;&lt;script&gt;<br>x</p><p itemtype="http://schema.skype.com/CodeBlockEditor" id="1">&nbsp;</p><pre class="language-plaintext" itemid="1"><code></code></pre>');
});

test('text derived from tables and headings keeps words apart', () => {
  assert.equal(htmlToText(toTeamsHtml('# T\n| A | B |\n|---|---|\n| c | d |')), 'T A B c d');
});
