import { marked } from "marked";
import DOMPurify from "dompurify";

marked.setOptions({ gfm: true, breaks: true });

/**
 * Turn a model answer into sanitized HTML.
 *
 * The answers come back as Markdown (`**结论**`, `- item`, `1.`, tables), and
 * rendering them as plain text shows the raw `**` asterisks — which is the main
 * thing that made the output look unfinished.
 *
 * Two safety notes:
 * - DOMPurify runs *last*, so anything the model emitted (or echoed from a
 *   document — the docx content is user-supplied and reaches the prompt) is
 *   stripped before it hits `v-html`.
 * - Citations like `[1]` are turned into badges after Markdown parsing, since
 *   `[1]` survives as literal text unless a link reference exists.
 */
export function renderMarkdown(source: string): string {
  const html = marked.parse(source, { async: false }) as string;

  const withCitations = html.replace(
    /\[(\d{1,2})\]/g,
    '<sup class="cite" data-cite="$1" title="跳到来源 $1">$1</sup>',
  );

  return DOMPurify.sanitize(withCitations);
}
