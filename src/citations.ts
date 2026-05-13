import type { CitationAppearance, CitationIconName } from "@microsoft/teams.api";
import { config } from "./config";
import type { OrchestratorDataPoint } from "./orchestratorClient";

/**
 * Inline marker emitted by the orchestrator alongside cited passages, e.g.
 *   [ResidentialTenancyAct-2010-042.pdf][Page37][Residential Tenancies Act 2010]
 *   [ResidentialTenancyAct-2010-042.pdf][Page124]
 *   [ResidentialTenancyAct-2010-042.pdf][Page124-125]
 *
 * Captures: 1) fileName, 2) page number (or first page of a range),
 * 3) optional document title. When the title is absent we fall back to
 * the file name.
 *
 * The leading `\s*` is intentional — it absorbs any whitespace/newlines
 * the orchestrator placed BEFORE the marker so the rewritten `[N]` chip
 * lands on the same line as the preceding sentence. Teams' citation
 * renderer only converts `[N]` into an interactive chip when the marker
 * is inline (not on its own line / paragraph).
 */
const CITATION_RE =
  /\s*\[([^\]\s]+\.[A-Za-z0-9]{1,5})\]\[Page\s*(\d+)(?:\s*-\s*\d+)?\](?:\[([^\]]+)\])?/g;

/** Teams hard limit per the AI-generated-content docs. */
const MAX_CITATIONS = 20;
const NAME_LIMIT = 80;
const ABSTRACT_LIMIT = 160;

export interface ParsedCitation {
  /** 1-based citation number, used both in text (`[N]`) and `addCitation`. */
  position: number;
  fileName: string;
  page: number;
  title: string;
}

export interface RewriteResult {
  /** Answer text with `[file][PageN][title]` markers replaced by `[N]`. */
  text: string;
  citations: ParsedCitation[];
}

/**
 * Replace orchestrator citation markers in `answer` with Teams-style
 * `[N]` references and return the unique citations in order.
 */
export function rewriteAnswerWithCitations(answer: string): RewriteResult {
  const citations: ParsedCitation[] = [];
  const seen = new Map<string, number>();

  // The orchestrator sometimes emits HTML (`<strong>`, `<em>`, `<br>`,
  // `<ul>/<li>`, etc.) instead of markdown. Teams' citation chip overlay
  // only kicks in on its default text renderer — once the body looks
  // like HTML, Teams switches to an HTML path that strips the chips.
  // Convert the small subset we see in practice down to markdown so the
  // default renderer keeps the `[N]` chips interactive.
  const normalised = flattenNestedLists(htmlToMarkdown(answer));

  const text = normalised.replace(
    CITATION_RE,
    (_match, fileName: string, pageStr: string, title: string | undefined) => {
      const page = Number(pageStr);
      const key = `${fileName.toLowerCase()}|${page}`;
      let position = seen.get(key);
      if (!position) {
        if (citations.length >= MAX_CITATIONS) {
          // Drop the marker entirely once we've hit Teams' cap.
          return "";
        }
        position = citations.length + 1;
        seen.set(key, position);
        citations.push({
          position,
          fileName,
          page,
          title: title?.trim() || fileName,
        });
      }
      return ` [${position}]`;
    }
  );

  return { text: text.trim(), citations };
}

/**
 * Minimal HTML → markdown conversion for orchestrator answers that come
 * back as HTML. We intentionally do NOT pull in a full HTML parser —
 * the orchestrator only emits a tiny tag set (`<strong>`, `<em>`,
 * `<br>`, `<p>`, `<ul>`, `<ol>`, `<li>`, `<h1>..<h6>`, `<code>`).
 * Anything else is stripped. HTML entities like `&amp;`, `&lt;`,
 * `&gt;`, `&quot;`, `&#39;`, `&nbsp;` are decoded.
 *
 * The goal is to keep Teams on its default text renderer (which honours
 * markdown AND overlays `[N]` citation chips). The moment Teams sees
 * raw HTML it switches to an HTML path that drops the chips.
 */
function htmlToMarkdown(input: string): string {
  if (!input || !/<[a-z!/]/i.test(input)) {
    return input;
  }
  let s = input;

  // Block-level: paragraphs and line breaks → newlines.
  s = s.replace(/<\s*br\s*\/?\s*>/gi, "\n");
  s = s.replace(/<\s*\/\s*p\s*>/gi, "\n\n");
  s = s.replace(/<\s*p\b[^>]*>/gi, "");

  // Headings → markdown headings.
  s = s.replace(/<\s*h([1-6])\b[^>]*>/gi, (_m, n) => "\n" + "#".repeat(Number(n)) + " ");
  s = s.replace(/<\s*\/\s*h[1-6]\s*>/gi, "\n");

  // Lists. Convert <li> to "- " for both ul and ol; numbered ordering
  // is good enough for chat rendering.
  s = s.replace(/<\s*\/\s*(ul|ol)\s*>/gi, "\n");
  s = s.replace(/<\s*(ul|ol)\b[^>]*>/gi, "\n");
  s = s.replace(/<\s*li\b[^>]*>/gi, "- ");
  s = s.replace(/<\s*\/\s*li\s*>/gi, "\n");

  // Inline emphasis.
  s = s.replace(/<\s*(strong|b)\b[^>]*>/gi, "**");
  s = s.replace(/<\s*\/\s*(strong|b)\s*>/gi, "**");
  s = s.replace(/<\s*(em|i)\b[^>]*>/gi, "*");
  s = s.replace(/<\s*\/\s*(em|i)\s*>/gi, "*");
  s = s.replace(/<\s*code\b[^>]*>/gi, "`");
  s = s.replace(/<\s*\/\s*code\s*>/gi, "`");

  // Strip anything else (best-effort).
  s = s.replace(/<\/?[a-z][^>]*>/gi, "");

  // Decode the handful of entities we actually see.
  s = s
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'");

  // Collapse runs of 3+ blank lines that the conversion can produce.
  s = s.replace(/\n{3,}/g, "\n\n");

  return s;
}

/**
 * Flatten the orchestrator's `N. **Title**:\n   - Body` pattern into a
 * single top-level bullet `- **Title**: Body`.
 *
 * Why: Teams' `[N]` citation chip overlay reliably fires on top-level
 * bullets / plain prose, but does NOT fire when the marker sits inside
 * a nested (indented) list item. The orchestrator commonly emits a
 * numbered list whose body is a single indented sub-bullet, which puts
 * every citation marker into a non-chippable position. Flattening to a
 * flat bullet list preserves the bold title + colon-separated body
 * structure while keeping every marker chippable.
 */
function flattenNestedLists(input: string): string {
  // Matches: "1. **Title**:\n   - Body..." up to the next numbered item,
  // blank line, or end of string. The body capture is everything until
  // the lookahead.
  const RE = /^(\d+)\.\s+(\*\*[^\n]+?\*\*)\s*:\s*\n[ \t]+-\s+([\s\S]+?)(?=\n[ \t]*(?:\d+\.\s|-\s|\n|$))/gm;
  return input.replace(RE, (_m, _n, title: string, body: string) => {
    // Collapse internal soft-wrap newlines/indents in the body so the
    // marker stays inline with the preceding sentence.
    const flatBody = body.replace(/\s*\n[ \t]*/g, " ").trim();
    return `- ${title}: ${flatBody}`;
  });
}

/**
 * Build an absolute URL pointing back at this bot's `/api/documents`
 * endpoint, which streams the blob via the bot's managed identity.
 * Returns `undefined` if `STORAGE_ACCOUNT` or the bot's public host
 * cannot be determined.
 */
export function buildDocumentUrl(
  fileName: string,
  page: number
): string | undefined {
  if (!config.storageAccount || !config.botBaseUrl) {
    return undefined;
  }
  const base = config.botBaseUrl.replace(/\/$/, "");
  // The orchestrator sometimes returns file names that are already
  // percent-encoded (e.g. "Residential%20Tenancies..."). Decode first
  // so we don't double-encode (`%20` → `%2520`). Falls back to the raw
  // value if decoding throws on a malformed sequence.
  let decoded = fileName;
  try {
    decoded = decodeURIComponent(fileName);
  } catch {
    /* leave as-is */
  }
  const name = encodeURIComponent(decoded);
  return `${base}/api/documents?name=${name}#page=${page}`;
}

/** Best-effort mapping of file extension to a Teams citation icon. */
export function iconForFile(fileName: string): CitationIconName {
  const ext = fileName.split(".").pop()?.toLowerCase() ?? "";
  switch (ext) {
    case "pdf":
      return "PDF";
    case "doc":
    case "docx":
      return "Microsoft Word";
    case "xls":
    case "xlsx":
    case "csv":
      return "Microsoft Excel";
    case "ppt":
    case "pptx":
      return "Microsoft PowerPoint";
    case "one":
      return "Microsoft OneNote";
    case "png":
    case "jpg":
    case "jpeg":
    case "bmp":
    case "tif":
    case "tiff":
      return "Image";
    case "gif":
      return "GIF";
    case "mp4":
    case "mov":
    case "avi":
    case "webm":
      return "Video";
    case "mp3":
    case "wav":
    case "m4a":
      return "Sound";
    case "zip":
    case "7z":
    case "rar":
      return "ZIP";
    case "json":
    case "ts":
    case "js":
    case "py":
    case "cs":
    case "java":
      return "Source Code";
    default:
      return "Text";
  }
}

function truncate(s: string, max: number): string {
  if (!s) return s;
  return s.length <= max ? s : s.slice(0, max - 1).trimEnd() + "\u2026";
}

/** Try to find the orchestrator data_points entry that matches a citation. */
function findDataPoint(
  citation: ParsedCitation,
  dataPoints: OrchestratorDataPoint[]
): OrchestratorDataPoint | undefined {
  const fnLower = citation.fileName.toLowerCase();
  return dataPoints.find((dp) => {
    const dpFile = (dp.fileName ?? "").toLowerCase();
    if (!dpFile) return false;
    if (dpFile !== fnLower) return false;
    if (dp.page != null && dp.page !== citation.page) return false;
    return true;
  });
}

/**
 * Build a stringified Adaptive Card to render inside the citation modal.
 * NOTE: Adaptive Cards in citation modals are currently in Teams public
 * developer preview. Clients without the preview will fall back to the
 * pop-up (without the "Open" button).
 */
function buildModalCard(
  citation: ParsedCitation,
  excerpt: string | undefined,
  url: string | undefined
): string {
  const card: Record<string, unknown> = {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: citation.title,
        weight: "Bolder",
        wrap: true,
        size: "Medium",
      },
      {
        type: "TextBlock",
        text: `${citation.fileName} \u2014 page ${citation.page}`,
        isSubtle: true,
        spacing: "None",
        wrap: true,
      },
      ...(excerpt
        ? [{ type: "TextBlock", text: excerpt, wrap: true, spacing: "Medium" }]
        : []),
    ],
    actions: url
      ? [
          {
            type: "Action.OpenUrl",
            title: "Open document",
            url,
          },
        ]
      : [],
  };
  return JSON.stringify(card);
}

/**
 * Build the `CitationAppearance` for a single parsed citation, including:
 * - GET URL via `DOCUMENT_URL_TEMPLATE` (Option A).
 * - Modal Adaptive Card with the source excerpt (Option B, dev preview).
 */
export function buildCitationAppearance(
  citation: ParsedCitation,
  dataPoints: OrchestratorDataPoint[]
): CitationAppearance {
  const url = buildDocumentUrl(citation.fileName, citation.page);
  const dp = findDataPoint(citation, dataPoints);
  const excerpt = typeof dp?.content === "string" ? dp.content : undefined;

  const appearance: CitationAppearance = {
    name: truncate(`${citation.title} (p.${citation.page})`, NAME_LIMIT),
    abstract: truncate(
      excerpt ?? `Page ${citation.page} of ${citation.fileName}`,
      ABSTRACT_LIMIT
    ),
    icon: iconForFile(citation.fileName),
    // Stringified Adaptive Card rendered in the citation modal.
    // Adaptive Cards in citation modals are in Teams public developer
    // preview; clients without preview will simply not show the modal.
    text: buildModalCard(citation, excerpt, url),
  };

  if (url) {
    appearance.url = url;
  }

  return appearance;
}

/** Coerce the orchestrator's loosely-typed `data_points` to an array. */
export function coerceDataPoints(raw: unknown): OrchestratorDataPoint[] {
  return Array.isArray(raw) ? (raw as OrchestratorDataPoint[]) : [];
}
