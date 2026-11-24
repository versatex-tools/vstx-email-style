var VSTX_SIGNATURE_ID = "vstx-signature";
var VSTX_BAR_ID = "vstx-bar";
var VSTX_LIST_CLASS = "vstx-list";
var VSTX_TABLE_CLASS = "vstx-table";
var SIGNATURE_SETTINGS_KEY = "vstxSignature";
var FEATURES_SETTINGS_KEY = "vstxFeatures";
var FEATURE_DEFAULTS = {
  headings: true,
  // Off by default: short plain lines between other text become headings too, for titles pasted without bold.
  plainTitles: false,
  callouts: true,
  figures: true,
  buttons: true,
  statuses: true,
  kpis: true,
  // Off by default: Outlook on the web draws its own mention chip over the tag when the email is read.
  mentions: false,
  agenda: true,
  shortReplies: true,
  attachmentCheck: true,
};
var OWN_MARKUP = /\bvstx-/;
var ASSET_BASE = "https://versatex-tools.github.io/vstx-email-style/assets/";
var LOGO_URL = ASSET_BASE + "vstx-logo-v2.png";
var LOGO_WIDTH = 42;
var LOGO_HEIGHT = 54;
var HEXAGONS_URL = ASSET_BASE + "vstx-hexagons.png";
var HEXAGONS_WIDTH = 52;
var HEXAGONS_HEIGHT = 49;
var FONT_STACK = "Roboto,Arial,sans-serif";
var INK = "#231f20";
var GREY = "#58595b";
var LIGHT_GREY = "#a7a9ac";
var YELLOW = "#fdc00f";
var PANEL = "#f4f4f4";
var HAIRLINE = "#e6e7e8";
var BODY_SIZE = "15px";
var BODY_LINE_HEIGHT = "1.6";
// Weight 700, never 900: mail clients don't load Roboto, and on Windows Arial at 900 turns into Arial Black.
var STRONG_WEIGHT = "700";
// Rounded corners show in Outlook on the web, on phones and in Gmail; classic Outlook draws them square.
var RADIUS = "6px";
var SMALL_RADIUS = "4px";
var PILL_RADIUS = "10px";
// Spacing steps of 8px. Components keep 8px around them; the blank lines people type (24px) add the rest.
var COMPONENT_MARGIN = "8px 0";
var BODY_TEXT = "font-family:" + FONT_STACK + ";color:" + INK + ";font-size:" + BODY_SIZE + ";line-height:" + BODY_LINE_HEIGHT;
var BULLET_SIZE = 18;
var BULLET_COLUMN = 22;
var BADGE_SIZE = 22;
var BADGE_CSS =
  "width:" + BADGE_SIZE + "px;height:" + BADGE_SIZE + "px;background:" + YELLOW + ";font-family:" + FONT_STACK +
  ";font-size:12px;font-weight:" + STRONG_WEIGHT + ";color:" + INK + ";line-height:" + BADGE_SIZE + "px;text-align:center;border-radius:" + SMALL_RADIUS;
var NUMBERED_TEXT_CSS = BODY_TEXT + ";vertical-align:top;padding:0 0 8px";
var REVIEW_TAG = /\[review\]/i;
var PLAIN_TAG = /\[plain\]/i;
var CONTROL_TAGS = /\[(?:review|plain)\]/gi;
var TEXT_TAGS = /<(\/?)(table|p|li|div|span|font|td|th|h[1-6])\b([^>]*)>/gi;
var BLOCK_TAGS = /^(p|li|div)$/;
var HEADING_TAGS = /^h[1-6]$/;
var MONOSPACE = /consolas|courier|monaco|menlo|cascadia|lucida console|monospace/i;
var LIST_OPENER = /<(ul|ol)\b([^>]*)>/gi;
var LEADING_HEXAGONS = /^((?:\s|<(?:div|p|span)\b[^>]*>)*)(?:<span\b[^>]*>&#11042;<\/span>(?:&nbsp;|\s)*)+/i;
var STYLE_ATTR = /\bstyle\s*=\s*("([^"]*)"|'([^']*)')/i;
var TABLE_CSS = "border-collapse:collapse;margin:8px 0 16px";
var HEADER_CELL_CSS =
  "background:" + INK + ";color:#ffffff;font-weight:700;padding:8px 12px;border:none;border-bottom:3px solid " + YELLOW + ";text-align:left;vertical-align:bottom";
var BODY_CELL_CSS = "padding:7px 12px;border:none;border-bottom:1px solid " + HAIRLINE + ";vertical-align:top";
// Everything from the first match onwards is left untouched: quoted replies, forwards and a Teams joining block.
var QUOTE_MARKERS = [
  /<div[^>]*\bid=["']?appendonsend\b/i,
  /<div[^>]*\bid=["']?divRplyFwdMsg\b/i,
  /<div[^>]*border-top:\s*solid\s+#E1E1E1/i,
  /<hr\b[^>]*>/i,
  /<div[^>]*>\s*<span[^>]*>\s*_{20,}/i,
  /<div[^>]*\bclass=["']?[^"'>]*\bme-email-/i,
];

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function findQuoteStart(html) {
  var start = -1;
  QUOTE_MARKERS.forEach(function (marker) {
    var match = marker.exec(html);
    if (match && (start === -1 || match.index < start)) start = match.index;
  });
  if (start !== -1) return start;
  var bodyEnd = html.search(/<\/body>/i);
  return bodyEnd === -1 ? html.length : bodyEnd;
}

// ---- Inline style helpers ----

function readStyle(attrs) {
  var match = STYLE_ATTR.exec(attrs || "");
  if (!match) return "";
  return (match[2] !== undefined ? match[2] : match[3]).replace(/&quot;/g, "'");
}

function parseDeclarations(style) {
  return style
    .split(";")
    .map(function (part) {
      var colon = part.indexOf(":");
      return colon === -1 ? null : [part.slice(0, colon).trim().toLowerCase(), part.slice(colon + 1).trim()];
    })
    .filter(function (declaration) {
      return declaration && declaration[0];
    });
}

function valueOf(declarations, property) {
  for (var i = declarations.length - 1; i >= 0; i--) {
    if (declarations[i][0] === property) return declarations[i][1];
  }
  return "";
}

function comparable(value) {
  return value.replace(/\s+/g, "").toLowerCase();
}

function withStyle(name, attrs, declarations, overrides) {
  var replaced = {};
  overrides.forEach(function (override) {
    replaced[override[0]] = true;
  });
  var style = declarations
    .filter(function (declaration) {
      return !replaced[declaration[0]];
    })
    .concat(overrides)
    .map(function (declaration) {
      return declaration[0] + ":" + declaration[1];
    })
    .join(";");
  var otherAttrs = (attrs || "").replace(STYLE_ATTR, "").replace(/\s{2,}/g, " ").replace(/\s+$/, "");
  // Word quotes font names with " inside single-quoted styles; inside style="…" they must become '.
  return "<" + name + otherAttrs + ' style="' + style.replace(/"/g, "'") + '">';
}

function mergeStyle(tag, name, attrs, css) {
  if (/WordSection/i.test(attrs || "")) return tag;
  return withStyle(name, attrs, parseDeclarations(readStyle(attrs)), parseDeclarations(css));
}

// ---- Text: restyle only the default look, keep deliberate formatting ----

function scanTextTags(html, visit) {
  var tables = [];
  var foreignTables = 0;
  var codeContext = [];
  return html.replace(TEXT_TAGS, function (tag, closing, name, attrs) {
    name = name.toLowerCase();
    if (name === "table") {
      if (closing) {
        var table = tables.pop();
        if (table && table.foreign) foreignTables--;
      } else {
        var foreign = attrs.indexOf(VSTX_LIST_CLASS) === -1 && attrs.indexOf(VSTX_TABLE_CLASS) === -1;
        // Lists, buttons and other tables the add-in builds (a branded data table is still the sender's).
        var built = OWN_MARKUP.test(attrs) && attrs.indexOf(VSTX_TABLE_CLASS) === -1;
        tables.push({ foreign: foreign, built: built });
        if (foreign) foreignTables++;
      }
      return tag;
    }
    if (closing) {
      codeContext.pop();
      return tag;
    }
    var inCode = MONOSPACE.test(attrs) || codeContext[codeContext.length - 1] === true;
    codeContext.push(inCode);
    if (/WordSection/i.test(attrs) || OWN_MARKUP.test(attrs)) return tag;
    var cellOfBuiltTable = /^t[dh]$/.test(name) && tables.length > 0 && tables[tables.length - 1].built;
    return visit(tag, name, attrs, foreignTables > 0, inCode, cellOfBuiltTable);
  });
}

// The look Outlook applies to plain typed text lives on block elements; anything else on a span is deliberate.
function defaultLook(html) {
  var look = { color: {}, size: {} };
  look.color[comparable(INK)] = true;
  look.size[BODY_SIZE] = true;
  scanTextTags(html, function (tag, name, attrs, inForeignTable, inCode) {
    if (inForeignTable || !BLOCK_TAGS.test(name)) return tag;
    var declarations = parseDeclarations(readStyle(attrs));
    var color = valueOf(declarations, "color");
    var size = valueOf(declarations, "font-size");
    if (color) look.color[comparable(color)] = true;
    if (size && !inCode) look.size[comparable(size)] = true;
    return tag;
  });
  return look;
}

// The add-in draws its hexagons and arrows in this font; restyling an email again must leave them in it.
var SYMBOL_FONT = /segoe ui symbol/i;

// Code keeps its font and size, but Outlook's default colour still becomes brand ink.
function brandText(html) {
  var look = defaultLook(html);
  return scanTextTags(html, function (tag, name, attrs, inForeignTable, inCode, cellOfBuiltTable) {
    var declarations = parseDeclarations(readStyle(attrs));
    var keepsFont = inCode || cellOfBuiltTable || SYMBOL_FONT.test(valueOf(declarations, "font-family") || "");
    var overrides = keepsFont ? [] : [["font-family", FONT_STACK]];
    if (!inForeignTable) {
      var isBlock = BLOCK_TAGS.test(name);
      var color = comparable(valueOf(declarations, "color"));
      var size = comparable(valueOf(declarations, "font-size"));
      if (color ? look.color[color] : isBlock) overrides.push(["color", INK]);
      if (!inCode && !HEADING_TAGS.test(name)) {
        if (size ? look.size[size] : isBlock) overrides.push(["font-size", BODY_SIZE]);
        if (isBlock) overrides.push(["line-height", BODY_LINE_HEIGHT]);
      }
    }
    // A tag that already has the look stays exactly as it is, so styling an email twice changes nothing.
    var changes = overrides.filter(function (override) {
      return comparable(valueOf(declarations, override[0])) !== comparable(override[1]);
    });
    return changes.length ? withStyle(name, attrs, declarations, overrides) : tag;
  });
}

// ---- Smart formatting: simple typing patterns become designed elements ----

// strong: ink label and a yellow edge; soft: grey label and edge. chip: the label sits on a small ink tag.
var CALLOUTS = {
  action: { label: "Action", strong: true, chip: true },
  actions: { label: "Actions", strong: true, chip: true },
  ask: { label: "Ask", strong: true },
  decision: { label: "Decision", strong: true, chip: true },
  deadline: { label: "Deadline", strong: true },
  important: { label: "Important", strong: true },
  risk: { label: "Risk", strong: true },
  note: { label: "Note", strong: false },
  fyi: { label: "FYI", strong: false },
  reminder: { label: "Reminder", strong: false },
  "next step": { label: "Next step", strong: false },
  "next steps": { label: "Next steps", strong: false },
};
// Fits the longest label ("Next steps") so every callout's text starts at the same edge.
var CALLOUT_LABEL_WIDTH = 92;
var CALLOUT_LINE =/^(actions?|ask|decision|deadline|important|risk|note|fyi|reminder|next steps?)\s*:\s*\S/i;
var STATUS_COLORS = {
  "on track": "#2e7d32",
  done: "#2e7d32",
  complete: "#2e7d32",
  completed: "#2e7d32",
  "at risk": "#e07b00",
  "in progress": "#e07b00",
  "off track": "#c62828",
  blocked: "#c62828",
  "not started": LIGHT_GREY,
};
var STATUS_TAG = /\[(on track|done|completed?|at risk|in progress|off track|blocked|not started)\]/gi;
// Money, percentages and big numbers; a range ("$14–16", "$14–$16", "10–20%") counts as one figure.
var FIGURE_NUMBER = "\\d+(?:,\\d{3})*(?:\\.\\d+)?";
var FIGURE_AMOUNT = "[$€£]\\s?" + FIGURE_NUMBER + "(?:\\s?(?:K|M|B|bn|million|billion)\\b)?";
var FIGURE_RANGE = "\\s?[–-]\\s?";
var KEY_FIGURE = new RegExp(
  "(?<![\\w$€£.,])(?:[+\\-−]\\s?)?(?:" +
    FIGURE_AMOUNT + "(?:" + FIGURE_RANGE + "(?:" + FIGURE_AMOUNT + "|" + FIGURE_NUMBER + "(?:\\s?(?:K|M|B|bn)\\b)?))?" +
    "|" + FIGURE_NUMBER + "(?:" + FIGURE_RANGE + FIGURE_NUMBER + ")?\\s?%" +
    "|" + FIGURE_NUMBER + "(?:(?:K|M|B|bn)\\b|\\s(?:million|billion)\\b))",
  "g"
);
var FIGURE_LIMIT = 8;
// Brand yellow #fdc00f at about a quarter strength, so ink text stays easy to read on it.
var FIGURE_HIGHLIGHT = "#fff0c2";
var INNERMOST_BLOCK = /<(div|p)\b([^>]*)>((?:(?!<\/?(?:div|p)\b)[\s\S])*?)<\/\1>/gi;
var BOLD_RUN = /<(b|strong)\b[^>]*>[\s\S]*?<\/\1>|<span\b[^>]*font-weight\s*:\s*(?:bold|[6-9]00)[^>]*>[\s\S]*?<\/span>/gi;
var BOLD_STYLE = /font-weight\s*:\s*(?:bold|[6-9]00)/i;
var NOT_HEADINGS = /^(hi|hello|hey|dear|thanks|thank you|regards|best|best regards|kind regards|cheers|sincerely)\b/i;
var LONE_LINK = /^((?:\s|&nbsp;|<(?!a\b)[^>]+>)*)<a\b([^>]*)>([\s\S]*?)<\/a>((?:\s|&nbsp;|<(?!a\b)[^>]+>)*)$/i;
var COMPONENT_MARKUP = /\bvstx-(?:figure|status|heading|callout-label|button|bar|signature|kpi)/;
var KPI_LINE = /^kpis?\s*:\s*\S/i;
var KPI_LABEL = /^kpis?\s*:\s*/i;
// Three tiles is the most a phone screen fits without scrolling sideways.
var KPI_PER_ROW = 3;
var AGENDA_TIME_WIDTH = 110;
var KPI_GAP = 2;
var NUMBER_BODY =
  "(?:[$€£]\\s?)?\\d[\\d,]*(?:\\.\\d+)?" +
  "(?:%|\\s?(?:pts|pp|bps|bn|[kmbx])\\b|\\s(?:million|billion|days?|hrs?|hours?|wks?|weeks?|months?|mins?|minutes?|yrs?|years?)\\b)?";
var KPI_VALUE = new RegExp("(?<![\\w.$€£])(?:[+\\-−]\\s?)?" + NUMBER_BODY, "i");
var KPI_CHANGE = new RegExp("(?<![\\w.$€£])[+\\-−]\\s?" + NUMBER_BODY, "i");
var VOID_TAGS = /^(br|img|hr|meta|link|input|col|wbr|area|base|source)$/i;

function roamingValue(key) {
  var roaming = typeof Office !== "undefined" && Office.context && Office.context.roamingSettings;
  var value = roaming ? roaming.get(key) : null;
  return value && typeof value === "object" ? value : {};
}

function featureSettings() {
  var saved = roamingValue(FEATURES_SETTINGS_KEY);
  var features = {};
  Object.keys(FEATURE_DEFAULTS).forEach(function (key) {
    features[key] = typeof saved[key] === "boolean" ? saved[key] : FEATURE_DEFAULTS[key];
  });
  return features;
}

// Character references that pasted or Word-made text can carry. Accented letters (&eacute; and the like) are built
// from a letter and its mark, so headings, tiles and buttons rebuilt from text never show a reference literally.
var NAMED_CHARACTERS = {
  nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'",
  lsquo: "‘", rsquo: "’", sbquo: "‚", ldquo: "“", rdquo: "”", bdquo: "„", ndash: "–", mdash: "—", hellip: "…",
  bull: "•", middot: "·", laquo: "«", raquo: "»", copy: "©", reg: "®", trade: "™", euro: "€", pound: "£", yen: "¥",
  cent: "¢", deg: "°", plusmn: "±", times: "×", divide: "÷", frac12: "½", frac14: "¼", frac34: "¾", sect: "§", para: "¶",
  szlig: "ß", aelig: "æ", AElig: "Æ", oslash: "ø", Oslash: "Ø", eth: "ð", ETH: "Ð", thorn: "þ", THORN: "Þ", iexcl: "¡", iquest: "¿",
};
var ACCENT_MARKS = { acute: "́", grave: "̀", circ: "̂", uml: "̈", tilde: "̃", ring: "̊", cedil: "̧" };

function decodeCharacter(reference, name) {
  var number = /^#(x?)([0-9a-f]+)$/i.exec(name);
  if (number) {
    var point = parseInt(number[2], number[1] ? 16 : 10);
    if (point === 160) return " ";
    return point > 0 && point <= 0x10ffff ? String.fromCodePoint(point) : reference;
  }
  if (Object.prototype.hasOwnProperty.call(NAMED_CHARACTERS, name)) return NAMED_CHARACTERS[name];
  var accented = /^([a-z])(acute|grave|circ|uml|tilde|ring|cedil)$/i.exec(name);
  return accented ? (accented[1] + ACCENT_MARKS[accented[2]]).normalize("NFC") : reference;
}

function plainText(html) {
  return html
    .replace(/<[^>]+>/g, "")
    .replace(/&(#x?[0-9a-f]+|[a-z][a-z0-9]*);/gi, decodeCharacter)
    .replace(/\s+/g, " ")
    .trim();
}

// Tag positions are worked out once per email rather than once per block: rescanning from the top for every block
// made Send slow down with the square of the email's size.
var structureIndex = { html: null, offsets: [], depths: [] };

function insideStructure(html, offset) {
  if (structureIndex.html !== html) {
    var tags = /<(\/?)(ul|ol|table)\b/gi;
    var offsets = [];
    var depths = [];
    var depth = 0;
    var match;
    while ((match = tags.exec(html))) {
      depth += match[1] ? -1 : 1;
      offsets.push(match.index);
      depths.push(depth);
    }
    structureIndex = { html: html, offsets: offsets, depths: depths };
  }
  var low = 0;
  var high = structureIndex.offsets.length;
  while (low < high) {
    var middle = (low + high) >> 1;
    if (structureIndex.offsets[middle] < offset) low = middle + 1;
    else high = middle;
  }
  return low > 0 && structureIndex.depths[low - 1] > 0;
}

// Classic Outlook (Word) gives an empty cell a full line of height unless its font size and line height match the cell.
function ruleCell(height, color, attrs) {
  var fill = color ? ' bgcolor="' + color + '"' : "";
  return (
    "<td" + attrs + ' height="' + height + '"' + fill +
    ' style="height:' + height + "px;" + (color ? "background:" + color + ";" : "") +
    "font-size:" + height + "px;line-height:" + height + 'px;mso-line-height-rule:exactly">&nbsp;</td>'
  );
}

function yellowRule(width, height) {
  return (
    '<table class="vstx-rule" role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse"><tr>' +
    ruleCell(height, YELLOW, ' width="' + width + '"').replace('style="', 'style="width:' + width + "px;") +
    "</tr></table>"
  );
}

// Classic Outlook ignores margins on tables, so components are spaced by a wrapper it does respect.
function spaced(margin, html) {
  return '<div class="vstx-space" style="margin:' + margin + '">' + html + "</div>";
}

// Headings and KPI tiles are rebuilt from a line's text alone, so a line holding links, @mentions, pictures,
// lists or tables keeps its own markup instead.
var KEEPS_ITS_MARKUP = /<(?:a|img|ul|ol|li|table|td|th)\b/i;

// A short line that is entirely bold, on its own, reads as a section heading.
function asHeading(attrs, inner, text) {
  if (KEEPS_ITS_MARKUP.test(inner)) return null;
  if (text.length < 2 || text.length > 70 || /[.!?,;]$/.test(text) || NOT_HEADINGS.test(text) || !/[a-z]/i.test(text)) return null;
  if (/<br\b/i.test(inner.replace(/(?:\s|<br\b[^>]*>)+$/i, ""))) return null;
  var wholeBlockBold = BOLD_STYLE.test(attrs);
  if (!wholeBlockBold && (inner.search(BOLD_RUN) === -1 || plainText(inner.replace(BOLD_RUN, "")))) return null;
  return headingHtml(text);
}

function headingHtml(text) {
  return (
    '<div class="vstx-heading" style="margin:16px 0 8px">' + yellowRule(24, 3) +
    '<div class="vstx-heading-text" style="font-family:' + FONT_STACK + ";font-size:18px;font-weight:" + STRONG_WEIGHT + ";line-height:1.3;color:" + INK + ';margin-top:8px">' +
    escapeHtml(text.replace(/\s*:\s*$/, "")) + "</div></div>"
  );
}

// Off by default (plainTitles): a short plain line with text before and after it reads as a section title, so titles
// pasted without bold still become headings. Greetings, sign-offs and lines ending in punctuation stay text.
var PLAIN_TITLE_MAX_WORDS = 6;
var PLAIN_TITLE_MAX_LENGTH = 40;
var NOT_PLAIN_TITLES = /^(good (?:morning|afternoon|evening|day)|morning|afternoon|evening|greetings|all|team|everyone|folks|ok|okay|yes|no|sure|noted|great|perfect|fyi|see (?:below|above|attached))\b/i;

function visibleText(html) {
  return plainText(html.replace(/<(style|script|head)\b[\s\S]*?<\/\1>/gi, " "));
}

function asPlainTitle(inner, text, whole, start, end) {
  if (text.length < 3 || text.length > PLAIN_TITLE_MAX_LENGTH || text.split(/\s+/).length > PLAIN_TITLE_MAX_WORDS) return null;
  if (!/^[A-ZÀ-ÖØ-Þ]/.test(text) || /[.!?,;:]$/.test(text) || NOT_HEADINGS.test(text) || NOT_PLAIN_TITLES.test(text)) return null;
  if (KEEPS_ITS_MARKUP.test(inner) || /<br\b/i.test(inner.replace(/(?:\s|<br\b[^>]*>)+$/i, ""))) return null;
  if (!visibleText(whole.slice(0, start)) || !visibleText(whole.slice(end))) return null;
  return headingHtml(text);
}

// "Action: …", "Decision: …", "Deadline: …" and similar become a labelled row.
function asCallout(inner, text) {
  var match = CALLOUT_LINE.exec(text);
  if (!match) return null;
  var keyword = match[1];
  var spec = CALLOUTS[keyword.toLowerCase().replace(/\s+/g, " ")];
  var label = new RegExp("^((?:\\s|<[^>]+>)*)" + keyword.split(/\s+/).join("(?:\\s|&nbsp;)+") + "\\s*:\\s*", "i");
  var body = inner.replace(label, "$1");
  if (!spec || body === inner) return null;
  // Classic Outlook merges neighbouring tables, so the gap between callouts is a row inside each one.
  return (
    '<table class="vstx-callout" role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">' +
    "<tr>" + ruleCell(8, "", ' colspan="2"') + "</tr><tr>" +
    '<td class="vstx-callout-label" width="' + CALLOUT_LABEL_WIDTH + '" bgcolor="' + PANEL + '" style="width:' + CALLOUT_LABEL_WIDTH + "px;background:" + PANEL +
    ";border-left:3px solid " + (spec.strong ? YELLOW : LIGHT_GREY) + ';padding:12px 6px 10px 12px;vertical-align:top;white-space:nowrap">' +
    calloutLabel(spec) + "</td>" +
    '<td bgcolor="' + PANEL + '" style="background:' + PANEL + ";border-radius:0 " + RADIUS + " " + RADIUS + ' 0;padding:10px 14px 10px 6px;vertical-align:top">' +
    '<div class="vstx-callout-text" style="' + BODY_TEXT + '">' + body + "</div></td>" +
    "</tr></table>"
  );
}

// A 20px label line starting 2px below the text's top padding centres on the text's first 24px line.
// The height sits on the label, not the cell: Word puts small text at the bottom of a tall cell line.
function calloutLabel(spec) {
  var letters = "font-family:" + FONT_STACK + ";font-size:11px;font-weight:" + STRONG_WEIGHT + ";letter-spacing:.08em;display:inline-block;line-height:20px;";
  var text = spec.label.toUpperCase();
  if (spec.chip) {
    return '<span style="' + letters + "background:" + INK + ";color:#ffffff;padding:0 7px;border-radius:" + SMALL_RADIUS + '">' + text + "</span>";
  }
  return '<span style="' + letters + "color:" + (spec.strong ? INK : GREY) + '">' + text + "</span>";
}

// "Spend $4.1M -8% vs plan [On track]" -> label "Spend", value "$4.1M", change "-8%", note "vs plan", status "on track".
function parseKpi(segment) {
  var status = null;
  var rest = segment.replace(STATUS_TAG, function (_tag, found) {
    status = status || found;
    return " ";
  });
  var value = KPI_VALUE.exec(rest);
  if (!value) return { label: "", value: rest.trim(), change: null, note: "", status: status, numeric: false };
  var label = rest.slice(0, value.index);
  var after = rest.slice(value.index + value[0].length);
  var change = KPI_CHANGE.exec(after);
  var note = change ? after.slice(0, change.index) + " " + after.slice(change.index + change[0].length) : after;
  var tidy = function (text) {
    return text.replace(/[()]/g, " ").replace(/\s+/g, " ").replace(/^[\s:,–—-]+|[\s:,–—-]+$/g, "");
  };
  label = tidy(label);
  note = tidy(note);
  if (!label && note) {
    label = note;
    note = "";
  }
  return { label: label, value: value[0].trim(), change: change ? change[0] : null, note: note, status: status, numeric: true };
}

function kpiTile(kpi, width) {
  var text = "font-family:" + FONT_STACK + ";";
  var lines = [];
  if (kpi.label) {
    lines.push(
      '<div class="vstx-kpi-label" style="' + text + "font-size:11px;font-weight:700;letter-spacing:.06em;color:" + GREY + ';line-height:1.35">' +
        escapeHtml(kpi.label.toUpperCase()) + "</div>"
    );
  }
  lines.push(
    '<div class="vstx-kpi-value" style="' + text + "font-size:" + (kpi.numeric ? "24px" : "16px") + ";font-weight:" + STRONG_WEIGHT + ";color:" + INK + ';line-height:1.2;padding-top:4px">' +
      escapeHtml(kpi.value) + "</div>"
  );
  if (kpi.change || kpi.note) {
    var change = "";
    if (kpi.change) {
      var falling = /^[\-−]/.test(kpi.change);
      change =
        "<span style=\"font-family:'Segoe UI Symbol',Arial,sans-serif;font-size:10px;color:" + INK + '">' + (falling ? "&#9660;" : "&#9650;") + "</span>&nbsp;" +
        '<span style="' + text + "font-weight:700;color:" + INK + '">' + escapeHtml(kpi.change.replace(/^[+\-−]\s?/, "")) + "</span>";
    }
    var note = kpi.note ? '<span style="' + text + "color:" + GREY + '">' + (change ? "&nbsp;" : "") + escapeHtml(kpi.note) + "</span>" : "";
    lines.push('<div class="vstx-kpi-change" style="' + text + "font-size:13px;color:" + INK + ';line-height:1.4;padding-top:2px">' + change + note + "</div>");
  }
  if (kpi.status) lines.push('<div class="vstx-kpi-status" style="padding-top:6px">' + statusPill("", kpi.status) + "</div>");
  return (
    '<td class="vstx-kpi-tile" width="' + width + '%" bgcolor="' + PANEL + '" style="width:' + width + "%;background:" + PANEL +
    ";border-top:3px solid " + YELLOW + ";border-radius:" + RADIUS + ';padding:10px 12px 12px;vertical-align:top">' + lines.join("") + "</td>"
  );
}

// "KPI: Revenue $1.2M +6% | Margin 18%" becomes a row of tiles. Arrows show direction only: whether a change
// is good depends on the metric (spend down is good news), so colour comes only from a status tag you add.
function asKpiRow(text) {
  if (!KPI_LINE.test(text)) return null;
  var kpis = text
    .replace(KPI_LABEL, "")
    .split("|")
    .map(function (segment) {
      return segment.trim();
    })
    .filter(Boolean)
    .map(parseKpi);
  if (!kpis.some(function (kpi) { return kpi.numeric; })) return null;
  var columns = Math.max(2, Math.min(kpis.length, KPI_PER_ROW));
  var width = Math.floor((100 - KPI_GAP * (columns - 1)) / columns);
  var gap = '<td class="vstx-kpi-gap" width="' + KPI_GAP + '%" style="width:' + KPI_GAP + '%;font-size:1px;line-height:1px">&nbsp;</td>';
  var rows = [];
  for (var start = 0; start < kpis.length; start += columns) {
    var cells = kpis.slice(start, start + columns).map(function (kpi) {
      return kpiTile(kpi, width);
    });
    while (cells.length < columns) cells.push('<td class="vstx-kpi-gap" width="' + width + '%">&nbsp;</td>');
    rows.push("<tr>" + ruleCell(8, "", ' colspan="' + (columns * 2 - 1) + '"') + "</tr><tr>" + cells.join(gap) + "</tr>");
  }
  return (
    '<table class="vstx-kpi" role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">' +
    rows.join("") +
    "<tr>" + ruleCell(8, "", ' colspan="' + (columns * 2 - 1) + '"') + "</tr></table>"
  );
}

// The site a browser would open, or "" when the address could mislead: a user name before "@"
// (https://sharepoint.com:x@evil.example opens evil.example) or a host that isn't plain ASCII (lookalike letters).
// Browsers end the host at a backslash too, so it is split the same way.
function hostOf(href) {
  var match = /^https?:\/\/([^\/\\?#]*)/i.exec(href);
  if (!match) return "";
  var host = match[1].toLowerCase().replace(/:\d*$/, "");
  return /^[a-z0-9.-]+$/.test(host) ? host.replace(/^www\./, "") : "";
}

function smartLinkLabel(href, host) {
  if (/(^|\.)powerbi\.com$/.test(host)) return "Open Power BI report";
  if (/(^|\.)sharepoint\.com$/.test(host)) {
    if (/\/:x:\//i.test(href)) return "Open the Excel workbook";
    if (/\/:w:\//i.test(href)) return "Open the Word document";
    if (/\/:p:\//i.test(href)) return "Open the presentation";
    if (/\/:b:\//i.test(href)) return "Open the PDF";
    if (/\/:f:\//i.test(href)) return "Open the folder";
    return "Open in SharePoint";
  }
  if (/(^|\.)teams\.(microsoft|live)\.com$/.test(host)) return "Open in Teams";
  if (/^forms\.(office|microsoft)\.com$/.test(host)) return "Open the form";
  if (/(^|\.)github\.com$/.test(host)) return "View on GitHub";
  return "Open link";
}

// Link text that names a site (with or without https://, zero-width characters ignored) gets a neutral label unless
// it names the site the link really opens: a button must never show one site and open another.
function looksLikeAddress(text, host) {
  var clean = text.replace(/[​-‏⁠﻿]/g, "");
  if (/https?\s*:\s*\/\/|www\./i.test(clean)) return true;
  var named = /(?:^|[^\w.-])((?:[a-z0-9-]+\.)+[a-z]{2,})(?=$|[^\w.-])/i.exec(clean);
  if (!named) return false;
  var site = named[1].toLowerCase().replace(/^www\./, "");
  return site !== host && host.slice(-(site.length + 1)) !== "." + site;
}

// A link alone on its own line becomes a button, with the website shown underneath so people know where it goes.
// Exactly one link and no picture: anything more would be lost from the button.
function asButton(inner) {
  var match = LONE_LINK.exec(inner);
  if (!match || (inner.match(/<a\b/gi) || []).length !== 1 || /<img\b/i.test(inner)) return null;
  var hrefMatch = /\bhref\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(match[2]);
  var href = hrefMatch ? hrefMatch[1] || hrefMatch[2] : "";
  var linkText = plainText(match[3]);
  var host = hostOf(href);
  if (!host || !linkText) return null;
  var label = looksLikeAddress(linkText, host) ? smartLinkLabel(href, host) : linkText;
  var text = "font-family:" + FONT_STACK + ";";
  return spaced(
    COMPONENT_MARGIN,
    '<table class="vstx-button" role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse"><tr>' +
      '<td bgcolor="' + INK + '" style="background:' + INK + ";border-bottom:3px solid " + YELLOW + ";border-radius:" + RADIUS + ';padding:11px 20px">' +
      '<a href="' + href + '" style="' + text + 'font-size:14px;font-weight:700;color:#ffffff;text-decoration:none">' + escapeHtml(label) + "&nbsp;&#8594;</a>" +
      "</td></tr></table>" +
      '<div class="vstx-button-caption" style="' + text + "font-size:12px;color:" + GREY + ';margin:4px 0 0">' + escapeHtml(host) + "</div>"
  );
}

function smartBlocks(html, features) {
  return html.replace(INNERMOST_BLOCK, function (block, name, attrs, inner, offset, whole) {
    if (OWN_MARKUP.test(attrs) || insideStructure(whole, offset)) return block;
    var text = plainText(inner);
    if (!text) return block;
    return (
      (features.buttons && asButton(inner)) ||
      (features.kpis && !KEEPS_ITS_MARKUP.test(inner) && asKpiRow(text)) ||
      (features.callouts && asCallout(inner, text)) ||
      (features.headings && asHeading(attrs, inner, text)) ||
      (features.plainTitles && asPlainTitle(inner, text, whole, offset, offset + block.length)) ||
      block
    );
  });
}

function isTextOffLimits(name, attrs, skipTables) {
  return (
    /^(a|head|style|script|title|code|pre)$/.test(name) ||
    (skipTables && name === "table") ||
    MONOSPACE.test(attrs) ||
    COMPONENT_MARKUP.test(attrs)
  );
}

// Applies transform to text between tags, skipping links, code, our own components and (optionally) tables.
function mapText(html, skipTables, transform) {
  var open = [];
  return html.replace(/(<[^>]+>)|([^<]+)/g, function (token, tag, text) {
    if (text !== undefined) {
      var skip = open.length > 0 && open[open.length - 1].skip;
      return skip ? text : transform(text);
    }
    var parts = /^<(\/?)([a-z][a-z0-9]*)\b([^>]*)>$/i.exec(tag);
    if (!parts) return tag;
    var name = parts[2].toLowerCase();
    if (parts[1]) {
      for (var i = open.length - 1; i >= 0; i--) {
        if (open[i].name === name) {
          open.length = i;
          break;
        }
      }
      return tag;
    }
    if (VOID_TAGS.test(name) || /\/>$/.test(tag)) return tag;
    var parentSkips = open.length > 0 && open[open.length - 1].skip;
    open.push({ name: name, skip: parentSkips || isTextOffLimits(name, parts[3], skipTables) });
    return tag;
  });
}

function statusPill(_tag, status) {
  var key = status.toLowerCase();
  var label = key.charAt(0).toUpperCase() + key.slice(1);
  return (
    '<span class="vstx-status" style="display:inline-block;padding:1px 9px 1px 7px;background:' + PANEL + ";border:1px solid #d3d2d2;border-radius:" + PILL_RADIUS + ";" +
    "font-family:" + FONT_STACK + ";font-size:12px;font-weight:700;line-height:1.5;color:" + INK + ';white-space:nowrap">' +
    '<span style="color:' + STATUS_COLORS[key] + ";font-family:'Segoe UI Symbol',Arial,sans-serif\">&#11042;</span>&nbsp;" + label + "</span>"
  );
}

function markStatuses(html) {
  return mapText(html, false, function (text) {
    return text.replace(STATUS_TAG, statusPill);
  });
}

// Key figures get a light yellow highlight; an email full of figures gets bold only, so the yellow never takes over.
// A background, not an underline: Outlook on the web strips borders from inline text when the email is read.
function markKeyFigures(html) {
  var count = 0;
  mapText(html, true, function (text) {
    count += (text.match(KEY_FIGURE) || []).length;
    return text;
  });
  var style = count > FIGURE_LIMIT ? "font-weight:700" : "font-weight:700;background:" + FIGURE_HIGHLIGHT + ";padding:0 2px";
  return mapText(html, true, function (text) {
    return text.replace(KEY_FIGURE, function (figure) {
      return '<span class="vstx-figure" style="' + style + '">' + figure + "</span>";
    });
  });
}

// ---- @mentions ----
// Outlook's mention is a mailto link whose id (OWAAM…) and address drive the "you were mentioned" notice,
// so the link itself is never changed: only its look, plus a styled wrapper around the name.
var MENTION_LINK = /<a\b([^>]*)>([\s\S]*?)<\/a>/gi;
var MAILTO_HREF = /\bhref\s*=\s*["']?mailto:/i;
var OUTLOOK_MENTION_ID = /\bid\s*=\s*["']?(?:x_)?OWAAM/i;

function markMentions(html) {
  return html.replace(MENTION_LINK, function (link, attrs, inner) {
    if (!MAILTO_HREF.test(attrs) || /vstx-mention/.test(inner)) return link;
    if (!OUTLOOK_MENTION_ID.test(attrs) && plainText(inner).charAt(0) !== "@") return link;
    var opening = mergeStyle("<a" + attrs + ">", "a", attrs, "color:" + INK + ";text-decoration:none");
    return (
      opening +
      '<span class="vstx-mention" style="font-family:' + FONT_STACK + ";font-weight:700;color:" + INK + ";background:" + PANEL +
      ";border-bottom:2px solid " + YELLOW + ';padding:0 4px;white-space:nowrap">' + inner + "</span></a>"
    );
  });
}

// ---- Timed agenda ----
// Colon times only: "1.50 per unit" on two lines must not become an agenda.
var CLOCK = "\\d{1,2}(?::\\d{2})?\\s*(?:[ap]\\.?m\\.?)?";
var AGENDA_LINE = new RegExp("^(" + CLOCK + ")(?:\\s*(?:-|–|—|to)\\s*(" + CLOCK + "))?(?=[\\s\\-–—:|•·])\\s*(?:[-–—:|•·]\\s*)?(\\S[\\s\\S]*)$", "i");
var AGENDA_PREFIX = new RegExp("^((?:\\s|<[^>]+>)*)" + CLOCK + "(?:\\s*(?:-|–|—|to)\\s*" + CLOCK + ")?\\s*(?:[-–—:|•·]\\s*)?", "i");

// "9:30", "9:30am", "2pm" -> minutes after midnight; a bare "9" isn't a time.
function clockMinutes(text) {
  var match = /^(\d{1,2})(?::(\d{2}))?\s*([ap])?/i.exec(text.trim());
  if (!match || (!match[2] && !match[3])) return null;
  var hour = parseInt(match[1], 10);
  var minute = match[2] ? parseInt(match[2], 10) : 0;
  if (hour > 23 || minute > 59 || (match[3] && hour > 12)) return null;
  if (match[3]) hour = (hour % 12) + (match[3].toLowerCase() === "p" ? 12 : 0);
  return hour * 60 + minute;
}

function agendaEntry(inner) {
  var parts = AGENDA_LINE.exec(plainText(inner));
  if (!parts) return null;
  var start = clockMinutes(parts[1]);
  var end = parts[2] ? clockMinutes(parts[2]) : null;
  if (start === null || (parts[2] && end === null)) return null;
  var item = inner.replace(AGENDA_PREFIX, "$1");
  if (item === inner || !plainText(item)) return null;
  var length = end !== null && end > start && end - start <= 720 ? end - start : null;
  return { from: parts[1].trim(), to: parts[2] ? parts[2].trim() : "", length: length, item: item };
}

function formatLength(minutes) {
  var hours = Math.floor(minutes / 60);
  var rest = minutes % 60;
  return (hours ? hours + " h" : "") + (hours && rest ? " " : "") + (rest ? rest + " min" : "");
}

function agendaTable(entries) {
  var text = "font-family:" + FONT_STACK + ";";
  var rows = entries.map(function (entry) {
    var clock = escapeHtml(entry.from) + (entry.to ? "&#8211;" + escapeHtml(entry.to) : "");
    var length = entry.length
      ? '<div class="vstx-agenda-length" style="' + text + "font-size:12px;color:" + GREY + ';line-height:1.4">' + formatLength(entry.length) + "</div>"
      : "";
    return (
      "<tr>" +
      '<td class="vstx-agenda-time" width="' + AGENDA_TIME_WIDTH + '" bgcolor="' + PANEL + '" style="width:' + AGENDA_TIME_WIDTH + "px;background:" + PANEL +
      ";border-left:3px solid " + YELLOW + ";border-bottom:2px solid #ffffff;padding:8px 12px;vertical-align:top;white-space:nowrap\">" +
      '<div class="vstx-agenda-clock" style="' + text + "font-size:14px;font-weight:700;color:" + INK + ';line-height:1.4">' + clock + "</div>" + length + "</td>" +
      '<td style="padding:8px 14px;vertical-align:top;border-bottom:1px solid ' + HAIRLINE + '">' +
      '<div class="vstx-agenda-text" style="' + BODY_TEXT + '">' + entry.item + "</div></td>" +
      "</tr>"
    );
  });
  return (
    '<table class="vstx-agenda" role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">' +
    "<tr>" + ruleCell(8, "", ' colspan="2"') + "</tr>" + rows.join("") + "<tr>" + ruleCell(8, "", ' colspan="2"') + "</tr></table>"
  );
}

// Two or more lines in a row that start with a time become an agenda; a single "10:00 call" stays a sentence.
function withAgenda(html) {
  var blocks = new RegExp(INNERMOST_BLOCK.source, "gi");
  var runs = [];
  var run = [];
  var match;
  var closeRun = function () {
    if (run.length >= 2) runs.push(run);
    run = [];
  };
  while ((match = blocks.exec(html))) {
    var entry = OWN_MARKUP.test(match[2]) || insideStructure(html, match.index) ? null : agendaEntry(match[3]);
    var adjacent = run.length && !html.slice(run[run.length - 1].end, match.index).trim();
    if (!entry || !adjacent) closeRun();
    if (entry) run.push({ start: match.index, end: blocks.lastIndex, entry: entry });
  }
  closeRun();
  for (var i = runs.length - 1; i >= 0; i--) {
    var first = runs[i][0];
    var last = runs[i][runs[i].length - 1];
    html = html.slice(0, first.start) + agendaTable(runs[i].map(function (line) { return line.entry; })) + html.slice(last.end);
  }
  return html;
}

function smartFormat(html) {
  var features = featureSettings();
  var result = features.agenda ? withAgenda(html) : html;
  result = smartBlocks(result, features);
  if (features.mentions) result = markMentions(result);
  if (features.statuses) result = markStatuses(result);
  if (features.figures) result = markKeyFigures(result);
  return result;
}

// ---- Lists ----
// Lists become small tables: classic Outlook (Word) draws its own "•" on every <li>, whatever the CSS says.

function hexagonCell(color) {
  return (
    '<td width="' + BULLET_COLUMN + '" style="width:' + BULLET_COLUMN + "px;vertical-align:top;padding:0;" +
    "font-size:" + BULLET_SIZE + "px;line-height:24px;mso-line-height-rule:exactly\">" +
    '<span style="color:' + color + ";font-size:" + BULLET_SIZE + "px;line-height:24px;font-family:'Segoe UI Symbol',Arial,sans-serif\">&#11042;</span></td>"
  );
}

function numberBadge(number) {
  return (
    '<table class="vstx-badge" role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse"><tr>' +
    '<td width="' + BADGE_SIZE + '" height="' + BADGE_SIZE + '" align="center" bgcolor="' + YELLOW + '" style="' + BADGE_CSS + '">' +
    '<span style="font-family:' + FONT_STACK + ";font-size:12px;font-weight:" + STRONG_WEIGHT + ";color:" + INK + '">' + number + "</span></td>" +
    "</tr></table>"
  );
}

function endOfList(html, from) {
  var tags = /<(\/?)(?:ul|ol)\b[^>]*>/gi;
  tags.lastIndex = from;
  var depth = 1;
  var match;
  while ((match = tags.exec(html))) {
    depth += match[1] ? -1 : 1;
    if (depth === 0) return { innerEnd: match.index, end: tags.lastIndex };
  }
  return null;
}

// Outlook places a sub-list between items (<ol><li>..</li><ul>..</ul><li>..</li></ol>), not inside one,
// so a sub-list found between items is attached to the item above it. Pasted lists can also carry text
// that isn't in any item (a line as a <div>, bare text, an unclosed <li>); that is kept with the item above.
function topLevelItems(inner) {
  var tags = /<(\/?)(ul|ol|li)\b[^>]*>/gi;
  var items = [];
  var leading = "";
  var depth = 0;
  var itemStart = -1;
  var subListStart = -1;
  var cursor = 0;
  var match;
  var attach = function (piece) {
    if (items.length) items[items.length - 1] += piece;
    else leading += piece;
  };
  var keepStray = function (to) {
    var stray = inner.slice(cursor, to);
    if (plainText(stray) || /<img\b/i.test(stray)) attach(stray);
  };
  var closeItem = function (at) {
    items.push((items.length ? "" : leading) + inner.slice(itemStart, at));
    itemStart = -1;
  };
  while ((match = tags.exec(inner))) {
    var closing = match[1] === "/";
    var isItem = match[2].toLowerCase() === "li";
    if (depth === 0 && itemStart === -1 && !closing) keepStray(match.index);
    if (!isItem) {
      if (!closing && depth === 0 && itemStart === -1) subListStart = match.index;
      depth += closing ? -1 : 1;
      if (closing && depth === 0 && subListStart !== -1) {
        attach(inner.slice(subListStart, tags.lastIndex));
        subListStart = -1;
        cursor = tags.lastIndex;
      }
    } else if (depth === 0 && !closing) {
      if (itemStart !== -1) closeItem(match.index);
      itemStart = tags.lastIndex;
    } else if (depth === 0 && itemStart !== -1) {
      closeItem(match.index);
      cursor = tags.lastIndex;
    }
  }
  if (itemStart !== -1) {
    closeItem(inner.length);
  } else {
    keepStray(inner.length);
  }
  return items;
}

// A list holding only other lists, with no item of its own: what Outlook makes of bullets indented twice.
function isParentlessList(inner) {
  if (topLevelItems(inner).length || !/<(?:ul|ol)\b/i.test(inner)) return false;
  return !plainText(withoutLists(inner));
}

function withoutLists(html) {
  return splitTopLevelLists(html)
    .map(function (part) {
      return part.inner === undefined ? part.text : "";
    })
    .join("");
}

function splitTopLevelLists(html) {
  var openers = new RegExp(LIST_OPENER.source, "gi");
  var parts = [];
  var cursor = 0;
  var open;
  while ((open = openers.exec(html))) {
    var close = endOfList(html, openers.lastIndex);
    if (!close) break;
    parts.push({ text: html.slice(cursor, open.index) });
    parts.push({ open: html.slice(open.index, openers.lastIndex), inner: html.slice(openers.lastIndex, close.innerEnd), close: html.slice(close.innerEnd, close.end) });
    cursor = close.end;
    openers.lastIndex = close.end;
  }
  parts.push({ text: html.slice(cursor) });
  return parts;
}

// Pasted sub-points often arrive as a parentless list right after the list they belong to (with only
// empty wrappers in between); they are moved into that list so they sit under its last item.
var ONLY_WRAPPERS = /^(?:\s|&nbsp;|<\/?(?:div|p|span)\b[^>]*>|<br\s*\/?>)*$/i;

function attachParentlessLists(html) {
  var parts = splitTopLevelLists(html);
  var previous = null;
  parts.forEach(function (part) {
    if (part.inner === undefined) {
      if (!ONLY_WRAPPERS.test(part.text)) previous = null;
      return;
    }
    if (previous && isParentlessList(part.inner)) {
      previous.inner += part.inner;
      part.moved = true;
      return;
    }
    previous = part;
  });
  return parts
    .map(function (part) {
      if (part.inner === undefined) return part.text;
      return part.moved ? "" : part.open + part.inner + part.close;
    })
    .join("");
}

function listTable(rows, margin) {
  return spaced(
    margin,
    '<table class="' + VSTX_LIST_CLASS + '" role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">' +
      rows.join("") +
      "</table>"
  );
}

function itemText(content, bulletDepth) {
  return '<div style="' + BODY_TEXT + '">' + withListTables(content.replace(LEADING_HEXAGONS, "$1"), true, bulletDepth) + "</div>";
}

function numberedTable(attrs, items) {
  var start = /\bstart=["']?(\d+)/i.exec(attrs);
  var number = start ? parseInt(start[1], 10) : 1;
  return listTable(
    items.map(function (content) {
      return (
        '<tr><td width="' + (BADGE_SIZE + 12) + '" style="width:' + BADGE_SIZE + 'px;vertical-align:top;padding:2px 12px 8px 0">' + numberBadge(number++) + "</td>" +
        '<td style="' + NUMBERED_TEXT_CSS + '">' + itemText(content, 0) + "</td></tr>"
      );
    }),
    COMPONENT_MARGIN
  );
}

function bulletTable(items, bulletDepth, nested) {
  var color = bulletDepth > 0 ? LIGHT_GREY : GREY;
  return listTable(
    items.map(function (content) {
      return "<tr>" + hexagonCell(color) + '<td style="vertical-align:top;padding:0 0 4px 6px">' + itemText(content, bulletDepth + 1) + "</td></tr>";
    }),
    nested ? "4px 0" : COMPONENT_MARGIN
  );
}

// Only a top-level numbered list keeps its numbers; every sub-list, numbered or not, gets hexagon bullets.
function withListTables(html, nested, bulletDepth) {
  var openers = new RegExp(LIST_OPENER.source, "gi");
  var output = "";
  var cursor = 0;
  var open;
  while ((open = openers.exec(html))) {
    var close = endOfList(html, openers.lastIndex);
    if (!close) break;
    var inner = html.slice(openers.lastIndex, close.innerEnd);
    var items = topLevelItems(inner);
    var numbered = open[1].toLowerCase() === "ol" && !nested;
    var table;
    if (items.length) table = numbered ? numberedTable(open[2], items) : bulletTable(items, bulletDepth || 0, nested);
    else table = isParentlessList(inner) ? withListTables(inner, nested, bulletDepth) : html.slice(open.index, close.end);
    output += html.slice(cursor, open.index) + table;
    cursor = close.end;
    openers.lastIndex = close.end;
  }
  return output + html.slice(cursor);
}

// ---- Tables ----

function endOfTable(html, from) {
  var tags = /<(\/?)table\b[^>]*>/gi;
  tags.lastIndex = from;
  var depth = 1;
  var match;
  while ((match = tags.exec(html))) {
    depth += match[1] ? -1 : 1;
    if (depth === 0) return tags.lastIndex;
  }
  return -1;
}

function eachTopLevelTable(html, transform) {
  var openers = /<table\b([^>]*)>/gi;
  var output = "";
  var cursor = 0;
  var open;
  while ((open = openers.exec(html))) {
    var end = endOfTable(html, openers.lastIndex);
    if (end === -1) break;
    output += html.slice(cursor, open.index) + transform(html.slice(open.index, end), open[1]);
    cursor = end;
    openers.lastIndex = end;
  }
  return output + html.slice(cursor);
}

function isOwnTable(attrs) {
  return OWN_MARKUP.test(attrs);
}

function withClass(openingTag, className) {
  if (/\bclass\s*=\s*"/i.test(openingTag)) return openingTag.replace(/\bclass\s*=\s*"/i, 'class="' + className + " ");
  return openingTag.replace(/^<table/i, '<table class="' + className + '"');
}

function withBgcolor(tag, color) {
  return tag.replace(/\sbgcolor\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/i, "").replace(/>$/, ' bgcolor="' + color + '">');
}

// Hexagon bullets only come from this add-in: a table holding them is one of its lists, copied from an earlier
// email with the class names Outlook on the web strips, so it is left as it is rather than branded as a data table.
var HEXAGON_BULLET = /&#11042;|⬢/;

// A table without colours of its own was inserted in Outlook; pasted Excel or Power BI tables keep their look.
function markInsertedTables(html) {
  return eachTopLevelTable(html, function (table, attrs) {
    if (isOwnTable(attrs)) return table;
    var openingEnd = table.indexOf(">") + 1;
    var pasted = /background|bgcolor|mso-|class=["']?xl/i.test(table) || HEXAGON_BULLET.test(table);
    if (pasted || /<table\b/i.test(table.slice(openingEnd))) return table;
    return withClass(table.slice(0, openingEnd), VSTX_TABLE_CLASS) + table.slice(openingEnd);
  });
}

function brandTables(html) {
  return eachTopLevelTable(html, function (table, attrs) {
    if (attrs.indexOf(VSTX_TABLE_CLASS) === -1) return table;
    var openingEnd = table.indexOf(">") + 1;
    var row = 0;
    var rows = table.slice(openingEnd).replace(/(<tr\b[^>]*>)([\s\S]*?)(<\/tr>)/gi, function (_row, open, cells, close) {
      var header = row === 0;
      var fill = header ? INK : row % 2 === 0 ? PANEL : "#ffffff";
      var css = header ? HEADER_CELL_CSS : BODY_CELL_CSS + ";background:" + fill;
      row++;
      var styled = cells.replace(/<(td|th)\b([^>]*)>/gi, function (tag, name, cellAttrs) {
        return withBgcolor(mergeStyle(tag, name, cellAttrs, css), fill);
      });
      return open + (header ? styled.replace(/([;"'\s])color:#231f20/gi, "$1color:#ffffff") : styled) + close;
    });
    return mergeStyle(table.slice(0, openingEnd), "table", attrs, TABLE_CSS) + rows;
  });
}

function restyleOwnPart(html) {
  return withListTables(attachParentlessLists(brandTables(smartFormat(brandText(markInsertedTables(html))))), false, 0);
}

// ---- Safety net: styling may move and decorate words, never drop them ----
var STYLING_ONLY_WORDS = { kpi: true, kpis: true };

function ownWords(html) {
  var text = plainText(html.replace(/<(style|script|head)\b[\s\S]*?<\/\1>/gi, " ").replace(/<[^>]+>/g, " "))
    .toLowerCase()
    .replace(/&[#\w]+;/g, " ")
    .replace(/(?:https?:\/\/|www\.)\S+/g, " ");
  return text.split(/[^a-z0-9À-ɏ]+/).filter(function (word) {
    return word && !STYLING_ONLY_WORDS[word];
  });
}

function keepsWords(before, after) {
  var available = {};
  ownWords(after).forEach(function (word) {
    available[word] = (available[word] || 0) + 1;
  });
  return ownWords(before).every(function (word) {
    return available[word]-- > 0;
  });
}

// If full styling would lose any of the sender's words, fall back to fonts and colours only;
// if even that would, the text goes exactly as written.
function safeRestyle(html) {
  var styled = restyleOwnPart(html);
  if (keepsWords(html, styled)) return styled;
  var basic = brandText(html);
  return keepsWords(html, basic) ? basic : html;
}

// ---- Frame: brand bar and signature ----

function brandBarHtml() {
  return (
    '<table id="' + VSTX_BAR_ID + '" role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">' +
    "<tr>" + ruleCell(6, INK, ' colspan="2"') + "</tr>" +
    "<tr>" + ruleCell(3, "", ' width="9%"') + ruleCell(3, YELLOW, "") + "</tr>" +
    "<tr>" + ruleCell(24, "", ' colspan="2"') + "</tr>" +
    "</table>"
  );
}

// The directory lists people as "Last, First"; a signature reads "First Last".
function profileName() {
  var profile = typeof Office !== "undefined" && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile;
  if (!profile) return "";
  var display = (profile.displayName || "").trim();
  var lastFirst = /^([^,]+),\s*(.+)$/.exec(display);
  if (lastFirst) {
    // "Reyes, Jordan (Contractor)" or "Reyes, Jordan - Finance": the note stays after the full name.
    var given = /^(.*?)(\s+(?:\(.*|[-–—|].*))?$/.exec(lastFirst[2]);
    return given[1] + " " + lastFirst[1] + (given[2] || "");
  }
  return display || (profile.emailAddress || "").split("@")[0];
}

// Until someone saves their own details, the signature carries just their name from Outlook.
function signatureDefaults() {
  return { name: profileName(), title: "", phone: "" };
}

function isSignatureSetUp() {
  return typeof roamingValue(SIGNATURE_SETTINGS_KEY).name === "string";
}

function signatureDetails() {
  var saved = roamingValue(SIGNATURE_SETTINGS_KEY);
  var defaults = signatureDefaults();
  var pick = function (key) {
    return typeof saved[key] === "string" ? saved[key].trim() : defaults[key];
  };
  return { name: pick("name") || defaults.name, title: pick("title"), phone: pick("phone") };
}

function signatureLines(details) {
  var text = "font-family:" + FONT_STACK + ";";
  var lines = ['<span style="' + text + "font-size:16px;font-weight:700;letter-spacing:.01em;color:" + INK + '">' + escapeHtml(details.name) + "</span>"];
  if (details.title) lines.push('<span style="' + text + "font-size:13px;color:" + INK + '">' + escapeHtml(details.title) + "</span>");
  if (details.phone) {
    lines.push('<span style="' + text + "font-size:13px;color:" + GREY + '">' + escapeHtml(details.phone).replace(/-/g, "&#8209;") + "</span>");
  }
  return lines.join("<br>");
}

function signatureHtml(details) {
  var text = "font-family:" + FONT_STACK + ";";
  details = details || signatureDetails();
  var panel = ' bgcolor="' + PANEL + '"';
  var panelCss = "background:" + PANEL + ";";
  return (
    '<table id="' + VSTX_SIGNATURE_ID + '" role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">' +
    "<tr>" + ruleCell(24, "", ' colspan="3"') + "</tr>" +
    "<tr>" +
    '<td width="' + LOGO_WIDTH + '"' + panel + ' style="' + panelCss + "border-radius:" + RADIUS + " 0 0 " + RADIUS + ";width:" + LOGO_WIDTH + 'px;padding:12px 14px 12px 16px;vertical-align:middle">' +
    '<img src="' + LOGO_URL + '" width="' + LOGO_WIDTH + '" height="' + LOGO_HEIGHT + '" alt="Versatex"' +
    ' style="display:block;width:' + LOGO_WIDTH + "px;height:" + LOGO_HEIGHT + "px;border:0;outline:none;" +
    text + "font-size:12px;font-weight:700;color:" + INK + '">' +
    "</td>" +
    "<td" + panel + ' style="' + panelCss + "vertical-align:middle;padding:10px 16px 10px 14px;border-left:3px solid " + YELLOW + '">' +
    signatureLines(details) +
    "</td>" +
    '<td align="right"' + panel + ' style="' + panelCss + "border-radius:0 " + RADIUS + " " + RADIUS + ' 0;vertical-align:middle;padding:10px 16px 10px 10px">' +
    '<img src="' + HEXAGONS_URL + '" width="' + HEXAGONS_WIDTH + '" height="' + HEXAGONS_HEIGHT + '" alt=""' +
    ' style="display:block;width:' + HEXAGONS_WIDTH + "px;height:" + HEXAGONS_HEIGHT + 'px;border:0;outline:none;margin-left:auto">' +
    "</td>" +
    "</tr></table>"
  );
}

// Once you've introduced yourself in a thread, later replies carry a one-line signature instead of the full panel.
function shortSignatureHtml(details) {
  var text = "font-family:" + FONT_STACK + ";font-size:13px;";
  var separator = '<span style="' + text + "color:" + LIGHT_GREY + '">&nbsp;&nbsp;|&nbsp;&nbsp;</span>';
  details = details || signatureDetails();
  var parts = ['<span style="' + text + "font-size:14px;font-weight:700;color:" + INK + '">' + escapeHtml(details.name) + "</span>"];
  if (details.title) parts.push('<span style="' + text + "color:" + INK + '">' + escapeHtml(details.title) + "</span>");
  if (details.phone) parts.push('<span style="' + text + "color:" + GREY + '">' + escapeHtml(details.phone).replace(/-/g, "&#8209;") + "</span>");
  return (
    '<table id="' + VSTX_SIGNATURE_ID + '" class="vstx-signature-short" role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">' +
    "<tr>" + ruleCell(16, "", "") + "</tr>" +
    '<tr><td style="border-left:3px solid ' + YELLOW + ';padding:2px 0 2px 12px;line-height:20px">' + parts.join(separator) + "</td></tr>" +
    "</table>"
  );
}

function signedEarlierInThread(quoted, name) {
  if (!name) return false;
  var signature = new RegExp(VSTX_SIGNATURE_ID + "[\\s\\S]{0,3000}?>\\s*" + escapeHtml(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\s*<", "i");
  return signature.test(quoted);
}

function insertAfterBodyTag(html, fragment) {
  var bodyTag = /<body[^>]*>/i.exec(html);
  if (!bodyTag) return fragment + html;
  var at = bodyTag.index + bodyTag[0].length;
  return html.slice(0, at) + fragment + html.slice(at);
}

// Only this email's own frame counts: a styled email pasted in from elsewhere carries x_-prefixed ids, and the
// words "vstx-signature" typed as text are just text.
var SIGNATURE_TABLE = /<table\b[^>]*\bid=["']?vstx-signature["'\s>]/i;
var BAR_TABLE = /<table\b[^>]*\bid=["']?vstx-bar["'\s>]/i;

// Each time an add-in sets the body, Outlook adds "x_" to every id and drops the classes and styles that restyling
// relies on to recognise its own work. Such a frame at the very top means this draft was styled already (Apply email
// style, then Send), and styling it again would add a second frame and double the accents.
var REWRITTEN_BAR_TABLE = /<table\b[^>]*\bid=["']?(?:x_)+vstx-bar["'\s>]/i;

function styledEarlierInThisDraft(ownPart) {
  var bar = REWRITTEN_BAR_TABLE.exec(ownPart);
  if (!bar) return false;
  var before = ownPart.slice(0, bar.index);
  return !visibleText(before) && !/<img\b/i.test(before);
}

// ---- A later pass in a draft Outlook's note says was styled ----

// Outlook rewrites the markup each time the body is set, so the frame alone can't say whose it is. A note kept on the
// draft can: with it, the frame is this email's own wherever it sits, nothing gets a second one, and only what was
// added since is styled. Without the note (older Outlook, or the note can't be read) a later pass falls back to
// styledEarlierInThisDraft.
var ANY_BAR_TABLE = /<table\b[^>]*\bid=["']?(?:x_)*vstx-bar["'\s>]/i;
var ANY_SIGNATURE_TABLE = /<table\b[^>]*\bid=["']?(?:x_)*vstx-signature["'\s>]/i;
var MARKUP_TOKEN = /<!--[\s\S]*?-->|<(\/?)([a-z][a-z0-9:]*)\b(?:[^>"']|"[^"]*"|'[^']*')*>/gi;

function colourPattern(hex) {
  var value = parseInt(hex.slice(1), 16);
  return hex + "|rgb\\(\\s*" + (value >> 16) + ",\\s*" + ((value >> 8) & 255) + ",\\s*" + (value & 255) + "\\s*\\)";
}

// Every block the add-in styles carries its font or one of its colours in its markup, before and after Outlook's
// rewrite, or a hexagon bullet.
var OWN_LOOK = new RegExp(
  [FONT_STACK.split(",")[0], "vstx-"].concat([INK, YELLOW, HAIRLINE, FIGURE_HIGHLIGHT].map(colourPattern)).join("|"),
  "i"
);

function hasOwnLook(part) {
  return OWN_LOOK.test((part.match(/<[^>]*>/g) || []).join(" ")) || HEXAGON_BULLET.test(part);
}

// The body's top-level elements and the text between them, in order. Null unless every tag closes where it should,
// so a later pass never moves half an element.
function topLevelParts(html) {
  var tokens = new RegExp(MARKUP_TOKEN.source, "gi");
  var parts = [];
  var open = [];
  var from = 0;
  var token;
  while ((token = tokens.exec(html))) {
    var name = (token[2] || "").toLowerCase();
    if (!name || VOID_TAGS.test(name) || /\/>$/.test(token[0])) continue;
    if (!token[1]) {
      if (!open.length && token.index > from) {
        parts.push(html.slice(from, token.index));
        from = token.index;
      }
      open.push(name);
      continue;
    }
    if (open.pop() !== name) return null;
    if (!open.length) {
      parts.push(html.slice(from, tokens.lastIndex));
      from = tokens.lastIndex;
    }
  }
  if (open.length) return null;
  if (from < html.length) parts.push(html.slice(from));
  return parts;
}

function hasContent(part) {
  return Boolean(visibleText(part)) || /<img\b/i.test(part);
}

function addedSinceStyling(part) {
  return hasContent(part) && !hasOwnLook(part);
}

// Runs of added blocks are styled together, with any unstyled blank lines between them, so a pasted section keeps
// its shape; blocks that already have the look are left byte for byte.
function withAddedPartsStyled(parts) {
  var result = [];
  var run = [];
  var flush = function () {
    if (run.some(addedSinceStyling)) result.push(safeRestyle(run.join("")));
    else result = result.concat(run);
    run = [];
  };
  parts.forEach(function (part) {
    if (hasOwnLook(part)) {
      flush();
      result.push(part);
    } else {
      run.push(part);
    }
  });
  flush();
  return result;
}

// Outlook puts a line typed at the very top of a styled draft inside the band's first cell, where it shows at 6px.
// Returns that text as lines of its own, [] for a clean band, or null if lifting it out would lose a word.
function linesTypedIntoBand(band) {
  if (!hasContent(band)) return [];
  var blocks = band.match(new RegExp(INNERMOST_BLOCK.source, "gi")) || [];
  var lines = blocks.filter(hasContent).map(function (block) {
    var inner = block.replace(/^<[^>]+>/, "").replace(/<\/[a-z]+>$/i, "");
    return "<div>" + (KEEPS_ITS_MARKUP.test(inner) ? inner : escapeHtml(visibleText(inner))) + "</div>";
  });
  if (!lines.length) lines = ["<div>" + escapeHtml(visibleText(band)) + "</div>"];
  return keepsWords(band, lines.join("")) ? lines : null;
}

// Only what comes before the signature is this email's own: Outlook's rewrite renames the reply and Teams markers, so
// what follows the signature may be the quoted thread and is never touched. Text above the band, or typed into it,
// moves just below it. Without a signature the email is left as it is. Null when the draft has no frame left at all.
function withAdditionsStyled(ownPart) {
  if (!ANY_BAR_TABLE.test(ownPart) && !ANY_SIGNATURE_TABLE.test(ownPart)) return null;
  var bodyTag = /<body[^>]*>/i.exec(ownPart);
  var from = bodyTag ? bodyTag.index + bodyTag[0].length : 0;
  var parts = topLevelParts(ownPart.slice(from));
  if (!parts) return ownPart;
  var bar = -1;
  var signature = -1;
  parts.forEach(function (part, index) {
    if (bar === -1 && part.search(ANY_BAR_TABLE) === 0) bar = index;
    if (signature === -1 && part.search(ANY_SIGNATURE_TABLE) === 0) signature = index;
  });
  if (signature === -1 || signature < bar) return ownPart;
  var band = bar === -1 ? [] : [parts[bar]];
  var lifted = bar === -1 ? [] : linesTypedIntoBand(parts[bar]);
  if (lifted === null) lifted = [];
  else if (lifted.length) band = [brandBarHtml()];
  var above = bar === -1 ? [] : parts.slice(0, bar);
  var framed = above
    .filter(function (part) {
      return !hasContent(part);
    })
    .concat(band, withAddedPartsStyled(lifted.concat(above.filter(hasContent), parts.slice(bar + 1, signature))), parts.slice(signature));
  return ownPart.slice(0, from) + framed.join("");
}

function signatureStart(ownPart) {
  var match = SIGNATURE_TABLE.exec(ownPart);
  return match ? match.index : -1;
}

function textStart(ownPart) {
  var bar = BAR_TABLE.exec(ownPart);
  if (bar) {
    var barClose = ownPart.indexOf("</table>", bar.index);
    if (barClose !== -1) return barClose + "</table>".length;
  }
  var bodyTag = /<body[^>]*>/i.exec(ownPart);
  return bodyTag ? bodyTag.index + bodyTag[0].length : 0;
}

// Returns the restyled email and whether this send added the frame (false when it only re-tidied the text).
// draft.styledBefore: Outlook's note on this draft says the add-in styled it on an earlier pass.
function styleEmail(html, draft) {
  var quoteStart = findQuoteStart(html);
  var ownPart = html.slice(0, quoteStart);
  var quoted = html.slice(quoteStart);
  var signatureAt = signatureStart(ownPart);
  if (signatureAt === -1 && draft && draft.styledBefore) {
    var tidied = withAdditionsStyled(ownPart);
    if (tidied !== null) return { html: tidied + quoted, framed: false };
  }
  if (signatureAt === -1 && styledEarlierInThisDraft(ownPart)) return { html: html, framed: false };
  if (signatureAt !== -1) {
    var from = Math.min(textStart(ownPart), signatureAt);
    return {
      html: ownPart.slice(0, from) + safeRestyle(ownPart.slice(from, signatureAt)) + ownPart.slice(signatureAt) + quoted,
      framed: false,
    };
  }
  var details = signatureDetails();
  var signature = featureSettings().shortReplies && signedEarlierInThread(quoted, details.name) ? shortSignatureHtml(details) : signatureHtml(details);
  return {
    html: insertAfterBodyTag(safeRestyle(ownPart), brandBarHtml()) + signature + quoted,
    framed: true,
  };
}

function applyVersatexStyle(html) {
  return styleEmail(html).html;
}

// ---- Outlook integration ----

function callAsync(invoke) {
  return new Promise(function (resolve, reject) {
    invoke(function (result) {
      if (result && result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(result ? result.error : undefined);
    });
  });
}

// ---- Notes Outlook keeps on this one draft (Mailbox 1.11) ----

var STYLED_NOTE = "vstxStyled";
var OPENING_SUBJECT_NOTE = "vstxOpeningSubject";

function draftNotes() {
  try {
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.11")) return null;
    return Office.context.mailbox.item.sessionData || null;
  } catch (error) {
    // Notes only refine a later pass; without them the add-in recognises its frame from the markup alone.
    return null;
  }
}

// Resolves to the note's text, or "" when there is none or Outlook can't read it.
function readDraftNote(name) {
  var notes = draftNotes();
  if (!notes) return Promise.resolve("");
  return callAsync(function (done) {
    notes.getAsync(name, done);
  }).then(
    function (value) {
      return typeof value === "string" ? value : "";
    },
    function () {
      return "";
    }
  );
}

// Never fails: a note Outlook won't keep just means the next pass falls back to the markup alone.
function writeDraftNote(name, value) {
  var notes = draftNotes();
  if (!notes) return Promise.resolve();
  return callAsync(function (done) {
    notes.setAsync(name, value, done);
  }).catch(function () {
    return undefined;
  });
}

// Resolves to "styled" (frame added), "refreshed" (text re-tidied), "unchanged" or "plain-text".
function restyleCurrentItem() {
  var body = Office.context.mailbox.item.body;
  return callAsync(function (done) {
    body.getTypeAsync(done);
  }).then(function (type) {
    if (type === Office.CoercionType.Text) return "plain-text";
    return readDraftNote(STYLED_NOTE).then(function (styledNote) {
      return callAsync(function (done) {
        body.getAsync(Office.CoercionType.Html, done);
      }).then(function (html) {
        var result = styleEmail(html, { styledBefore: Boolean(styledNote) });
        if (result.html === html) return "unchanged";
        return callAsync(function (done) {
          body.setAsync(result.html, { coercionType: Office.CoercionType.Html }, done);
        })
          .then(function () {
            return writeDraftNote(STYLED_NOTE, "1");
          })
          .then(function () {
            return result.framed ? "styled" : "refreshed";
          });
      });
    });
  });
}

function describeError(error) {
  if (error && error.message) return error.message;
  return error ? String(error) : "Outlook didn't give a reason.";
}

function canOfferSendAnyway() {
  return Office.context.requirements.isSetSupported("Mailbox", "1.14");
}

var ATTACHMENT_WORDS = /\b(?:attach(?:ed|es|ing|ment|ments)?|enclosed)(?![\wÀ-ɏ])/i;
var SHARED_FILE_LINK = /href=["']?https?:\/\/[^"'\s>]*(?:sharepoint\.com|onedrive\.live\.com|1drv\.ms)/i;
var PLAIN_TEXT_QUOTE = /^\s*(?:-{2,}\s*Original Message|_{10,}|From:\s)/im;

// Only your own words count: the quoted thread is full of other people's "see attached".
function mentionsMissingAttachment(item) {
  var body = item.body;
  return callAsync(function (done) {
    body.getTypeAsync(done);
  })
    .then(function (type) {
      var coercion = type === Office.CoercionType.Text ? Office.CoercionType.Text : Office.CoercionType.Html;
      return callAsync(function (done) {
        body.getAsync(coercion, done);
      }).then(function (content) {
        if (coercion === Office.CoercionType.Text) {
          var quoteAt = content.search(PLAIN_TEXT_QUOTE);
          return { words: quoteAt === -1 ? content : content.slice(0, quoteAt), sharedLink: false };
        }
        var own = content.slice(0, findQuoteStart(content));
        return { words: visibleText(own), sharedLink: SHARED_FILE_LINK.test(own) };
      });
    })
    .then(function (own) {
      if (!ATTACHMENT_WORDS.test(own.words) || own.sharedLink || !Office.context.requirements.isSetSupported("Mailbox", "1.8")) return false;
      return callAsync(function (done) {
        item.getAttachmentsAsync(done);
      }).then(function (attachments) {
        return !attachments.some(function (attachment) {
          return !attachment.isInline;
        });
      });
    });
}

function withoutControlTags(subject) {
  return subject.replace(CONTROL_TAGS, " ").replace(/\s{2,}/g, " ").trim();
}

function hasControlTag(text) {
  return REVIEW_TAG.test(text) || PLAIN_TAG.test(text);
}

// A reply or forward can arrive with someone else's [review] or [plain] in the subject. Only tags you add steer the
// add-in and are removed; once the subject as opened has been edited, every tag counts again.
function ownSubjectTags(subject, openingSubject) {
  var at = openingSubject ? subject.indexOf(openingSubject) : -1;
  if (at === -1) return { typed: subject, clean: withoutControlTags(subject) };
  var before = subject.slice(0, at);
  var after = subject.slice(at + openingSubject.length);
  var typed = before + " " + after;
  if (!hasControlTag(typed)) return { typed: typed, clean: subject };
  return {
    typed: typed,
    clean: [withoutControlTags(before), openingSubject, withoutControlTags(after)].filter(Boolean).join(" "),
  };
}

// The [review] and [plain] subject tags steer the add-in and are removed before the email goes out.
// The attachment check runs after styling, so "Send anyway" still sends the styled email.
function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var subject = "";
  var tags = { typed: "", clean: "" };
  // Outlook must hear back exactly once, even if reporting the outcome itself throws.
  var finished = false;
  var complete = function (outcome) {
    if (finished) return;
    finished = true;
    event.completed(outcome);
  };
  callAsync(function (done) {
    item.subject.getAsync(done);
  })
    .then(function (value) {
      subject = value || "";
      return readDraftNote(OPENING_SUBJECT_NOTE);
    })
    .then(function (openingSubject) {
      tags = ownSubjectTags(subject, openingSubject);
      return PLAIN_TAG.test(tags.typed) ? "skipped" : restyleCurrentItem();
    })
    .then(function (status) {
      if (tags.clean === subject) return status;
      return callAsync(function (done) {
        item.subject.setAsync(tags.clean, done);
      }).then(function () {
        return status;
      });
    })
    .then(function (status) {
      var checkAttachment = featureSettings().attachmentCheck && canOfferSendAnyway();
      // A reminder isn't worth blocking Send over: if the lookup fails, the email goes out.
      var lookup = checkAttachment
        ? mentionsMissingAttachment(item).catch(function () {
            return false;
          })
        : Promise.resolve(false);
      return lookup.then(function (missing) {
        return { status: status, missingAttachment: missing };
      });
    })
    .then(function (outcome) {
      if (outcome.missingAttachment) {
        complete({
          allowEvent: false,
          errorMessage: "Your email mentions an attachment, but nothing is attached. Select Don't send to add it, or send it as it is.",
          sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
        });
        return;
      }
      var wantsReview = outcome.status === "styled" && REVIEW_TAG.test(tags.typed);
      if (!wantsReview) {
        complete({ allowEvent: true });
        return;
      }
      complete({
        allowEvent: false,
        errorMessage: "Versatex style applied. Check how the email looks, then select Send again.",
      });
    })
    .catch(function (error) {
      var offersSendAnyway = false;
      try {
        offersSendAnyway = canOfferSendAnyway();
      } catch (requirementError) {
        // Without a way to ask, the email goes out as it is rather than being held.
        offersSendAnyway = false;
      }
      if (!offersSendAnyway) {
        complete({ allowEvent: true });
        return;
      }
      complete({
        allowEvent: false,
        errorMessage: "Couldn't apply the Versatex style to this email: " + describeError(error),
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
      });
    });
}

var SETTINGS_BUTTON = { message: "vstxSettingsButton", appointment: "vstxMeetingSettingsButton" };

// Noting a subject that arrives with a tag lets Send tell someone else's tags from yours. Never fails.
function noteOpeningSubject() {
  return callAsync(function (done) {
    Office.context.mailbox.item.subject.getAsync(done);
  })
    .then(function (subject) {
      return subject && hasControlTag(subject) ? writeDraftNote(OPENING_SUBJECT_NOTE, subject) : undefined;
    })
    .catch(function () {
      // Without the note every tag in the subject counts at Send.
      return undefined;
    });
}

// First use: until someone saves their details, each new email offers a one-click way to the settings pane.
function onNewComposeHandler(event) {
  var finished = false;
  var done = function () {
    if (finished) return;
    finished = true;
    event.completed();
  };
  var offerSetup = function () {
    try {
      if (isSignatureSetUp() || !Office.context.requirements.isSetSupported("Mailbox", "1.10")) {
        done();
        return;
      }
      var item = Office.context.mailbox.item;
      var isInvite = item.itemType === Office.MailboxEnums.ItemType.Appointment;
      showSetupPrompt(item, isInvite ? SETTINGS_BUTTON.appointment : SETTINGS_BUTTON.message, done);
    } catch (error) {
      // A prompt Outlook rejects, or settings it can't read, are skipped; the email itself is unaffected.
      done();
    }
  };
  if (draftNotes()) noteOpeningSubject().then(offerSetup);
  else offerSetup();
}

function showSetupPrompt(item, commandId, done) {
  item.notificationMessages.replaceAsync(
    "vstx-setup",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      message: "Finish your Versatex signature: add your title and phone once, and every email carries it.",
      icon: "Icon.16x16",
      actions: [
        {
          actionText: "Open settings",
          actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
          commandId: commandId,
          contextData: "{}",
        },
      ],
    },
    done
  );
}

function notify(message) {
  return callAsync(function (done) {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "vstx",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message.slice(0, 150),
        icon: "Icon.16x16",
        persistent: false,
      },
      done
    );
  });
}

var PREVIEW_MESSAGES = {
  styled: "Versatex style applied.",
  refreshed: "Versatex style refreshed for your latest edits.",
  unchanged: "This email already has the Versatex style.",
  "plain-text": "This is a plain-text email, so the Versatex style isn't applied.",
};

function previewStyle(event) {
  restyleCurrentItem()
    .then(function (status) {
      return notify(PREVIEW_MESSAGES[status]);
    })
    .catch(function (error) {
      return notify("Couldn't restyle: " + describeError(error)).catch(function () {
        return undefined;
      });
    })
    .then(function () {
      event.completed();
    });
}

// The settings pane loads this file only to preview the signature, and declares VSTX_SETTINGS_PANE first.
if (typeof VSTX_SETTINGS_PANE === "undefined") {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("previewStyle", previewStyle);
  Office.actions.associate("onNewComposeHandler", onNewComposeHandler);
}
