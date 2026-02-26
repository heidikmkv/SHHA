function onOpen() {
  DocumentApp.getUi()
    .createMenu('GRIT')
    .addItem('Open Markdown Sidebar', 'showGritSidebar')
    .addToUi();
}

function showGritSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GRIT Markdown');
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Supported GRIT Markdown (inserter):
 *  == Section Name        => Title
 *  # Article Title        => Heading 1
 *  ## Section Header      => Heading 2
 *  ### Subsection         => Heading 3
 *  #### Minor header      => Heading 4
 *  blank line             => paragraph break
 *  - item or * item       => bulleted list
 *  1. item                => numbered list
 *  Inline: **bold**, _italic_
 */
function insertGritMarkdown(markdownText) {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  if (!cursor) {
    throw new Error('Click in the document where you want to insert content, then try again.');
  }

  const body = doc.getBody();

  // Find stable paragraph anchor
  const insertionPoint = cursor.getElement();
  let parent = insertionPoint;
  while (parent && parent.getType() !== DocumentApp.ElementType.PARAGRAPH && parent.getParent()) {
    parent = parent.getParent();
  }

  // Fallback: append to end
  if (!parent || parent.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    insertAtEnd_(body, markdownText);
    return;
  }

  const anchorParagraph = parent.asParagraph();
  const anchorIndex = body.getChildIndex(anchorParagraph);
  let insertIndex = anchorIndex + 1;

  const lines = normalizeNewlines_(markdownText).split('\n');

  for (const raw of lines) {
    const line = String(raw || '').replace(/\s+$/g, '');

    // Blank line
    if (line.trim() === '') {
      body.insertParagraph(insertIndex++, '');
      continue;
    }

    // == Section Title → Docs Title style
    if (/^==\s+/.test(line)) {
      const p = body.insertParagraph(insertIndex++, line.replace(/^==\s+/, ''));
      p.setHeading(DocumentApp.ParagraphHeading.TITLE);
      applyInlineFormatting_(p);
      continue;
    }

    // Headings in descending specificity
    if (line.startsWith('#### ')) {
      const p = body.insertParagraph(insertIndex++, line.substring(5));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      applyInlineFormatting_(p);
      continue;
    }

    if (line.startsWith('### ')) {
      const p = body.insertParagraph(insertIndex++, line.substring(4));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      applyInlineFormatting_(p);
      continue;
    }

    if (line.startsWith('## ')) {
      const p = body.insertParagraph(insertIndex++, line.substring(3));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      applyInlineFormatting_(p);
      continue;
    }

    if (line.startsWith('# ')) {
      const p = body.insertParagraph(insertIndex++, line.substring(2));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      applyInlineFormatting_(p);
      continue;
    }

    // Bulleted list
    if (/^[-*]\s+/.test(line)) {
      const li = body.insertListItem(insertIndex++, line.replace(/^[-*]\s+/, ''));
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
      applyInlineFormatting_(li);
      continue;
    }

    // Numbered list
    if (/^\d+\.\s+/.test(line)) {
      const li = body.insertListItem(insertIndex++, line.replace(/^\d+\.\s+/, ''));
      li.setGlyphType(DocumentApp.GlyphType.NUMBER);
      applyInlineFormatting_(li);
      continue;
    }

    // Normal paragraph
    const p = body.insertParagraph(insertIndex++, line);
    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    applyInlineFormatting_(p);
  }
}

/** Append fallback */
function insertAtEnd_(body, markdownText) {
  const lines = normalizeNewlines_(markdownText).split('\n');

  for (const raw of lines) {
    const line = String(raw || '').replace(/\s+$/g, '');

    if (line.trim() === '') {
      body.appendParagraph('');
      continue;
    }

    if (/^==\s+/.test(line)) {
      const p = body.appendParagraph(line.replace(/^==\s+/, ''));
      p.setHeading(DocumentApp.ParagraphHeading.TITLE);
      applyInlineFormatting_(p);
      continue;
    }

    if (line.startsWith('#### ')) {
      const p = body.appendParagraph(line.substring(5));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      applyInlineFormatting_(p);
      continue;
    }

    if (line.startsWith('### ')) {
      const p = body.appendParagraph(line.substring(4));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      applyInlineFormatting_(p);
      continue;
    }

    if (line.startsWith('## ')) {
      const p = body.appendParagraph(line.substring(3));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      applyInlineFormatting_(p);
      continue;
    }

    if (line.startsWith('# ')) {
      const p = body.appendParagraph(line.substring(2));
      p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      applyInlineFormatting_(p);
      continue;
    }

    if (/^[-*]\s+/.test(line)) {
      const li = body.appendListItem(line.replace(/^[-*]\s+/, ''));
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
      applyInlineFormatting_(li);
      continue;
    }

    if (/^\d+\.\s+/.test(line)) {
      const li = body.appendListItem(line.replace(/^\d+\.\s+/, ''));
      li.setGlyphType(DocumentApp.GlyphType.NUMBER);
      applyInlineFormatting_(li);
      continue;
    }

    const p = body.appendParagraph(line);
    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    applyInlineFormatting_(p);
  }
}

/**
 * Export the doc (or selection if present) into GRIT Markdown.
 * Returns a string. Sidebar uses it to trigger a download.
 */
function exportGritMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const sel = doc.getSelection();

  // If there is a selection, export only selected paragraphs; otherwise export whole body
  if (sel) {
    const els = sel.getRangeElements();
    const paras = extractParagraphsFromRangeElements_(els);
    return paragraphsToGritMarkdown_(paras);
  }

  const body = doc.getBody();
  const paras = extractParagraphsFromBody_(body);
  return paragraphsToGritMarkdown_(paras);
}

/** Gather all Paragraph/ListItem elements from the document body, including nested containers (e.g., tables). */
function extractParagraphsFromBody_(body) {
  const out = [];
  collectParagraphLikeElements_(body, out);
  return out;
}

function collectParagraphLikeElements_(container, out) {
  if (!container || typeof container.getNumChildren !== 'function') return;

  const n = container.getNumChildren();
  for (let i = 0; i < n; i++) {
    const child = container.getChild(i);
    const t = child.getType();

    if (t === DocumentApp.ElementType.PARAGRAPH) {
      out.push(child.asParagraph());
      continue;
    }

    if (t === DocumentApp.ElementType.LIST_ITEM) {
      out.push(child.asListItem());
      continue;
    }

    collectParagraphLikeElements_(child, out);
  }
}

/** Extract paragraphs/list items touched by selection range elements. */
function extractParagraphsFromRangeElements_(rangeElements) {
  const seen = new Set();
  const out = [];

  for (const re of rangeElements) {
    let el = re.getElement();

    // Climb to paragraph/list item if needed
    while (el && el.getParent && el.getType() !== DocumentApp.ElementType.PARAGRAPH
           && el.getType() !== DocumentApp.ElementType.LIST_ITEM) {
      el = el.getParent();
    }
    if (!el) continue;

    if (seen.has(el)) continue;
    seen.add(el);

    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) out.push(el.asParagraph());
    if (el.getType() === DocumentApp.ElementType.LIST_ITEM) out.push(el.asListItem());
  }

  return out;
}

/** Convert paragraph/listitem array into GRIT Markdown text */
function paragraphsToGritMarkdown_(paras) {
  const lines = [];
  let inBullet = false;
  let inNumber = false;
  let numberCounter = 1;

  const flushLists = () => {
    if (inBullet || inNumber) lines.push(''); // blank line after list block
    inBullet = false;
    inNumber = false;
    numberCounter = 1;
  };

  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    const type = p.getType();

    // Handle list items
    if (type === DocumentApp.ElementType.LIST_ITEM) {
      const li = p.asListItem();
      const glyph = li.getGlyphType();
      const text = elementTextToMarkdown_(li);

      if (glyph === DocumentApp.GlyphType.BULLET) {
        if (inNumber) flushLists();
        inBullet = true;
        lines.push(`- ${text}`);
        continue;
      }

      if (glyph === DocumentApp.GlyphType.NUMBER) {
        if (inBullet) flushLists();
        inNumber = true;
        lines.push(`${numberCounter}. ${text}`);
        numberCounter++;
        continue;
      }

      // Unknown list glyph => treat like bullet
      if (inNumber) flushLists();
      inBullet = true;
      lines.push(`- ${text}`);
      continue;
    }

    // Non-list paragraph => close any open list block
    if (inBullet || inNumber) flushLists();

    const para = p.asParagraph();
    const heading = para.getHeading();
    const text = elementTextToMarkdown_(para);

    // Empty paragraph => blank line
    if (!text || text.trim() === '') {
      lines.push('');
      continue;
    }

    if (heading === DocumentApp.ParagraphHeading.TITLE) {
      lines.push(`== ${text}`);
      lines.push('');
      continue;
    }
    if (heading === DocumentApp.ParagraphHeading.HEADING1) {
      lines.push(`# ${text}`);
      lines.push('');
      continue;
    }
    if (heading === DocumentApp.ParagraphHeading.HEADING2) {
      lines.push(`## ${text}`);
      lines.push('');
      continue;
    }
    if (heading === DocumentApp.ParagraphHeading.HEADING3) {
      lines.push(`### ${text}`);
      lines.push('');
      continue;
    }
    if (heading === DocumentApp.ParagraphHeading.HEADING4) {
      lines.push(`#### ${text}`);
      lines.push('');
      continue;
    }

    // Normal paragraph
    lines.push(text);
    lines.push('');
  }

  // Trim trailing blank lines
  while (lines.length && lines[lines.length - 1] === '') lines.pop();

  return lines.join('\n');
}

/**
 * Convert a Paragraph or ListItem to markdown text, preserving **bold** and _italic_.
 * (Intentionally simple; no links/tables/nesting.)
 */
function elementTextToMarkdown_(paragraphOrListItem) {
  const textEl = paragraphOrListItem.editAsText();
  const full = textEl.getText();
  if (!full) return '';

  const indices = textEl.getTextAttributeIndices();
  indices.push(full.length); // sentinel

  let out = '';

  for (let i = 0; i < indices.length - 1; i++) {
    const start = indices[i];
    const end = indices[i + 1];
    const chunk = full.substring(start, end);
    if (chunk === '') continue;

    const attrs = textEl.getAttributes(start);

    const bold = !!attrs[DocumentApp.Attribute.BOLD];
    const italic = !!attrs[DocumentApp.Attribute.ITALIC];

    // Basic escaping so markdown markers in text are less likely to break formatting
    let safe = chunk
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n');

    // Wrap styles
    if (bold && italic) safe = `**_${safe}_**`;
    else if (bold) safe = `**${safe}**`;
    else if (italic) safe = `_${safe}_`;

    out += safe;
  }

  // Collapse any accidental double spaces at ends
  return out.replace(/\s+$/g, '');
}

/** Minimal inline formatting for inserter: **bold**, _italic_ (non-nested) */
function applyInlineFormatting_(paragraphOrListItem) {
  const textEl = paragraphOrListItem.editAsText();
  const original = textEl.getText();
  if (!original) return;

  const tokens = [];
  let out = '';
  let i = 0;

  while (i < original.length) {
    if (original.startsWith('**', i)) {
      const end = original.indexOf('**', i + 2);
      if (end !== -1) {
        const content = original.substring(i + 2, end);
        const start = out.length;
        out += content;
        tokens.push({ type: 'BOLD', start, end: out.length - 1 });
        i = end + 2;
        continue;
      }
    }

    if (original[i] === '_') {
      const end = original.indexOf('_', i + 1);
      if (end !== -1) {
        const content = original.substring(i + 1, end);
        const start = out.length;
        out += content;
        tokens.push({ type: 'ITALIC', start, end: out.length - 1 });
        i = end + 1;
        continue;
      }
    }

    out += original[i];
    i++;
  }

  if (out === original) return;

  textEl.setText(out);

  for (const t of tokens) {
    if (t.start > t.end) continue;
    if (t.type === 'BOLD') textEl.setBold(t.start, t.end, true);
    if (t.type === 'ITALIC') textEl.setItalic(t.start, t.end, true);
  }
}

function normalizeNewlines_(s) {
  return String(s || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}