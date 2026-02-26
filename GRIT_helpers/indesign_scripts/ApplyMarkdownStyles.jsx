#target "InDesign"

(function () {
  // =========================
  // 1) GRIT STYLE MAP
  // =========================
  var STYLE_MAP = {
    sectionTitle: "Heading 1",           // == Section Title
    h1: "Heading 1",                     // # Heading 1
    h2: "Heading 2",                     // ## Heading 2
    h3: "Heading 2",                     // ### Heading 3
    h4: "Normal Small",                  // #### Heading 4
    bullet: "Bullet List",               // - item / * item
    number: "Number List",               // 1. item
    body: "Normal"                       // normal paragraph
  };

  var INLINE_STYLE_MAP = {
    bold: "Bold",
    italic: "Italic",
    boldItalic: "Bold Italic"
  };

  // =========================
  // 2) SELECT MARKDOWN FILE
  // =========================
  var mdFile = File.openDialog("Select a GRIT Markdown file", function (f) {
    return f instanceof Folder || /\.(md|txt)$/i.test(f.name);
  }, false);
  if (!mdFile) return;

  mdFile.encoding = "UTF-8";
  mdFile.open("r");
  var markdown = mdFile.read();
  mdFile.close();

  // =========================
  // 3) RUN
  // =========================
  if (app.documents.length === 0) {
    alert("Open a document first.");
    return;
  }

  var doc = app.activeDocument;
  var story = getTargetStory();
  if (!story) {
    alert("Select a text frame or insertion point first.");
    return;
  }

  validateStyles(doc, STYLE_MAP, INLINE_STYLE_MAP);

  app.doScript(function () {
    applyMarkdownToStory(doc, story, markdown, STYLE_MAP, INLINE_STYLE_MAP);
  }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Apply Simple Markdown");

  function getTargetStory() {
    if (app.selection.length === 0) return null;
    var s = app.selection[0];

    if (s.hasOwnProperty("parentStory")) return s.parentStory;
    if (s.constructor && s.constructor.name === "TextFrame") return s.parentStory;
    return null;
  }

  // --- Recursive style finder (searches inside groups) ---
  function findParagraphStyle(doc, name) {
    var st = doc.paragraphStyles.itemByName(name);
    if (st.isValid) return st;
    return searchParagraphStyleGroups(doc.paragraphStyleGroups, name);
  }

  function searchParagraphStyleGroups(groups, name) {
    for (var i = 0; i < groups.length; i++) {
      var g = groups[i];
      var st = g.paragraphStyles.itemByName(name);
      if (st.isValid) return st;
      var nested = searchParagraphStyleGroups(g.paragraphStyleGroups, name);
      if (nested) return nested;
    }
    return null;
  }

  function findCharacterStyle(doc, name) {
    var st = doc.characterStyles.itemByName(name);
    if (st.isValid) return st;
    return searchCharacterStyleGroups(doc.characterStyleGroups, name);
  }

  function searchCharacterStyleGroups(groups, name) {
    for (var i = 0; i < groups.length; i++) {
      var g = groups[i];
      var st = g.characterStyles.itemByName(name);
      if (st.isValid) return st;
      var nested = searchCharacterStyleGroups(g.characterStyleGroups, name);
      if (nested) return nested;
    }
    return null;
  }

  function validateStyles(doc, pMap, cMap) {
    var key;
    for (key in pMap) {
      if (pMap.hasOwnProperty(key)) {
        if (!findParagraphStyle(doc, pMap[key]))
          throw Error("Missing paragraph style: " + pMap[key]);
      }
    }
    for (key in cMap) {
      if (cMap.hasOwnProperty(key)) {
        if (!findCharacterStyle(doc, cMap[key]))
          throw Error("Missing character style: " + cMap[key]);
      }
    }
  }

  function applyMarkdownToStory(doc, story, md, pMap, cMap) {
    var lines = String(md || "")
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .split("\n");

    // Strip ALL blank/whitespace-only lines — no empty paragraphs
    var content = [];
    for (var j = 0; j < lines.length; j++) {
      if (/\S/.test(lines[j])) content.push(lines[j]);
    }

    for (var i = 0; i < content.length; i++) {
      var raw = content[i];
      var parsed = parseBlockLine(raw, pMap);
      var inline = parseInlineSimple(parsed.text);

      // insert paragraph at end of story
      var ip = story.insertionPoints[-1];
      ip.contents = inline.clean + "\r";

      var p = story.paragraphs[-1];
      p.appliedParagraphStyle = findParagraphStyle(doc, parsed.pStyleName);

      // apply character styles
      applyInlineSpans(doc, p, inline.spans, cMap);
    }
  }

  function parseBlockLine(line, pMap) {
    var s = String(line || "");
    if (/^\s*$/.test(s)) return { pStyleName: pMap.body, text: "" };

    if (/^==\s+/.test(s)) return { pStyleName: pMap.sectionTitle, text: s.replace(/^==\s+/, "") };
    if (/^####\s+/.test(s)) return { pStyleName: pMap.h4, text: s.replace(/^####\s+/, "") };
    if (/^###\s+/.test(s)) return { pStyleName: pMap.h3, text: s.replace(/^###\s+/, "") };
    if (/^##\s+/.test(s)) return { pStyleName: pMap.h2, text: s.replace(/^##\s+/, "") };
    if (/^#\s+/.test(s)) return { pStyleName: pMap.h1, text: s.replace(/^#\s+/, "") };
    if (/^[-*]\s+/.test(s)) return { pStyleName: pMap.bullet, text: s.replace(/^[-*]\s+/, "") };
    if (/^\d+\.\s+/.test(s)) return { pStyleName: pMap.number, text: s.replace(/^\d+\.\s+/, "") };

    return { pStyleName: pMap.body, text: s };
  }

  // Supports non-nested **bold** and _italic_
  function parseInlineSimple(text) {
    var src = String(text || "");
    var out = "";
    var spans = []; // {start, endExclusive, kind}
    var i = 0;

    while (i < src.length) {
      if (src.substr(i, 2) === "**") {
        var bEnd = src.indexOf("**", i + 2);
        if (bEnd !== -1) {
          var b = src.substring(i + 2, bEnd);
          var bStart = out.length;
          out += b;
          spans.push({ start: bStart, end: out.length, kind: "bold" });
          i = bEnd + 2;
          continue;
        }
      }

      if (src.charAt(i) === "_") {
        var iEnd = src.indexOf("_", i + 1);
        if (iEnd !== -1) {
          var it = src.substring(i + 1, iEnd);
          var itStart = out.length;
          out += it;
          spans.push({ start: itStart, end: out.length, kind: "italic" });
          i = iEnd + 1;
          continue;
        }
      }

      out += src.charAt(i);
      i++;
    }

    return { clean: out, spans: spans };
  }

  function applyInlineSpans(doc, paragraph, spans, cMap) {
    for (var i = 0; i < spans.length; i++) {
      var span = spans[i];
      if (span.end <= span.start) continue;

      var styleName = (span.kind === "bold") ? cMap.bold : cMap.italic;
      var style = findCharacterStyle(doc, styleName);

      // paragraph.characters uses local character indexing
      var first = span.start;
      var last = span.end - 1;
      if (last < first) continue;

      paragraph.characters.itemByRange(first, last).appliedCharacterStyle = style;
    }
  }

})();

