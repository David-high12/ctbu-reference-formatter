(() => {
  "use strict";

  const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const xmlSpace = "http://www.w3.org/XML/1998/namespace";
  const contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
  const wordType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
  const maxFileBytes = 30 * 1024 * 1024;
  const manualLabel = /^[ \t\u00a0]*([\[［]\d{1,6}[\]］])[ \t\u00a0]*/u;
  const referenceHeading = /^参\s*考\s*文\s*献\s*[:：]?$/u;
  const followingHeading = /^(?:致\s*谢|后\s*记|攻读.*期间.*(?:成果|论文))\s*[:：]?$|^附\s*录(?:$|[\s:：A-Za-z0-9一二三四五六七八九十(（])/u;
  const protectedAnchors = new Set(["bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd"]);

  const nodes = (element) => Array.from(element?.children ?? []);
  const child = (element, kind) => nodes(element).find((node) => node.namespaceURI === w && node.localName === kind);
  const attr = (element, name = "val") => element?.getAttributeNS(w, name);
  const descendant = (element, kind) => Array.from(element?.getElementsByTagNameNS(w, kind) ?? []);
  const paragraphText = (paragraph) => descendant(paragraph, "t").map((node) => node.textContent).join("");

  function parseXml(text) {
    if (/<!DOCTYPE\b/iu.test(text)) throw new Error("文档含有暂不支持的 XML 声明。");
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("Word 文件结构损坏，无法读取内容。");
    return doc;
  }

  async function readPart(zip, name) {
    const part = zip.file(name);
    if (!part) return null;
    const text = await part.async("string");
    return parseXml(text);
  }

  function loadStyles(document) {
    return new Map(descendant(document, "style").map((style) => [attr(style, "styleId"), style]));
  }

  function paragraphStyles(paragraph, styles) {
    const chain = [];
    let id = attr(child(child(paragraph, "pPr"), "pStyle"));
    while (id && styles.has(id) && !chain.includes(styles.get(id))) {
      const style = styles.get(id);
      chain.push(style);
      id = attr(child(style, "basedOn"));
    }
    return chain;
  }

  function isContentsEntry(paragraph, styles) {
    return paragraphStyles(paragraph, styles).some((style) => /^toc\b|^toc[0-9]|^目录/iu.test(attr(child(style, "name")) ?? ""))
      || descendant(paragraph, "instrText").some((field) => /\bTOC\b|\bPAGEREF\s+_Toc/iu.test(field.textContent))
      || descendant(paragraph, "hyperlink").some((link) => /^_Toc/iu.test(attr(link, "anchor") ?? ""));
  }

  function isOutlineHeading(paragraph, styles) {
    const properties = [paragraph, ...paragraphStyles(paragraph, styles)];
    return properties.some((element) => /^[0-8]$/u.test(attr(child(child(element, "pPr"), "outlineLvl")) ?? ""));
  }

  function loadNumbering(document) {
    const abstract = new Map(descendant(document, "abstractNum").map((num) => [attr(num, "abstractNumId"), num]));
    const definitions = new Map();
    for (const num of descendant(document, "num")) {
      const scheme = abstract.get(attr(child(num, "abstractNumId")));
      const level = nodes(scheme).find((node) => node.localName === "lvl" && attr(node, "ilvl") === "0");
      const override = nodes(num).find((node) => node.localName === "lvlOverride" && attr(node, "ilvl") === "0");
      const rule = child(override, "lvl") ?? level;
      if (attr(child(rule, "numFmt")) !== "decimal") continue;
      const pattern = attr(child(rule, "lvlText")) ?? "";
      const start = Number(attr(child(override, "startOverride")) ?? attr(child(rule, "start")) ?? 1);
      if (/^(?:\[%1\]|［%1］|%1[.)、]|%1)$/u.test(pattern) && Number.isSafeInteger(start) && start >= 0) {
        definitions.set(attr(num, "numId"), { pattern, next: start });
      }
    }
    return definitions;
  }

  function mapNumbering(doc, definitions, styles) {
    const labels = new Map();
    for (const paragraph of descendant(doc, "p")) {
      const numbering = [paragraph, ...paragraphStyles(paragraph, styles)]
        .map((element) => child(child(element, "pPr"), "numPr")).filter(Boolean);
      const id = numbering.map((entry) => attr(child(entry, "numId"))).find((value) => value != null);
      const level = numbering.map((entry) => attr(child(entry, "ilvl"))).find((value) => value != null) ?? "0";
      const definition = definitions.get(id);
      if (level === "0" && definition) {
        labels.set(paragraph, definition.pattern.replace("%1", String(definition.next++)));
      }
    }
    return labels;
  }

  function readParagraph(paragraph) {
    const chunks = [];
    const anchorOffsets = [];
    let text = "";
    let plain = true;
    for (const element of nodes(paragraph)) {
      if (element.namespaceURI !== w) {
        plain = false;
        continue;
      }
      if (element.localName === "pPr" || element.localName === "proofErr") continue;
      if (protectedAnchors.has(element.localName)) {
        anchorOffsets.push(text.length);
        continue;
      }
      if (element.localName !== "r") {
        plain = false;
        continue;
      }
      for (const part of nodes(element)) {
        if (part.namespaceURI === w && part.localName === "rPr") {
          const hidden = child(part, "vanish");
          if (hidden && !["0", "false", "off"].includes(attr(hidden))) plain = false;
          continue;
        }
        const value = part.namespaceURI === w && (part.localName === "t" ? part.textContent : part.localName === "tab" ? "\t" : null);
        if (typeof value !== "string") {
          plain = false;
        } else if (value.length) {
          chunks.push({ node: part, start: text.length, end: text.length + value.length });
          text += value;
        }
      }
    }
    return { plain, text, chunks, anchorOffsets };
  }

  function bibliographyEntries(doc, styles, labels) {
    const body = child(doc.documentElement, "body");
    if (!body) throw new Error("找不到 Word 正文。");
    const blocks = nodes(body);
    const headings = blocks.filter((p) => p.namespaceURI === w && p.localName === "p"
      && referenceHeading.test(paragraphText(p).trim()) && !isContentsEntry(p, styles));
    if (headings.length > 1) throw new Error("发现多个“参考文献”标题。请确认文末只有一个独立的参考文献区。");
    if (!headings.length) throw new Error("没有找到独立的“参考文献”标题。请检查文末标题是否单独成段。");
    const entries = [];
    let stoppedAtUnnumbered = false;
    for (let index = blocks.indexOf(headings[0]) + 1; index < blocks.length; index++) {
      const paragraph = blocks[index];
      if (paragraph.namespaceURI !== w || paragraph.localName !== "p") break;
      const visible = paragraphText(paragraph).trim();
      const manual = manualLabel.exec(visible);
      const unnumbered = manual ? visible.slice(manual[0].length) : visible;
      if (followingHeading.test(unnumbered) || isOutlineHeading(paragraph, styles)) break;
      const content = readParagraph(paragraph);
      const automatic = labels.get(paragraph);
      if (automatic || manual) {
        const prefix = !automatic && manualLabel.exec(content.text);
        entries.push({
          paragraphIndex: index,
          number: automatic ?? manual[1],
          content,
          text: prefix ? content.text.slice(prefix[0].length) : content.text,
          start: prefix ? prefix[0].length : 0,
          manualAndAutomatic: Boolean(automatic && manual)
        });
      } else if (visible || !content.plain) {
        stoppedAtUnnumbered = true;
        break;
      }
    }
    if (!entries.length) throw new Error("参考文献标题后没有识别到连续的编号条目。请检查是否使用了 [1] 形式的手打编号或 Word 数字列表。");
    return { entries, stoppedAtUnnumbered };
  }

  function setText(node, value) {
    node.textContent = value;
    if (/^\s|\s$/u.test(value)) node.setAttributeNS(xmlSpace, "xml:space", "preserve");
  }

  function applyEdits(entry, edits) {
    for (const edit of edits.toSorted((a, b) => b.start - a.start || b.end - a.end)) {
      const start = entry.start + edit.start;
      const end = entry.start + edit.end;
      if (start === end) {
        const chunk = entry.content.chunks.find((chunk) => chunk.start < start && start <= chunk.end && chunk.node.localName === "t")
          ?? entry.content.chunks.find((chunk) => chunk.start === start && chunk.node.localName === "t");
        const text = chunk.node.textContent;
        const offset = start - chunk.start;
        setText(chunk.node, text.slice(0, offset) + edit.replacement + text.slice(offset));
        continue;
      }
      let replacement = edit.replacement;
      for (const chunk of entry.content.chunks) {
        if (chunk.end <= start) continue;
        if (chunk.start >= end) break;
        const text = chunk.node.textContent;
        const left = Math.max(0, start - chunk.start);
        const right = Math.min(chunk.end, end) - chunk.start;
        setText(chunk.node, text.slice(0, left) + replacement + text.slice(right));
        replacement = "";
      }
    }
  }

  async function formatWord(bytes) {
    if (bytes.byteLength > maxFileBytes) throw new Error("Word 文件超过 30 MB，请选择更小的 .docx。");
    let zip;
    try {
      zip = await window.JSZip.loadAsync(bytes);
    } catch {
      throw new Error("无法打开这个 Word 文件。请在 Word 或 WPS 里另存为普通 .docx 后重试。");
    }
    if (!zip.file("word/document.xml") || !zip.file("[Content_Types].xml")) {
      throw new Error("无法打开这个文件。请选择标准的 .docx 文件。");
    }
    const xmlParts = new Set(["word/document.xml", "word/styles.xml", "word/numbering.xml", "[Content_Types].xml"]);
    for (const part of Object.values(zip.files)) {
      if (part.name.startsWith("_xmlsignatures/")) throw new Error("此文档有数字签名。请由签名人提供可编辑副本。");
      if (xmlParts.has(part.name) && part._data?.uncompressedSize > 12 * 1024 * 1024) {
        throw new Error("Word 文件的文本结构过大，浏览器无法整理。");
      }
    }
    const types = await readPart(zip, "[Content_Types].xml");
    const valid = Array.from(types.getElementsByTagNameNS(contentTypes, "Override")).some((part) => (
      part.getAttribute("PartName") === "/word/document.xml" && part.getAttribute("ContentType") === wordType
    ));
    if (!valid) throw new Error("请选择普通的 .docx。含宏的 Word 文件和 Word 模板暂不支持。");
    const [doc, styleDoc, numberingDoc] = await Promise.all([
      readPart(zip, "word/document.xml"), readPart(zip, "word/styles.xml"), readPart(zip, "word/numbering.xml")
    ]);
    const styles = loadStyles(styleDoc);
    const labels = mapNumbering(doc, loadNumbering(numberingDoc), styles);
    const { entries, stoppedAtUnnumbered } = bibliographyEntries(doc, styles, labels);
    let changed = 0;
    const reviews = [];
    for (const entry of entries) {
      const citation = !entry.content.plain || entry.manualAndAutomatic
        ? { reason: "含复杂 Word 格式或重复编号，已保留原文" }
        : window.formatCitation(entry.text);
      const crossesAnchor = citation.edits?.some(({ start, end }) => {
        const left = entry.start + start;
        const right = entry.start + end;
        return entry.content.anchorOffsets.some((offset) => (
          (left < offset && offset < right)
          || (left === right && left === offset && 0 < offset && offset < entry.content.text.length)
        ));
      });
      const reason = citation.reason ?? (crossesAnchor ? "修改会跨过书签或批注的边界，已保留原文" : null);
      if (reason) {
        reviews.push({ number: entry.number, paragraphIndex: entry.paragraphIndex, reason });
      } else if (citation.edits.length) {
        applyEdits(entry, citation.edits);
        changed++;
      }
    }
    let output = bytes;
    if (changed) {
      zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
      output = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    }
    const references = entries.map(({ number, paragraphIndex }) => ({ number, paragraphIndex }));
    return { bytes: output, references, changed, reviews, stoppedAtUnnumbered };
  }

  window.formatWord = formatWord;
})();
