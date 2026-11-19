const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const { pathToFileURL } = require("node:url");
const { test, before, after } = require("node:test");

const modules = path.resolve(path.dirname(process.execPath), "..", "node_modules");
const { chromium } = require(path.join(modules, "playwright"));
const JSZip = require(path.join(modules, "jszip"));
const pagePath = path.join(__dirname, "..", "dist", "index.html");
const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const types = "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>";

let browser;
before(async () => {
  browser = await chromium.launch({ channel: "msedge" });
});
after(async () => {
  await browser?.close();
});

function paragraph(text, properties = "") {
  return "<w:p>" + properties + "<w:r><w:t>" + text + "</w:t></w:r></w:p>";
}

async function documentFile(body) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", types);
  zip.file("word/document.xml", '<w:document xmlns:w="' + w + '"><w:body>' + body + "</w:body></w:document>");
  zip.file("word/styles.xml", '<w:styles xmlns:w="' + w + '"><w:style w:type="paragraph" w:styleId="t1"><w:name w:val="toc 1"/></w:style></w:styles>');
  return zip.generateAsync({ type: "nodebuffer" });
}

async function openTool(t, viewport = { width: 1366, height: 900 }) {
  const page = await browser.newPage({ viewport, acceptDownloads: true });
  t.after(() => page.close());
  const requests = [];
  page.on("request", (request) => {
    if (/^https?:/.test(request.url())) requests.push(request.url());
  });
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(pathToFileURL(pagePath).href);
  t.after(() => {
    assert.deepEqual(requests, []);
    assert.deepEqual(errors, []);
  });
  return page;
}

async function selectFile(page, buffer, name = "论文.docx") {
  await page.locator('input[type="file"]').setInputFiles({
    name,
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer
  });
  await page.getByRole("button", { name: "整理参考文献", exact: true }).click();
}

async function downloadFile(page, saveAs) {
  const download = page.waitForEvent("download");
  await page.getByRole("button", { name: "下载 Word 副本" }).click();
  const item = await download;
  assert.match(item.suggestedFilename(), /_参考文献整理\.docx$/u);
  if (saveAs) await item.saveAs(saveAs);
  return JSZip.loadAsync(await fs.readFile(await item.path()), { checkCRC32: true });
}

async function readParagraphs(page, zip) {
  const xml = await zip.file("word/document.xml").async("string");
  return page.evaluate(({ xml, w }) => {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const body = doc.getElementsByTagNameNS(w, "body")[0];
    const serializer = new XMLSerializer();
    return Array.from(body.children, (p) => {
      const element = (root, name) => Array.from(root?.children ?? []).find((child) => child.namespaceURI === w && child.localName === name);
      const properties = element(p, "pPr");
      const textRun = Array.from(p.children).find((run) => element(run, "t"));
      const runProperties = element(textRun, "rPr");
      const withoutFormatting = p.cloneNode(true);
      for (const kind of ["pPr", "rPr"]) {
        for (const node of Array.from(withoutFormatting.getElementsByTagNameNS(w, kind))) node.remove();
      }
      return {
        text: Array.from(p.getElementsByTagNameNS(w, "t"), (t) => t.textContent).join(""),
        xml: serializer.serializeToString(p),
        contentXml: serializer.serializeToString(withoutFormatting),
        bookmarks: Array.from(p.getElementsByTagNameNS(w, "bookmarkStart"), (b) => b.getAttributeNS(w, "name")),
        listId: properties?.getElementsByTagNameNS(w, "numId")[0]?.getAttributeNS(w, "val") ?? null,
        align: element(properties, "jc")?.getAttributeNS(w, "val"),
        beforeLines: element(properties, "spacing")?.getAttributeNS(w, "beforeLines"),
        font: element(runProperties, "rFonts")?.getAttributeNS(w, "eastAsia"),
        size: element(runProperties, "sz")?.getAttributeNS(w, "val")
      };
    });
  }, { xml, w });
}

test("在整份 Word 中只整理文末可识别的手打编号条目", async (t) => {
  const source = await documentFile(
    paragraph("参考文献", '<w:pPr><w:pStyle w:val="t1"/></w:pPr>') +
    paragraph("正文保留引用 [7]。") +
    paragraph("参考文献") +
    '<w:p><w:bookmarkStart w:id="7" w:name="Ref7"/><w:r><w:t>[7] 张三， 李四。示例题名［j］。期刊名，2025，08（2）：12－18。</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>' +
    paragraph("[8] 李四. 另一个题名[J]. 期刊名, 2025, 8(2).") +
    paragraph("[9] 孙乙。学位论文题名［D］．导师：王五．示例大学，2024。") +
    paragraph("[1] 附录 A 原文摘录") +
    paragraph("[2] 赵甲。这里属于附录［J］。期刊名，2024，2（3）：1－2。")
  );
  const page = await openTool(t);
  await selectFile(page, source);
  const output = await downloadFile(page);
  const input = await JSZip.loadAsync(source);
  const before = await readParagraphs(page, input);
  const after = await readParagraphs(page, output);
  for (const i of [0, 1, 6, 7]) assert.deepEqual(after[i], before[i]);
  assert.equal(after[4].contentXml, before[4].contentXml);
  assert.equal(after[3].text, "[7] 张三, 李四. 示例题名[J]. 期刊名, 2025, 08(2): 12-18.");
  assert.deepEqual(after[3].bookmarks, ["Ref7"]);
  assert.equal(after[5].text, "[9] 孙乙. 学位论文题名[D]. 导师: 王五. 示例大学, 2024.");
  await page.getByText(/缺少页码或文章编号/u).waitFor();
});

test("识别学院样例的 18 条自动编号条目并维持书签和编号定义", {
  skip: !process.env.REFERENCE_SAMPLE
}, async (t) => {
  const source = await fs.readFile(process.env.REFERENCE_SAMPLE);
  const page = await openTool(t);
  await selectFile(page, source, "学院样例.docx");
  await page.getByText("18", { exact: true }).waitFor();
  const report = await page.locator(".metrics dd").allTextContents();
  t.diagnostic("样例识别、已调整、待复核条目数 " + report.join(" / "));
  const qa = process.env.REFERENCE_QA;
  if (qa) {
    await fs.mkdir(qa, { recursive: true });
    await page.screenshot({ path: path.join(qa, "desktop-result.png"), fullPage: true });
  }
  const input = await JSZip.loadAsync(source);
  const output = await downloadFile(page, qa ? path.join(qa, "verification.docx") : undefined);
  const before = await readParagraphs(page, input);
  const after = await readParagraphs(page, output);
  assert.equal(after.filter((p) => p.listId === "1").length, 18);
  assert.deepEqual(after.map((p) => p.listId), before.map((p) => p.listId));
  assert.deepEqual(after.map((p) => p.bookmarks), before.map((p) => p.bookmarks));
  assert.ok(after.some((p, index) => p.text !== before[index].text));
  for (const [name, entry] of Object.entries(input.files)) {
    if (entry.dir || name === "word/document.xml") continue;
    assert.deepEqual(await output.file(name).async("nodebuffer"), await entry.async("nodebuffer"), name);
  }
  assert.equal(after[0].contentXml, before[0].contentXml);
});

test("遇到多个可能的参考文献区时提示检查并阻止导出", async (t) => {
  const source = await documentFile(
    paragraph("参考文献") + paragraph("[1] 不应猜[J]. 期刊, 2025, 8(2): 12-18.") +
    paragraph("参考文献") + paragraph("[1] 另一组[J]. 期刊, 2025, 8(2): 12-18.")
  );
  const page = await openTool(t, { width: 390, height: 844 });
  await selectFile(page, source);
  await page.getByRole("alert").filter({ hasText: "多个" }).waitFor();
  assert.equal(await page.getByRole("button", { name: "下载 Word 副本" }).isVisible(), false);
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), 390);
});

test("保留条目内书签覆盖的文字并提示跨过书签边界的修改", async (t) => {
  const source = await documentFile(
    paragraph("参考文献") +
    '<w:p><w:r><w:t xml:space="preserve">[1] 张三。示例题名 </w:t></w:r>' +
    '<w:bookmarkStart w:id="1" w:name="Type"/><w:r><w:t>［J］</w:t></w:r><w:bookmarkEnd w:id="1"/>' +
    '<w:r><w:t>。刊名，2025，1（2）：3－4。</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>[2] 张三. 题名[J]. 刊名</w:t></w:r>' +
    '<w:bookmarkStart w:id="2" w:name="Separator"/><w:r><w:t>，</w:t></w:r><w:bookmarkEnd w:id="2"/>' +
    '<w:r><w:t xml:space="preserve"> 2025, 1(2): 3-4.</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>[3] 张三. 题名</w:t></w:r>' +
    '<w:bookmarkStart w:id="3" w:name="Marker"/><w:r><w:t>[J]</w:t></w:r><w:bookmarkEnd w:id="3"/>' +
    '<w:r><w:t>刊名, 2025, 1(2): 3-4.</w:t></w:r></w:p>'
  );
  const page = await openTool(t);
  await selectFile(page, source);
  const output = await downloadFile(page);
  const original = await readParagraphs(page, await JSZip.loadAsync(source));
  const result = await readParagraphs(page, output);
  assert.equal(result[1].text, "[1] 张三. 示例题名[J]. 刊名, 2025, 1(2): 3-4.");
  assert.match(result[1].contentXml, /w:name="Type"\/><w:r><w:t>\[J\]<\/w:t><\/w:r><w:bookmarkEnd/u);
  assert.equal(result[2].contentXml, original[2].contentXml);
  assert.equal(result[3].contentXml, original[3].contentXml);
  assert.equal(await page.getByText(/书签或批注的边界/u).count(), 2);
});

test("题名中的英文句点无法替代作者分隔符", async (t) => {
  const source = await documentFile(
    paragraph("参考文献") +
    paragraph("[1] 这是一篇关于 U.S. 临床数据的题名［J］。刊名，2025，1（2）：3－4。") +
    paragraph("[2] J. K. Smith, Mei Zhang. Journal article［j］。Periodical，2025，8（2）：21－25。") +
    paragraph("[3] Clinical Trial A. Dose escalation［J］。Journal，2025，8（2）：1－2。") +
    paragraph("[4] 第一作者简介")
  );
  const page = await openTool(t);
  await selectFile(page, source);
  const output = await downloadFile(page);
  const original = await readParagraphs(page, await JSZip.loadAsync(source));
  const result = await readParagraphs(page, output);
  assert.equal(result[1].contentXml, original[1].contentXml);
  assert.equal(result[2].text, "[2] J. K. Smith, Mei Zhang. Journal article[J]. Periodical, 2025, 8(2): 21-25.");
  assert.deepEqual(result[3], original[3]);
  assert.deepEqual(result[4], original[4]);
  assert.equal(await page.getByText(/无法确认作者与题名的分隔/u).count(), 2);
});

test("按学院 2026 规范设置标题、条目排版并留住原来的分页、字号之外的属性", async (t) => {
  const source = await documentFile(
    paragraph("正文内容", '<w:pPr><w:jc w:val="both"/></w:pPr>') +
    paragraph("参考文献", '<w:pPr><w:pageBreakBefore/><w:spacing w:after="80"/></w:pPr>') +
    '<w:p><w:pPr><w:spacing w:line="460" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr><w:r>' +
    '<w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/><w:i/><w:sz w:val="18"/></w:rPr>' +
    '<w:lastRenderedPageBreak/><w:t>[1] SMITH J A, WANG H, et al. Test article［J］。Journal，2020，11：123。</w:t></w:r></w:p>' +
    paragraph("致谢")
  );
  const page = await openTool(t);
  await selectFile(page, source);
  const output = await downloadFile(page);
  const original = await readParagraphs(page, await JSZip.loadAsync(source));
  const result = await readParagraphs(page, output);
  assert.deepEqual(result[0], original[0]);
  assert.equal(result[1].align, "center");
  assert.equal(result[1].beforeLines, "200");
  assert.equal(result[1].font, "黑体");
  assert.equal(result[1].size, "28");
  assert.match(result[1].xml, /<w:pageBreakBefore\/>.*w:after="80"/u);
  assert.match(result[1].xml, /<w:contextualSpacing w:val="0"\/>/u);
  assert.equal(result[2].text, "[1] SMITH J A, WANG H, et al. Test article[J]. Journal, 2020, 11: 123.");
  assert.equal(result[2].align, "left");
  assert.equal(result[2].font, "仿宋");
  assert.equal(result[2].size, "24");
  assert.match(result[2].xml, /w:line="460" w:lineRule="exact"/u);
  assert.match(result[2].xml, /<w:i\/>.*<w:lastRenderedPageBreak\/>/u);
  assert.doesNotMatch(result[2].xml, /eastAsiaTheme/u);
  assert.deepEqual(result[3], original[3]);
});

test("整理会议论文与书籍，并保留出处字段和无页码的会议项", async (t) => {
  const source = await documentFile(
    paragraph("参考文献") +
    paragraph("[1] SMITH J. Example paper［C］ ／／ Example Proceedings．2025，30：42－48。") +
    paragraph("[2] SMITH J. Unpaginated paper[C] //Example Conference. 2025.") +
    paragraph("[3] 张三。示例书名［M］。北京：示例出版社，2022。") +
    paragraph("[4] SMITH J. Example volume［M］。4th ed. New York：Sample Publishing，2022：14－21。")
  );
  const page = await openTool(t);
  await selectFile(page, source);
  const output = await downloadFile(page);
  const result = await readParagraphs(page, output);
  assert.deepEqual(result.slice(1).map((p) => p.text), [
    "[1] SMITH J. Example paper[C]//Example Proceedings. 2025, 30: 42-48.",
    "[2] SMITH J. Unpaginated paper[C]//Example Conference. 2025.",
    "[3] 张三. 示例书名[M]. 北京: 示例出版社, 2022.",
    "[4] SMITH J. Example volume[M]. 4th ed. New York: Sample Publishing, 2022: 14-21."
  ]);
  assert.equal(await page.locator("#review-count").textContent(), "0");

  const zipBytes = await output.generateAsync({ type: "nodebuffer" });
  await selectFile(page, zipBytes);
  await page.locator("#result").waitFor();
  assert.equal(await page.locator("#changed").textContent(), "0");
  assert.equal(await page.locator("#review-count").textContent(), "0");
});

test("2026 正本样例识别三条自动编号，保留编号和致谢", {
  skip: !process.env.REFERENCE_TEMPLATE
}, async (t) => {
  const source = await fs.readFile(process.env.REFERENCE_TEMPLATE);
  const page = await openTool(t);
  await selectFile(page, source, "设计说明书样例.docx");
  const output = await downloadFile(page);
  assert.equal(await page.locator("#total").textContent(), "3");
  const before = await readParagraphs(page, await JSZip.loadAsync(source));
  const after = await readParagraphs(page, output);
  const heading = before.findIndex((p) => p.text.trim() === "参考文献");
  assert.notEqual(heading, -1);
  assert.deepEqual(after.slice(0, heading), before.slice(0, heading));
  assert.deepEqual(after.slice(heading + 4), before.slice(heading + 4));
  assert.deepEqual(after.map((p) => p.listId), before.map((p) => p.listId));
});

test("实际论文中的 21 条文献只在定位信息待核实时提示复核", {
  skip: !process.env.REFERENCE_COURSE_PAPER
}, async (t) => {
  const source = await fs.readFile(process.env.REFERENCE_COURSE_PAPER);
  const page = await openTool(t);
  await selectFile(page, source);
  const output = await downloadFile(page);
  const report = await page.locator(".metrics dd").allTextContents();
  t.diagnostic("课堂论文识别、著录已调整、待复核条目数 " + report.join(" / "));
  assert.equal(report[0], "21");
  const reviews = await page.locator("#review-list li").allTextContents();
  assert.equal(reviews.length, 1);
  assert.match(reviews[0], /^\[21\]/u);
  const before = await readParagraphs(page, await JSZip.loadAsync(source));
  const after = await readParagraphs(page, output);
  const heading = before.findIndex((p) => p.text.trim() === "参考文献");
  assert.notEqual(heading, -1);
  assert.deepEqual(after.slice(0, heading), before.slice(0, heading));
  assert.match(after[heading + 5].text, /^\[5\].+\[C\]\/\/Proceedings/u);
  assert.match(after[heading + 15].xml, /w:lastRenderedPageBreak/u);
  assert.equal(after[heading + 21].text, before[heading + 21].text);
});

test("拒绝非 Word 文件后可以更换文件并继续整理", async (t) => {
  const page = await openTool(t);
  await selectFile(page, Buffer.from("invalid archive"));
  await page.getByRole("alert").filter({ hasText: "无法打开" }).waitFor();
  const source = await documentFile(paragraph("参考文献") + paragraph("[1] 李四. 题名[J]. 刊名, 2025, (2): 3-4."));
  await selectFile(page, source);
  await downloadFile(page);
});
