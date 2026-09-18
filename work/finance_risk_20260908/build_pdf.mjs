import { chromium } from "playwright";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const base = "C:/Users/HZJ/Desktop/Git/work/finance_risk_20260908";
const input = path.join(base, "deliverables", "03-会计职业风险实战手册.html");
const outputDir = path.join(base, "output", "pdf");
const output = path.join(outputDir, "03-会计职业风险实战手册.pdf");
await fs.mkdir(outputDir, { recursive: true });

const browser = await chromium.launch({
  headless: true,
  executablePath: "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
});
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
await page.goto(pathToFileURL(input).href, { waitUntil: "load" });

await page.evaluate(() => {
  // PDF only: use ASCII hyphens as required by the PDF production standard.
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) walker.currentNode.nodeValue = walker.currentNode.nodeValue.replace(/[‐‑‒–—―]/g, "-");

  const headings = [...document.querySelectorAll("main > section > h2")];
  const toc = document.createElement("section");
  toc.id = "pdf-toc";
  toc.className = "pdf-toc";
  toc.innerHTML = `<h2>目录</h2><p class="meta">可使用PDF阅读器的书签侧栏或全文搜索快速定位；风险场景按R编号查找，案例按C编号查找。</p><ol>${headings.map(h => `<li><a href="#${h.parentElement.id}">${h.textContent}</a></li>`).join("")}</ol><h3>重点专题入口</h3><ul><li><a href="#R05">R05 虚开增值税专用发票</a></li><li><a href="#R06">R06 虚开增值税普通发票及其他发票</a></li><li><a href="#chapter-5">第5章 常用法条与后果</a></li><li><a href="#chapter-6">第6章 发现问题后的补救SOP</a></li><li><a href="#chapter-7">第7章 官方案例</a></li></ul>`;
  document.querySelector("#chapter-1").before(toc);
});

await page.addStyleTag({ content: `
  @media print {
    @page { size: A4; margin: 18mm 17mm 18mm 17mm; }
    body { font-size: 10.5pt; line-height: 1.72; }
    main { margin: 0; padding: 0; max-width: none; }
    main > h1 { font-size: 25pt; margin-top: 18mm; }
    main > p { max-width: 170mm; }
    .pdf-toc { display: block; break-before: page; break-after: page; }
    .pdf-toc h2 { break-before: auto; border-top: none; margin-top: 0; }
    .pdf-toc ol { list-style: none; padding-left: 0; }
    .pdf-toc ul { padding-left: 7mm; }
    .pdf-toc li { margin: 3mm 0; }
    .pdf-toc a { color: #17365D; text-decoration: none; }
    h2 { font-size: 17pt; color: #17365D; }
    h3 { font-size: 12.5pt; }
    a { color: #235793; text-decoration: underline; }
    .risk > h3, .case > h3, .source > h3 { border-bottom: 0.3mm solid #D5DBE5; padding-bottom: 1.5mm; }
    .tag { border-width: 0.25mm; }
    footer { border-top: 0.3mm solid #D5DBE5; padding-top: 3mm; }
  }
` });

await page.pdf({
  path: output,
  format: "A4",
  printBackground: true,
  preferCSSPageSize: true,
  displayHeaderFooter: true,
  outline: true,
  tagged: true,
  headerTemplate: `<div style="width:100%;font-family:Arial,sans-serif;font-size:8px;color:#6B7280;padding:0 17mm;text-align:right">M12-001 会计职业风险实战手册</div>`,
  footerTemplate: `<div style="width:100%;font-family:Arial,sans-serif;font-size:8px;color:#6B7280;padding:0 17mm;text-align:center"><span class="pageNumber"></span> / <span class="totalPages"></span></div>`,
});
await browser.close();
console.log(output);
