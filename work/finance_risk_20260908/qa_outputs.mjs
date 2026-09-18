import { chromium } from "playwright";
import fs from "node:fs/promises";
import { pathToFileURL } from "node:url";

const base="C:/Users/HZJ/Desktop/Git/work/finance_risk_20260908";
const html=base+"/deliverables/03-会计职业风险实战手册.html";
const browser=await chromium.launch({headless:true,executablePath:"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"});
const page=await browser.newPage({viewport:{width:1440,height:1000},deviceScaleFactor:1});
const errors=[];
page.on("pageerror",e=>errors.push(String(e)));
await page.goto(pathToFileURL(html).href,{waitUntil:"load"});
const summary=await page.evaluate(()=>({
  title:document.title,
  h2:[...document.querySelectorAll("h2")].map(x=>x.textContent.trim()),
  riskArticles:document.querySelectorAll("article.risk").length,
  caseArticles:document.querySelectorAll("article.case").length,
  sourceArticles:document.querySelectorAll("article.source").length,
  tocLinks:document.querySelectorAll("aside nav a").length,
  brokenInternal:[...document.querySelectorAll('a[href^="#"]')].map(a=>a.getAttribute("href")).filter(h=>!document.querySelector(h)),
  blankLinks:[...document.querySelectorAll("a")].filter(a=>!a.textContent.trim()||!a.getAttribute("href")).length
}));
await page.screenshot({path:base+"/preview-handbook-top.png",fullPage:false});
await page.locator("#chapter-4").scrollIntoViewIfNeeded();
await page.screenshot({path:base+"/preview-handbook-detail.png",fullPage:false});
await browser.close();
await fs.writeFile(base+"/html_qa.json",JSON.stringify({...summary,pageErrors:errors},null,2),"utf8");
console.log(JSON.stringify({...summary,pageErrors:errors}));
