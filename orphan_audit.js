'use strict';

/*
===========================================================
 ORPHAN SHEET BITBUCKET ACCESS AUDIT
 Enhanced with:
 ✔ html/has_access & html/no_access
 ✔ png/has_access & png/no_access
 ✔ DOCX generation for both types
===========================================================
*/

const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const axios = require("axios");
const puppeteer = require("puppeteer");
const officegen = require("officegen");
require("dotenv").config();

// -------------------------------------------------------
// CONFIG & DIRECTORIES
// -------------------------------------------------------
const INPUT_DIR = path.join(__dirname, "input_files");
const OUTPUT_DIR = path.join(__dirname, "output_files");

const HTML_DIR = path.join(OUTPUT_DIR, "html");
const HTML_HAS = path.join(HTML_DIR, "has_access");
const HTML_NO = path.join(HTML_DIR, "no_access");

const PNG_DIR = path.join(OUTPUT_DIR, "png");
const PNG_HAS = path.join(PNG_DIR, "has_access");
const PNG_NO = path.join(PNG_DIR, "no_access");

const DOC_DIR = path.join(OUTPUT_DIR, "doc");

const INPUT_XLSX = path.join(INPUT_DIR, "Orphan_Decision_Sheet_Dummy.xlsx");
const UNIQUE_CSV = path.join(OUTPUT_DIR, "orphan_unique_rows.csv");
const FORMATTED_CSV = path.join(OUTPUT_DIR, "formatted_orphan_rows.csv");

const ACCESS_CSV = path.join(OUTPUT_DIR, "orphan_access_results.csv");
const NO_ACCESS_CSV = path.join(OUTPUT_DIR, "orphan_no_access_results.csv");

const URL = process.env.BB_URL || "localhost:7990";
const USERNAME = process.env.BB_USERNAME || "admin";
const KEYNAME = process.env.BB_KEYNAME || "REPLACE_ME";

// optional image processing for tight-crop
let sharp;
try { sharp = require('sharp'); } catch (e) { sharp = null; }

const OUTER_PAD = 8;
const CROP_COLOR_THRESHOLD = 200; // 0-255, lower => more pixels considered non-white

// -------------------------------------------------------
function ensureDir(d) { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); }

ensureDir(INPUT_DIR);
ensureDir(OUTPUT_DIR);

ensureDir(HTML_DIR);
ensureDir(HTML_HAS);
ensureDir(HTML_NO);

ensureDir(PNG_DIR);
ensureDir(PNG_HAS);
ensureDir(PNG_NO);

ensureDir(DOC_DIR);

// -------------------------------------------------------
function trim(v) { return String(v || "").trim(); }
function csvRow(a) { return a.map(x => String(x).replace(/\n/g, " ")).join(","); }

function timestamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} `
       + `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

function safeTs() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}_`
       + `${p(d.getHours())}-${p(d.getMinutes())}-${p(d.getSeconds())}`;
}

// -------------------------------------------------------
// STEP 1 – Load XLSX
// -------------------------------------------------------
function loadXlsx() {
  console.log("[STEP 1] Loading XLSX...");
  const wb = xlsx.readFile(INPUT_XLSX);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return xlsx.utils.sheet_to_json(ws, { defval: "" });
}

// -------------------------------------------------------
// STEP 2 – Extract unique User SSO rows
// -------------------------------------------------------
function extractUniqueUsers(rows) {
  console.log("[STEP 2] Extracting unique users...");
  const header = Object.keys(rows[0]);
  const seen = new Set();
  const unique = [];

  for (const r of rows) {
    const sso = trim(r["User SSO"]);
    if (!sso) continue;
    if (!seen.has(sso)) {
      seen.add(sso);
      unique.push(r);
    }
  }

  const out = [csvRow(header)];
  unique.forEach(r => out.push(csvRow(header.map(h => r[h]))));
  fs.writeFileSync(UNIQUE_CSV, out.join("\n"));

  console.log("[INFO] Unique rows → " + UNIQUE_CSV);
  return unique;
}

// -------------------------------------------------------
// STEP 3 – Format rows into final CSV
// -------------------------------------------------------
function formatRows(uniqueRows) {
  console.log("[STEP 3] Formatting entitlement → ProjectKey + AccessPermission");
  const header = ["User SSO","Account ID","Project Key","Access Permission"];
  const out = [csvRow(header)];

  uniqueRows.forEach(r => {
    const user = trim(r["User SSO"]);
    const acc = trim(r["Account ID"]);
    const ent = trim(r["Entitlement Description"] || r["Entitlement Desription"]);

    let pk = "", perm = "";
    if (ent.includes(":")) {
      const last = ent.split(":").pop().trim();
      pk = trim(last.split("-")[0]);
      perm = trim(last.split("-")[1]);
    }

    out.push(csvRow([user, acc, pk, perm]));
  });

  fs.writeFileSync(FORMATTED_CSV, out.join("\n"));
  console.log("[INFO] Formatted CSV → " + FORMATTED_CSV);
}

// -------------------------------------------------------
// STEP 4 – Access Check + HTML + PNG
// -------------------------------------------------------
async function accessCheck() {
  console.log("[STEP 4] Access Check with HTML + PNG …");

  fs.writeFileSync(ACCESS_CSV,
    csvRow(["User SSO","Account ID","Project Key","Access Permission",
            "Access Status","Timestamp","HTML","PNG"]) + "\n"
  );
  fs.writeFileSync(NO_ACCESS_CSV,
    csvRow(["User SSO","Account ID","Project Key","Access Permission",
            "Access Status","Timestamp","HTML","PNG"]) + "\n"
  );

  const lines = fs.readFileSync(FORMATTED_CSV,"utf8").trim().split(/\r?\n/).slice(1);

  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();

  for (const l of lines) {
    const [user, acc, project, perm] = l.split(",").map(trim);
    if (!user || !project) continue;

    console.log(`[CHECK] ${user} → ${project}`);

    const apiUrl =
      `http://${URL}/rest/api/1.0/projects/${project}/permissions/users?filter=${encodeURIComponent(user)}`;

    let apiData = { values: [] };
    try {
      const r = await axios.get(apiUrl, {
        auth: { username: USERNAME, password: KEYNAME }
      });
      apiData = r.data;
    } catch {}

    const hasAccess = (apiData.values || []).length > 0;
    const ts = timestamp();
    const sTs = safeTs();

    const htmlFile = hasAccess
      ? path.join(HTML_HAS, `${user}_${project}_${sTs}.html`)
      : path.join(HTML_NO,  `${user}_${project}_${sTs}.html`);

    const pngFile = hasAccess
      ? path.join(PNG_HAS, `${user}_${project}_${sTs}.png`)
      : path.join(PNG_NO,  `${user}_${project}_${sTs}.png`);


    // Generate HTML Evidence (wrap content in #evidence for tight clipping)
    const html = `
<html>
<head>
  <meta charset="utf-8" />
  <style>
    body { font-family: monospace; margin:8px; background: #fff; color:#000 }
    #evidence { display:inline-block; background:#fff; color:#000; }
    pre { display:inline-block; white-space:pre-wrap; word-break:break-word; max-width:1000px; }
    h2,h3 { margin:6px 0 }
  </style>
</head>
<body>
<div id="evidence">
<b>User:</b> ${user}<br>
<b>Account ID:</b> ${acc}<br>
<b>Project Key:</b> ${project}<br>
<b>Timestamp:</b> ${ts}<br>
<h3>API URL</h3><pre>${apiUrl}</pre>
<h3>API Response</h3>
<pre>${JSON.stringify(apiData,null,2)}</pre>
</div>
</body>
</html>`;
    fs.writeFileSync(htmlFile, html);

    // Screenshot: try to clip to the #evidence bounding box (preferred)
    await page.goto("file://" + htmlFile, { waitUntil: "networkidle0" });
    try { await page.waitForSelector('#evidence', { timeout: 2000 }); } catch (e) {}

    try {
      const box = await page.evaluate(() => {
        const el = document.getElementById('evidence');
        if (!el) return null;
        const r = el.getBoundingClientRect();
        return { x: r.x, y: r.y, width: r.width, height: r.height, dpr: window.devicePixelRatio || 1 };
      });

      if (box && box.width > 0 && box.height > 0) {
        const dpr = box.dpr || 1;
        const clipX = Math.max(0, Math.floor(box.x * dpr) - OUTER_PAD);
        const clipY = Math.max(0, Math.floor(box.y * dpr) - OUTER_PAD);
        const clipW = Math.ceil(box.width * dpr) + OUTER_PAD * 2;
        const clipH = Math.ceil(box.height * dpr) + OUTER_PAD * 2;

        const pageSize = await page.evaluate(() => ({ w: document.documentElement.scrollWidth, h: document.documentElement.scrollHeight }));
        const maxW = Math.ceil(pageSize.w * (box.dpr || 1));
        const maxH = Math.ceil(pageSize.h * (box.dpr || 1));
        const finalW = Math.min(clipW, Math.max(1, maxW - clipX));
        const finalH = Math.min(clipH, Math.max(1, maxH - clipY));

        await page.screenshot({ path: pngFile, clip: { x: clipX, y: clipY, width: finalW, height: finalH } });
      } else {
        await page.screenshot({ path: pngFile, fullPage: true });
        if (sharp) {
          try { await sharp(pngFile).trim().toFile(pngFile + '.tmp'); fs.renameSync(pngFile + '.tmp', pngFile); }
          catch (e) { /* ignore */ }
        }
      }
    } catch (err) {
      // fallback to full page screenshot
      try { await page.screenshot({ path: pngFile, fullPage: true }); } catch (e) {}
    }

    const logRow = csvRow([
      user, acc, project, perm,
      hasAccess ? "HAS_ACCESS" : "NO_ACCESS",
      ts, htmlFile, pngFile
    ]) + "\n";

    if (hasAccess) fs.appendFileSync(ACCESS_CSV, logRow);
    else fs.appendFileSync(NO_ACCESS_CSV, logRow);
  }

  await browser.close();

  console.log("[INFO] Access check done.");
}

// -------------------------------------------------------
// STEP 5 – Generate DOCX
// -------------------------------------------------------
async function generateDocx(dirPath, outName) {
  const images = fs.readdirSync(dirPath).filter(f => f.endsWith(".png"));
  if (!images.length) return;

  console.log(`Generating DOCX → ${outName}`);

  const docx = officegen("docx");
  docx.on("error", console.error);

  const pTitle = docx.createP({ align: "center" });
  pTitle.addText(outName.replace(".docx",""), { bold:true, font_size:28 });

  docx.createP();
  docx.createP({ align:"center" })
    .addText("Generated: " + new Date().toLocaleString(), { font_size: 14 });

  docx.createP().addText("", { pageBreakBefore:true });

  images.forEach((img, idx) => {
    docx.createP({ align:"center" }).addImage(path.join(dirPath, img));
    if (idx < images.length - 1) docx.createP().addText("", { pageBreakBefore:true });
  });

  const outPath = path.join(DOC_DIR, outName);
  const out = fs.createWriteStream(outPath);

  return new Promise((resolve, reject) => {
    out.on("close", () => {
      console.log("✔ DOCX created → " + outPath);
      resolve();
    });
    out.on("error", reject);
    docx.generate(out);
  });
}

// -------------------------------------------------------
// MAIN
// -------------------------------------------------------
async function main() {
  const rows = loadXlsx();
  const unique = extractUniqueUsers(rows);
  formatRows(unique);
  await accessCheck();

  await generateDocx(PNG_HAS, "Orphan_Has_Access_Report.docx");
  await generateDocx(PNG_NO,  "Orphan_No_Access_Report.docx");

  console.log("===============================================");
  console.log("PROCESS COMPLETE");
  console.log("Reports saved in → output_files/doc/");
  console.log("===============================================");
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
