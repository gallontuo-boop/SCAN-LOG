/**********************************************************************
 * 出貨掃描存證 — Google Apps Script 後端
 * 功能：
 *  1. 接收手機端每一筆掃描，即時寫入 Google Sheet，蓋「伺服器時間戳」
 *  2. 失敗紀錄的存證照片自動存到 Google Drive
 *  3. 結案時計算整份紀錄的 SHA-256 數位指紋，寄信到指定信箱
 *  4. 場次結案後鎖定，之後再收到同場次資料一律拒收（不可修改）
 *
 * 安裝方式請見「設定說明」文件。
 **********************************************************************/

/* ═══════ 設定區：只需要改這裡 ═══════ */
const CONFIG = {
  // 收件人，多人用逗號分隔
  RECIPIENTS: "your-name@example.com, manager@example.com",
  // 信件主旨開頭
  SUBJECT_PREFIX: "[出貨掃描存證]",
  // 存證照片的 Drive 資料夾名稱（不存在會自動建立）
  PHOTO_FOLDER: "出貨掃描存證照片",
};
/* ═══════════════════════════════════ */

const SHEET_DETAIL = "掃描明細";
const SHEET_SESSION = "場次總表";

function doGet(e) {
  // 場次查詢：在後端網址後加 ?q=場次編號或PO或品牌 即可查詢歷史紀錄
  if (e && e.parameter && e.parameter.q !== undefined) {
    return searchPage(String(e.parameter.q || "").trim());
  }
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, message: "出貨掃描後端運作中", serverTime: fmt(new Date()) })
  ).setMimeType(ContentService.MimeType.JSON);
}

/* ── 場次查詢頁（輸入場次ID／PO#／GT#／品牌關鍵字） ── */
function searchPage(q) {
  const sess = getSheet(SHEET_SESSION, null);
  let rowsHtml = "";
  if (q && sess.getLastRow() > 1) {
    const data = sess.getDataRange().getValues();
    const header = data[0];
    const hits = data.slice(1).filter(function (r) {
      return [r[1], r[2], r[3], r[4], r[5]].some(function (v) {
        return String(v).toLowerCase().indexOf(q.toLowerCase()) >= 0;
      });
    });
    if (hits.length) {
      rowsHtml = "<table border='1' cellspacing='0' cellpadding='6' style='border-collapse:collapse;font-size:13px'><tr>" +
        header.map(function (h) { return "<th style='background:#f0f0f0'>" + esc(h) + "</th>"; }).join("") + "</tr>" +
        hits.map(function (r) {
          return "<tr>" + r.map(function (v) { return "<td>" + esc(fmtCell(v)) + "</td>"; }).join("") + "</tr>";
        }).join("") + "</table>";
    } else {
      rowsHtml = "<p>查無符合「" + esc(q) + "」的場次。No results.</p>";
    }
  }
  const html = "<html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<title>場次查詢</title></head><body style='font-family:sans-serif;padding:16px'>" +
    "<h2>出貨掃描 場次查詢 Session Lookup</h2>" +
    "<form method='get'><input name='q' value='" + esc(q) + "' placeholder='場次編號 / PO# / GT# / 品牌' " +
    "style='padding:10px;font-size:16px;width:70%'> <button style='padding:10px 16px;font-size:16px'>查詢 Search</button></form>" +
    "<div style='margin-top:16px;overflow:auto'>" + rowsHtml + "</div>" +
    "<p style='color:#888;font-size:12px;margin-top:20px'>完整明細請見 Google Sheet「掃描明細」分頁。</p>" +
    "</body></html>";
  return HtmlService.createHtmlOutput(html);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "scan")     return json(handleScan(data));
    if (data.action === "finalize") return json(handleFinalize(data));
    return json({ ok: false, error: "未知的 action" });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ── 每一筆掃描：即時寫入，蓋伺服器時間 ── */
function handleScan(data) {
  const s = data.session, r = data.record;
  if (isSessionClosed(s.id)) {
    return { ok: false, error: "此場次已結案，拒絕新增紀錄" };
  }
  let photoUrl = "";
  if (r.photoBase64) {
    photoUrl = savePhoto(s.id, r.seq, r.photoBase64);
  }
  const sheet = getSheet(SHEET_DETAIL, [
    "伺服器時間", "場次ID", "序號", "條碼內容", "結果",
    "手機時間", "單筆備註", "照片連結",
    "品牌", "Style#", "PO#", "GT#", "預計箱數", "掃描人員"
  ]);
  const serverTime = fmt(new Date());
  sheet.appendRow([
    serverTime, s.id, r.seq, r.code, r.result,
    r.deviceTime, r.note || "", photoUrl,
    s.brand, s.style || "", s.po || "", s.gt || "", s.qty, s.operator
  ]);
  return { ok: true, serverTime: serverTime };
}

/* ── 結案：算指紋、寄信、鎖定場次 ── */
function handleFinalize(data) {
  const s = data.session;
  if (isSessionClosed(s.id)) {
    return { ok: false, error: "此場次已結案，不可重複結案" };
  }
  const detail = getSheet(SHEET_DETAIL, null);
  const rows = detail.getDataRange().getValues()
    .filter(row => String(row[1]) === s.id);
  if (!rows.length) return { ok: false, error: "找不到此場次的掃描紀錄" };

  // 依序號排序，組出標準字串後計算 SHA-256
  rows.sort((a, b) => a[2] - b[2]);
  const canon = rows.map(row =>
    [row[2], row[3], row[4], fmtCell(row[0]), fmtCell(row[5]), row[6]].join("|")
  ).join("\n") + "\n" +
    [s.id, s.brand, s.style || "", s.po || "", s.gt || "", s.qty, s.operator, s.startAt, data.endAt].join("|");
  const hash = sha256(canon);

  const ok  = rows.filter(r => r[4] === "成功").length;
  const bad = rows.filter(r => r[4] === "失敗").length;
  const dup = rows.filter(r => r[4] === "重複").length;
  const mis = rows.filter(r => r[4] === "不符").length;
  const missing = data.missing || [];          // 漏掃清單（有啟用箱號比對時）
  const missCount = missing.reduce(function (a, m) { return a + (m.qty || 1); }, 0);
  const closedAt = fmt(new Date());

  // 寫入場次總表（此列存在 = 場次鎖定）
  const sess = getSheet(SHEET_SESSION, [
    "結案伺服器時間", "場次ID", "品牌", "Style#", "PO#", "GT#", "預計箱數",
    "掃描人員", "開始時間", "手機結案時間", "成功", "失敗", "重複", "不符", "漏掃",
    "漏掃箱號", "SHA-256 指紋", "場次備註"
  ]);
  sess.appendRow([
    closedAt, s.id, s.brand, s.style || "", s.po || "", s.gt || "", s.qty,
    s.operator, s.startAt, data.endAt, ok, bad, dup, mis, missCount,
    missing.map(function (m) { return m.code + (m.qty > 1 ? "×" + m.qty : ""); }).join("\n"),
    hash, s.note || ""
  ]);

  sendReport(s, rows, { ok: ok, bad: bad, dup: dup, mis: mis }, missing, hash, closedAt, data.endAt);
  return { ok: true, hash: hash, message: "報告已寄至：" + CONFIG.RECIPIENTS };
}

/* ── 寄出報告信（含 PDF 附件與內嵌失敗照片） ── */
function sendReport(s, rows, counts, missing, hash, closedAt, endAt) {
  const total = s.hasList ? s.listTotal : Number(s.qty);
  const mismatch = counts.ok !== total;
  const trs = rows.map(function (row) {
    const color = row[4] === "成功" ? "#1a7f37" : row[4] === "重複" ? "#b45309" :
                  row[4] === "不符" ? "#a021a0" : "#c0392b";
    return "<tr>" +
      "<td style='padding:5px 8px;border:1px solid #ccc;text-align:center'>" + row[2] + "</td>" +
      "<td style='padding:5px 8px;border:1px solid #ccc;font-family:monospace'>" + esc(row[3]) + "</td>" +
      "<td style='padding:5px 8px;border:1px solid #ccc;color:" + color + ";font-weight:bold'>" + row[4] + "</td>" +
      "<td style='padding:5px 8px;border:1px solid #ccc;font-family:monospace'>" + fmtCell(row[0]) + "</td>" +
      "<td style='padding:5px 8px;border:1px solid #ccc;font-family:monospace'>" + fmtCell(row[5]) + "</td>" +
      "<td style='padding:5px 8px;border:1px solid #ccc'>" + esc(row[6] || "") +
        (row[7] && String(row[7]).indexOf("http") === 0 ? " <a href='" + row[7] + "'>照片</a>" : "") + "</td>" +
      "</tr>";
  }).join("");

  // 漏掃清單區塊
  let missHtml = "";
  if (s.hasList) {
    if (missing.length) {
      missHtml = "<h3 style='margin:16px 0 6px;color:#c0392b'>漏掃箱號 Missing Codes（" + missing.length + " 項）</h3>" +
        "<p style='font-family:monospace;background:#fff3f3;border:1px solid #e0b4b4;padding:8px;line-height:1.6'>" +
        missing.map(function (m) { return esc(m.code) + (m.qty > 1 ? "　×" + m.qty : ""); }).join("<br>") + "</p>";
    } else {
      missHtml = "<p style='color:#1a7f37;font-weight:bold'>清單比對：全部箱號皆已掃到，無漏掃。All codes on list scanned.</p>";
    }
    if (counts.mis) {
      missHtml += "<p style='color:#a021a0;font-weight:bold'>注意：有 " + counts.mis +
        " 筆「不符」：掃到的條碼不在清單內，可能貼錯箱。Codes scanned but NOT on list.</p>";
    }
  }

  const html =
    "<div style='font-family:sans-serif;font-size:14px;color:#222'>" +
    "<h2 style='margin:0 0 4px'>出貨掃描存證報告 Shipment Scan Report</h2>" +
    "<p style='margin:0 0 14px;color:#666'>本信件由系統於結案時自動產生並寄出，內容不可修改。Auto-generated, immutable record.</p>" +
    "<table style='border-collapse:collapse;margin-bottom:14px'>" +
    kvRow("場次編號 Session", s.id) + kvRow("品牌 Brand", s.brand) +
    (s.style ? kvRow("型號 Style #", s.style) : "") +
    (s.po ? kvRow("PO#", s.po) : "") +
    (s.gt ? kvRow("GT#", s.gt) : "") +
    kvRow("掃描人員 Operator", s.operator) +
    kvRow("開始掃描時間 Start", s.startAt) +
    kvRow("結案時間（手機）End (device)", endAt) +
    kvRow("結案時間（伺服器）End (server)", closedAt) +
    kvRow("預計箱數 Expected Cartons", total) +
    kvRow("掃描結果 Result", "成功 OK " + counts.ok + "／重複 Dup " + counts.dup +
          "／不符 NotOnList " + counts.mis + "／失敗 Failed " + counts.bad) +
    (s.note ? kvRow("備註 Remarks", s.note) : "") +
    "</table>" +
    (mismatch
      ? "<p style='color:#c0392b;font-weight:bold'>注意：成功筆數（" + counts.ok +
        "）與預計箱數（" + total + "）不一致，請查明。Count mismatch.</p>"
      : "<p style='color:#1a7f37;font-weight:bold'>成功筆數與預計箱數一致。Counts match.</p>") +
    missHtml +
    "<h3 style='margin:16px 0 6px'>掃描明細 Scan Details（每筆均含伺服器時間戳）</h3>" +
    "<table style='border-collapse:collapse;font-size:13px'>" +
    "<tr style='background:#f0f0f0'>" +
    "<th style='padding:5px 8px;border:1px solid #ccc'>序</th>" +
    "<th style='padding:5px 8px;border:1px solid #ccc'>條碼內容 Code</th>" +
    "<th style='padding:5px 8px;border:1px solid #ccc'>結果 Result</th>" +
    "<th style='padding:5px 8px;border:1px solid #ccc'>伺服器時間 Server Time</th>" +
    "<th style='padding:5px 8px;border:1px solid #ccc'>手機時間 Device Time</th>" +
    "<th style='padding:5px 8px;border:1px solid #ccc'>備註／照片 Note/Photo</th></tr>" +
    trs + "</table>" +
    buildPhotoHtml(rows).html +
    "<h3 style='margin:16px 0 6px'>紀錄數位指紋 SHA-256</h3>" +
    "<p style='font-family:monospace;word-break:break-all;background:#f6f6f6;padding:8px;border:1px solid #ddd'>" +
    hash + "</p>" +
    "<p style='color:#666;font-size:12px'>日後自 Google Sheet 匯出之 Excel／PDF，重算指紋與本信相符即可證明內容未經修改。</p>" +
    "</div>";

  const subject = CONFIG.SUBJECT_PREFIX + " " + s.brand +
    (s.po ? " / " + s.po : "") + (s.style ? " / " + s.style : "") + " / " + s.id +
    "（成功" + counts.ok + "・失敗" + counts.bad + (counts.mis ? "・不符" + counts.mis : "") + "）";

  const options = { htmlBody: html };
  // 內嵌失敗照片
  const photos = buildPhotoHtml(rows).inline;
  if (Object.keys(photos).length) options.inlineImages = photos;
  // 附上 PDF 版報告
  try {
    const pdf = Utilities.newBlob(html, "text/html", "report.html")
      .getAs("application/pdf").setName("掃描存證報告_" + s.id + ".pdf");
    options.attachments = [pdf];
  } catch (e) { /* PDF 轉換失敗時仍寄出 HTML 信 */ }

  GmailApp.sendEmail(CONFIG.RECIPIENTS, subject, "請以支援 HTML 的信箱檢視本報告。", options);
}

/* 失敗照片：內嵌於信件（最多 8 張，其餘以連結呈現） */
function buildPhotoHtml(rows) {
  const inline = {};
  let html = "";
  let n = 0;
  rows.forEach(function (row) {
    const url = String(row[7] || "");
    if (url.indexOf("http") !== 0 || n >= 8) return;
    const idMatch = url.match(/[-\w]{25,}/);
    if (!idMatch) return;
    try {
      const blob = DriveApp.getFileById(idMatch[0]).getBlob();
      const cid = "photo" + n;
      inline[cid] = blob;
      html += "<div style='margin:6px 0'><b>序 " + row[2] + "（" + fmtCell(row[0]) + "）</b><br>" +
        "<img src='cid:" + cid + "' style='max-width:340px;border:1px solid #ccc'></div>";
      n++;
    } catch (e) {}
  });
  if (html) html = "<h3 style='margin:16px 0 6px'>掃描失敗存證照片 Failure Photos</h3>" + html;
  return { html: html, inline: inline };
}

/* ── 工具函式 ── */
function isSessionClosed(sessionId) {
  const sess = getSheet(SHEET_SESSION, null);
  if (sess.getLastRow() < 2) return false;
  return sess.getRange(2, 2, sess.getLastRow() - 1, 1).getValues()
    .some(function (r) { return String(r[0]) === sessionId; });
}

function savePhoto(sessionId, seq, base64) {
  try {
    const folders = DriveApp.getFoldersByName(CONFIG.PHOTO_FOLDER);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.PHOTO_FOLDER);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), "image/jpeg",
      sessionId + "_" + seq + "_" + fmt(new Date()).replace(/[: ]/g, "-") + ".jpg");
    const file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) { return "照片儲存失敗:" + e; }
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function sha256(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, "0"); }).join("");
}

function fmt(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}
function fmtCell(v) {
  return (v instanceof Date) ? fmt(v) : String(v);
}
function esc(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
  });
}
function kvRow(k, v) {
  return "<tr><td style='padding:3px 12px 3px 0;color:#666'>" + k +
    "</td><td style='padding:3px 0;font-weight:bold'>" + esc(v) + "</td></tr>";
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
