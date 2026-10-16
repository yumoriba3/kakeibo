// マイ出納帳 ⇔ Googleスプレッドシート 連携用プログラム
// スプレッドシートの「拡張機能」→「Apps Script」に、このファイルの中身をすべて貼り付けて使います。
// 手順の詳細は SETUP_GOOGLE.md を見てください。

// ▼ ここを、自分だけの合言葉に書き換えてください（英数字で12文字以上。手順書の例をそのまま使わないでください）
const TOKEN = "ここを自分だけの合言葉に変更";

const TYPE_LABEL = { exp: "支出", inc: "収入" };
const SCRIPT_VERSION = 2; // アプリ側で「プログラムが古い版か」を判定するための番号
const NO_SUB = "（小分類なし）";

/** 項目 c の小分類の並び（設定の順 → 設定に無い名前 → 小分類なし） */
function subOrder_(cats, t, c, items) {
  const base = (((cats.sub || {})[t] || {})[c] || []).slice();
  items.forEach(function (e) { if (e.s && base.indexOf(e.s) < 0) base.push(e.s); });
  if (items.some(function (e) { return !e.s; })) base.push("");
  return base;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (TOKEN.indexOf("ここを") === 0) return out_({ ok: false, error: "token_not_set" });
    if (req.token !== TOKEN) return out_({ ok: false, error: "token" });
    if (req.action === "ping") return out_({ ok: true });
    if (req.action === "save") { save_(req.data); return out_({ ok: true, savedAt: new Date().toISOString() }); }
    if (req.action === "load") return out_({ ok: true, data: load_() });
    return out_({ ok: false, error: "action" });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out_({ ok: true, message: "マイ出納帳の連携用URLです。このURLをアプリに入力してください。" });
}

function out_(obj) {
  obj.v = SCRIPT_VERSION;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** アプリから届いた全データを保存し、見やすい表を作り直す */
function save_(text) {
  const data = JSON.parse(text);
  const docs = data.docs || {};

  // 1) 復元用の元データ（非表示シートに分割して保存）
  const raw = sheet_("_backup");
  raw.clear();
  const chunks = [];
  // 先頭に「#」を付け、数式や数値として解釈されないようにする
  for (let i = 0; i < text.length; i += 40000) chunks.push(["#" + text.slice(i, i + 40000)]);
  if (chunks.length) raw.getRange(1, 1, chunks.length, 1).setNumberFormat("@").setValues(chunks);

  // 2) 明細シート
  let entries = [];
  Object.keys(docs).forEach(function (k) {
    if (k.indexOf("months/") === 0) entries = entries.concat(docs[k].entries || []);
  });
  entries.sort(function (a, b) { return a.d < b.d ? -1 : a.d > b.d ? 1 : 0; });
  const detail = sheet_("明細");
  detail.clear();
  const rows = [["日付", "区分", "項目", "小分類", "金額", "メモ", "定期"]].concat(entries.map(function (e) {
    const p = e.d.split("-");
    return [new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])), TYPE_LABEL[e.t] || e.t, e.c, e.s || "", e.a, e.m || "", e.r ? "定期" : ""];
  }));
  detail.getRange(1, 1, rows.length, 7).setValues(rows);
  detail.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#E0EFF0");
  if (rows.length > 1) {
    detail.getRange(2, 1, rows.length - 1, 1).setNumberFormat("yyyy/mm/dd");
    detail.getRange(2, 5, rows.length - 1, 1).setNumberFormat("#,##0");
  }
  detail.setFrozenRows(1);

  // 3) 月別集計シート（項目 × 月）
  const months = [];
  const sums = {};
  entries.forEach(function (e) {
    const ym = e.d.slice(0, 7);
    if (months.indexOf(ym) < 0) months.push(ym);
    const k = e.t + "|" + e.c + "|" + ym;
    sums[k] = (sums[k] || 0) + e.a;
    const ks = e.t + "|" + e.c + "|" + (e.s || "") + "|" + ym;
    sums[ks] = (sums[ks] || 0) + e.a;
    const kt = e.t + "||" + ym;
    sums[kt] = (sums[kt] || 0) + e.a;
  });
  months.sort();
  const cats = docs["settings/categories"] || { exp: [], inc: [] };
  const head = ["項目"].concat(months.map(function (m) { return Number(m.slice(0, 4)) + "年" + Number(m.slice(5)) + "月"; }), ["合計"]);
  const width = head.length;
  const line = function (label, vals) { return [label].concat(vals, [vals.reduce(function (a, b) { return a + b; }, 0)]); };
  const blank = function () { const r = []; for (let i = 0; i < width; i++) r.push(""); return r; };
  const table = [head];
  const styleRows = { section: [], total: [], sub: [] };
  ["inc", "exp"].forEach(function (t) {
    const list = (cats[t] || []).slice();
    entries.forEach(function (e) { if (e.t === t && list.indexOf(e.c) < 0) list.push(e.c); });
    const sec = blank(); sec[0] = "【" + TYPE_LABEL[t] + "】"; table.push(sec); styleRows.section.push(table.length);
    list.forEach(function (c) {
      table.push(line(c, months.map(function (m) { return sums[t + "|" + c + "|" + m] || 0; })));
      const items = entries.filter(function (e) { return e.t === t && e.c === c; });
      if (!items.some(function (e) { return e.s; })) return;
      subOrder_(cats, t, c, items).forEach(function (sn) {
        const vals = months.map(function (m) { return sums[t + "|" + c + "|" + sn + "|" + m] || 0; });
        if (vals.some(function (v) { return v; })) { table.push(line("　└ " + (sn || NO_SUB), vals)); styleRows.sub.push(table.length); }
      });
    });
    table.push(line(TYPE_LABEL[t] + "合計", months.map(function (m) { return sums[t + "||" + m] || 0; }))); styleRows.total.push(table.length);
    table.push(blank());
  });
  table.push(line("収支", months.map(function (m) { return (sums["inc||" + m] || 0) - (sums["exp||" + m] || 0); }))); styleRows.total.push(table.length);
  const summary = sheet_("月別集計");
  summary.clear();
  summary.getRange(1, 1, table.length, width).setValues(table);
  if (width > 1) summary.getRange(2, 2, table.length - 1, width - 1).setNumberFormat("#,##0;▲#,##0;-");
  summary.getRange(1, 1, 1, width).setFontWeight("bold").setBackground("#E0EFF0");
  styleRows.section.forEach(function (r) { summary.getRange(r, 1).setFontWeight("bold"); });
  styleRows.total.forEach(function (r) { summary.getRange(r, 1, 1, width).setFontWeight("bold").setBackground("#F2F4F6"); });
  styleRows.sub.forEach(function (r) { summary.getRange(r, 1, 1, width).setFontColor("#5D6A7E"); });
  summary.setFrozenRows(1);
  summary.setFrozenColumns(1);
  summary.getRange(table.length + 2, 1).setValue("最終更新：" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm"));

  // 4) 内訳シート（月 → 項目 → 日付・店・金額。新しい月が上）
  const breakdown = writeBreakdown_(entries, cats);

  // 5) シートの並びを整える
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(summary); ss.moveActiveSheet(1);
  ss.setActiveSheet(breakdown); ss.moveActiveSheet(2);
  ss.setActiveSheet(detail); ss.moveActiveSheet(3);
  ["シート1", "Sheet1"].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  if (!raw.isSheetHidden()) raw.hideSheet();
}

/** 内訳シート：月ごと・項目ごとに、日付と店（メモ）と金額を並べる */
function writeBreakdown_(entries, cats) {
  const sh = sheet_("内訳");
  sh.clear();
  const byMonth = {};
  entries.forEach(function (e) { const ym = e.d.slice(0, 7); (byMonth[ym] = byMonth[ym] || []).push(e); });
  const months = Object.keys(byMonth).sort().reverse();
  const rows = [["項目", "日付", "店・メモ", "金額"]];
  const fmt = { month: [], cat: [], sub: [], date: [] };
  months.forEach(function (ym) {
    const list = byMonth[ym];
    rows.push([Number(ym.slice(0, 4)) + "年" + Number(ym.slice(5)) + "月", "", "", ""]); fmt.month.push(rows.length);
    ["exp", "inc"].forEach(function (t) {
      const order = (cats[t] || []).slice();
      list.forEach(function (e) { if (e.t === t && order.indexOf(e.c) < 0) order.push(e.c); });
      order.forEach(function (c) {
        const items = list.filter(function (e) { return e.t === t && e.c === c; });
        if (!items.length) return;
        const total = items.reduce(function (a, e) { return a + e.a; }, 0);
        rows.push([(t === "inc" ? "収入：" : "") + c + "（" + items.length + "件）", "", "", total]); fmt.cat.push(rows.length);
        const pushItems = function (list2) {
          list2.forEach(function (e) {
            const p = e.d.split("-");
            rows.push(["", new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])), e.m || "（メモなし）", e.a]); fmt.date.push(rows.length);
          });
        };
        if (!items.some(function (e) { return e.s; })) { pushItems(items); return; }
        subOrder_(cats, t, c, items).forEach(function (sn) {
          const sub = items.filter(function (e) { return (e.s || "") === sn; });
          if (!sub.length) return;
          rows.push(["　" + (sn || NO_SUB) + "（" + sub.length + "件）", "", "", sub.reduce(function (a, e) { return a + e.a; }, 0)]); fmt.sub.push(rows.length);
          pushItems(sub);
        });
      });
    });
    rows.push(["", "", "", ""]);
  });
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#E0EFF0");
  if (rows.length > 1) sh.getRange(2, 4, rows.length - 1, 1).setNumberFormat("#,##0");
  fmt.month.forEach(function (r) { sh.getRange(r, 1, 1, 4).setFontWeight("bold").setFontSize(12).setBackground("#16233A").setFontColor("#FFFFFF"); });
  fmt.cat.forEach(function (r) { sh.getRange(r, 1, 1, 4).setFontWeight("bold").setBackground("#F2F4F6"); });
  fmt.sub.forEach(function (r) { sh.getRange(r, 1, 1, 4).setFontWeight("bold").setFontColor("#0E6B74"); });
  fmt.date.forEach(function (r) { sh.getRange(r, 2).setNumberFormat("m/d（ddd）"); });
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 260); sh.setColumnWidth(4, 110);
  return sh;
}

/** 復元用：保存してある元データを返す */
function load_() {
  const raw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_backup");
  if (!raw || raw.getLastRow() === 0) return "";
  return raw.getRange(1, 1, raw.getLastRow(), 1).getValues().map(function (r) { return String(r[0]).slice(1); }).join("");
}
