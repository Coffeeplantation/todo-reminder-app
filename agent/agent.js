/**
 * agent.js — Todoリマインダー 常駐エージェント（完全ローカル・依存パッケージなし）
 *
 * 役割（30分ごと＋起動時に自動実行）:
 *  1. Outlookの指定メールフォルダを読む（agent/read-mail.ps1、COM経由）
 *  2. 設定された予定表リンク＋メール本文中のリンクを直接取得する（第三者プロキシなし）
 *  3. ローカルAI（Ollama）で予定・タスク・締め切りを抽出する
 *  4. 出典（メール件名/差出人/受信日時、リンクURL、取得日時）付きで受信箱（inbox）に蓄積する
 *     → ブラウザアプリが受信箱をポーリングしてタスクボードへ自動追加する
 *  5. アプリから受け取った最新状態（タスク・名簿・設定）をもとにリマインダーメールを
 *     Outlookで自動送信、または下書き作成する（agent/send-mail.ps1）
 *
 * プライバシー: 127.0.0.1 のみで待ち受け。外部への通信は「ユーザーが登録したリンク先URL」への
 * 直接取得（読み込みのみ）だけ。メール・タスク内容がクラウドへ送信されることはない。
 *
 * 起動: node agent/agent.js   （常駐化は「エージェント常駐登録.ps1」参照）
 */
"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { execFile } = require("child_process");

const ROOT = path.join(__dirname, "..");
const BROWSER_DIR = path.join(ROOT, "browser");
const DATA_DIR = path.join(__dirname, "data");
const PORT = 8787;
const OLLAMA_URL = "http://localhost:11434";

// ---- 永続データ（agent/data/*.json） ----
fs.mkdirSync(DATA_DIR, { recursive: true });

function loadJson(name, fallback) {
  // メモ帳やPowerShellで編集されたファイルはBOM付きUTF-8のことがあるため取り除いてから解析する
  try { return JSON.parse(fs.readFileSync(path.join(DATA_DIR, name), "utf8").replace(/^﻿/, "")); }
  catch (_) { return fallback; }
}
function saveJson(name, data) {
  fs.writeFileSync(path.join(DATA_DIR, name), JSON.stringify(data, null, 1), "utf8");
}

const DEFAULT_CONFIG = {
  enabled: true,        // 自動収集サイクルのON/OFF
  mailEnabled: true,    // メールボックスを読むか（Outlookプロファイル未設定のPCではOFFにする）
  mailFolder: "受信トレイ", // 読み取るOutlookフォルダ名
  mailDays: 7,          // 何日前までのメールを対象にするか
  links: [],            // 定期取得する予定表リンク（URLの配列）
  followLinks: true,    // メール本文中のリンクも読みに行くか
  intervalMin: 30,      // 収集サイクルの間隔（分）
  autoSend: false       // true=リマインダーを実際に送信 / false=Outlookの下書きに保存
};

let config = Object.assign({}, DEFAULT_CONFIG, loadJson("config.json", {}));
let inbox = loadJson("inbox.json", []);          // アプリ未取り込みの抽出済み予定
let processed = loadJson("processed.json", { mailIds: [], linkHash: {}, taskKeys: [] });
let snapshot = loadJson("snapshot.json", null);  // アプリから受け取った最新状態
let sentLog = loadJson("sent.json", {});         // リマインダー送信済み記録 {taskId: {first,last}}
let log = loadJson("log.json", []);

let lastRun = null;
let running = false;
let timer = null;

function addLog(kind, msg) {
  log.unshift({ at: Date.now(), kind, msg: String(msg).slice(0, 500) });
  log = log.slice(0, 200);
  saveJson("log.json", log);
  console.log(`[${new Date().toLocaleTimeString()}] ${kind}: ${msg}`);
}

// ---- 日付ヘルパー ----
const pad2 = n => String(n).padStart(2, "0");
const fmtDate = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
function todayDate() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }
function parseDue(s) {
  const m = String(s || "").match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
const daysBetween = (a, b) => Math.round((b - a) / 86400000);

// ---- PowerShell 実行 ----
function runPs(script, args) {
  return new Promise(resolve => {
    execFile("powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path.join(__dirname, script), ...args],
      { maxBuffer: 64 * 1024 * 1024, timeout: 180000 },
      (err, stdout) => {
        if (err && !stdout) { resolve({ error: "ps_error", message: err.message }); return; }
        try { resolve(JSON.parse(String(stdout).trim())); }
        catch (e) { resolve({ error: "parse_error", message: String(stdout).slice(0, 300) }); }
      });
  });
}

// ---- リンク取得（入力URLへ直接。読み込みのみ・外部送信なし） ----
async function fetchLinkText(url) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), 20000);
  try {
    const res = await fetch(url, { signal: ctrl.signal, redirect: "follow" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const ct = res.headers.get("content-type") || "";
    let text = await res.text();
    if (/html/i.test(ct) || /^\s*</.test(text)) {
      // HTMLはスクリプト・スタイルを除去してテキスト化
      text = text
        .replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<style[\s\S]*?<\/style>/gi, " ")
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
        .replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n");
    }
    return text.slice(0, 20000);
  } finally { clearTimeout(t); }
}

// ---- ローカルAI（Ollama）抽出 ----
function extractionPrompt() {
  const today = fmtDate(todayDate());
  const year = today.slice(0, 4);
  return `あなたは文書やメールから予定・タスク・締め切りを抽出するアシスタントです。次の形式のJSONだけを返してください（説明文は書かない）:
{"tasks":[{"name":"タスク名","due":"yyyy-mm-dd","assignee":"担当者名（不明なら空文字）","participants":["参加者名"]}]}
今日の日付は ${today} です。「来週金曜」などの相対表現は具体的な日付に変換してください。期日が読み取れない場合 due は空文字にしてください。

例：
入力「件名: 会議資料の件\n本文: 8月3日までに企画書を提出してください。担当は田中さんです。会議は8月5日14時からです。」
出力:
{"tasks":[{"name":"企画書の提出","due":"${year}-08-03","assignee":"田中","participants":[]},{"name":"会議","due":"${year}-08-05","assignee":"","participants":[]}]}

「〜までに提出」「〜日開催」「期限」「締め切り」などはすべてタスクとして抽出してください。`;
}

async function ollamaExtract(text) {
  const model = ((snapshot && snapshot.settings && snapshot.settings.ollamaModel) || "qwen2.5:3b").trim();
  const res = await fetch(`${OLLAMA_URL}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model, stream: false, format: "json",
      options: { num_ctx: 8192, temperature: 0 },
      messages: [
        { role: "system", content: extractionPrompt() },
        { role: "user", content: String(text).slice(0, 6000) }
      ]
    })
  });
  if (!res.ok) throw new Error(`Ollamaエラー（${res.status}）`);
  const data = await res.json();
  const m = String((data.message && data.message.content) || "").match(/\{[\s\S]*\}/);
  if (!m) return [];
  try {
    const parsed = JSON.parse(m[0]);
    return Array.isArray(parsed.tasks) ? parsed.tasks : [];
  } catch (_) { return []; }
}

// ---- 抽出結果を受信箱へ（重複排除＋出典付き） ----
function taskKey(name, due) { return `${String(name).trim()}|${String(due || "").trim()}`; }

function addToInbox(tasks, source) {
  let added = 0;
  for (const t of tasks) {
    const name = String(t.name || "").trim();
    if (!name) continue;
    const due = /^\d{4}-\d{2}-\d{2}$/.test(t.due || "") ? t.due : "";
    const key = taskKey(name, due);
    if (processed.taskKeys.includes(key)) continue; // 過去に配信済み
    if (inbox.some(x => taskKey(x.name, x.due) === key)) continue;
    if (snapshot && (snapshot.tasks || []).some(x => taskKey(x.name, x.due) === key)) continue;
    inbox.push({
      uid: crypto.randomUUID(),
      name, due,
      assignee: String(t.assignee || "").trim(),
      participants: Array.isArray(t.participants) ? t.participants.map(p => String(p).trim()).filter(Boolean) : [],
      source: Object.assign({ at: Date.now() }, source),
      extractedAt: Date.now()
    });
    processed.taskKeys.push(key);
    added++;
  }
  processed.taskKeys = processed.taskKeys.slice(-5000);
  return added;
}

// ---- 収集サイクル ----
async function collectMails() {
  const r = await runPs("read-mail.ps1", ["-Folder", config.mailFolder, "-Days", String(config.mailDays), "-MaxCount", "100"]);
  if (r.error) { addLog("mail", `メール読み取り失敗: ${r.error} ${r.message || r.folder || ""}`); return 0; }
  // ConvertTo-Json は要素が1件だと配列にならないため、必ず配列に揃える
  const list = Array.isArray(r.mails) ? r.mails : (r.mails ? [r.mails] : []);
  const mails = list.filter(m => m && m.id && !processed.mailIds.includes(m.id));
  let added = 0;
  for (const m of mails) {
    try {
      const tasks = await ollamaExtract(`件名: ${m.subject}\n差出人: ${m.from}\n受信日時: ${m.received}\n本文:\n${m.body}`);
      added += addToInbox(tasks, {
        type: "mail",
        label: `📧 メール「${m.subject || "（件名なし）"}」`,
        detail: `差出人: ${m.from} / 受信: ${String(m.received).replace("T", " ")} / フォルダ: ${r.folder}`
      });
      // メール本文中のリンクも読む（最大3件/通）
      if (config.followLinks) {
        const urls = [...new Set((String(m.body).match(/https?:\/\/[^\s<>"'）)】\]]+/g) || []))].slice(0, 3);
        for (const url of urls) added += await collectLink(url, `メール「${m.subject}」記載のリンク`);
      }
      processed.mailIds.push(m.id);
    } catch (e) {
      addLog("mail", `抽出失敗（${m.subject}）: ${e.message}`);
      if (/Ollama/.test(e.message) || /fetch failed/.test(e.message)) break; // Ollama停止中は打ち切り
    }
  }
  processed.mailIds = processed.mailIds.slice(-5000);
  if (mails.length) addLog("mail", `メール ${mails.length} 通を確認し ${added} 件の新しい予定を抽出`);
  return added;
}

async function collectLink(url, viaLabel) {
  try {
    const text = await fetchLinkText(url);
    const hash = crypto.createHash("sha1").update(text).digest("hex");
    if (processed.linkHash[url] === hash) return 0; // 前回から変化なし
    const tasks = await ollamaExtract(`以下はウェブページ（${url}）の本文です。\n${text}`);
    const added = addToInbox(tasks, {
      type: "link",
      label: `🔗 ${viaLabel || "予定表リンク"}`,
      url
    });
    processed.linkHash[url] = hash;
    if (added) addLog("link", `${url} から ${added} 件の予定を抽出`);
    return added;
  } catch (e) {
    addLog("link", `リンク取得失敗（${url}）: ${e.message}`);
    return 0;
  }
}

// ---- リマインダー（自動送信 or 下書き作成） ----
function applyTemplate(tpl, ctx) {
  return String(tpl)
    .replace(/\{タスク名\}/g, ctx.name).replace(/\{期日\}/g, ctx.due)
    .replace(/\{残り日数\}/g, String(ctx.daysLeft)).replace(/\{担当者\}/g, ctx.assignee)
    .replace(/\{参加者\}/g, ctx.participants).replace(/\{宛先名\}/g, ctx.toNames)
    .replace(/\{署名\}/g, ctx.signature);
}

async function runReminders() {
  if (!snapshot || !Array.isArray(snapshot.tasks)) return; // アプリを一度開くとスナップショットが届く
  const s = snapshot.settings || {};
  const contacts = snapshot.contacts || [];
  const emailOf = name => {
    const c = contacts.find(c => (c.name || "").trim() === (name || "").trim());
    return c ? (c.email || "").trim() : "";
  };
  const today = todayDate();
  const items = [];

  for (const task of snapshot.tasks) {
    if (task.status === "完了" || task.archived) continue;
    const due = parseDue(task.due);
    if (!due) continue;
    const reminderDays = task.reminderDays > 0 ? task.reminderDays : (s.reminderDays || 3);
    const firstDate = new Date(due); firstDate.setDate(firstDate.getDate() - reminderDays);
    const lastDate = new Date(due); lastDate.setDate(lastDate.getDate() - Math.max(0, s.lastReminderDays ?? 1));
    const rec = sentLog[task.id] || {};
    const appSent = task.autoMail || {}; // アプリ側で作成済みの分は二重にしない
    let kind = null;
    if (daysBetween(lastDate, today) >= 0 && !rec.last && !appSent.last) kind = "直前";
    else if (daysBetween(firstDate, today) >= 0 && !rec.first && !appSent.first) kind = "指定日";
    if (!kind) continue;

    const names = [task.assignee, ...(task.participants || [])].map(x => (x || "").trim()).filter(Boolean);
    const uniqNames = [...new Set(names)];
    const emails = uniqNames.map(emailOf).filter(Boolean);
    const fixed = (s.mailTo || "").split(/[,、;\s]+/).filter(Boolean);
    let to, cc;
    if (fixed.length) { to = fixed.join(", "); cc = emails.join(", "); }
    else if (emails.length) { to = emails[0]; cc = emails.slice(1).join(", "); }
    else { addLog("remind", `宛先メール未登録のためスキップ: ${task.name}`); continue; }

    const daysLeft = daysBetween(today, due);
    const ctx = {
      name: task.name, due: task.due || "", daysLeft,
      assignee: task.assignee || "", participants: (task.participants || []).join("、"),
      toNames: uniqNames.join("様、") + (uniqNames.length ? "様" : ""),
      signature: s.signature || "Todoリマインダーシステム"
    };
    const subject = (s.mailSubject || "").trim()
      ? applyTemplate(s.mailSubject, ctx)
      : `【リマインド${kind === "直前" ? "・直前" : ""}】${ctx.name}（期日：${ctx.due}）`;
    const body = (s.mailBody || "").trim()
      ? applyTemplate(s.mailBody, ctx)
      : `${ctx.toNames}\n\n「${ctx.name}」の期日（${ctx.due}）まで残り ${daysLeft} 日です。\nご対応をお願いいたします。\n\n${ctx.signature}`;

    items.push({ taskId: task.id, kind, to, cc, subject, body, from: s.mailFrom || "" });
  }

  if (!items.length) return;
  const jsonPath = path.join(DATA_DIR, "outbox.json");
  fs.writeFileSync(jsonPath, JSON.stringify({ send: !!config.autoSend, items }, null, 1), "utf8");
  const r = await runPs("send-mail.ps1", ["-JsonPath", jsonPath]);
  if (r.error) { addLog("remind", `メール${config.autoSend ? "送信" : "下書き"}失敗: ${r.message}`); return; }
  for (const it of items) {
    const rec = sentLog[it.taskId] || {};
    if (it.kind === "直前") { rec.last = fmtDate(today); if (!rec.first) rec.first = fmtDate(today); }
    else rec.first = fmtDate(today);
    sentLog[it.taskId] = rec;
  }
  saveJson("sent.json", sentLog);
  addLog("remind", `リマインダー ${r.ok} 件を${config.autoSend ? "自動送信" : "Outlookの下書きに作成"}しました` +
    (r.errors && r.errors.length ? `（失敗: ${r.errors.join(" / ")}）` : ""));
}

async function cycle(reason) {
  if (running) return;
  running = true;
  try {
    addLog("cycle", `収集サイクル開始（${reason}）`);
    let added = 0;
    if (config.enabled) {
      if (config.mailEnabled) added += await collectMails();
      for (const url of config.links || []) added += await collectLink(url, "予定表リンク");
      if (config.mailEnabled) await runReminders(); // 送信もOutlook経由のためメール連携OFF時はスキップ
    }
    saveJson("inbox.json", inbox);
    saveJson("processed.json", processed);
    lastRun = Date.now();
    addLog("cycle", `収集サイクル完了（新規 ${added} 件、受信箱 ${inbox.length} 件）`);
  } catch (e) {
    addLog("error", e.message);
  } finally {
    running = false;
  }
}

function scheduleTimer() {
  if (timer) clearInterval(timer);
  timer = setInterval(() => cycle("定期実行"), Math.max(5, config.intervalMin) * 60000);
}

// ---- HTTPサーバー（127.0.0.1のみ・アプリ配信＋API） ----
const MIME = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css", ".png": "image/png", ".jpg": "image/jpeg", ".svg": "image/svg+xml", ".json": "application/json" };

function readBody(req) {
  return new Promise(resolve => {
    let buf = "";
    req.on("data", c => { buf += c; if (buf.length > 32 * 1024 * 1024) req.destroy(); });
    req.on("end", () => { try { resolve(JSON.parse(buf || "{}")); } catch (_) { resolve({}); } });
  });
}

function json(res, code, data) {
  res.writeHead(code, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*", // 127.0.0.1のみで待ち受けるため実質ローカル限定
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS"
  });
  res.end(JSON.stringify(data));
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://localhost");
  const p = url.pathname;

  if (req.method === "OPTIONS") { json(res, 204, {}); return; }

  // ---- API ----
  if (p === "/api/agent/status") {
    json(res, 200, {
      running: true, busy: running, lastRun, config,
      inboxCount: inbox.length, hasSnapshot: !!snapshot,
      nextRun: lastRun ? lastRun + config.intervalMin * 60000 : null
    });
    return;
  }
  if (p === "/api/agent/config") {
    if (req.method === "POST") {
      const body = await readBody(req);
      const allowed = ["enabled", "mailEnabled", "mailFolder", "mailDays", "links", "followLinks", "intervalMin", "autoSend"];
      for (const k of allowed) if (body[k] !== undefined) config[k] = body[k];
      config.mailDays = Math.max(1, parseInt(config.mailDays, 10) || 7);
      config.intervalMin = Math.max(5, parseInt(config.intervalMin, 10) || 30);
      config.links = (Array.isArray(config.links) ? config.links : []).map(s => String(s).trim()).filter(s => /^https?:\/\//.test(s));
      saveJson("config.json", config);
      scheduleTimer();
      addLog("config", "設定を更新しました");
    }
    json(res, 200, config);
    return;
  }
  if (p === "/api/agent/inbox") { json(res, 200, { items: inbox }); return; }
  if (p === "/api/agent/consume" && req.method === "POST") {
    const body = await readBody(req);
    const uids = new Set(body.uids || []);
    const before = inbox.length;
    inbox = inbox.filter(x => !uids.has(x.uid));
    saveJson("inbox.json", inbox);
    if (before !== inbox.length) addLog("inbox", `アプリが ${before - inbox.length} 件をタスクボードへ取り込みました`);
    json(res, 200, { remaining: inbox.length });
    return;
  }
  if (p === "/api/agent/state" && req.method === "POST") {
    const body = await readBody(req);
    if (Array.isArray(body.tasks)) {
      snapshot = { tasks: body.tasks, contacts: body.contacts || [], settings: body.settings || {}, at: Date.now() };
      saveJson("snapshot.json", snapshot);
    }
    json(res, 200, { ok: true });
    return;
  }
  if (p === "/api/agent/log") { json(res, 200, { log: log.slice(0, 50) }); return; }
  if (p === "/api/agent/run" && req.method === "POST") {
    cycle("手動実行"); // 完了を待たず即応答
    json(res, 200, { started: true });
    return;
  }

  // ---- 静的配信（browser/ 配下のみ） ----
  let file = p === "/" ? "/index.html" : decodeURIComponent(p);
  const full = path.join(BROWSER_DIR, file);
  if (!full.startsWith(BROWSER_DIR)) { res.writeHead(403); res.end(); return; }
  fs.readFile(full, (err, buf) => {
    if (err) { res.writeHead(404); res.end("Not Found"); return; }
    res.writeHead(200, { "Content-Type": MIME[path.extname(full).toLowerCase()] || "application/octet-stream" });
    res.end(buf);
  });
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`Todoリマインダー 常駐エージェント起動: http://127.0.0.1:${PORT}/`);
  addLog("start", `エージェント起動（間隔 ${config.intervalMin} 分、メールフォルダ「${config.mailFolder}」）`);
  scheduleTimer();
  setTimeout(() => cycle("起動時"), 5000);
});
