/* PROJEKTITA lauke – telefono / GPS valdiklio programėlė (PWA).
 *
 * Programėlė NIEKADA nerašo į PROJEKTITA duomenų bazę. Ji tik deda mažus, nepriklausomus įrašus
 * (failai + <id>.json manifestas, įrašomas PASKUTINIS) į OneDrive aplanką Telefono_inbox; kompiuteryje
 * PROJEKTITA sistemos „Telefono dėžutė“ juos peržiūri ir (po patvirtinimo) sudeda į projektą.
 * Projektų ir riboženklių taškų sąrašus programėlė skaito iš to paties aplanko (_projektai),
 * kurį atnaujina kompiuterio programa.
 *
 * Prisijungimas: standartinis OAuth 2.0 authorization code + PKCE (be išorinių bibliotekų).
 */
"use strict";
(() => {
const VERSION = "1.4.0";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = "openid profile offline_access User.Read Files.ReadWrite";
const DEFAULT_NOTE_CATS = ["Problema", "Pastaba", "Užduotis", "Į ką atkreipti dėmesį", "Kita"];
const DEFAULT_EXPENSES = [
  { name: "Omniva siunta", unit_price: 2.2, party: "Omniva" },
  { name: "Pašto išlaidos", unit_price: null, party: "AB Lietuvos paštas" },
  { name: "Parkavimo išlaidos", unit_price: null, party: "JUDU – parkavimas" },
  { name: "Transportas – kuras", unit_price: null, party: "" },
  { name: "Plastikinis riboženklis", unit_price: 1.57, party: "UAB Plastmasės fabrikas" },
  { name: "Kitos projekto išlaidos", unit_price: null, party: "" },
];
const BIG_FILE = 3_500_000;           // didesni failai keliami per upload session
const CHUNK = 327680 * 8;             // 2,5 MB (320 KiB kartotinis)

/* ------------------------------------------------------------------ pagalbinės */
const $ = (id) => document.getElementById(id);
const pad = (n) => String(n).padStart(2, "0");
function localISO(d = new Date()) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
const todayISO = () => localISO().slice(0, 10);
function slugPy(v, limit = 40) {   // tas pats, kas Python phone_inbox._slug
  let t = String(v == null ? "" : v).trim().replace(/[\\/:*?"<>|\r\n\t]+/g, " ");
  t = t.replace(/\s+/g, "_").replace(/^[._ ]+|[._ ]+$/g, "");
  return t.slice(0, limit) || "x";
}
const encPath = (p) => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");
const num = (v) => { const t = String(v == null ? "" : v).replace(",", ".").replace(/\s/g, ""); const n = parseFloat(t); return isFinite(n) ? n : null; };
const fmt2 = (n) => (n == null ? "" : (Math.round(n * 100) / 100).toFixed(2).replace(".", ","));
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
function b64url(buf) {
  let s = ""; const b = new Uint8Array(buf);
  for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function b64urlDecodeToString(t) {
  t = t.replace(/-/g, "+").replace(/_/g, "/"); while (t.length % 4) t += "=";
  const bin = atob(t); const bytes = Uint8Array.from(bin, (c) => c.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}
function randStr(n) { const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a.buffer).replace(/[-_]/g, "x").slice(0, n); }

/* ------------------------------------------------------------------ žurnalas */
function log(msg) {
  try {
    const arr = JSON.parse(localStorage.getItem("prj_log") || "[]");
    arr.push(`${localISO().replace("T", " ")}  ${msg}`);
    localStorage.setItem("prj_log", JSON.stringify(arr.slice(-60)));
  } catch (_) { /* ignoruojam */ }
}

/* ------------------------------------------------------------------ nustatymai */
const CFG0 = Object.assign({ clientId: "", tenantId: "", inboxPath: "GEO_programs/PROJEKTITA sistema/Telefono_inbox" }, window.PRJ_CONFIG || {});
function loadCfg() {
  let s = {};
  try { s = JSON.parse(localStorage.getItem("prj_settings") || "{}"); } catch (_) { /* tuščia */ }
  const c = Object.assign({}, CFG0, s);
  if (!c.device) { c.device = "tel-" + randStr(3).toLowerCase(); saveCfg(c); }
  c.recents = Array.isArray(c.recents) ? c.recents : [];
  return c;
}
function saveCfg(c) { try { localStorage.setItem("prj_settings", JSON.stringify(c)); } catch (_) { /* ignoruojam */ } }
let cfg = loadCfg();
const cfgReady = () => !!(cfg.clientId && cfg.tenantId);
const inboxPath = () => cfg.inboxPath.replace(/^\/+|\/+$/g, "");

/* ------------------------------------------------------------------ IndexedDB */
const DB = (() => {
  let dbp = null;
  function open() {
    if (dbp) return dbp;
    dbp = new Promise((res, rej) => {
      const r = indexedDB.open("prj_phone", 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore("queue", { keyPath: "id" });
        d.createObjectStore("cache", { keyPath: "key" });
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return dbp;
  }
  async function tx(store, mode, fn) {
    const d = await open();
    return new Promise((res, rej) => {
      const t = d.transaction(store, mode);
      const req = fn(t.objectStore(store));
      t.oncomplete = () => res(req && req.result !== undefined ? req.result : undefined);
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error);
    });
  }
  return {
    get: (s, k) => tx(s, "readonly", (st) => st.get(k)),
    put: (s, v) => tx(s, "readwrite", (st) => st.put(v)),
    del: (s, k) => tx(s, "readwrite", (st) => st.delete(k)),
    all: (s) => tx(s, "readonly", (st) => st.getAll()),
  };
})();

/* ------------------------------------------------------------------ prisijungimas (OAuth code + PKCE) */
const Auth = (() => {
  const KEY = "prj_auth";
  const PENDING = "prj_auth_pending";
  const redirectUri = () => location.origin + location.pathname;
  const authBase = () => `https://login.microsoftonline.com/${encodeURIComponent(cfg.tenantId)}/oauth2/v2.0`;
  const read = () => { try { return JSON.parse(localStorage.getItem(KEY) || "null"); } catch (_) { return null; } };
  const write = (v) => { try { v ? localStorage.setItem(KEY, JSON.stringify(v)) : localStorage.removeItem(KEY); } catch (_) { /* */ } };

  function authError(msg) { const e = new Error(msg); e.authRequired = true; return e; }
  function decodeClaims(idToken) {
    try { return JSON.parse(b64urlDecodeToString(idToken.split(".")[1])); } catch (_) { return {}; }
  }
  function storeTokens(j, prev) {
    const claims = j.id_token ? decodeClaims(j.id_token) : {};
    write({
      access_token: j.access_token,
      refresh_token: j.refresh_token || (prev && prev.refresh_token) || "",
      expires_at: Date.now() + (Number(j.expires_in || 3600) - 120) * 1000,
      name: claims.name || claims.preferred_username || (prev && prev.name) || "",
      user: claims.preferred_username || (prev && prev.user) || "",
    });
  }
  async function signIn() {
    if (!cfgReady()) throw new Error("Pirmiausia įveskite Client ID ir Tenant ID (Nustatymai).");
    const verifier = randStr(64);
    const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
    const state = randStr(24);
    localStorage.setItem(PENDING, JSON.stringify({ verifier, state, at: Date.now() }));
    const p = new URLSearchParams({
      client_id: cfg.clientId, response_type: "code", redirect_uri: redirectUri(),
      response_mode: "fragment", scope: SCOPES, state,
      code_challenge: challenge, code_challenge_method: "S256", prompt: "select_account",
    });
    log("Nukreipiama į Microsoft prisijungimą");
    location.assign(`${authBase()}/authorize?${p}`);
  }
  // Grįžus iš Microsoft: #code=...&state=...
  async function handleRedirect() {
    const h = location.hash.replace(/^#/, "");
    if (!h || !(h.includes("code=") || h.includes("error="))) return null;
    const q = new URLSearchParams(h);
    history.replaceState(null, "", location.pathname + location.search);
    if (q.get("error")) { const m = `${q.get("error")}: ${q.get("error_description") || ""}`; log("Prisijungimo klaida: " + m); throw new Error(m); }
    let pend = null;
    try { pend = JSON.parse(localStorage.getItem(PENDING) || "null"); } catch (_) { /* */ }
    localStorage.removeItem(PENDING);
    if (!pend || pend.state !== q.get("state")) throw new Error("Prisijungimo būsena nesutampa – bandykite dar kartą.");
    const r = await fetch(`${authBase()}/token`, {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: cfg.clientId, grant_type: "authorization_code", code: q.get("code"),
        redirect_uri: redirectUri(), code_verifier: pend.verifier, scope: SCOPES,
      }),
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) { const m = `${j.error || r.status}: ${j.error_description || ""}`; log("Token klaida: " + m); throw new Error(m); }
    storeTokens(j, null);
    log("Prisijungta: " + (read().user || ""));
    return read();
  }
  async function refresh(t) {
    let r;
    try {
      r = await fetch(`${authBase()}/token`, {
        method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ client_id: cfg.clientId, grant_type: "refresh_token", refresh_token: t.refresh_token, scope: SCOPES }),
      });
    } catch (e) { const er = new Error("Nėra ryšio"); er.network = true; throw er; }
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      log(`Atnaujinimo klaida: ${j.error || r.status}`);
      if (j.error === "invalid_grant" || j.error === "interaction_required" || r.status === 400 || r.status === 401) write(null);
      throw authError("Reikia iš naujo prisijungti");
    }
    storeTokens(j, t);
    return read();
  }
  async function getToken() {
    if (!cfgReady()) throw authError("Nenustatyta Microsoft programa");
    const t = read();
    if (!t || !t.access_token) throw authError("Neprisijungta");
    if (t.expires_at > Date.now()) return t.access_token;
    if (!t.refresh_token) { write(null); throw authError("Reikia iš naujo prisijungti"); }
    return (await refresh(t)).access_token;
  }
  function invalidate() { const t = read(); if (t) { t.expires_at = 0; write(t); } }
  return { signIn, handleRedirect, getToken, read, invalidate, signOut: () => write(null), signedIn: () => !!(read() && read().refresh_token) };
})();

/* ------------------------------------------------------------------ Graph */
function httpErr(r, body) {
  const e = new Error(`HTTP ${r.status}${body ? " – " + String(body).slice(0, 160) : ""}`);
  e.status = r.status; return e;
}
async function graphFetch(url, init = {}, opts = {}) {
  let attempt = 0;
  for (;;) {
    const tok = await Auth.getToken();
    let r;
    try {
      r = await fetch(url, Object.assign({}, init, { headers: Object.assign({}, init.headers || {}, { Authorization: "Bearer " + tok }) }));
    } catch (e) { if (opts.noNetErr) throw e; const er = new Error("Nėra ryšio"); er.network = true; throw er; }
    if (r.status === 401 && attempt === 0) { Auth.invalidate(); attempt++; continue; }
    if ((r.status === 429 || r.status === 503) && attempt < 3) {
      const wait = Math.min(15, Number(r.headers.get("Retry-After")) || 2 * (attempt + 1));
      log(`Graph ${r.status}, laukiama ${wait}s`); await new Promise((s) => setTimeout(s, wait * 1000)); attempt++; continue;
    }
    return r;
  }
}
async function graphPut(rel, blob, type) {
  const base = `${GRAPH}/me/drive/root:/${encPath(rel)}`;
  if (blob.size >= BIG_FILE) {
    const r = await graphFetch(`${base}:/createUploadSession`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }),
    });
    if (!r.ok) throw httpErr(r, await r.text());
    const { uploadUrl } = await r.json();
    for (let off = 0; off < blob.size; off += CHUNK) {
      const end = Math.min(off + CHUNK, blob.size);
      let rr;
      try {
        rr = await fetch(uploadUrl, { method: "PUT", headers: { "Content-Range": `bytes ${off}-${end - 1}/${blob.size}` }, body: blob.slice(off, end) });
      } catch (e) { const er = new Error("Nėra ryšio"); er.network = true; throw er; }
      if (!rr.ok && rr.status !== 202) throw httpErr(rr, await rr.text());
    }
    return;
  }
  const r = await graphFetch(`${base}:/content?@microsoft.graph.conflictBehavior=replace`, {
    method: "PUT", headers: { "Content-Type": type || "application/octet-stream" }, body: blob,
  });
  if (!r.ok) throw httpErr(r, await r.text());
}
async function graphGetJson(rel) {
  const url = `${GRAPH}/me/drive/root:/${encPath(rel)}:/content`;
  let r;
  try { r = await graphFetch(url, { method: "GET" }, { noNetErr: true }); }
  catch (e) {
    if (e.authRequired) throw e;
    // kai kuriose naršyklėse peradresavimas į atsisiuntimo nuorodą blokuojamas CORS – bandome per downloadUrl
    log("GET /content nepavyko (" + e.message + "), bandoma per downloadUrl");
    const m = await graphFetch(`${GRAPH}/me/drive/root:/${encPath(rel)}?select=id,@microsoft.graph.downloadUrl`, { method: "GET" });
    if (m.status === 404) return null;
    if (!m.ok) throw httpErr(m, await m.text());
    const meta = await m.json();
    const dl = meta["@microsoft.graph.downloadUrl"];
    if (!dl) throw new Error("Nėra atsisiuntimo nuorodos");
    const d = await fetch(dl);
    if (!d.ok) throw httpErr(d);
    return await d.json();
  }
  if (r.status === 404) return null;
  if (!r.ok) throw httpErr(r, await r.text());
  return await r.json();
}

/* ------------------------------------------------------------------ nuotraukos */
async function prepareImage(file, maxDim = 2000, quality = 0.85) {
  try {
    let bmp;
    if (window.createImageBitmap) bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
    else {
      bmp = await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = URL.createObjectURL(file); });
    }
    const w = bmp.width, h = bmp.height; const k = Math.min(1, maxDim / Math.max(w, h));
    const cw = Math.max(1, Math.round(w * k)), ch = Math.max(1, Math.round(h * k));
    const cv = document.createElement("canvas"); cv.width = cw; cv.height = ch;
    cv.getContext("2d").drawImage(bmp, 0, 0, cw, ch);
    if (bmp.close) bmp.close();
    const blob = await new Promise((res) => cv.toBlob(res, "image/jpeg", quality));
    if (blob && blob.size > 0) return blob;
  } catch (e) { log("Nuotraukos apdorojimas nepavyko: " + e.message); }
  return file;
}

/* ------------------------------------------------------------------ eilė */
function newId() {
  const d = new Date();
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}T${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}_${slugPy(cfg.device, 12)}_${randStr(4).toLowerCase()}`;
}
async function enqueue(type, fields, files, summary, extra = {}) {
  const id = newId();
  const names = {}; const fileList = [];
  files.forEach((f, i) => {
    const orig = f.origName || null;
    const ext = (orig && /\.[A-Za-z0-9]{1,5}$/.test(orig)) ? orig.slice(orig.lastIndexOf(".")).toLowerCase() : ".jpg";
    const name = `${id}_${i + 1}${ext}`;
    if (orig) names[name] = orig;
    fileList.push({ name, blob: f.blob, type: f.blob.type || (ext === ".jpg" ? "image/jpeg" : "application/octet-stream") });
  });
  const manifest = Object.assign({ version: 1, id, type, device: cfg.device, created: localISO() }, fields, { files: fileList.map((f) => f.name) });
  if (Object.keys(names).length) manifest.names = names;
  const item = Object.assign({ id, created: manifest.created, type, summary, project_no: fields.project_no || "", manifest, files: fileList, state: "pending", error: "", attempts: 0, sentAt: "" }, extra);
  try { await DB.put("queue", item); }
  catch (e) { delete item.handles; log("Failų nuorodų išsaugoti nepavyko: " + e.message); await DB.put("queue", item); }
  log(`Į eilę: ${type} ${summary}`);
  updateBadge();
  flush();
  return item;
}
let flushing = false;
async function sendItem(it) {
  it.state = "sending"; it.error = ""; await DB.put("queue", it);
  const base = inboxPath();
  for (const f of it.files) await graphPut(`${base}/${f.name}`, f.blob, f.type);
  const mb = new Blob([JSON.stringify(it.manifest, null, 1)], { type: "application/json" });
  await graphPut(`${base}/${it.id}.json`, mb, "application/json");   // manifestas PASKUTINIS
  it.askDelete = !!(it.handles && it.handles.length);
  it.state = "sent"; it.sentAt = localISO(); it.files = it.files.map((f) => ({ name: f.name })); it.error = "";
  await DB.put("queue", it);
  log("Išsiųsta: " + it.id);
}
async function flush() {
  if (flushing || !navigator.onLine || !cfgReady() || !Auth.signedIn()) { updateBadge(); return; }
  flushing = true; updateBadge();
  try {
    const tried = new Set(); let stop = false;
    while (!stop) {   // siuntimo metu įdėti nauji įrašai paimami tame pačiame cikle
      const items = (await DB.all("queue")).filter((i) => i.state !== "sent" && !tried.has(i.id)).sort((a, b) => a.created < b.created ? -1 : 1);
      if (!items.length) break;
      const it = items[0]; tried.add(it.id);
      try { await sendItem(it); }
      catch (e) {
        it.attempts = (it.attempts || 0) + 1;
        if (e.authRequired) { it.state = "pending"; await DB.put("queue", it); log("Reikia prisijungti"); stop = true; }
        else if (e.network) { it.state = "pending"; await DB.put("queue", it); log("Nėra ryšio"); stop = true; }
        else { it.state = "error"; it.error = e.message; await DB.put("queue", it); log(`Klaida (${it.id}): ${e.message}`); }
      }
      updateBadge();
      if ($("v-queue") && !$("v-queue").hidden) renderQueue();
    }
  } finally { flushing = false; updateBadge(); if (!$("v-queue").hidden) renderQueue(); maybeAskDelete(); }
}
/* Po sėkmingo išsiuntimo – klausiame, ar ištrinti originalius failus telefone (jei failai pasirinkti per sistemos failų langą). */
let delAsking = false;
async function maybeAskDelete() {
  if (delAsking || document.visibilityState !== "visible") return;
  try {
    const it = (await DB.all("queue")).find((i) => i.state === "sent" && i.askDelete && i.handles && i.handles.length);
    if (it) askDeleteFiles(it);
  } catch (e) { log("Trynimo klausimas: " + e.message); }
}
function askDeleteFiles(it) {
  if (delAsking) return; delAsking = true;
  const box = document.createElement("div"); box.className = "sheet";
  box.innerHTML = `<div class="sheet-card"><div class="sheet-head"><strong>Failai išsiųsti</strong></div>` +
    `<p>Ištrinti šiuos failus iš telefono?</p><ul class="list">${it.handles.map((h) => `<li><div class="t">${esc(h.name)}</div></li>`).join("")}</ul>` +
    `<div class="row"><button type="button" class="btn primary" data-y>Taip, ištrinti</button><button type="button" class="btn ghost" data-n>Ne, palikti</button></div>` +
    `<p class="muted small" data-m></p></div>`;
  document.body.appendChild(box);
  const close = () => { box.remove(); delAsking = false; };
  box.querySelector("[data-n]").onclick = async () => { it.askDelete = false; await DB.put("queue", it); close(); if (!$("v-queue").hidden) renderQueue(); };
  box.querySelector("[data-y]").onclick = async () => {
    box.querySelector("[data-y]").disabled = true;
    const left = []; let ok = 0; const errs = [];
    for (const h of it.handles) {
      try {
        if (h.requestPermission) { const r = await h.requestPermission({ mode: "readwrite" }); if (r !== "granted") throw new Error("nėra leidimo"); }
        if (typeof h.remove !== "function") throw new Error("naršyklė netrina failų");
        await h.remove(); ok++;
      } catch (e) { left.push(h); errs.push(`${h.name}: ${e.message}`); log(`Trynimas ${h.name}: ${e.message}`); }
    }
    it.handles = left; it.askDelete = false; await DB.put("queue", it);
    close();
    if (!left.length) toast(`Ištrinta failų: ${ok}`);
    else toast(`Ištrinta ${ok}, nepavyko ${left.length} (${errs[0]}). Likusius ištrinkite per „Failai“.`, true);
    if (!$("v-queue").hidden) renderQueue();
  };
}
async function purgeOldSent() {
  const lim = Date.now() - 7 * 86400000;
  for (const it of await DB.all("queue")) if (it.state === "sent" && Date.parse(it.sentAt) < lim) await DB.del("queue", it.id);
}

/* ------------------------------------------------------------------ duomenys iš kompiuterio */
async function cacheGet(key) { const r = await DB.get("cache", key); return r ? r.value : null; }
async function cachePut(key, value) { await DB.put("cache", { key, value, at: localISO() }); }
async function refreshProjects(silent) {
  try {
    const j = await graphGetJson(`${inboxPath()}/_projektai/projektai.json`);
    if (!j) { if (!silent) toast("Kompiuteris dar nepaskelbė projektų sąrašo (atidarykite „Telefono dėžutė“ PROJEKTITA programoje).", true); return null; }
    await cachePut("projects", j); log(`Projektų sąrašas atnaujintas (${(j.projects || []).length})`);
    if (!silent) toast("Projektų sąrašas atnaujintas");
    return j;
  } catch (e) {
    log("Projektų sąrašas: " + e.message);
    if (!silent && !e.authRequired) toast("Nepavyko atnaujinti: " + e.message, true);
    if (!silent && e.authRequired) toast("Pirmiausia prisijunkite (Nustatymai)", true);
    return null;
  }
}
async function refreshPoints(prj, silent) {
  try {
    const j = await graphGetJson(`${inboxPath()}/_projektai/${slugPy(prj, 30)}_taskai.json`);
    if (!j) { if (!silent) toast("Šio projekto taškų sąrašo nėra – kompiuteryje atidarykite KADMAT taškus.", true); return null; }
    await cachePut("points:" + prj, j); log(`Taškai ${prj}: ${(j.points || []).length}`);
    return j;
  } catch (e) { log("Taškai: " + e.message); if (!silent) toast("Nepavyko gauti taškų: " + e.message, true); return null; }
}

/* ------------------------------------------------------------------ UI pagrindas */
const views = ["home", "note", "marker", "expense", "files", "contact", "queue", "settings"];
const stack = [];
function show(name, push = true) {
  views.forEach((v) => { $("v-" + v).hidden = v !== name; });
  $("title").textContent = $("v-" + name).dataset.title;
  if (push) stack.push(name);
  $("btn-back").hidden = name === "home";
  window.scrollTo(0, 0);
  const cur = name;
  if (cur === "queue") renderQueue();
  if (cur === "settings") renderSettings();
  if (cur === "home") renderHome();
}
function back() { stack.pop(); const prev = stack[stack.length - 1] || "home"; show(prev, false); }
let toastTimer = null;
function toast(msg, bad) {
  const t = $("toast"); t.textContent = msg; t.className = "toast" + (bad ? " bad" : ""); t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, bad ? 5000 : 2600);
}
async function updateBadge() {
  const p = $("pill");
  let items = []; try { items = await DB.all("queue"); } catch (_) { /* */ }
  const open = items.filter((i) => i.state !== "sent");
  const err = open.filter((i) => i.state === "error").length;
  let cls = "ok", txt = "Viskas išsiųsta";
  if (!navigator.onLine) { cls = "warn"; txt = open.length ? `Be ryšio · eilėje ${open.length}` : "Be ryšio"; }
  else if (cfgReady() && !Auth.signedIn()) { cls = "warn"; txt = open.length ? `Prisijunkite · eilėje ${open.length}` : "Prisijunkite"; }
  else if (!cfgReady()) { cls = "warn"; txt = "Reikia nustatymų"; }
  else if (err) { cls = "bad"; txt = `Klaida · ${err}`; }
  else if (flushing) { cls = "warn"; txt = "Siunčiama…"; }
  else if (open.length) { cls = "warn"; txt = `Eilėje ${open.length}`; }
  p.className = "pill " + cls; p.textContent = txt;
}
function renderHome() {
  const b = $("home-banner");
  if (!cfgReady()) { b.hidden = false; b.className = "banner"; b.textContent = "Pirmas kartas: atidarykite „Nustatymai“ ir įveskite Microsoft programos duomenis. Fotografuoti ir rašyti galite jau dabar – įrašai lauks eilėje."; }
  else if (!Auth.signedIn()) { b.hidden = false; b.className = "banner"; b.textContent = "Neprisijungta – įrašai kaupiami eilėje ir bus išsiųsti prisijungus (Nustatymai → Prisijungti)."; }
  else b.hidden = true;
  DB.all("queue").then((items) => {
    const open = items.filter((i) => i.state !== "sent").length;
    $("home-queue-title").textContent = open ? `Siuntimo eilė · ${open}` : "Siuntimo eilė";
    $("home-queue-sub").textContent = open ? "yra neišsiųstų įrašų" : "viskas išsiųsta";
  });
  $("home-ver").textContent = `v${VERSION} · ${cfg.device}`;
}

/* ------------------------------------------------------------------ projekto parinkimas */
let pickerCb = null, pickerNav = false;
async function openPicker(cb, { optional = false, title = "Projektas", nav = false } = {}) {
  pickerCb = cb; pickerNav = nav;
  $("picker-title").textContent = title;
  $("picker").hidden = false;
  $("picker-q").value = ""; $("picker-manual").value = "";
  const data = await cacheGet("projects");
  $("picker-note").textContent = data ? "" : "Projektų sąrašo dar nėra – jį galima gauti prisijungus (Nustatymai → Atnaujinti projektų sąrašą) arba įrašyti PRJ numerį ranka.";
  renderPicker();
  if (!data && cfgReady() && Auth.signedIn()) { const j = await refreshProjects(true); if (j) { $("picker-note").textContent = ""; renderPicker(); } }
  if (optional && !$("picker-none")) {
    const li = document.createElement("li"); li.id = "picker-none"; li.innerHTML = '<div class="t">Be projekto</div>';
    li.onclick = () => closePicker(""); $("picker-list").prepend(li);
  }
}
function closePicker(val) { $("picker").hidden = true; const cb = pickerCb; pickerCb = null; if (cb && val !== undefined) cb(val); }
async function renderPicker() {
  const data = await cacheGet("projects");
  const q = $("picker-q").value.trim().toLowerCase();
  const showDone = $("picker-completed").checked;
  const list = $("picker-list"); list.innerHTML = "";
  let projects = (data && data.projects) || [];
  projects = projects.filter((p) => showDone || !p.completed);
  if (q) projects = projects.filter((p) => `${p.project_no} ${p.client} ${p.place} ${p.cad_no}`.toLowerCase().includes(q));
  const rec = cfg.recents;
  projects = projects.slice().sort((a, b) => {
    const ra = rec.indexOf(a.project_no), rb = rec.indexOf(b.project_no);
    if (ra !== rb) return (ra < 0 ? 99 : ra) - (rb < 0 ? 99 : rb);
    return 0;
  });
  for (const p of projects.slice(0, 80)) {
    const li = document.createElement("li");
    li.innerHTML = `<div class="t">${pickerNav ? "🚗 " : ""}${esc(p.project_no)}${p.type ? " · " + esc(p.type) : ""}${p.completed ? " · užbaigtas" : ""}${pickerNav && !hasLoc(p) ? " · vieta nenustatyta" : ""}</div>` +
      `<div class="s">${esc([p.client, p.place].filter(Boolean).join(" · "))}</div>`;
    if (pickerNav && !hasLoc(p)) li.style.opacity = ".55";
    li.onclick = () => closePicker(p.project_no);
    list.appendChild(li);
  }
  if (!projects.length && data) { const li = document.createElement("li"); li.className = "muted"; li.textContent = "Nieko nerasta."; list.appendChild(li); }
}
function hasLoc(p) { return p && typeof p.lat === "number" && typeof p.lon === "number"; }
async function navigateToProject(prj) {
  const data = await cacheGet("projects");
  const p = ((data && data.projects) || []).find((x) => x.project_no === String(prj).toUpperCase());
  if (!p) { toast("Projekto nėra sąraše – atnaujinkite projektų sąrašą (Nustatymai).", true); return; }
  if (!hasLoc(p)) { toast("Šio projekto vieta dar nenustatyta (nėra objektai.shp taško).", true); return; }
  window.location.href = `https://www.google.com/maps/dir/?api=1&destination=${p.lat.toFixed(7)},${p.lon.toFixed(7)}&travelmode=driving`;
}
function rememberProject(prj) {
  if (!prj) return;
  cfg.recents = [prj].concat(cfg.recents.filter((x) => x !== prj)).slice(0, 8); saveCfg(cfg);
}
function setPick(btn, prj, emptyText) {
  btn.textContent = prj || emptyText; btn.classList.toggle("empty", !prj); btn.classList.toggle("filled", !!prj);
}

/* ------------------------------------------------------------------ nuotraukų įvedimas */
let camTarget = null;   // callback(files[])
function takePhotos(cb, { camera = true, multiple = false } = {}) {
  camTarget = cb;
  const inp = camera ? $("in-cam") : $("in-gal");
  inp.value = ""; inp.click();
}
async function onPhotoInput(ev) {
  const files = Array.from(ev.target.files || []); ev.target.value = "";
  const cb = camTarget; camTarget = null;
  if (!files.length || !cb) return;
  const out = [];
  for (const f of files) out.push({ blob: await prepareImage(f), preview: null });
  cb(out);
}
function renderThumbs(box, arr, onRemove) {
  box.innerHTML = "";
  arr.forEach((it, i) => {
    if (!it.url) it.url = URL.createObjectURL(it.blob);
    const d = document.createElement("div"); d.className = "thumb";
    d.innerHTML = `<img alt="" src="${it.url}"><button type="button" aria-label="Pašalinti">×</button>`;
    d.querySelector("button").onclick = () => { URL.revokeObjectURL(it.url); arr.splice(i, 1); onRemove(); };
    box.appendChild(d);
  });
}

/* ------------------------------------------------------------------ PASTABA */
const note = { project: "", photos: [] };
function initNote() {
  note.project = ""; note.photos = [];
  $("note-title").value = ""; $("note-body").value = "";
  setPick($("note-project"), "", "Pasirinkite projektą…");
  renderThumbs($("note-thumbs"), note.photos, () => renderThumbs($("note-thumbs"), note.photos, () => {}));
  cacheGet("projects").then((d) => {
    const cats = (d && d.note_categories && d.note_categories.length) ? d.note_categories : DEFAULT_NOTE_CATS;
    const sel = $("note-cat"); sel.innerHTML = cats.map((c) => `<option>${esc(c)}</option>`).join("");
    sel.value = cats.includes("Pastaba") ? "Pastaba" : cats[0];
  });
}
function bindNote() {
  $("note-project").onclick = () => openPicker((p) => { note.project = p; rememberProject(p); setPick($("note-project"), p, "Pasirinkite projektą…"); });
  const addPhotos = (arr) => { note.photos.push(...arr); renderThumbs($("note-thumbs"), note.photos, () => renderThumbs($("note-thumbs"), note.photos, () => {})); };
  $("note-cam").onclick = () => takePhotos(addPhotos, { camera: true });
  $("note-gal").onclick = () => takePhotos(addPhotos, { camera: false, multiple: true });
  $("note-save").onclick = async () => {
    if (!note.project) return toast("Pasirinkite projektą", true);
    const title = $("note-title").value.trim(), body = $("note-body").value.trim();
    if (!title && !body && !note.photos.length) return toast("Įrašykite temą, aprašymą arba pridėkite nuotrauką", true);
    await enqueue("note", { project_no: note.project, category: $("note-cat").value, title, body, note_date: todayISO() },
      note.photos.map((p) => ({ blob: p.blob })), `${note.project}: ${title || body.slice(0, 40) || "nuotrauka"}${note.photos.length ? ` (+${note.photos.length} foto)` : ""}`);
    toast("Pastaba įrašyta į eilę"); stack.length = 0; show("home");
  };
}

/* ------------------------------------------------------------------ RIBOŽYMIO FOTO */
const marker = { project: "", data: null };
async function localPointCounts(prj, since) {
  const out = {};
  for (const it of await DB.all("queue")) {
    if (it.type !== "marker_photo" || it.project_no !== prj) continue;
    if (since && it.created <= since) continue;
    const nr = it.manifest.point_name; out[nr] = (out[nr] || 0) + 1;
  }
  return out;
}
async function renderMarker() {
  const body = $("marker-body"); body.hidden = !marker.project;
  if (!marker.project) return;
  const d = marker.data; const box = $("marker-chips"); box.innerHTML = "";
  const flt = $("marker-filter").value.trim().toLowerCase();
  const pts = (d && d.points) || [];
  const local = await localPointCounts(marker.project, d && d.updated);
  let withPhoto = 0;
  const sorted = pts.slice().sort((a, b) => { const x = parseFloat(a.nr), y = parseFloat(b.nr); return (isNaN(x) || isNaN(y)) ? String(a.nr).localeCompare(String(b.nr)) : x - y; });
  for (const p of sorted) {
    const cnt = (Number(p.photos) || 0) + (local[p.nr] || 0);
    if (cnt) withPhoto++;
    if (flt && !(`${p.nr} ${p.kodas}`.toLowerCase().includes(flt))) continue;
    const b = document.createElement("button"); b.type = "button"; b.className = "chip" + (cnt ? " done" : "");
    b.innerHTML = `<b>${esc(p.nr)}</b><small>${esc(p.kodas || "")}</small>${cnt ? `<span class="n">✓ ${cnt}</span>` : ""}`;
    b.onclick = () => shootPoint(p.nr, p.kodas || "");
    box.appendChild(b);
  }
  $("marker-info").textContent = pts.length ? `${marker.project} · taškų ${pts.length} · su nuotr. ${withPhoto}${d && d.updated ? " · sąrašas " + d.updated.slice(0, 16).replace("T", " ") : ""}` : marker.project;
  $("marker-manual").hidden = !!pts.length;
  $("marker-manual-msg").textContent = pts.length ? "" : "Šiam projektui taškų sąrašo nėra (kompiuteryje nenuskaitytas KADMAT DWG). Taško Nr. galite įrašyti ranka.";
  if (pts.length) { $("marker-manual").hidden = false; $("marker-manual-msg").textContent = "Taško nėra sąraše? Įrašykite Nr. ranka:"; }
}
async function loadMarkerProject(prj) {
  marker.project = prj; rememberProject(prj); setPick($("marker-project"), prj, "Pasirinkite projektą…");
  marker.data = await cacheGet("points:" + prj); renderMarker();
  if (cfgReady() && Auth.signedIn() && navigator.onLine) { const j = await refreshPoints(prj, true); if (j) { marker.data = j; renderMarker(); } }
}
function shootPoint(nr, kodas) {
  if (!marker.project) return toast("Pasirinkite projektą", true);
  takePhotos(async (arr) => {
    for (const a of arr) {
      await enqueue("marker_photo", { project_no: marker.project, point_name: String(nr), code: kodas, taken: todayISO() }, [{ blob: a.blob }], `${marker.project}: taškas ${nr}${kodas ? " (" + kodas + ")" : ""}`);
    }
    toast(`Taško ${nr} nuotrauka įrašyta į eilę`); renderMarker();
  }, { camera: true });
}
function bindMarker() {
  $("marker-project").onclick = () => openPicker(loadMarkerProject, { title: "Projektas (KADMAT)" });
  $("marker-filter").oninput = renderMarker;
  $("marker-refresh").onclick = async () => { if (!marker.project) return; const j = await refreshPoints(marker.project, false); if (j) { marker.data = j; renderMarker(); toast("Taškai atnaujinti"); } };
  $("marker-nr-go").onclick = () => { const nr = $("marker-nr").value.trim(); if (!nr) return toast("Įrašykite taško Nr.", true); shootPoint(nr, ""); $("marker-nr").value = ""; };
}
function initMarker() { marker.project = ""; marker.data = null; setPick($("marker-project"), "", "Pasirinkite projektą…"); $("marker-body").hidden = true; $("marker-filter").value = ""; }

/* ------------------------------------------------------------------ IŠLAIDA */
const exp = { project: "", photos: [], cats: DEFAULT_EXPENSES };
function recalcExpense(src) {
  const q = num($("exp-qty").value) || 1, u = num($("exp-unit").value), a = num($("exp-amount").value);
  if (src === "amount") { if (a != null) $("exp-unit").value = fmt2(a / q); }
  else if (u != null) $("exp-amount").value = fmt2(q * u);
}
function onExpenseCat() {
  const c = exp.cats.find((x) => x.name === $("exp-cat").value);
  if (!c) return;
  if (c.unit_price != null) { $("exp-unit").value = fmt2(c.unit_price); recalcExpense("unit"); }
  if (c.party && !$("exp-party").value.trim()) $("exp-party").value = c.party;
}
function initExpense() {
  exp.project = ""; exp.photos = [];
  setPick($("exp-project"), "", "Be projekto");
  ["exp-unit", "exp-amount", "exp-party", "exp-doc", "exp-comment"].forEach((i) => { $(i).value = ""; });
  $("exp-qty").value = "1"; $("exp-date").value = todayISO();
  renderThumbs($("exp-thumbs"), exp.photos, () => {});
  cacheGet("projects").then((d) => {
    exp.cats = (d && d.expense_categories && d.expense_categories.length) ? d.expense_categories : DEFAULT_EXPENSES;
    const sel = $("exp-cat"); sel.innerHTML = exp.cats.map((c) => `<option>${esc(c.name)}</option>`).join("");
    sel.value = exp.cats.find((c) => c.name === "Omniva siunta") ? "Omniva siunta" : exp.cats[0].name;
    $("exp-unit").value = ""; $("exp-amount").value = ""; $("exp-party").value = ""; onExpenseCat();
  });
}
function bindExpense() {
  $("exp-project").onclick = () => openPicker((p) => { exp.project = p; if (p) rememberProject(p); setPick($("exp-project"), p, "Be projekto"); }, { optional: true });
  $("exp-cat").onchange = () => { $("exp-party").value = ""; $("exp-unit").value = ""; $("exp-amount").value = ""; onExpenseCat(); };
  $("exp-qty").oninput = () => recalcExpense("unit");
  $("exp-unit").oninput = () => recalcExpense("unit");
  $("exp-amount").oninput = () => recalcExpense("amount");
  const addPhotos = (arr) => { exp.photos.push(...arr); renderThumbs($("exp-thumbs"), exp.photos, () => renderThumbs($("exp-thumbs"), exp.photos, () => {})); };
  $("exp-cam").onclick = () => takePhotos(addPhotos, { camera: true });
  $("exp-gal").onclick = () => takePhotos(addPhotos, { camera: false });
  $("exp-save").onclick = async () => {
    const amount = num($("exp-amount").value);
    if (amount == null) return toast("Įrašykite sumą (arba kiekį ir vnt. kainą)", true);
    const qty = num($("exp-qty").value) || 1; const unit = num($("exp-unit").value);
    const fields = {
      project_no: exp.project, date: $("exp-date").value || todayISO(), category: $("exp-cat").value,
      counterparty: $("exp-party").value.trim(), quantity: qty, unit_price: unit != null ? unit : amount / qty, amount,
      document_no: $("exp-doc").value.trim(), comment: $("exp-comment").value.trim(),
    };
    await enqueue("expense", fields, exp.photos.map((p) => ({ blob: p.blob })), `${fields.category} ${fmt2(amount)} €${exp.project ? " · " + exp.project : ""}`);
    toast("Išlaida įrašyta į eilę"); stack.length = 0; show("home");
  };
}

/* ------------------------------------------------------------------ KONTAKTAS */
const ct = { project: "" };
function initContact() {
  ct.project = ""; setPick($("ct-project"), "", "Pasirinkite projektą…");
  ["ct-name", "ct-role", "ct-phone", "ct-email", "ct-notes"].forEach((i) => { $(i).value = ""; });
  $("ct-pick").hidden = !(navigator.contacts && navigator.contacts.select);
}
function bindContact() {
  $("ct-project").onclick = () => openPicker((p) => { ct.project = p || ""; if (p) rememberProject(p); setPick($("ct-project"), ct.project, "Pasirinkite projektą…"); });
  $("ct-pick").onclick = async () => {
    try {
      let props = ["name", "tel", "email"];
      try { const sup = await navigator.contacts.getProperties(); props = props.filter((x) => sup.includes(x)); } catch (_) { /* paliekame visus */ }
      const r = await navigator.contacts.select(props, { multiple: false });
      if (!r || !r.length) return;
      const c = r[0];
      if (c.name && c.name[0]) $("ct-name").value = c.name[0];
      if (c.tel && c.tel[0]) $("ct-phone").value = String(c.tel[0]).replace(/\s+/g, " ").trim();
      if (c.email && c.email[0]) $("ct-email").value = c.email[0];
    } catch (e) { log("Kontaktų pasirinkimas: " + e.name + " " + e.message); toast(`Nepavyko atidaryti telefono kontaktų (${e.name || "klaida"}). Įrašykite ranka.`, true); }
  };
  $("ct-save").onclick = async () => {
    const name = $("ct-name").value.trim();
    if (!name) return toast("Įrašykite vardą", true);
    if (!ct.project) return toast("Pasirinkite projektą", true);
    const fields = { project_no: ct.project, name, role: $("ct-role").value.trim(), phone: $("ct-phone").value.trim(), email: $("ct-email").value.trim(), notes: $("ct-notes").value.trim() };
    await enqueue("contact", fields, [], `${ct.project}: ${name}${fields.role ? " (" + fields.role + ")" : ""}${fields.phone ? " " + fields.phone : ""}`);
    toast("Kontaktas įrašytas į eilę"); stack.length = 0; show("home");
  };
}

/* ------------------------------------------------------------------ MATAVIMŲ FAILAI */
const mf = { project: "", files: [] };
function renderFiles() {
  const ul = $("files-list"); ul.innerHTML = "";
  mf.files.forEach((f, i) => {
    const li = document.createElement("li");
    li.innerHTML = `<div class="t">${esc(f.name)}</div><div class="s">${(f.size / 1024).toFixed(1)} KB</div><div class="acts"><button type="button" class="btn small ghost">Pašalinti</button></div>`;
    li.querySelector("button").onclick = () => { mf.files.splice(i, 1); renderFiles(); };
    ul.appendChild(li);
  });
}
function initFiles() { mf.project = ""; mf.files = []; setPick($("files-project"), "", "Pasirinkite projektą…"); renderFiles(); }
function bindFiles() {
  $("files-project").onclick = () => openPicker((p) => { mf.project = p; rememberProject(p); setPick($("files-project"), p, "Pasirinkite projektą…"); });
  const afterAdd = () => {
    renderFiles();
    // jei failo vardas yra PRJ numeris ir projektas dar nepasirinktas – pasiūlome
    if (!mf.project) {
      const m = mf.files.map((f) => f.name.replace(/\.[^.]+$/, "").toUpperCase()).find((n) => /^PRJ\d+(-\d+)?$/.test(n));
      if (m) { mf.project = m; setPick($("files-project"), m, "Pasirinkite projektą…"); }
    }
  };
  $("files-pick").onclick = async () => {
    if (window.showOpenFilePicker) {   // leidžia vėliau paklausti, ar ištrinti originalus telefone
      try {
        const hs = await window.showOpenFilePicker({ multiple: true, types: [{ description: "SurPad failai", accept: { "text/csv": [".csv"], "text/html": [".htm", ".html"] } }] });
        for (const h of hs) { const f = await h.getFile(); f._h = h; mf.files.push(f); }
        afterAdd(); return;
      } catch (e) {
        if (e && e.name === "AbortError") return;
        log("Failų langas: " + (e && e.message) + " – naudojamas įprastas");
      }
    }
    $("in-files").value = ""; $("in-files").click();
  };
  $("in-files").onchange = (ev) => {
    for (const f of Array.from(ev.target.files || [])) mf.files.push(f);
    ev.target.value = ""; afterAdd();
  };
  $("files-save").onclick = async () => {
    if (!mf.project) return toast("Pasirinkite projektą", true);
    if (!mf.files.length) return toast("Pasirinkite bent vieną failą", true);
    await enqueue("measure_files", { project_no: mf.project },
      mf.files.map((f) => ({ blob: f, origName: f.name })), `${mf.project}: ${mf.files.map((f) => f.name).join(", ")}`,
      { handles: mf.files.map((f) => f._h).filter(Boolean) });
    toast("Failai įrašyti į eilę"); stack.length = 0; show("home");
  };
}

/* ------------------------------------------------------------------ EILĖ */
const STATE_TXT = { pending: "Laukia", sending: "Siunčiama…", sent: "Išsiųsta", error: "Klaida" };
const TYPE_TXT = { note: "Pastaba", expense: "Išlaida", measure_files: "Matavimų failai", marker_photo: "Riboženklio foto", contact: "Kontaktas" };
async function renderQueue() {
  const ul = $("queue-list"); ul.innerHTML = "";
  const items = (await DB.all("queue")).sort((a, b) => a.created < b.created ? 1 : -1);
  $("queue-msg").textContent = !navigator.onLine ? "Be ryšio – įrašai bus išsiųsti, kai ryšys atsiras."
    : (!cfgReady() ? "Pirmiausia užpildykite Nustatymus." : (!Auth.signedIn() ? "Neprisijungta – paspauskite Nustatymai → Prisijungti." : ""));
  if (!items.length) { const li = document.createElement("li"); li.className = "muted"; li.textContent = "Eilė tuščia."; ul.appendChild(li); return; }
  for (const it of items) {
    const li = document.createElement("li");
    li.innerHTML = `<div class="t">${esc(TYPE_TXT[it.type] || it.type)}</div><div class="s">${esc(it.summary)}</div>` +
      `<div class="s">${esc(it.created.replace("T", " ").slice(0, 16))}${it.error ? " · " + esc(it.error) : ""}</div>` +
      `<span class="st ${it.state}">${STATE_TXT[it.state] || it.state}</span>` +
      (it.state !== "sent" ? '<div class="acts"><button type="button" class="btn small ghost" data-act="del">Ištrinti</button></div>'
        : (it.handles && it.handles.length ? '<div class="acts"><button type="button" class="btn small ghost" data-act="rm">Ištrinti failus iš telefono</button></div>' : ""));
    const del = li.querySelector('[data-act="del"]');
    const rm = li.querySelector('[data-act="rm"]');
    if (rm) rm.onclick = () => askDeleteFiles(it);
    if (del) del.onclick = async () => { if (confirm("Ištrinti šį neišsiųstą įrašą?")) { await DB.del("queue", it.id); renderQueue(); updateBadge(); } };
    ul.appendChild(li);
  }
}

/* ------------------------------------------------------------------ NUSTATYMAI */
async function renderSettings() {
  const t = Auth.read();
  $("set-account").textContent = !cfgReady() ? "Microsoft programa dar nenustatyta."
    : (Auth.signedIn() ? `Prisijungta: ${t.name || t.user || "paskyra"}` : "Neprisijungta.");
  $("set-login").hidden = !cfgReady() || Auth.signedIn();
  $("set-logout").hidden = !Auth.signedIn();
  $("set-device").value = cfg.device; $("set-client").value = cfg.clientId; $("set-tenant").value = cfg.tenantId; $("set-path").value = cfg.inboxPath;
  const pr = await DB.get("cache", "projects");
  $("set-data").textContent = pr ? `Projektų sąrašas: ${(pr.value.projects || []).length} įrašų, kompiuteris paskelbė ${String(pr.value.updated || "").replace("T", " ").slice(0, 16)}.` : "Projektų sąrašas dar neatsisiųstas.";
  $("set-log").textContent = JSON.parse(localStorage.getItem("prj_log") || "[]").join("\n") || "(žurnalas tuščias)";
}
function bindSettings() {
  $("set-save").onclick = () => {
    cfg.device = slugPy($("set-device").value, 20) || cfg.device;
    cfg.clientId = $("set-client").value.trim(); cfg.tenantId = $("set-tenant").value.trim();
    cfg.inboxPath = $("set-path").value.trim() || CFG0.inboxPath; saveCfg(cfg); toast("Nustatymai išsaugoti"); renderSettings(); updateBadge();
  };
  $("set-login").onclick = () => Auth.signIn().catch((e) => toast(e.message, true));
  $("set-logout").onclick = () => { Auth.signOut(); toast("Atsijungta"); renderSettings(); updateBadge(); };
  $("set-refresh").onclick = async () => { await refreshProjects(false); renderSettings(); };
  $("set-copylog").onclick = async () => { try { await navigator.clipboard.writeText($("set-log").textContent); toast("Nukopijuota"); } catch (_) { toast("Nepavyko nukopijuoti", true); } };
  $("set-clearlog").onclick = () => { localStorage.removeItem("prj_log"); renderSettings(); };
}

/* ------------------------------------------------------------------ paleidimas */
function applyCfgFragment() {
  const m = location.hash.match(/^#cfg=(.+)$/);
  if (!m) return;
  try {
    const j = JSON.parse(b64urlDecodeToString(m[1]));
    if (j.clientId) cfg.clientId = String(j.clientId).trim();
    if (j.tenantId) cfg.tenantId = String(j.tenantId).trim();
    if (j.inboxPath) cfg.inboxPath = String(j.inboxPath).trim();
    saveCfg(cfg); log("Nustatymai gauti iš nuorodos"); setTimeout(() => toast("Nustatymai gauti iš nuorodos"), 300);
  } catch (e) { log("Blogas #cfg: " + e.message); }
  history.replaceState(null, "", location.pathname + location.search);
}
async function boot() {
  $("btn-back").onclick = back;
  $("home-nav").onclick = () => openPicker((prj) => { if (prj) navigateToProject(prj); }, { title: "Naviguoti į projektą", nav: true });
  document.querySelectorAll("[data-go]").forEach((b) => {
    b.onclick = () => {
      const g = b.dataset.go;
      if (g === "note") initNote(); if (g === "marker") initMarker(); if (g === "expense") initExpense(); if (g === "files") initFiles(); if (g === "contact") initContact();
      show(g);
    };
  });
  $("pill").onclick = () => { if (cfgReady() && !Auth.signedIn()) show("settings"); else { flush(); show("queue"); } };
  $("picker-close").onclick = () => closePicker(undefined);
  $("picker-q").oninput = renderPicker; $("picker-completed").onchange = renderPicker;
  $("picker-manual-ok").onclick = () => { const v = $("picker-manual").value.trim().toUpperCase(); if (v) closePicker(v); };
  $("picker").onclick = (e) => { if (e.target.id === "picker") closePicker(undefined); };
  $("in-cam").onchange = onPhotoInput; $("in-gal").onchange = onPhotoInput;
  $("queue-send").onclick = async () => { await flush(); renderQueue(); };
  $("queue-clean").onclick = async () => { for (const it of await DB.all("queue")) if (it.state === "sent") await DB.del("queue", it.id); renderQueue(); updateBadge(); };
  bindNote(); bindMarker(); bindExpense(); bindFiles(); bindContact(); bindSettings();
  window.addEventListener("online", () => { updateBadge(); flush(); });
  window.addEventListener("offline", updateBadge);
  window.addEventListener("popstate", () => {});
  applyCfgFragment();
  try { await Auth.handleRedirect(); } catch (e) { setTimeout(() => toast("Prisijungti nepavyko: " + e.message, true), 400); }
  show("home");
  await purgeOldSent();
  updateBadge();
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch((e) => log("SW: " + e.message));
  if (cfgReady() && Auth.signedIn() && navigator.onLine) { refreshProjects(true); flush(); }
  setInterval(() => { if (navigator.onLine) flush(); }, 60000);
  document.addEventListener("visibilitychange", maybeAskDelete);
  setTimeout(maybeAskDelete, 1500);
}
document.addEventListener("DOMContentLoaded", boot);
window.__prj = { slugPy, num, flush, enqueue, DB, Auth };   // diagnostikai / testams
})();
