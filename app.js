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
const VERSION = "1.6.1";
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
const views = ["home", "note", "marker", "expense", "files", "contact", "field", "queue", "settings"];
const stack = [];
function show(name, push = true) {
  views.forEach((v) => { $("v-" + v).hidden = v !== name; });
  $("title").textContent = $("v-" + name).dataset.title;
  if (push) stack.push(name);
  $("btn-back").hidden = name === "home";
  window.scrollTo(0, 0);
  const cur = name;
  if (cur !== "field") { if (AR.on) closeAR(); stopFieldWatch(); fld.mapOpen = false; }
  if (cur === "queue") renderQueue();
  if (cur === "settings") renderSettings();
  if (cur === "home") renderHome();
}
function back() {
  if (AR.on) { closeAR(); return; }
  if (fld.mapOpen) { closeFieldMap(); return; }
  stack.pop(); const prev = stack[stack.length - 1] || "home"; show(prev, false); }
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

/* ------------------------------------------------------------------ lauko žemėlapis (paketas iš PC + GPS/RTK padėtis) */
/* Telefonas paprašo PC paketo (OneDrive _lauko_uzklausa), PC paruošia vieno projekto sklypą, kaimynus ir objekto
   taškus (_laukas), telefonas atsisiunčia, IŠTRINA kopiją iš OneDrive ir laiko tik vietinėje atmintyje (IndexedDB)
   iki „Baigti ir ištrinti“ arba FIELD_KEEP_DAYS dienų. Padėtis imama iš naršyklės geolokacijos – ji rodo ir telefono GPS,
   ir GNSS Master „Mock Location“ (RTK) padėtį. */
const FIELD_KEEP_DAYS = 3;
const FIELD_WAIT_MAX_MS = 10 * 60 * 1000;
const fld = { test: false, prj: "", mapOpen: false, waiting: false, cancel: false, watch: null, pos: null, lastAt: 0, wake: null, ro: null, raf: 0 };
const MAP = { lat0: 0, lon0: 0, kx: 1, ky: 111195, cx: 0, cy: 0, scale: 1, polys: [], pts: [], marks: [], W: 0, H: 0, dpr: 1, follow: false, ptrs: new Map(), pinch: 0, ctx: null, pkg: null };
const fieldKey = (prj) => "field:" + prj;
const fieldRel = (sub, prj) => `${inboxPath()}/${sub}/${slugPy(prj, 30)}.json`;

async function graphDelete(rel) {
  const r = await graphFetch(`${GRAPH}/me/drive/root:/${encPath(rel)}`, { method: "DELETE" });
  if (!r.ok && r.status !== 404) throw httpErr(r, await r.text());
}
async function fieldAll() {
  let all = []; try { all = await DB.all("cache"); } catch (_) { /* */ }
  return all.filter((x) => String(x.key).startsWith("field:")).map((x) => x.value).filter(Boolean)
    .sort((a, b) => String(b.fetched || "").localeCompare(String(a.fetched || "")));
}
async function fieldDelete(prj, { remote = true } = {}) {
  try { await DB.del("cache", fieldKey(prj)); } catch (_) { /* */ }
  if (remote && cfgReady() && Auth.signedIn() && navigator.onLine) {
    for (const sub of ["_laukas", "_lauko_uzklausa"]) { try { await graphDelete(fieldRel(sub, prj)); } catch (_) { /* nieko */ } }
  }
}
async function fieldPurgeOld() {
  const lim = Date.now() - FIELD_KEEP_DAYS * 86400000; let n = 0;
  for (const p of await fieldAll()) {
    const t = Date.parse(p.fetched || "");
    if (!isFinite(t) || t < lim) { await fieldDelete(p.project_no, { remote: false }); n++; }
  }
  if (n) setTimeout(() => toast(`Seni lauko duomenys ištrinti (${n})`), 600);
}
async function fieldClearAll() {
  const all = await fieldAll();
  for (const p of all) await fieldDelete(p.project_no);
  try { await DB.del("cache", "field_req"); } catch (_) { /* */ }
  fld.cancel = true; stopFieldWatch();
  return all.length;
}

/* Bandomasis režimas: be projekto ir be kompiuterio. Aplink dabartinę padėtį sukuriamas 20 × 30 m „sklypas“,
   kad būtų galima išbandyti padėtį, tikslumą, atstumą iki ribos ir „Žymą čia“. Niekur nesaugoma. */
function fieldTest() {
  if (!("geolocation" in navigator)) { toast("Naršyklė nepalaiko vietos nustatymo", true); return; }
  $("field-test").disabled = true; $("field-test").textContent = "Gaunama padėtis…";
  const reset = () => { $("field-test").disabled = false; $("field-test").textContent = "Testas be projekto (mano vieta)"; };
  navigator.geolocation.getCurrentPosition((p) => {
    reset();
    const lat = p.coords.latitude, lon = p.coords.longitude;
    const ky = 111195, kx = ky * Math.cos(lat * Math.PI / 180), dx = 10 / kx, dy = 15 / ky;
    const ring = [[lon - dx, lat - dy], [lon + dx, lat - dy], [lon + dx, lat + dy], [lon - dx, lat + dy]];
    openFieldMap({ project_no: "TESTAS", fetched: localISO(), points: [], notes: ["Bandomasis 20 × 30 m sklypas aplink jūsų padėtį"],
      parcels: [{ cad_no: "Bandomasis sklypas", role: "main", rings: [ring] }],
      bbox: [lon - dx, lat - dy, lon + dx, lat + dy] }, { test: true });
  }, (e) => { reset(); toast("Padėtis negauta: " + e.message + ". Leiskite naršyklei naudoti vietą.", true); },
  { enableHighAccuracy: true, maximumAge: 0, timeout: 20000 });
}
function initField() {
  fld.prj = ""; fld.mapOpen = false;
  setPick($("field-project"), "", "Pasirinkite projektą…");
  $("field-start").hidden = false; $("field-map-wrap").hidden = true;
  renderFieldStart();
  resumeFieldWait();
}
async function renderFieldStart() {
  const ul = $("field-saved"); ul.innerHTML = "";
  const all = await fieldAll();
  $("field-saved-h").hidden = !all.length; $("field-clear").hidden = !all.length;
  for (const p of all) {
    const li = document.createElement("li");
    li.innerHTML = `<div class="t">${esc(p.project_no)}</div><div class="s">gautas ${esc(String(p.fetched || "").slice(5, 16).replace("T", " "))} · sklypų ${(p.parcels || []).length} · taškų ${(p.points || []).length}</div>` +
      `<div class="acts"><button type="button" class="btn small primary">Atidaryti</button><button type="button" class="btn small ghost">Ištrinti</button></div>`;
    const b = li.querySelectorAll("button");
    b[0].onclick = () => openFieldMap(p);
    b[1].onclick = async () => { if (confirm(`Ištrinti ${p.project_no} lauko paketą iš telefono?`)) { await fieldDelete(p.project_no); renderFieldStart(); toast("Ištrinta"); } };
    ul.appendChild(li);
  }
}
function setWaitBanner(txt, bad) {
  const b = $("field-wait"); b.hidden = !txt; b.className = "banner " + (bad ? "bad" : "info"); $("field-wait-txt").textContent = txt || "";
  $("field-cancel").hidden = !fld.waiting;
}
async function fieldRequest() {
  const prj = fld.prj;
  if (!prj) { toast("Pirmiausia pasirinkite projektą", true); return; }
  if (!cfgReady() || !Auth.signedIn()) { toast("Pirmiausia prisijunkite (Nustatymai)", true); return; }
  if (!navigator.onLine) { toast("Nėra ryšio", true); return; }
  const reqId = newId();
  try {
    await graphPut(fieldRel("_lauko_uzklausa", prj), new Blob([JSON.stringify({ version: 1, req_id: reqId, project_no: prj, device: cfg.device, created: localISO() })], { type: "application/json" }), "application/json");
  } catch (e) { toast("Nepavyko išsiųsti užklausos: " + e.message, true); return; }
  await cachePut("field_req", { prj, reqId, at: Date.now() });
  rememberProject(prj);
  fld.cancel = false;
  resumeFieldWait();
}
async function resumeFieldWait() {
  if (fld.waiting) return;
  const rq = await cacheGet("field_req");
  if (!rq) { setWaitBanner(""); return; }
  if (Date.now() - rq.at > FIELD_WAIT_MAX_MS) { await DB.del("cache", "field_req"); setWaitBanner("Kompiuteris neatsakė per 10 min. PROJEKTITA programa turi būti atidaryta kompiuteryje. Bandykite dar kartą.", true); return; }
  fld.waiting = true; fld.cancel = false;
  const t0 = rq.at;
  try {
    for (;;) {
      if (fld.cancel) { setWaitBanner(""); break; }
      const sec = Math.round((Date.now() - t0) / 1000);
      if (Date.now() - t0 > FIELD_WAIT_MAX_MS) { await DB.del("cache", "field_req"); fld.waiting = false; setWaitBanner("Kompiuteris neatsakė per 10 min. PROJEKTITA programa turi būti atidaryta kompiuteryje. Bandykite dar kartą.", true); break; }
      setWaitBanner(`Laukiama kompiuterio (${rq.prj})… ${Math.floor(sec / 60)}:${pad(sec % 60)}. Paprastai 1–3 min.`);
      let j = null;
      try { j = await graphGetJson(fieldRel("_laukas", rq.prj)); }
      catch (e) { if (e.authRequired) { fld.waiting = false; setWaitBanner("Reikia prisijungti (Nustatymai).", true); return; } log("Lauko paketas: " + e.message); }
      if (j && j.req_id === rq.reqId) {
        await DB.del("cache", "field_req");
        fld.waiting = false;
        try { await graphDelete(fieldRel("_laukas", rq.prj)); await graphDelete(fieldRel("_lauko_uzklausa", rq.prj)); } catch (e) { log("Paketo šalinimas OneDrive: " + e.message); }
        if (j.error) { setWaitBanner(j.error, true); return; }
        j.fetched = localISO();
        await cachePut(fieldKey(j.project_no || rq.prj), j);
        setWaitBanner("");
        toast(`Paketas ${rq.prj} gautas ir ištrintas iš OneDrive`);
        if (!$("v-field").hidden && !fld.mapOpen) renderFieldStart();
        return;
      }
      for (let i = 0; i < 5 && !fld.cancel; i++) {
        await new Promise((s) => setTimeout(s, 1000));
        const sc = Math.round((Date.now() - t0) / 1000);
        setWaitBanner(`Laukiama kompiuterio (${rq.prj})… ${Math.floor(sc / 60)}:${pad(sc % 60)}. Paprastai 1–3 min.`);
      }
    }
  } finally { fld.waiting = false; $("field-cancel").hidden = true; }
}

/* ---- žemėlapis ---- */
function toM(lon, lat) { return [(lon - MAP.lon0) * MAP.kx, (lat - MAP.lat0) * MAP.ky]; }
function openFieldMap(pkg, opts = {}) {
  MAP.pkg = pkg; fld.mapOpen = true; fld.prj = pkg.project_no; fld.test = !!opts.test; MAP.marks = [];
  $("field-done").textContent = fld.test ? "Uždaryti testą" : "Baigti ir ištrinti";
  const bb = pkg.bbox || (() => { const p = (pkg.parcels[0] || {}).rings; return p ? [p[0][0][0], p[0][0][1], p[0][0][0], p[0][0][1]] : [25.28, 54.69, 25.28, 54.69]; })();
  MAP.lon0 = (bb[0] + bb[2]) / 2; MAP.lat0 = (bb[1] + bb[3]) / 2;
  MAP.kx = MAP.ky * Math.cos(MAP.lat0 * Math.PI / 180);
  MAP.polys = (pkg.parcels || []).map((p) => {
    const rings = p.rings.map((r) => r.map(([lo, la]) => toM(lo, la)));
    let sx = 0, sy = 0, n = 0; rings[0].forEach(([x, y]) => { sx += x; sy += y; n++; });
    return { cad: p.cad_no, role: p.role, rings, c: [sx / n, sy / n] };
  });
  MAP.pts = (pkg.points || []).map((p) => ({ label: p.label, xy: toM(p.lon, p.lat) }));
  $("field-start").hidden = true; $("field-map-wrap").hidden = false;
  $("title").textContent = pkg.project_no;
  const notes = (pkg.notes || []).join(" · ");
  $("field-legend").textContent = (notes ? notes + " · " : "") + "Mėlyna – projekto sklypas, pilka – kaimynai, raudona – objekto taškas, oranžinis – jūsų žymos, mėlynas apskritimas – jūsų padėtis.";
  initCanvas(); fitMap(); startFieldWatch(); drawMap(); updateFieldInfo();
}
function closeFieldMap() {
  fld.mapOpen = false; stopFieldWatch();
  $("field-map-wrap").hidden = true; $("field-start").hidden = false;
  $("title").textContent = $("v-field").dataset.title;
  renderFieldStart();
}
function initCanvas() {
  const cv = $("field-canvas");
  if (!cv._bound) {
    cv._bound = true;
    cv.style.touchAction = "none";
    cv.addEventListener("pointerdown", (e) => { cv.setPointerCapture(e.pointerId); MAP.ptrs.set(e.pointerId, [e.clientX, e.clientY]); MAP.pinch = 0; });
    cv.addEventListener("pointermove", (e) => {
      if (!MAP.ptrs.has(e.pointerId)) return;
      const prev = MAP.ptrs.get(e.pointerId); MAP.ptrs.set(e.pointerId, [e.clientX, e.clientY]);
      if (MAP.ptrs.size === 1) { MAP.cx -= (e.clientX - prev[0]) / MAP.scale; MAP.cy += (e.clientY - prev[1]) / MAP.scale; MAP.follow = false; scheduleDraw(); }
      else if (MAP.ptrs.size === 2) {
        const [a, b] = [...MAP.ptrs.values()]; const d = Math.hypot(a[0] - b[0], a[1] - b[1]);
        if (MAP.pinch) { const r = cv.getBoundingClientRect(); zoomAt((a[0] + b[0]) / 2 - r.left, (a[1] + b[1]) / 2 - r.top, d / MAP.pinch); }
        MAP.pinch = d; MAP.follow = false;
      }
    });
    const up = (e) => { MAP.ptrs.delete(e.pointerId); MAP.pinch = 0; };
    cv.addEventListener("pointerup", up); cv.addEventListener("pointercancel", up);
    cv.addEventListener("wheel", (e) => { e.preventDefault(); const r = cv.getBoundingClientRect(); zoomAt(e.clientX - r.left, e.clientY - r.top, e.deltaY < 0 ? 1.25 : 0.8); }, { passive: false });
    $("field-zin").onclick = () => zoomAt(MAP.W / 2, MAP.H / 2, 1.5);
    $("field-zout").onclick = () => zoomAt(MAP.W / 2, MAP.H / 2, 1 / 1.5);
    $("field-center").onclick = () => { if (!fld.pos) { toast("Padėtis dar negauta", true); return; } MAP.follow = true; centerOnUser(); };
    $("field-fit").onclick = () => { MAP.follow = false; fitMap(); drawMap(); };
    $("field-bg").onclick = () => { cfg.mapBg = !bgOn(); saveCfg(cfg); toast(cfg.mapBg ? "OSM fonas įjungtas" : "OSM fonas išjungtas"); drawMap(); };
    if (window.ResizeObserver) { fld.ro = new ResizeObserver(() => { sizeCanvas(); drawMap(); }); fld.ro.observe(cv.parentElement); }
    window.addEventListener("orientationchange", () => setTimeout(() => { sizeCanvas(); fitMap(); drawMap(); }, 300));
  }
  sizeCanvas();
}
function sizeCanvas() {
  const cv = $("field-canvas"), box = cv.parentElement.getBoundingClientRect();
  MAP.dpr = window.devicePixelRatio || 1; MAP.W = Math.max(200, Math.round(box.width)); MAP.H = Math.max(200, Math.round(box.height));
  cv.width = MAP.W * MAP.dpr; cv.height = MAP.H * MAP.dpr; cv.style.width = MAP.W + "px"; cv.style.height = MAP.H + "px";
  MAP.ctx = cv.getContext("2d");
}
function fitMap() {
  let xs = [], ys = [];
  const main = MAP.polys.filter((p) => p.role === "main");
  (main.length ? main : MAP.polys).forEach((p) => p.rings.forEach((r) => r.forEach(([x, y]) => { xs.push(x); ys.push(y); })));
  MAP.pts.forEach((p) => { xs.push(p.xy[0]); ys.push(p.xy[1]); });
  if (!xs.length) { MAP.cx = 0; MAP.cy = 0; MAP.scale = 2; return; }
  const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
  MAP.cx = (x0 + x1) / 2; MAP.cy = (y0 + y1) / 2;
  const w = Math.max(x1 - x0, 10), h = Math.max(y1 - y0, 10);
  MAP.scale = Math.min(MAP.W / (w * 1.4), MAP.H / (h * 1.4));
  MAP.scale = Math.max(0.02, Math.min(300, MAP.scale));
}
function zoomAt(sx, sy, f) {
  const ns = Math.max(0.02, Math.min(400, MAP.scale * f)); f = ns / MAP.scale;
  const wx = MAP.cx + (sx - MAP.W / 2) / MAP.scale, wy = MAP.cy - (sy - MAP.H / 2) / MAP.scale;
  MAP.scale = ns; MAP.cx = wx - (sx - MAP.W / 2) / ns; MAP.cy = wy + (sy - MAP.H / 2) / ns;
  scheduleDraw();
}
function centerOnUser() { if (!fld.pos) return; const [x, y] = toM(fld.pos.lon, fld.pos.lat); MAP.cx = x; MAP.cy = y; scheduleDraw(); }
function scheduleDraw() { if (fld.raf) return; fld.raf = requestAnimationFrame(() => { fld.raf = 0; drawMap(); }); }
const sx_ = (x) => MAP.W / 2 + (x - MAP.cx) * MAP.scale;
const sy_ = (y) => MAP.H / 2 - (y - MAP.cy) * MAP.scale;
/* OSM fono plytelės (Web Mercator, 256 px). Piešiamos po vektoriais; be ryšio tiesiog nematomos. */
const TILE_URL = (z, x, y) => `https://tile.openstreetmap.org/${z}/${x}/${y}.png`;
const tiles = new Map();
const bgOn = () => cfg.mapBg !== false;
const lon2x = (lon, z) => (lon + 180) / 360 * Math.pow(2, z);
const lat2y = (lat, z) => { const r = lat * Math.PI / 180; return (1 - Math.log(Math.tan(r) + 1 / Math.cos(r)) / Math.PI) / 2 * Math.pow(2, z); };
const x2lon = (x, z) => x / Math.pow(2, z) * 360 - 180;
const y2lat = (y, z) => Math.atan(Math.sinh(Math.PI * (1 - 2 * y / Math.pow(2, z)))) * 180 / Math.PI;
function getTile(z, x, y) {
  const k = z + "/" + x + "/" + y; let t = tiles.get(k);
  if (!t) {
    if (tiles.size > 300) { const first = tiles.keys().next().value; tiles.delete(first); }
    t = { img: new Image(), ok: false };
    t.img.onload = () => { t.ok = true; scheduleDraw(); };
    t.img.onerror = () => { t.fail = true; };
    t.img.src = TILE_URL(z, x, y);
    tiles.set(k, t);
  }
  return t;
}
function drawTiles(c) {
  if (!bgOn() || !navigator.onLine) return;
  const lonAt = (sx) => MAP.lon0 + (MAP.cx + (sx - MAP.W / 2) / MAP.scale) / MAP.kx;
  const latAt = (sy) => MAP.lat0 + (MAP.cy - (sy - MAP.H / 2) / MAP.scale) / MAP.ky;
  const z = Math.max(2, Math.min(19, Math.round(Math.log2(156543.03 * Math.cos(MAP.lat0 * Math.PI / 180) * MAP.scale))));
  const x0 = Math.floor(lon2x(lonAt(0), z)), x1 = Math.floor(lon2x(lonAt(MAP.W), z));
  const y0 = Math.floor(lat2y(latAt(0), z)), y1 = Math.floor(lat2y(latAt(MAP.H), z));
  if ((x1 - x0 + 1) * (y1 - y0 + 1) > 60) return;
  const n = Math.pow(2, z);
  for (let ty = y0; ty <= y1; ty++) for (let tx = x0; tx <= x1; tx++) {
    if (ty < 0 || ty >= n) continue;
    const t = getTile(z, ((tx % n) + n) % n, ty);
    if (!t.ok) continue;
    const [ax, ay] = toM(x2lon(tx, z), y2lat(ty, z)), [bx, by] = toM(x2lon(tx + 1, z), y2lat(ty + 1, z));
    const px = sx_(ax), py = sy_(ay);
    c.drawImage(t.img, px, py, sx_(bx) - px + 0.5, sy_(by) - py + 0.5);
  }
  c.font = "10px system-ui,sans-serif"; c.textAlign = "right";
  const txt = "© OpenStreetMap", w = c.measureText(txt).width;
  c.fillStyle = "rgba(255,255,255,.75)"; c.fillRect(MAP.W - w - 10, MAP.H - 16, w + 10, 16);
  c.fillStyle = "#333"; c.fillText(txt, MAP.W - 5, MAP.H - 5);
}
function cssVar(n, d) { const v = getComputedStyle(document.documentElement).getPropertyValue(n).trim(); return v || d; }
function drawMap() {
  const c = MAP.ctx; if (!c || !fld.mapOpen) return;
  const dark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const prim = cssVar("--primary", "#393187"), txt = cssVar("--text", "#1b1c3a");
  c.setTransform(MAP.dpr, 0, 0, MAP.dpr, 0, 0);
  c.fillStyle = dark ? "#0f0f20" : "#f7f8fb"; c.fillRect(0, 0, MAP.W, MAP.H);
  drawTiles(c);
  const path = (rings) => { c.beginPath(); rings.forEach((r) => { r.forEach(([x, y], i) => { i ? c.lineTo(sx_(x), sy_(y)) : c.moveTo(sx_(x), sy_(y)); }); c.closePath(); }); };
  c.lineJoin = "round";
  for (const role of ["neighbor", "main"]) {
    for (const p of MAP.polys) {
      if (p.role !== role) continue;
      path(p.rings);
      if (role === "main") { c.fillStyle = dark ? "rgba(141,134,232,.25)" : "rgba(57,49,135,.16)"; c.fill("evenodd"); c.strokeStyle = prim; c.lineWidth = 2.5; }
      else { c.fillStyle = dark ? "rgba(160,160,190,.08)" : "rgba(120,120,150,.08)"; c.fill("evenodd"); c.strokeStyle = dark ? "#8a8cab" : "#7b7e99"; c.lineWidth = 1.2; }
      c.stroke();
      if (role === "main" && MAP.scale > 0.6) { c.fillStyle = prim; p.rings[0].forEach(([x, y]) => { c.beginPath(); c.arc(sx_(x), sy_(y), 3, 0, 6.2832); c.fill(); }); }
      if (MAP.scale > 0.25 || role === "main") {
        c.fillStyle = role === "main" ? prim : (dark ? "#b4b6d4" : "#555873"); c.font = (role === "main" ? "bold 13px" : "11px") + " system-ui,sans-serif"; c.textAlign = "center";
        c.fillText(p.cad, sx_(p.c[0]), sy_(p.c[1]));
      }
    }
  }
  for (const p of MAP.pts) {
    const x = sx_(p.xy[0]), y = sy_(p.xy[1]);
    c.fillStyle = "#e8201f"; c.strokeStyle = "#fff"; c.lineWidth = 2; c.beginPath(); c.arc(x, y, 7, 0, 6.2832); c.fill(); c.stroke();
    c.fillStyle = txt; c.font = "bold 12px system-ui,sans-serif"; c.textAlign = "left"; c.fillText(p.label, x + 10, y + 4);
  }
  MAP.marks.forEach((m) => {
    const x = sx_(m.xy[0]), y = sy_(m.xy[1]);
    c.fillStyle = "#ff8a00"; c.strokeStyle = "#fff"; c.lineWidth = 2; c.beginPath(); c.moveTo(x, y - 9); c.lineTo(x + 9, y); c.lineTo(x, y + 9); c.lineTo(x - 9, y); c.closePath(); c.fill(); c.stroke();
    c.fillStyle = txt; c.font = "bold 12px system-ui,sans-serif"; c.textAlign = "left"; c.fillText(m.label, x + 12, y + 4);
  });
  if (fld.pos) {
    const [ux, uy] = toM(fld.pos.lon, fld.pos.lat), x = sx_(ux), y = sy_(uy), r = Math.max(fld.pos.acc * MAP.scale, 0);
    c.fillStyle = "rgba(20,110,255,.15)"; c.strokeStyle = "rgba(20,110,255,.6)"; c.lineWidth = 1;
    c.beginPath(); c.arc(x, y, r, 0, 6.2832); c.fill(); c.stroke();
    c.fillStyle = "#1470ff"; c.strokeStyle = "#fff"; c.lineWidth = 2.5; c.beginPath(); c.arc(x, y, 7, 0, 6.2832); c.fill(); c.stroke();
  }
  // mastelio liniuotė ir šiaurė
  const nice = [0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50, 100, 200, 500, 1000]; let m = nice[0];
  for (const v of nice) { if (v * MAP.scale <= 110) m = v; }
  c.strokeStyle = txt; c.fillStyle = txt; c.lineWidth = 2; c.beginPath(); c.moveTo(12, MAP.H - 14); c.lineTo(12 + m * MAP.scale, MAP.H - 14); c.stroke();
  c.font = "12px system-ui,sans-serif"; c.textAlign = "left"; c.fillText(m >= 1 ? m + " m" : (m * 100) + " cm", 12, MAP.H - 20);
  c.textAlign = "center"; c.font = "bold 13px system-ui,sans-serif"; c.fillText("▲ Š", MAP.W - 22, 20);
}

/* ---- padėtis ir atstumai ---- */
function startFieldWatch() {
  stopFieldWatch();
  if (!("geolocation" in navigator)) { $("field-info").textContent = "Naršyklė nepalaiko vietos nustatymo."; return; }
  fld.watch = navigator.geolocation.watchPosition((p) => {
    fld.pos = { lat: p.coords.latitude, lon: p.coords.longitude, acc: p.coords.accuracy }; fld.lastAt = Date.now();
    if (MAP.follow) centerOnUser();
    updateFieldInfo(); scheduleDraw();
  }, (e) => { $("field-info").textContent = "Vietos klaida: " + e.message + ". Leiskite naršyklei naudoti vietą."; },
  { enableHighAccuracy: true, maximumAge: 0, timeout: 20000 });
  try { if (navigator.wakeLock) navigator.wakeLock.request("screen").then((w) => { fld.wake = w; }).catch(() => {}); } catch (_) { /* */ }
}
function stopFieldWatch() {
  if (fld.watch != null) { navigator.geolocation.clearWatch(fld.watch); fld.watch = null; }
  if (fld.wake) { try { fld.wake.release(); } catch (_) { /* */ } fld.wake = null; }
}
function fmtM(v) { return v < 1 ? Math.round(v * 100) + " cm" : v.toFixed(v < 10 ? 2 : 1).replace(".", ",") + " m"; }
function segDist(px, py, ax, ay, bx, by) {
  const dx = bx - ax, dy = by - ay, l2 = dx * dx + dy * dy;
  let t = l2 ? ((px - ax) * dx + (py - ay) * dy) / l2 : 0; t = Math.max(0, Math.min(1, t));
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}
function inRings(px, py, rings) {
  let ins = false;
  for (const r of rings) for (let i = 0, j = r.length - 1; i < r.length; j = i++) {
    const [xi, yi] = r[i], [xj, yj] = r[j];
    if ((yi > py) !== (yj > py) && px < (xj - xi) * (py - yi) / (yj - yi) + xi) ins = !ins;
  }
  return ins;
}
function boundaryStats(px, py) {
  let best = null;
  for (const p of MAP.polys) {
    if (p.role !== "main") continue;
    let d = Infinity, v = Infinity;
    for (const r of p.rings) for (let i = 0; i < r.length; i++) {
      const a = r[i], b = r[(i + 1) % r.length];
      d = Math.min(d, segDist(px, py, a[0], a[1], b[0], b[1])); v = Math.min(v, Math.hypot(px - a[0], py - a[1]));
    }
    const inside = inRings(px, py, p.rings);
    if (!best || (inside && !best.inside) || (inside === best.inside && d < best.d)) best = { cad: p.cad, d, v, inside };
  }
  return best;
}
function updateFieldInfo() {
  const el = $("field-info");
  if (!fld.pos) { el.innerHTML = '<span class="muted">Laukiama GPS padėties…</span>'; return; }
  const a = fld.pos.acc;
  const cls = a <= 0.05 ? "ok" : a <= 0.5 ? "warn" : "bad";
  const q = a <= 0.05 ? "RTK fiksuotas" : a <= 0.5 ? "RTK float / DGPS" : a <= 5 ? "įprastas GPS" : "silpnas signalas";
  const [x, y] = toM(fld.pos.lon, fld.pos.lat), b = boundaryStats(x, y);
  let line = "";
  if (b) line = `${b.inside ? "Sklypo viduje" : "Už sklypo ribų"} · iki ribos <b>${fmtM(b.d)}</b> · iki artimiausio posūkio taško ${fmtM(b.v)}`;
  else if (MAP.pts.length) { const p = MAP.pts[0]; line = `Iki objekto taško ${fmtM(Math.hypot(x - p.xy[0], y - p.xy[1]))}`; }
  if (MAP.marks.length) { const m = MAP.marks[MAP.marks.length - 1]; line += `${line ? " · " : ""}iki žymos ${m.label} <b>${fmtM(Math.hypot(x - m.xy[0], y - m.xy[1]))}</b>`; }
  el.innerHTML = `<div><span class="acc ${cls}">±${a < 1 ? Math.round(a * 100) + " cm" : a.toFixed(1).replace(".", ",") + " m"}</span> <span class="muted">${q}</span></div><div>${line}</div>` +
    (Date.now() - fld.lastAt > 6000 ? '<div class="small warnTxt">Padėtis nebeatnaujinama (patikrinkite GPS / GNSS Master).</div>' : "") +
    (a > 1 ? '<div class="small warnTxt">Su tokiu tikslumu ribos išnešti negalima – tik orientavimuisi.</div>' : "");
}
setInterval(() => { if (fld.mapOpen && fld.pos) updateFieldInfo(); }, 2000);

function bindField() {
  $("field-project").onclick = () => openPicker((prj) => { if (prj) { fld.prj = prj; setPick($("field-project"), prj, "Pasirinkite projektą…"); } }, { title: "Lauko paketas" });
  $("field-get").onclick = fieldRequest;
  $("field-test").onclick = fieldTest;
  $("field-mark").onclick = () => {
    if (!fld.pos) { toast("Padėtis dar negauta", true); return; }
    MAP.marks.push({ label: "Z" + (MAP.marks.length + 1), xy: toM(fld.pos.lon, fld.pos.lat) });
    toast(`Žyma Z${MAP.marks.length} padėta. Eikite toliau ir žiūrėkite atstumą.`); updateFieldInfo(); scheduleDraw();
  };
  $("field-cancel").onclick = () => { fld.cancel = true; DB.del("cache", "field_req"); setWaitBanner(""); };
  $("field-clear").onclick = async () => { if (confirm("Ištrinti VISUS lauko duomenis iš telefono?")) { const n = await fieldClearAll(); renderFieldStart(); toast(`Ištrinta paketų: ${n}`); } };
  $("field-done").onclick = async () => {
    const prj = fld.prj;
    if (fld.test) { closeFieldMap(); return; }
    if (!confirm(`Baigti darbą ir ištrinti ${prj} duomenis iš telefono?`)) return;
    await fieldDelete(prj); closeFieldMap(); toast("Lauko duomenys ištrinti");
  };
  $("set-fieldclear").onclick = async () => { if (confirm("Ištrinti VISUS lauko duomenis iš telefono?")) { const n = await fieldClearAll(); toast(`Ištrinta paketų: ${n}`); } };
}

/* ------------------------------------------------------------------ AR kameros vaizdas */
/* Kamera + sklypo ribos / taškai ant jos. Padėtis – iš GPS/RTK (fld.pos), kryptis – iš telefono kompaso ir giroskopo
   (deviceorientation). Kompaso paklaida kompensuojama „Kryptis pagal tašką“ kalibravimu. Tai vaizdinė pagalba, ne matavimas. */
const AR = { on: false, stream: null, R: null, off: 0, fov: 62, h: 1.5, raf: 0, absolute: false, evAt: 0, W: 0, H: 0, dpr: 1, hudAt: 0, lis: [], ev: { abs: 0, rel: 0, nul: 0 }, sensor: null, via: "", perm: "" };
function arLoadPrefs() { if (typeof cfg.arFov === "number") AR.fov = cfg.arFov; if (typeof cfg.arH === "number") AR.h = cfg.arH; }
function rotFromEuler(aD, bD, gD) {
  const d = Math.PI / 180, a = aD * d, b = bD * d, g = gD * d;
  const cA = Math.cos(a), sA = Math.sin(a), cB = Math.cos(b), sB = Math.sin(b), cG = Math.cos(g), sG = Math.sin(g);
  return [
    [cA * cG - sA * sB * sG, -cB * sA, cA * sG + cG * sA * sB],
    [cG * sA + cA * sB * sG, cA * cB, sA * sG - cA * cG * sB],
    [-cB * sG, sB, cB * cG],
  ];
}
function arOrthonorm(R) {   // stulpeliai: x=R[.][0], y=R[.][1], z=R[.][2]
  const col = (j) => [R[0][j], R[1][j], R[2][j]];
  const nrm = (v) => { const l = Math.hypot(...v) || 1; return v.map((t) => t / l); };
  const dot = (p, q) => p[0] * q[0] + p[1] * q[1] + p[2] * q[2];
  let z = nrm(col(2)); let y = col(1); const k = dot(y, z); y = nrm([y[0] - k * z[0], y[1] - k * z[1], y[2] - k * z[2]]);
  const x = [y[1] * z[2] - y[2] * z[1], y[2] * z[0] - y[0] * z[2], y[0] * z[1] - y[1] * z[0]];
  return [[x[0], y[0], z[0]], [x[1], y[1], z[1]], [x[2], y[2], z[2]]];
}
function arOnOrientation(e) {
  if (e.type === "deviceorientationabsolute") AR.ev.abs++; else AR.ev.rel++;
  if (e.alpha == null || e.beta == null || e.gamma == null) { AR.ev.nul++; return; }
  if (e.type === "deviceorientation" && AR.absolute) return;            // jau gaunamas absoliutus kompasas
  AR.via = "event";
  if (e.type === "deviceorientationabsolute") AR.absolute = true;
  const R = rotFromEuler(e.alpha, e.beta, e.gamma);
  if (!AR.R) AR.R = R;
  else { const k = 0.35; AR.R = arOrthonorm(AR.R.map((row, i) => row.map((v, j) => v + k * (R[i][j] - v)))); }
  AR.evAt = Date.now();
}
function arQuatToR(q) {   // [x,y,z,w] -> posūkio matrica (įrenginys -> Žemė ENU)
  const [x, y, z, w] = q;
  return [
    [1 - 2 * (y * y + z * z), 2 * (x * y - z * w), 2 * (x * z + y * w)],
    [2 * (x * y + z * w), 1 - 2 * (x * x + z * z), 2 * (y * z - x * w)],
    [2 * (x * z - y * w), 2 * (y * z + x * w), 1 - 2 * (x * x + y * y)],
  ];
}
async function arStartSensor() {   // atsarginis kelias: Generic Sensor API (Chromium naršyklės)
  if (AR.sensor || typeof AbsoluteOrientationSensor === "undefined") return false;
  try {
    for (const n of ["accelerometer", "gyroscope", "magnetometer"]) {
      try { const r = await navigator.permissions.query({ name: n }); if (r.state === "denied") { AR.perm = "blokuota: " + n; return false; } } catch (e) { /* nežinomas leidimas */ }
    }
    const s = new AbsoluteOrientationSensor({ frequency: 30, referenceFrame: "device" });
    s.addEventListener("reading", () => {
      if (!s.quaternion) return;
      const R = arQuatToR(s.quaternion);
      if (!AR.R) AR.R = R; else AR.R = arOrthonorm(AR.R.map((row, i) => row.map((v, j) => v + 0.35 * (R[i][j] - v))));
      AR.absolute = true; AR.evAt = Date.now(); AR.via = "sensor";
    });
    s.addEventListener("error", (ev) => { AR.perm = "jutiklio klaida: " + (ev.error && ev.error.name || "?"); });
    s.start(); AR.sensor = s; return true;
  } catch (e) { AR.perm = "jutiklis: " + e.name; return false; }
}
function arYaw() {   // kameros kryptis (azimutas, laipsniais pagal laikrodį nuo šiaurės), be poslinkio
  if (!AR.R) return null;
  const fE = -AR.R[0][2], fN = -AR.R[1][2];
  return ((Math.atan2(fE, fN) * 180 / Math.PI) + 360) % 360;
}
/* Pasaulio vektorius (E, N, U metrai nuo telefono) -> kameros koordinatės (x dešinėn, y aukštyn, gylis pirmyn) */
function arCam(E, N, U) {
  const o = AR.off * Math.PI / 180, c = Math.cos(o), s = Math.sin(o);
  const e2 = E * c - N * s, n2 = E * s + N * c, R = AR.R;
  const dx = e2 * R[0][0] + n2 * R[1][0] + U * R[2][0];
  const dy = e2 * R[0][1] + n2 * R[1][1] + U * R[2][1];
  const dz = e2 * R[0][2] + n2 * R[1][2] + U * R[2][2];
  return [dx, dy, -dz];
}
function arScreen(p) {   // p = kameros koordinatės, gylis > 0
  const f = (AR.H / 2) / Math.tan(AR.fov * Math.PI / 360);
  return [AR.W / 2 + f * p[0] / p[2], AR.H / 2 - f * p[1] / p[2]];
}
function arPoint(wx, wy, ux, uy, ground = true) { return arCam(wx - ux, wy - uy, ground ? -AR.h : 0); }
function arTargets(ux, uy) {   // kalibravimo / etikečių taikiniai
  const t = [];
  MAP.polys.filter((p) => p.role === "main").forEach((p) => p.rings[0].forEach(([x, y], i) => t.push({ label: "Posūkio taškas " + (i + 1), short: "P" + (i + 1), x, y })));
  MAP.pts.forEach((p) => t.push({ label: "Objekto taškas " + p.label, short: p.label, x: p.xy[0], y: p.xy[1], obj: true }));
  MAP.marks.forEach((p) => t.push({ label: "Žyma " + p.label, short: p.label, x: p.xy[0], y: p.xy[1], obj: true }));
  t.forEach((q) => { q.d = Math.hypot(q.x - ux, q.y - uy); });
  return t.sort((a, b) => a.d - b.d);
}
function arDraw() {
  AR.raf = requestAnimationFrame(arDraw);
  const cv = $("ar-canvas"), c = cv.getContext("2d");
  const W = cv.clientWidth, H = cv.clientHeight, dpr = window.devicePixelRatio || 1;
  if (cv.width !== Math.round(W * dpr) || cv.height !== Math.round(H * dpr)) { cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr); }
  AR.W = W; AR.H = H; AR.dpr = dpr;
  c.setTransform(dpr, 0, 0, dpr, 0, 0); c.clearRect(0, 0, W, H);
  const now = Date.now();
  if (now - AR.hudAt > 400) {
    AR.hudAt = now;
    const yaw = arYaw(), st = [];
    if (!AR.R || now - AR.evAt > 2000) st.push("jutiklių duomenų nėra (įvykiai " + AR.ev.abs + "/" + AR.ev.rel + ", tušti " + AR.ev.nul + (AR.perm ? ", " + AR.perm : "") + ")");
    else st.push((AR.absolute ? "kompasas " : "kompasas (nepatikimas) ") + Math.round((yaw + AR.off + 360) % 360) + "°");
    st.push("poslinkis " + (AR.off > 0 ? "+" : "") + AR.off.toFixed(0) + "°");
    $("ar-hud").innerHTML = $("field-info").innerHTML + `<div class="small">${st.join(" · ")} · FOV ${AR.fov}° · aukštis ${AR.h.toFixed(1).replace(".", ",")} m</div>`;
  }
  c.strokeStyle = "rgba(255,255,255,.7)"; c.lineWidth = 1.5;
  c.beginPath(); c.moveTo(W / 2 - 12, H / 2); c.lineTo(W / 2 + 12, H / 2); c.moveTo(W / 2, H / 2 - 12); c.lineTo(W / 2, H / 2 + 12); c.stroke();
  if (!AR.R || !fld.pos) {
    c.fillStyle = "rgba(0,0,0,.55)"; c.fillRect(0, H / 2 - 28, W, 56);
    c.fillStyle = "#fff"; c.font = "bold 15px system-ui,sans-serif"; c.textAlign = "center";
    c.fillText(!fld.pos ? "Laukiama GPS padėties…" : "Laukiama kompaso / giroskopo…", W / 2, H / 2 + 5); return;
  }
  const [ux, uy] = toM(fld.pos.lon, fld.pos.lat), NEAR = 0.3;
  const line = (a, b, col, lw) => {
    let pa = arPoint(a[0], a[1], ux, uy), pb = arPoint(b[0], b[1], ux, uy);
    if (pa[2] < NEAR && pb[2] < NEAR) return;
    if (pa[2] < NEAR) { const t = (NEAR - pa[2]) / (pb[2] - pa[2]); pa = pa.map((v, i) => v + t * (pb[i] - v)); }
    else if (pb[2] < NEAR) { const t = (NEAR - pb[2]) / (pa[2] - pb[2]); pb = pb.map((v, i) => v + t * (pa[i] - v)); }
    const s1 = arScreen(pa), s2 = arScreen(pb);
    c.strokeStyle = col; c.lineWidth = lw; c.beginPath(); c.moveTo(s1[0], s1[1]); c.lineTo(s2[0], s2[1]); c.stroke();
  };
  c.lineCap = "round";
  for (const role of ["neighbor", "main"]) for (const p of MAP.polys) {
    if (p.role !== role) continue;
    for (const r of p.rings) for (let i = 0; i < r.length; i++) {
      const a = r[i], b = r[(i + 1) % r.length];
      if (role === "main") { line(a, b, "rgba(0,0,0,.6)", 6); line(a, b, "#ffd400", 3); } else line(a, b, "rgba(255,255,255,.55)", 1.5);
    }
  }
  const lab = (t, col) => {
    const p = arPoint(t.x, t.y, ux, uy); if (p[2] < NEAR) return;
    const [sx, sy] = arScreen(p); if (sx < -40 || sx > AR.W + 40 || sy < -40 || sy > AR.H + 40) return;
    c.fillStyle = col; c.strokeStyle = "#fff"; c.lineWidth = 2; c.beginPath(); c.arc(sx, sy, t.obj ? 9 : 6, 0, 6.2832); c.fill(); c.stroke();
    c.font = "bold 14px system-ui,sans-serif"; c.textAlign = "left"; const tx = t.short + " · " + fmtM(t.d), w = c.measureText(tx).width;
    c.fillStyle = "rgba(0,0,0,.65)"; c.fillRect(sx + 10, sy - 18, w + 10, 22); c.fillStyle = "#fff"; c.fillText(tx, sx + 15, sy - 2);
  };
  for (const t of arTargets(ux, uy)) if (t.d < 200) lab(t, t.obj ? "#e8201f" : "#ffd400");
}
function arPickTarget() {
  if (!fld.pos) { toast("Padėtis dar negauta", true); return; }
  const [ux, uy] = toM(fld.pos.lon, fld.pos.lat), ul = $("ar-pick-list"); ul.innerHTML = "";
  arTargets(ux, uy).slice(0, 8).forEach((t) => {
    const li = document.createElement("li");
    li.innerHTML = `<div class="t">${esc(t.label)}</div><div class="s">${fmtM(t.d)} nuo jūsų</div>`;
    li.onclick = () => {
      const yaw = arYaw(); if (yaw == null) { toast("Nėra kompaso duomenų", true); return; }
      const brg = (Math.atan2(t.x - ux, t.y - uy) * 180 / Math.PI + 360) % 360;
      let off = brg - yaw; while (off > 180) off -= 360; while (off < -180) off += 360;
      AR.off = off; $("ar-pick").hidden = true; toast(`Kryptis sulygiuota (poslinkis ${off.toFixed(0)}°)`);
    };
    ul.appendChild(li);
  });
  $("ar-pick").hidden = false;
}
async function openAR() {
  if (!fld.mapOpen) return;
  arLoadPrefs(); AR.R = null; AR.absolute = false; AR.off = 0; AR.ev = { abs: 0, rel: 0, nul: 0 }; AR.via = ""; AR.perm = "";
  try {   // iOS ir kai kurios naršyklės reikalauja leidimo judesio jutikliams (turi būti iškviesta paspaudimo metu)
    if (typeof DeviceOrientationEvent !== "undefined" && typeof DeviceOrientationEvent.requestPermission === "function") {
      const r = await DeviceOrientationEvent.requestPermission(); if (r !== "granted") AR.perm = "leidimas: " + r;
    }
  } catch (e) { AR.perm = "leidimas: " + e.name; }
  $("ar").hidden = false; AR.on = true;
  $("ar-msg").textContent = "Paleidžiama kamera…";
  try {
    AR.stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false });
    const v = $("ar-video"); v.srcObject = AR.stream; await v.play().catch(() => {});
    $("ar-msg").textContent = "";
  } catch (e) { $("ar-msg").textContent = "Kamera nepasiekiama: " + e.message + ". Leiskite naršyklei naudoti kamerą."; }
  const add = (n) => { window.addEventListener(n, arOnOrientation, true); AR.lis.push(n); };
  add("deviceorientationabsolute"); add("deviceorientation");
  if (!AR.raf) arDraw();
  setTimeout(() => { if (AR.on && !AR.R) arStartSensor(); }, 1500);
}
function closeAR() {
  AR.on = false; cancelAnimationFrame(AR.raf); AR.raf = 0;
  AR.lis.forEach((n) => window.removeEventListener(n, arOnOrientation, true)); AR.lis = [];
  if (AR.sensor) { try { AR.sensor.stop(); } catch (e) { /* */ } AR.sensor = null; }
  if (AR.stream) { AR.stream.getTracks().forEach((t) => t.stop()); AR.stream = null; }
  $("ar-video").srcObject = null; $("ar").hidden = true; $("ar-pick").hidden = true;
}
function bindAR() {
  $("field-ar").onclick = openAR;
  $("ar-close").onclick = closeAR;
  $("ar-cal").onclick = arPickTarget;
  $("ar-pick-close").onclick = () => { $("ar-pick").hidden = true; };
  $("ar-left").onclick = () => { AR.off += 1; };
  $("ar-right").onclick = () => { AR.off -= 1; };
  $("ar-fov-").onclick = () => { AR.fov = Math.max(40, AR.fov - 2); cfg.arFov = AR.fov; saveCfg(cfg); };
  $("ar-fov+").onclick = () => { AR.fov = Math.min(100, AR.fov + 2); cfg.arFov = AR.fov; saveCfg(cfg); };
  $("ar-h").onclick = () => { const o = [1.2, 1.4, 1.5, 1.6, 1.8]; let i = 0; o.forEach((v, k) => { if (Math.abs(v - AR.h) < Math.abs(o[i] - AR.h)) i = k; }); AR.h = o[(i + 1) % o.length]; cfg.arH = AR.h; saveCfg(cfg); toast(`Telefono aukštis ${AR.h.toFixed(1).replace(".", ",")} m`); };
}
window.__ar = { rotFromEuler, arCam, arScreen, AR, arYaw, setR: (R) => { AR.R = R; } };

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
      if (g === "note") initNote(); if (g === "marker") initMarker(); if (g === "expense") initExpense(); if (g === "files") initFiles(); if (g === "contact") initContact(); if (g === "field") initField();
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
  bindNote(); bindMarker(); bindExpense(); bindFiles(); bindContact(); bindField(); bindAR(); bindSettings();
  window.addEventListener("online", () => { updateBadge(); flush(); });
  window.addEventListener("offline", updateBadge);
  window.addEventListener("popstate", () => {});
  applyCfgFragment();
  try { await Auth.handleRedirect(); } catch (e) { setTimeout(() => toast("Prisijungti nepavyko: " + e.message, true), 400); }
  show("home");
  await purgeOldSent();
  try { await fieldPurgeOld(); } catch (e) { log("Lauko duomenų valymas: " + e.message); }
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
