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
const VERSION = "1.9.0";
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

async function refreshBuildings(prj, silent) {
  try {
    const j = await graphGetJson(`${inboxPath()}/_projektai/${slugPy(prj, 30)}_pastatai.json`);
    if (!j) { if (!silent) toast("Šio projekto pastatų (sluoksniai 2601–2699) sąrašo nėra – kompiuteryje atidarykite KADMAT taškus.", true); return null; }
    await cachePut("buildings:" + prj, j); log(`Pastatai ${prj}: ${(j.buildings || []).length}`);
    return j;
  } catch (e) { log("Pastatai: " + e.message); if (!silent && !e.authRequired) toast("Nepavyko gauti pastatų: " + e.message, true); return null; }
}

/* ------------------------------------------------------------------ UI pagrindas */
const views = ["home", "note", "marker", "expense", "files", "contact", "field", "fm", "queue", "settings"];
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
/* „Atnaujinti versiją“: atsisiunčia naujausius programėlės failus (apeidama naršyklės ir paslaugos talpyklą), išvalo
   programos talpyklą ir perkrauna. Pastabų, siuntimo eilės ir nustatymų (localStorage / IndexedDB) NEliečia. */
async function updateApp() {
  const btn = $("home-update"), label = btn.textContent; btn.disabled = true; btn.textContent = "Tikrinama…";
  const done = () => { btn.disabled = false; btn.textContent = label; };
  try {
    const base = location.href.split("#")[0].split("?")[0].replace(/[^/]*$/, "");
    let remote = "";
    for (const f of ["", "index.html", "app.css", "config.js", "app.js", "sw.js", "manifest.webmanifest"]) {
      const r = await fetch(base + f, { cache: "reload" });
      if (!r.ok) throw new Error((f || "/") + " " + r.status);
      if (f === "app.js") remote = ((await r.text()).match(/const VERSION = "([\d.]+)"/) || [])[1] || "";
    }
    if (remote && remote === VERSION) { toast("Jau naujausia versija (v" + VERSION + ")"); done(); return; }
    btn.textContent = "Atnaujinama…";
    if ("serviceWorker" in navigator) for (const r of await navigator.serviceWorker.getRegistrations()) await r.unregister();
    if (window.caches) for (const k of await caches.keys()) await caches.delete(k);
    toast("Rasta v" + (remote || "?") + " – perkraunama…");
    setTimeout(() => location.reload(), 700);
  } catch (e) { toast("Nepavyko atnaujinti (nėra ryšio?): " + e.message, true); done(); }
}
function bindSettings() {
  $("home-update").onclick = updateApp;
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
const MAP = { lat0: 0, lon0: 0, kx: 1, ky: 111195, cx: 0, cy: 0, scale: 1, polys: [], pts: [], lp: [], lpSel: null, tap: null, marks: [], W: 0, H: 0, dpr: 1, follow: false, ptrs: new Map(), pinch: 0, ctx: null, pkg: null };
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
/* „Mano vieta“: telefonas nusiunčia dabartinę padėtį, kompiuteris „Ribos“ SHP (sklypai.shp) suranda sklypą ir
   besiribojančius sklypus (tarpas ≤ 0,5 m), o jei SHP / sklypo nėra – grąžina priežastį. */
async function fieldHere() {
  if (!cfgReady() || !Auth.signedIn()) { toast("Pirmiausia prisijunkite (Nustatymai)", true); return; }
  if (!navigator.onLine) { toast("Nėra ryšio", true); return; }
  if (!("geolocation" in navigator)) { toast("Naršyklė nepalaiko vietos nustatymo", true); return; }
  const b = $("field-here"), label = b.textContent; b.disabled = true; b.textContent = "Gaunama padėtis…";
  const reset = () => { b.disabled = false; b.textContent = label; };
  navigator.geolocation.getCurrentPosition(async (p) => {
    const reqId = newId(), prj = "VIETA";
    try {
      await graphPut(fieldRel("_lauko_uzklausa", prj), new Blob([JSON.stringify({ version: 2, mode: "here", req_id: reqId, project_no: prj, lat: p.coords.latitude, lon: p.coords.longitude, acc: Math.round(p.coords.accuracy), device: cfg.device, created: localISO() })], { type: "application/json" }), "application/json");
    } catch (e) { reset(); toast("Nepavyko išsiųsti užklausos: " + e.message, true); return; }
    reset();
    await cachePut("field_req", { prj, reqId, at: Date.now() });
    fld.cancel = false;
    resumeFieldWait();
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
    li.innerHTML = `<div class="t">${esc(p.title || p.project_no)}</div><div class="s">gautas ${esc(String(p.fetched || "").slice(5, 16).replace("T", " "))} · sklypų ${(p.parcels || []).length} · taškų ${(p.points || []).length}${(p.local_points || []).length ? " · vietinių " + p.local_points.length : ""}</div>` +
      `<div class="acts"><button type="button" class="btn small primary">Atidaryti</button><button type="button" class="btn small ghost">Ištrinti</button></div>`;
    const b = li.querySelectorAll("button");
    b[0].onclick = () => openFieldMap(p);
    b[1].onclick = async () => { if (confirm(`Ištrinti „${p.title || p.project_no}“ lauko paketą iš telefono?`)) { await fieldDelete(p.project_no); renderFieldStart(); toast("Ištrinta"); } };
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
      setWaitBanner(`Laukiama kompiuterio (${rq.prj === "VIETA" ? "mano vieta" : rq.prj})… ${Math.floor(sec / 60)}:${pad(sec % 60)}. Paprastai 1–3 min.`);
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
        toast(`Paketas ${j.title || rq.prj} gautas ir ištrintas iš OneDrive`);
        if (!$("v-field").hidden && !fld.mapOpen) { if (j.mode === "here") openFieldMap(j); else renderFieldStart(); }
        return;
      }
      for (let i = 0; i < 5 && !fld.cancel; i++) {
        await new Promise((s) => setTimeout(s, 1000));
        const sc = Math.round((Date.now() - t0) / 1000);
        setWaitBanner(`Laukiama kompiuterio (${rq.prj === "VIETA" ? "mano vieta" : rq.prj})… ${Math.floor(sc / 60)}:${pad(sc % 60)}. Paprastai 1–3 min.`);
      }
    }
  } finally { fld.waiting = false; $("field-cancel").hidden = true; }
}

/* ---- žemėlapis ---- */
function toM(lon, lat) { return [(lon - MAP.lon0) * MAP.kx, (lat - MAP.lat0) * MAP.ky]; }
function openFieldMap(pkg, opts = {}) {
  MAP.pkg = pkg; fld.mapOpen = true; fld.prj = pkg.project_no; fld.test = !!opts.test; MAP.marks = [];
  $("field-done").textContent = fld.test ? (opts.lpOnly ? "Uždaryti" : "Uždaryti testą") : "Baigti ir ištrinti";
  const bb = pkg.bbox || (() => { const p = (pkg.parcels[0] || {}).rings; return p ? [p[0][0][0], p[0][0][1], p[0][0][0], p[0][0][1]] : [25.28, 54.69, 25.28, 54.69]; })();
  MAP.lon0 = (bb[0] + bb[2]) / 2; MAP.lat0 = (bb[1] + bb[3]) / 2;
  MAP.kx = MAP.ky * Math.cos(MAP.lat0 * Math.PI / 180);
  MAP.polys = (pkg.parcels || []).map((p) => {
    const rings = p.rings.map((r) => r.map(([lo, la]) => toM(lo, la)));
    let sx = 0, sy = 0, n = 0; rings[0].forEach(([x, y]) => { sx += x; sy += y; n++; });
    return { cad: p.cad_no, role: p.role, rings, c: [sx / n, sy / n] };
  });
  MAP.pts = (pkg.points || []).map((p) => ({ label: p.label, xy: toM(p.lon, p.lat) }));
  MAP.lp = (pkg.local_points || []).map((p) => ({ xy: toM(p.lon, p.lat), at: p.at || {} })); MAP.lpSel = null; hideLpPop();
  $("field-lp").hidden = !MAP.lp.length;
  MAP.terr = pkg.terrain && Array.isArray(pkg.terrain.pts) && pkg.terrain.pts.length >= 3 ? pkg.terrain.pts.map(([lo, la, z]) => { const [x, y] = toM(lo, la); return [x, y, z]; }) : null;
  MAP.terrMemo = new Map();
  $("field-start").hidden = true; $("field-map-wrap").hidden = false;
  $("title").textContent = pkg.title || pkg.project_no;
  const notes = (pkg.notes || []).join(" · ");
  MAP.legendBase = (notes ? notes + " · " : "") + "Mėlyna – projekto sklypas, pilka – kaimynai, raudona – objekto taškas, oranžinis – jūsų žymos, mėlynas apskritimas – jūsų padėtis.";
  updateLegend();
  initCanvas(); fitMap(); startFieldWatch(); drawMap(); updateFieldInfo(); lpAttach(pkg);
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
    cv.addEventListener("pointerdown", (e) => { cv.setPointerCapture(e.pointerId); MAP.ptrs.set(e.pointerId, [e.clientX, e.clientY]); MAP.pinch = 0; MAP.tap = MAP.ptrs.size === 1 ? { x: e.clientX, y: e.clientY, t: Date.now(), moved: false } : null; });
    cv.addEventListener("pointermove", (e) => {
      if (!MAP.ptrs.has(e.pointerId)) return;
      if (MAP.tap && Math.hypot(e.clientX - MAP.tap.x, e.clientY - MAP.tap.y) > 9) MAP.tap.moved = true;
      const prev = MAP.ptrs.get(e.pointerId); MAP.ptrs.set(e.pointerId, [e.clientX, e.clientY]);
      if (MAP.ptrs.size === 1) { MAP.cx -= (e.clientX - prev[0]) / MAP.scale; MAP.cy += (e.clientY - prev[1]) / MAP.scale; MAP.follow = false; scheduleDraw(); }
      else if (MAP.ptrs.size === 2) {
        const [a, b] = [...MAP.ptrs.values()]; const d = Math.hypot(a[0] - b[0], a[1] - b[1]);
        if (MAP.pinch) { const r = cv.getBoundingClientRect(); zoomAt((a[0] + b[0]) / 2 - r.left, (a[1] + b[1]) / 2 - r.top, d / MAP.pinch); }
        MAP.pinch = d; MAP.follow = false;
      }
    });
    const up = (e) => {
      const tap = MAP.tap; MAP.tap = null;
      if (tap && !tap.moved && e.type === "pointerup" && MAP.ptrs.size === 1 && Date.now() - tap.t < 700) { const r = cv.getBoundingClientRect(); fieldTap(e.clientX - r.left, e.clientY - r.top); }
      MAP.ptrs.delete(e.pointerId); MAP.pinch = 0;
    };
    cv.addEventListener("pointerup", up); cv.addEventListener("pointercancel", up);
    cv.addEventListener("wheel", (e) => { e.preventDefault(); const r = cv.getBoundingClientRect(); zoomAt(e.clientX - r.left, e.clientY - r.top, e.deltaY < 0 ? 1.25 : 0.8); }, { passive: false });
    $("field-zin").onclick = () => zoomAt(MAP.W / 2, MAP.H / 2, 1.5);
    $("field-zout").onclick = () => zoomAt(MAP.W / 2, MAP.H / 2, 1 / 1.5);
    $("field-center").onclick = () => { if (!fld.pos) { toast("Padėtis dar negauta", true); return; } MAP.follow = true; centerOnUser(); };
    $("field-fit").onclick = () => { MAP.follow = false; fitMap(); drawMap(); };
    $("field-lp").onclick = () => { cfg.mapLp = !lpOn(); saveCfg(cfg); if (!lpOn()) hideLpPop(); toast(cfg.mapLp ? "Vietiniai punktai rodomi" : "Vietiniai punktai paslėpti"); drawMap(); };
    $("field-lp-x").onclick = () => { hideLpPop(); drawMap(); };
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
  if (!xs.length) { MAP.cx = 0; MAP.cy = 0; MAP.scale = Math.max(0.5, Math.min(MAP.W, MAP.H) / 160); return; }
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
  drawLocalPoints(c);
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

/* ---- vietiniai geodezijos punktai (Vilniaus_m_vietiniai_p.shp; stilius kaip QGIS Tmz_punktai.qml) ---- */
const LP_COL = { "1": ["#166dd7", 1, "rastas"], "2": ["#b80000", 1, "nerastas"], "3": ["#fafd21", 1, "pažeistas"], "5": ["#b977d7", 0.3, "istorinis"] };
const LP_ORDER = ["PPAV", "DATA", "Bukle", "Abrisas", "PTIP", "PKLAS", "ZTIP", "X", "Y", "H_VIETINE", "H_BALTIJOS", "RASTVIETOV", "Matavo", "photo", "ID"];
const lpOn = () => cfg.mapLp !== false;
const lpCol = (at) => LP_COL[String(at.Bukle || "").trim()] || ["#ffffff", 1, ""];
/* QGIS užrašas: PPAV, eilutė DATA + " m", eilutė " (a)" kai Abrisas netuščias */
function lpLabel(at) {
  const L = [];
  if (at.PPAV) L.push(at.PPAV);
  if (at.DATA) L.push(at.DATA + " m");
  if (String(at.Abrisas || "").trim()) L.push("(a)");
  return L;
}
function drawLocalPoints(c) {
  if (!MAP.lp.length || !lpOn()) return;
  const vis = [];
  for (const p of MAP.lp) {
    const x = sx_(p.xy[0]), y = sy_(p.xy[1]);
    if (x < -40 || y < -40 || x > MAP.W + 40 || y > MAP.H + 40) continue;
    vis.push([p, x, y]);
  }
  const labels = vis.length <= 40 || MAP.scale >= 1.5;
  c.lineJoin = "round";
  for (const [p, x, y] of vis) {
    const col = lpCol(p.at);
    c.globalAlpha = col[1]; c.fillStyle = col[0]; c.strokeStyle = "#232323"; c.lineWidth = 1;
    c.beginPath(); c.arc(x, y, 5.5, 0, 6.2832); c.fill(); c.stroke();
    c.globalAlpha = 1;
  }
  if (labels) {
    c.font = "11px system-ui,sans-serif"; c.textAlign = "left"; c.lineWidth = 3; c.strokeStyle = "rgba(255,255,255,.85)"; c.fillStyle = "#e31a1c";
    for (const [p, x, y] of vis) lpLabel(p.at).forEach((t, i) => { c.strokeText(t, x + 9, y - 2 + i * 12); c.fillText(t, x + 9, y - 2 + i * 12); });
  }
  if (MAP.lpSel) {
    const x = sx_(MAP.lpSel.xy[0]), y = sy_(MAP.lpSel.xy[1]);
    c.strokeStyle = "#ff8a00"; c.lineWidth = 3; c.beginPath(); c.arc(x, y, 10, 0, 6.2832); c.stroke();
  }
}
function hideLpPop() { MAP.lpSel = null; const el = $("field-lp-pop"); if (el) el.hidden = true; }
function showLpPop(p) {
  const at = p.at || {}, col = lpCol(at), keys = LP_ORDER.filter((k) => at[k] != null && at[k] !== "").concat(Object.keys(at).filter((k) => !LP_ORDER.includes(k)));
  let dist = "";
  if (fld.pos) { const [ux, uy] = toM(fld.pos.lon, fld.pos.lat); dist = ` · nuo jūsų ${fmtM(Math.hypot(ux - p.xy[0], uy - p.xy[1]))}`; }
  const rows = keys.map((k) => {
    let v = String(at[k]);
    if (k === "Bukle" && col[2]) v += " (" + col[2] + ")";
    if (/^[xy]1$/i.test(k) && v.trim() !== "" && isFinite(Number(v.replace(",", ".")))) v = fmt2(Number(v.replace(",", ".")));   // tik rodymui: X1/Y1 su 2 skaitmenimis po kablelio
    const val = /^https?:\/\//i.test(v) ? `<a href="${esc(v)}" target="_blank" rel="noopener">${esc(v)}</a>` : esc(v);
    return `<tr><td>${esc(k)}</td><td>${val}</td></tr>`;
  }).join("");
  $("field-lp-body").innerHTML = `<div><span class="dot" style="background:${col[0]};opacity:${Math.max(col[1], 0.5)}"></span><b>${esc(at.PPAV || at.ID || "Vietinis punktas")}</b>${col[2] ? " · " + esc(col[2]) : ""}<span class="muted small">${esc(dist)}</span></div>` +
    (rows ? `<table>${rows}</table>` : '<div class="muted small">Savybių nėra.</div>') + (p.idx != null ? lpEditorHtml(p) : "");
  $("field-lp-pop").hidden = false;
  lpBindEditor(p);
}
function fieldTap(sx, sy) {
  if (!MAP.lp.length || !lpOn()) { hideLpPop(); return; }
  let best = null, bd = 30;
  for (const p of MAP.lp) { const d = Math.hypot(sx_(p.xy[0]) - sx, sy_(p.xy[1]) - sy); if (d < bd) { bd = d; best = p; } }
  if (!best) { if (MAP.lpSel) { hideLpPop(); scheduleDraw(); } return; }
  MAP.lpSel = best; showLpPop(best); scheduleDraw();
}

/* ---- vietinių punktų duomenys iš OneDrive (telefonas skaito SHP/DBF tiesiai per Microsoft Graph; kompiuterio nereikia) ---- */
const LPD = { dir: "NAVAROTAI/Vietiniai punktai/Vilniaus m", base: "vilniaus_m_vietiniai_p", pts: null, meta: null, promise: null, maxAgeDays: 14 };
function updateLegend() {
  $("field-legend").textContent = (MAP.legendBase || "") + (MAP.lp.length ? ` Vietiniai punktai (${MAP.lp.length}): mėlyna – rastas, raudona – nerastas, geltona – pažeistas, violetinė – istorinis, balta – būklė nenurodyta; užrašas: PPAV, DATA m, (a) – yra abrisas. Palieskite punktą – savybės; 📌 – rodyti/slėpti.` : "");
}
function lks94ToWgs(E, N) {
  const a = 6378137.0, f = 1 / 298.257222101, e2 = f * (2 - f), ep2 = e2 / (1 - e2), k0 = 0.9998, E0 = 500000.0, lon0 = 24 * Math.PI / 180;
  const M = N / k0, mu = M / (a * (1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256)), e1 = (1 - Math.sqrt(1 - e2)) / (1 + Math.sqrt(1 - e2));
  const J1 = 3 * e1 / 2 - 27 * e1 ** 3 / 32, J2 = 21 * e1 ** 2 / 16 - 55 * e1 ** 4 / 32, J3 = 151 * e1 ** 3 / 96, J4 = 1097 * e1 ** 4 / 512;
  const fp = mu + J1 * Math.sin(2 * mu) + J2 * Math.sin(4 * mu) + J3 * Math.sin(6 * mu) + J4 * Math.sin(8 * mu);
  const sinf = Math.sin(fp), cosf = Math.cos(fp), tanf = Math.tan(fp), C1 = ep2 * cosf * cosf, T1 = tanf * tanf;
  const N1 = a / Math.sqrt(1 - e2 * sinf * sinf), R1 = a * (1 - e2) / Math.pow(1 - e2 * sinf * sinf, 1.5), D = (E - E0) / (N1 * k0);
  const lat = fp - (N1 * tanf / R1) * (D ** 2 / 2 - (5 + 3 * T1 + 10 * C1 - 4 * C1 ** 2 - 9 * ep2) * D ** 4 / 24 + (61 + 90 * T1 + 298 * C1 + 45 * T1 ** 2 - 252 * ep2 - 3 * C1 ** 2) * D ** 6 / 720);
  const lon = lon0 + (D - (1 + 2 * T1 + C1) * D ** 3 / 6 + (5 - 2 * C1 + 28 * T1 - 3 * C1 ** 2 + 8 * ep2 + 24 * T1 ** 2) * D ** 5 / 120) / cosf;
  return [lat * 180 / Math.PI, lon * 180 / Math.PI];
}
function lpParseShp(buf, shxBuf) {
  const dv = new DataView(buf), out = [];
  const rec = (off) => { const p = off + 8; if (p + 20 > buf.byteLength) return null; const t = dv.getInt32(p, true); return (t === 1 || t === 11 || t === 21) ? [dv.getFloat64(p + 4, true), dv.getFloat64(p + 12, true)] : null; };
  if (shxBuf) { const sv = new DataView(shxBuf), n = Math.floor((shxBuf.byteLength - 100) / 8); for (let i = 0; i < n; i++) out.push(rec(sv.getInt32(100 + i * 8, false) * 2)); }
  else { let p = 100; while (p + 8 <= buf.byteLength) { const w = dv.getInt32(p + 4, false); out.push(rec(p)); p += 8 + w * 2; } }
  return out;
}
function lpParseDbf(buf, cpg) {
  const u8 = new Uint8Array(buf), dv = new DataView(buf), nrec = dv.getUint32(4, true), hlen = dv.getUint16(8, true), rlen = dv.getUint16(10, true);
  const fields = [];
  for (let p = 32; p < hlen - 1 && u8[p] !== 0x0D; p += 32) { let nm = ""; for (let k = 0; k < 11 && u8[p + k]; k++) nm += String.fromCharCode(u8[p + k]); fields.push([nm.trim(), u8[p + 16]]); }
  const cp = /1250/.test(cpg || "") ? "windows-1250" : /1252/.test(cpg || "") ? "windows-1252" : "windows-1257";
  let u8d, cpd; try { u8d = new TextDecoder("utf-8", { fatal: true }); cpd = new TextDecoder(cp); } catch (_) { u8d = null; cpd = new TextDecoder(); }
  const dec = (a, b) => {
    let ascii = true; for (let i = a; i < b; i++) if (u8[i] > 127) { ascii = false; break; }
    if (ascii) { let s = ""; for (let i = a; i < b; i++) s += String.fromCharCode(u8[i]); return s.replace(/\0/g, "").trim(); }
    const sub = u8.subarray(a, b); let s; try { s = u8d ? u8d.decode(sub) : cpd.decode(sub); } catch (_) { s = cpd.decode(sub); } return s.replace(/\0/g, "").trim();
  };
  const rows = []; let p = hlen;
  for (let i = 0; i < nrec; i++, p += rlen) {
    if (p + rlen > buf.byteLength) break;
    if (u8[p] === 0x2A) { rows.push(null); continue; }
    let pos = p + 1; const row = {};
    for (const [nm, ln] of fields) { const v = dec(pos, pos + ln); pos += ln; if (v !== "") row[nm] = v; }
    rows.push(row);
  }
  return rows;
}
function lpEn(a, b) {
  if (!isFinite(a) || !isFinite(b)) return null;
  for (const [e, n] of [[a, b], [b, a]]) if (e > 100000 && e < 900000 && n > 5500000 && n < 6500000) return [e, n];
  return null;
}
function lpBuild(shpBuf, dbfBuf, shxBuf, cpg) {
  const geo = lpParseShp(shpBuf, shxBuf), rows = lpParseDbf(dbfBuf, cpg), pts = []; let g = 0, a = 0, bad = 0;
  const num = (v) => parseFloat(String(v == null ? "" : v).replace(",", "."));
  const pick = (at, k) => { const lk = k.toLowerCase(); for (const key in at) if (key.toLowerCase() === lk) return at[key]; return null; };
  for (let i = 0; i < rows.length; i++) {
    const at = rows[i]; if (!at) continue;
    let en = geo[i] ? lpEn(geo[i][0], geo[i][1]) : null;
    if (en) g++; else {
      for (const [kx, ky] of [["x_lks", "y_lks"], ["X1", "Y1"], ["X", "Y"]]) { en = lpEn(num(pick(at, kx)), num(pick(at, ky))); if (en) { a++; break; } }
    }
    if (!en) { bad++; continue; }
    const [lat, lon] = lks94ToWgs(en[0], en[1]);
    pts.push([Math.round(lat * 1e7) / 1e7, Math.round(lon * 1e7) / 1e7, at, i]);
  }
  return { pts, stats: { total: rows.length, geo: g, attr: a, bad } };
}
async function lpDownload(itemUrl) {
  const m = await graphFetch(`${itemUrl}?select=id,size,eTag,@microsoft.graph.downloadUrl`, { method: "GET" });
  if (m.status === 404) return null;
  if (!m.ok) throw httpErr(m, await m.text());
  const meta = await m.json(), dl = meta["@microsoft.graph.downloadUrl"];
  let r;
  if (dl) r = await fetch(dl); else r = await graphFetch(`${itemUrl}/content`, { method: "GET" });
  if (!r.ok) throw new Error("Atsisiuntimas: HTTP " + r.status);
  return { buf: await r.arrayBuffer(), etag: meta.eTag || "" };
}
async function lpFindFiles() {
  /* 1) numatytas aplankas, 2) paieška visame OneDrive */
  let items = null;
  let r = await graphFetch(`${GRAPH}/me/drive/root:/${encPath(LPD.dir)}:/children?$select=id,name,size,eTag,parentReference&$top=200`, { method: "GET" });
  if (r.ok) items = (await r.json()).value || [];
  const pickSet = (list) => {
    const lc = (list || []).map((x) => ({ x, n: String(x.name || "").toLowerCase() }));
    const shp = lc.find((i) => i.n.startsWith(LPD.base) && i.n.endsWith(".shp")) || lc.find((i) => i.n.includes("vietiniai_p") && i.n.endsWith(".shp"));
    if (!shp) return null;
    const stem = shp.n.slice(0, -4), f = (ext) => (lc.find((i) => i.n === stem + ext) || {}).x;
    return { shp: shp.x, dbf: f(".dbf"), shx: f(".shx"), cpg: f(".cpg") };
  };
  let set = pickSet(items);
  if (!set) {
    r = await graphFetch(`${GRAPH}/me/drive/root/search(q='vietiniai_p')?$select=id,name,parentReference&$top=50`, { method: "GET" });
    if (r.ok) {
      const found = ((await r.json()).value || []).filter((x) => /vietiniai_p.*\.shp$/i.test(x.name || ""));
      if (found.length) {
        const pid = found[0].parentReference && found[0].parentReference.id;
        if (pid) { const c = await graphFetch(`${GRAPH}/me/drive/items/${pid}/children?$select=id,name,size,eTag,parentReference&$top=200`, { method: "GET" }); if (c.ok) set = pickSet((await c.json()).value || []); }
      }
    }
  }
  if (!set || !set.dbf) throw new Error("OneDrive nerastas Vilniaus_m_vietiniai_p.shp / .dbf (aplankas NAVAROTAI/Vietiniai punktai/Vilniaus m)");
  return set;
}
/* X1/Y1 pakeitimų aptikimas atnaujinant punktus (koordinates keičia tik QGIS; telefonas jų nekeičia): grąžina laukiančių žurnalo įrašų sąrašą. */
const lpKey = (at) => `${at.ID == null ? "" : at.ID}|${at.PPAV == null ? "" : at.PPAV}`;
const lpNum = (v) => { const n = Number(String(v == null ? "" : v).replace(",", ".")); return v === "" || v == null || !isFinite(n) ? null : n; };
function lpXyChanges(oldPts, newPts, pending) {
  const out = new Map((pending || []).map((c) => [c.key, c]));
  if (oldPts && oldPts.length) {
    const old = new Map(oldPts.map((q) => [lpKey(q[2]), q[2]]));
    for (const q of newPts) {
      const a = q[2], b = old.get(lpKey(a)); if (!b) continue;
      const nx = lpNum(a.X1), ny = lpNum(a.Y1), ox = lpNum(b.X1), oy = lpNum(b.Y1);
      if (nx == null || ny == null || ox == null || oy == null) continue;
      if (Math.abs(nx - ox) >= 0.005 || Math.abs(ny - oy) >= 0.005) out.set(lpKey(a), { key: lpKey(a), ppav: a.PPAV || a.ID || "", date: localISO().slice(0, 10) });
    }
  }
  return [...out.values()];
}
/* Pastabų žurnalas: įrašai atskirti „; “; netilpus į 254 baitus išmetami seniausi. */
function lpAppendNote(existing, entry) {
  const parts = String(existing || "").split(";").map((x) => x.trim()).filter(Boolean);
  if (parts[parts.length - 1] !== entry) parts.push(entry);
  const len = (a) => new TextEncoder().encode(a.join("; ")).length;
  while (parts.length > 1 && len(parts) > LP_NOTE.len) parts.shift();
  return parts.join("; ");
}
async function lpLogXY() {
  const list = (LPD.meta && LPD.meta.xy) || []; if (!list.length) return 0;
  const byKey = new Map((LPD.pts || []).map((q) => [lpKey(q[2]), q]));
  const items = list.filter((c) => byKey.has(c.key)).map((c) => { const q = byKey.get(c.key); return { p: { idx: q[3], at: q[2] }, field: LP_NOTE.field, value: (row) => lpAppendNote(row[LP_NOTE.field], `${c.date} X1/Y1 pakeistos`) }; });
  if (!items.length) { LPD.meta.xy = []; await cachePut("lpdata", { meta: LPD.meta, pts: LPD.pts }); return 0; }
  const names = list.slice(0, 6).map((c) => c.ppav).join(", ") + (list.length > 6 ? ` ir dar ${list.length - 6}` : "");
  if (!confirm(`Pastebėta pakeistos X1/Y1 koordinatės (${list.length}): ${names}.\nĮrašyti žurnalo įrašą „data X1/Y1 pakeistos“ į lauką ${LP_NOTE.field}?` + (LPD.meta.note ? "" : `\n(SHP lauko „${LP_NOTE.field}“ dar nėra – jis bus sukurtas.)`))) return -1;
  const r = await lpSaveMany(items);
  const okKeys = new Set(r.done.map((x) => lpKey(x.it.p.at)));
  LPD.meta.xy = list.filter((c) => !okKeys.has(c.key));
  await cachePut("lpdata", { meta: LPD.meta, pts: LPD.pts });
  if (r.failed.length) log("X1/Y1 žurnalas: nepavyko " + r.failed.length + " – " + r.failed[0].err);
  return r.done.length;
}
async function lpFetchAll(silent) {
  if (!cfgReady() || !Auth.signedIn()) throw new Error("Reikia prisijungti (Nustatymai)");
  if (!navigator.onLine) throw new Error("Nėra ryšio");
  if (!silent) toast("Kraunami vietiniai punktai iš OneDrive…");
  const set = await lpFindFiles();
  const get = async (x) => x ? lpDownload(`${GRAPH}/me/drive/items/${x.id}`) : null;
  const dbf = await get(set.dbf), shp = await get(set.shp), shx = await get(set.shx), cpg = await get(set.cpg);
  if (!dbf || !shp) throw new Error("Nepavyko atsisiųsti SHP / DBF");
  const cpgText = cpg ? new TextDecoder().decode(cpg.buf) : "";
  const res = lpBuild(shp.buf, dbf.buf, shx ? shx.buf : null, cpgText);
  if (!res.pts.length) throw new Error(`SHP neturi tinkamų LKS94 koordinačių (įrašų ${res.stats.total})`);
  let oldPts = LPD.pts, oldMeta = LPD.meta;
  if (!oldPts || !oldPts.length) { try { const c = await cacheGet("lpdata"); if (c && c.pts && c.meta && c.meta.v === 2) { oldPts = c.pts; oldMeta = c.meta; } } catch (_) { /* */ } }
  const meta = { v: 2, ts: Date.now(), name: set.shp.name, etag: (shp.etag || "") + "|" + (dbf.etag || ""), stats: res.stats, note: lpDbfHasField(dbf.buf, LP_NOTE.field),
    dbf: { id: set.dbf.id, name: set.dbf.name, parent: (set.dbf.parentReference || {}).id || "", cpg: cpgText },
    xy: lpXyChanges(oldPts, res.pts, (oldMeta && oldMeta.xy) || []) };
  await cachePut("lpdata", { meta, pts: res.pts });
  LPD.pts = res.pts; LPD.meta = meta;
  log(`Vietiniai punktai: ${res.pts.length} (įrašų ${res.stats.total}, geometrija ${res.stats.geo}, atributai ${res.stats.attr}, be koordinačių ${res.stats.bad})`);
  return res.pts;
}
/* Grąžina visus punktus: atmintyje → IndexedDB → OneDrive. Jei talpykla sena (> maxAgeDays) ir yra ryšys – atnaujina fone. */
function lpLoad(force) {
  if (LPD.pts && !force) return Promise.resolve(LPD.pts);
  if (LPD.promise) return LPD.promise;
  LPD.promise = (async () => {
    if (!force) {
      const c = await cacheGet("lpdata");
      if (c && c.pts && c.pts.length && c.meta && c.meta.v === 2) {
        LPD.pts = c.pts; LPD.meta = c.meta || {};
        if (navigator.onLine && cfgReady() && Auth.signedIn() && Date.now() - ((c.meta || {}).ts || 0) > LPD.maxAgeDays * 86400000) lpFetchAll(true).then(() => toast("Vietinių punktų duomenys atnaujinti")).catch((e) => log("Vietinių punktų atnaujinimas: " + e.message));
        return LPD.pts;
      }
    }
    return lpFetchAll(false);
  })().finally(() => { LPD.promise = null; });
  return LPD.promise;
}
function lpNear(lat, lon, rad, max) {
  const dlat = rad / 111195, dlon = rad / (111195 * Math.cos(lat * Math.PI / 180)), out = [];
  for (const p of LPD.pts || []) if (Math.abs(p[0] - lat) <= dlat && Math.abs(p[1] - lon) <= dlon) out.push(p);
  if (out.length > max) { out.sort((a, b) => Math.hypot(a[0] - lat, (a[1] - lon) * 0.58) - Math.hypot(b[0] - lat, (b[1] - lon) * 0.58)); out.length = max; }
  return out;
}
async function lpAttach(pkg) {
  if (pkg.local_points && pkg.local_points.length) { updateLegend(); return; }
  try {
    const pts = await lpLoad(false);
    if (!pts || !fld.mapOpen || MAP.pkg !== pkg) return;
    const bb = pkg.bbox, lat = MAP.lat0, lon = MAP.lon0;
    let rad = 450;
    if (bb) rad += Math.max((bb[2] - bb[0]) * MAP.kx, (bb[3] - bb[1]) * MAP.ky) / 2;
    const near = lpNear(lat, lon, rad, 1500);
    MAP.lp = near.map(([la, lo, at, idx]) => ({ xy: toM(lo, la), at, idx })); MAP.lpSel = null;
    $("field-lp").hidden = !MAP.lp.length; updateLegend(); scheduleDraw();
    if (!MAP.lp.length) toast("Šalia vietinių punktų nėra (≤ " + Math.round(rad) + " m)");
  } catch (e) { log("Vietiniai punktai: " + e.message); if (fld.mapOpen) toast("Vietinių punktų nėra: " + e.message, true); }
}
/* Žemėlapis tik su vietiniais punktais aplink mano vietą – kompiuterio nereikia */
function lpNearMe() {
  if (!("geolocation" in navigator)) { toast("Naršyklė nepalaiko vietos nustatymo", true); return; }
  const b = $("field-lpnear"), label = b.textContent; b.disabled = true; b.textContent = "Gaunama padėtis…";
  const reset = () => { b.disabled = false; b.textContent = label; };
  navigator.geolocation.getCurrentPosition((p) => {
    reset();
    const lat = p.coords.latitude, lon = p.coords.longitude, dy = 80 / 111195, dx = 80 / (111195 * Math.cos(lat * Math.PI / 180));
    openFieldMap({ project_no: "VIETINIAI", title: "Vietiniai punktai", fetched: localISO(), points: [], parcels: [], notes: [], bbox: [lon - dx, lat - dy, lon + dx, lat + dy] }, { test: true, lpOnly: true });
  }, (e) => { reset(); toast("Padėtis negauta: " + e.message + ". Leiskite naršyklei naudoti vietą.", true); },
  { enableHighAccuracy: true, maximumAge: 0, timeout: 20000 });
}
/* ---- punkto redagavimas telefone: RASTVIETOV reikšmė (rašoma į DBF OneDrive) ir nuotraukos (OneDrive aplankas „Qgis punktai“, vardas pagal PPAV) ---- */
const LPP = { dir: "Qgis punktai", list: null, at: 0 };
const lpFileKey = (s) => String(s == null ? "" : s).replace(/[\\/:*?"<>|]/g, "_").trim();
async function lpPhotoIndex(force) {
  if (!force && LPP.list && Date.now() - LPP.at < 300000) return LPP.list;
  let url = `${GRAPH}/me/drive/root:/${encPath(LPP.dir)}:/children?$select=id,name,webUrl,file&$top=999`; const all = [];
  while (url) {
    const r = await graphFetch(url, { method: "GET" });
    if (r.status === 404) break;
    if (!r.ok) throw httpErr(r, await r.text());
    const j = await r.json(); all.push(...(j.value || []).filter((x) => x.file)); url = j["@odata.nextLink"] || null;
  }
  LPP.list = all; LPP.at = Date.now(); return all;
}
function lpPhotosOf(list, ppav) {
  const k = lpFileKey(ppav).toLowerCase(); if (!k) return [];
  return list.filter((x) => { const stem = x.name.replace(/\.[^.]+$/, "").toLowerCase(); return stem === k || (stem.startsWith(k) && /^(_\d+| \(\d+\))$/.test(stem.slice(k.length))); });
}
async function lpAddPhoto(p, file) {
  const ppav = lpFileKey(p.at.PPAV || p.at.ID);
  if (!ppav) throw new Error("Punktas neturi PPAV – vardo nuotraukai nėra");
  if (!navigator.onLine) throw new Error("Nėra ryšio");
  const taken = new Set((await lpPhotoIndex(true)).map((x) => x.name.toLowerCase()));
  const blob = await prepareImage(file, 2000, 0.85);
  const ext = blob.type === "image/jpeg" ? ".jpg" : ((String(file.name || "").match(/\.[A-Za-z0-9]+$/) || [".jpg"])[0].toLowerCase());
  let name = ppav + ext, n = 1;
  while (taken.has(name.toLowerCase())) { n++; name = `${ppav}_${n}${ext}`; }
  await graphPut(`${LPP.dir}/${name}`, blob, blob.type || "image/jpeg");
  LPP.at = 0;
  let link = "";
  try { const it = (await lpPhotoIndex(true)).find((x) => x.name === name); link = it && it.webUrl ? it.webUrl : ""; } catch (_) { /* */ }
  return { name, link };
}
/* Nuotraukos nuoroda → atributas photo (jei tuščias; jei jau užpildytas – klausiama). Grąžina tekstą apie rezultatą. */
async function lpLinkPhoto(p, r) {
  if (!r.link) return "nuorodos į photo neįrašyta (nerasta OneDrive nuoroda)";
  if (p.at.photo && p.at.photo !== r.link && !confirm(`Laukas photo jau užpildytas:\n${p.at.photo}\n\nPakeisti nauja nuoroda?`)) return "photo paliktas nepakeistas";
  let val = r.link;
  try { await lpSaveAttr(p, "photo", val); return "nuoroda įrašyta į photo"; }
  catch (e) {
    if (/Per ilga/.test(e.message)) { val = `${LPP.dir}/${r.name}`; try { await lpSaveAttr(p, "photo", val); return "pilna nuoroda netilpo – į photo įrašytas kelias " + val; } catch (e2) { return "photo neįrašytas: " + e2.message; } }
    return "photo neįrašytas: " + e.message;
  }
}
const LP_EDITABLE = ["rastvietov", "photo", "pastaba"];   // koordinatės ir kiti laukai telefone NIEKADA nekeičiami – tik peržiūra
const LP_NOTE = { field: "Pastaba", len: 254 };   // naujas tekstinis atributas pastaboms (sukuriamas pirmo išsaugojimo metu)
function lpDbfHasField(buf, name) {
  const u8 = new Uint8Array(buf), hlen = new DataView(buf).getUint16(8, true);
  for (let p = 32; p < hlen - 1 && u8[p] !== 0x0D; p += 32) {
    let nm = ""; for (let k = 0; k < 11 && u8[p + k]; k++) nm += String.fromCharCode(u8[p + k]);
    if (nm.trim().toLowerCase() === name.toLowerCase()) return true;
  }
  return false;
}
/* Prideda tekstinį (C) lauką į DBF pabaigoje: antraštėje įrašomas aprašas, kiekvieno įrašo gale pridedami tarpai. Kiti duomenys nekeičiami. Jei laukas jau yra – grąžina tą patį buferį. */
function lpDbfAddField(buf, name, len) {
  if (lpDbfHasField(buf, name)) return buf;
  const u8 = new Uint8Array(buf), dv = new DataView(buf), nrec = dv.getUint32(4, true), hlen = dv.getUint16(8, true), rlen = dv.getUint16(10, true);
  let p = 32; while (p < hlen - 1 && u8[p] !== 0x0D) p += 32;
  if (u8[p] !== 0x0D) throw new Error("Netinkama DBF antraštė – lauko sukurti negalima");
  const nh = hlen + 32, nr = rlen + len, body = hlen + nrec * rlen;
  if (body > buf.byteLength) throw new Error("DBF failas nepilnas – lauko sukurti negalima");
  if (nh > 65535 || nr > 65535 || !/^[A-Za-z][A-Za-z0-9_]{0,9}$/.test(name)) throw new Error("Lauko sukurti negalima");
  const out = new Uint8Array(nh + nrec * nr + (buf.byteLength - body)), odv = new DataView(out.buffer);
  out.set(u8.subarray(0, p), 0);
  for (let i = 0; i < name.length; i++) out[p + i] = name.charCodeAt(i);
  out[p + 11] = 0x43; out[p + 16] = len;                                   // tipas C, ilgis
  if (u8[0] === 0x30 || u8[0] === 0x31 || u8[0] === 0x32) odv.setUint32(p + 12, rlen, true);   // FoxPro: lauko poslinkis įraše
  out.set(u8.subarray(p, hlen), p + 32);                                    // terminatorius ir kas už jo
  odv.setUint16(8, nh, true); odv.setUint16(10, nr, true);
  for (let i = 0; i < nrec; i++) { const src = hlen + i * rlen, dst = nh + i * nr; out.set(u8.subarray(src, src + rlen), dst); out.fill(0x20, dst + rlen, dst + nr); }
  out.set(u8.subarray(body), nh + nrec * nr);                              // EOF žymė
  return out.buffer;
}
function lpDbfSet(buf, cpg, idx, ident, field, value, rows) {
  if (!LP_EDITABLE.includes(String(field).toLowerCase())) throw new Error("Laukas " + field + " telefone nekeičiamas");
  const u8 = new Uint8Array(buf), dv = new DataView(buf), nrec = dv.getUint32(4, true), hlen = dv.getUint16(8, true), rlen = dv.getUint16(10, true);
  let f = null, off = 1;
  for (let p = 32; p < hlen - 1 && u8[p] !== 0x0D; p += 32) {
    let nm = ""; for (let k = 0; k < 11 && u8[p + k]; k++) nm += String.fromCharCode(u8[p + k]);
    const len = u8[p + 16];
    if (nm.trim().toLowerCase() === field.toLowerCase()) f = { type: String.fromCharCode(u8[p + 11]), len, dec: u8[p + 17], off };
    off += len;
  }
  if (!f) throw new Error("DBF nėra lauko " + field);
  if (f.type !== "C" && f.type !== "N") throw new Error(`Laukas ${field} turi nepalaikomą tipą (${f.type})`);
  if (idx < 0 || idx >= nrec) throw new Error("Įrašas nerastas – atnaujinkite punktus (Nustatymai)");
  const rec = hlen + idx * rlen;
  if (rec + rlen > buf.byteLength || u8[rec] === 0x2A) throw new Error("Įrašas pašalintas arba pasikeitė – atnaujinkite punktus (Nustatymai)");
  const cur = (rows || lpParseDbf(buf, cpg))[idx] || {};
  for (const k of ["PPAV", "ID"]) if (String(cur[k] || "") !== String(ident[k] || "")) throw new Error("SHP pasikeitė kitur (neatitinka punktas) – atnaujinkite punktus (Nustatymai)");
  if (f.type === "N") {      /* skaitinis laukas (pvz. RASTVIETOV): sveikas skaičius, dešinėje lygiuotas, tuščias = tarpai */
    const v = String(value).trim();
    if (v !== "" && !(f.dec ? /^-?\d+([.,]\d+)?$/ : /^-?\d+$/).test(v)) throw new Error(`Laukas ${field} turi būti ${f.dec ? "skaičius" : "sveikas skaičius"}`);
    const s = v === "" ? "" : (f.dec ? Number(v.replace(",", ".")).toFixed(f.dec) : String(parseInt(v, 10)));
    if (s.length > f.len) throw new Error(`Per didelė reikšmė: laukas ${field} talpina ${f.len} simbolių`);
    const st = rec + f.off, pad = s.padStart(f.len, " ");
    for (let i = 0; i < f.len; i++) u8[st + i] = pad.charCodeAt(i);
    return;
  }
  let utf8 = /utf/i.test(cpg || "");
  if (!utf8 && !/125\d/.test(cpg || "")) {
    try { let hi = false; for (let i = hlen; i < u8.length; i++) if (u8[i] > 127) { hi = true; break; } if (hi) { new TextDecoder("utf-8", { fatal: true }).decode(u8.subarray(hlen)); utf8 = true; } } catch (_) { utf8 = false; }
  }
  const bytes = []; let cut = false;
  if (utf8) { const te = new TextEncoder(); for (const ch of value) { const b = te.encode(ch); if (bytes.length + b.length > f.len) { cut = true; break; } bytes.push(...b); } }
  else {
    const dec = new TextDecoder(/1250/.test(cpg || "") ? "windows-1250" : /1252/.test(cpg || "") ? "windows-1252" : "windows-1257"), map = new Map();
    for (let b = 128; b < 256; b++) { const c = dec.decode(new Uint8Array([b])); if (!map.has(c)) map.set(c, b); }
    for (const ch of value) {
      const b = ch.charCodeAt(0) < 128 ? ch.charCodeAt(0) : map.get(ch);
      if (b == null) throw new Error(`Simbolio „${ch}“ negalima įrašyti į DBF koduotę`);
      if (bytes.length + 1 > f.len) { cut = true; break; } bytes.push(b);
    }
  }
  if (cut) throw new Error(`Per ilga reikšmė: laukas ${field} talpina ${f.len} baitų`);
  const start = rec + f.off;
  for (let i = 0; i < f.len; i++) u8[start + i] = i < bytes.length ? bytes[i] : 0x20;
}
async function lpCommitDbf(m, d, orig) {   // dienos atsarginė kopija (tame pačiame aplanke) + įrašymas su If-Match
  const itemUrl = `${GRAPH}/me/drive/items/${m.id}`;
  const stamp = localISO().slice(0, 10).replace(/-/g, "");
  const b = await graphFetch(`${GRAPH}/me/drive/items/${m.parent}:/${encodeURIComponent(m.name)}.bak${stamp}:/content?@microsoft.graph.conflictBehavior=fail`,
    { method: "PUT", headers: { "Content-Type": "application/octet-stream" }, body: new Blob([orig]) });
  if (!b.ok && b.status !== 409) throw new Error("Nepavyko sukurti atsarginės DBF kopijos (HTTP " + b.status + ") – pakeitimas neįrašytas");
  const r = await graphFetch(`${itemUrl}/content`, { method: "PUT", headers: { "Content-Type": "application/octet-stream", "If-Match": d.etag }, body: new Blob([d.buf]) });
  if (r.status === 412) throw new Error("SHP pakeistas kitur – atnaujinkite punktus ir bandykite dar kartą");
  if (!r.ok) throw httpErr(r, await r.text());
}
async function lpSaveAttr(p, field, value) {
  if (!navigator.onLine) throw new Error("Nėra ryšio");
  const m = LPD.meta && LPD.meta.dbf;
  if (!m || !m.id || !m.parent) throw new Error("Trūksta DBF nuorodos – atnaujinkite punktus (Nustatymai)");
  const itemUrl = `${GRAPH}/me/drive/items/${m.id}`;
  const d = await lpDownload(itemUrl);
  if (!d) throw new Error("DBF OneDrive nerastas");
  const orig = d.buf.slice(0);
  let created = false;
  if (String(field).toLowerCase() === LP_NOTE.field.toLowerCase() && !lpDbfHasField(d.buf, LP_NOTE.field)) { d.buf = lpDbfAddField(d.buf, LP_NOTE.field, LP_NOTE.len); created = true; }
  lpDbfSet(d.buf, m.cpg, p.idx, p.at, field, value);              // klaidos – dar prieš bet kokį įrašymą
  await lpCommitDbf(m, d, orig);
  if (value) p.at[field] = value; else delete p.at[field];
  log(`${field} pakeista: ${p.at.PPAV || p.at.ID} → ${value || "(tuščia)"}${created ? " (sukurtas naujas laukas)" : ""}`);
  if (created) LPD.meta.note = true;
  await cachePut("lpdata", { meta: LPD.meta, pts: LPD.pts });
  return { created };
}
/* Keli laukai vienu įrašymu (viena atsarginė kopija, vienas PUT). items: [{p:{idx,at}, field, value, alt}] – alt: atsarginė reikšmė, jei value netilpo. */
async function lpSaveMany(items) {
  if (!navigator.onLine) throw new Error("Nėra ryšio");
  const m = LPD.meta && LPD.meta.dbf;
  if (!m || !m.id || !m.parent) throw new Error("Trūksta DBF nuorodos – atnaujinkite punktus (Nustatymai)");
  const d = await lpDownload(`${GRAPH}/me/drive/items/${m.id}`);
  if (!d) throw new Error("DBF OneDrive nerastas");
  const orig = d.buf.slice(0), done = [], failed = [];
  if (items.some((it) => String(it.field).toLowerCase() === LP_NOTE.field.toLowerCase()) && !lpDbfHasField(d.buf, LP_NOTE.field)) d.buf = lpDbfAddField(d.buf, LP_NOTE.field, LP_NOTE.len);
  const rows = lpParseDbf(d.buf, m.cpg);
  for (const it of items) {
    let err = null;
    for (const v0 of [it.value, it.alt]) {
      if (v0 == null) continue;
      const v = typeof v0 === "function" ? v0(rows[it.p.idx] || {}) : v0;
      try { lpDbfSet(d.buf, m.cpg, it.p.idx, it.p.at, it.field, v, rows); done.push({ it, value: v }); err = null; break; }
      catch (e) { err = e; if (!/Per ilga/.test(e.message)) break; }
    }
    if (err) failed.push({ it, err: err.message });
  }
  if (done.length) {
    await lpCommitDbf(m, d, orig);
    for (const x of done) { if (x.value) x.it.p.at[x.it.field] = x.value; else delete x.it.p.at[x.it.field]; }
    log(`SHP: įrašyta laukų ${done.length}${failed.length ? ", nepavyko " + failed.length : ""}`);
    if (items.some((it) => String(it.field).toLowerCase() === LP_NOTE.field.toLowerCase())) LPD.meta.note = true;
    await cachePut("lpdata", { meta: LPD.meta, pts: LPD.pts });
  }
  return { done, failed };
}
function lpBestPhoto(list, ppav) {
  const k = lpFileKey(ppav).toLowerCase(); if (!k) return null;
  const rank = (x) => { const stem = x.name.replace(/\.[^.]+$/, "").toLowerCase(); const n = stem.slice(k.length).match(/\d+/); return stem === k ? 0 : (n ? Number(n[0]) : 999); };
  const arr = lpPhotosOf(list, ppav).filter((x) => x.webUrl).sort((a, b) => rank(a) - rank(b) || a.name.localeCompare(b.name));
  return arr[0] || null;
}
/* Visų punktų, kurių photo tuščias, bet OneDrive aplanke „Qgis punktai“ yra nuotrauka (vardas pagal PPAV), nuorodos įrašomos į photo. */
async function lpLinkAll() {
  const info = $("set-lpinfo"), btn = $("set-lplink");
  if (!navigator.onLine || !cfgReady() || !Auth.signedIn()) { toast("Reikia ryšio ir prisijungimo", true); return; }
  if (!LPD.pts || !LPD.pts.length) { toast("Pirmiausia atnaujinkite vietinius punktus", true); return; }
  btn.disabled = true; info.textContent = "Ieškoma nuotraukų…";
  try {
    const list = await lpPhotoIndex(true), items = []; let had = 0, none = 0;
    for (const q of LPD.pts) {
      const at = q[2], ph = lpBestPhoto(list, at.PPAV || at.ID);
      if (!ph) { none++; continue; }
      if (at.photo) { had++; continue; }
      items.push({ p: { idx: q[3], at }, field: "photo", value: ph.webUrl, alt: `${LPP.dir}/${ph.name}` });
    }
    if (!items.length) { info.textContent = `Naujų nuorodų nėra (photo jau užpildyta: ${had}, be nuotraukos: ${none})`; return; }
    if (!confirm(`Nuotraukų rasta ${items.length} punktams, kurių photo tuščias (dar ${had} jau užpildyti).\nĮrašyti nuorodas į SHP (OneDrive)?`)) { info.textContent = ""; return; }
    info.textContent = "Rašoma…";
    const r = await lpSaveMany(items);
    info.textContent = `Susieta: ${r.done.length}; jau turėjo: ${had}; be nuotraukos: ${none}` + (r.failed.length ? `; nepavyko: ${r.failed.length} (${r.failed[0].err})` : "");
    toast(`Nuotraukų nuorodos įrašytos: ${r.done.length}`);
    if (MAP.lpSel) showLpPop(MAP.lpSel);
  } catch (e) { info.textContent = "Klaida: " + e.message; toast(e.message, true); log("Susiejimas: " + e.message); }
  finally { btn.disabled = false; }
}
function lpEditorHtml(p) {
  const vals = [...new Set((LPD.pts || []).map((q) => q[2].RASTVIETOV).filter(Boolean))].sort((a, b) => Number(a) - Number(b) || String(a).localeCompare(String(b))).slice(0, 80);
  return `<div class="lp-edit"><label for="lp-rv" class="small muted">RASTVIETOV</label><div class="row"><input id="lp-rv" list="lp-rv-list" inputmode="numeric" maxlength="12" value="${esc(p.at.RASTVIETOV || "")}">` +
    `<datalist id="lp-rv-list">${vals.map((v) => `<option value="${esc(v)}">`).join("")}</datalist><button type="button" class="btn small primary" id="lp-rv-save">Saugoti</button></div></div>` +
    `<div class="lp-edit"><label for="lp-note" class="small muted">Pastaba</label><textarea id="lp-note" rows="2" maxlength="240" placeholder="Pastaba apie punktą…">${esc(p.at.Pastaba || "")}</textarea><button type="button" class="btn small primary" id="lp-note-save">Saugoti pastabą</button></div>` +
    `<div class="lp-ph"><div id="lp-ph-list" class="muted small">Nuotraukos: tikrinama…</div><button type="button" class="btn small" id="lp-ph-add">📷 Pridėti nuotrauką</button><input type="file" id="lp-ph-file" accept="image/*" hidden></div>`;
}
async function lpShowPhotos(p) {
  const el = $("lp-ph-list"); if (!el) return;
  if (!navigator.onLine || !cfgReady() || !Auth.signedIn()) { el.textContent = "Nuotraukų sąrašui reikia ryšio ir prisijungimo"; return; }
  try {
    const list = lpPhotosOf(await lpPhotoIndex(false), p.at.PPAV || p.at.ID), cur = $("lp-ph-list");
    if (MAP.lpSel !== p || !cur) return;
    const best = lpBestPhoto(list, p.at.PPAV || p.at.ID);
    cur.innerHTML = (list.length ? "Nuotraukos: " + list.map((x) => `<a href="${esc(x.webUrl)}" target="_blank" rel="noopener">${esc(x.name)}</a>`).join(" · ") : "Nuotraukų nėra") +
      (best && !p.at.photo ? ` <button type="button" class="btn small" id="lp-ph-link">Susieti su photo</button>` : "");
    const lb = $("lp-ph-link");
    if (lb) lb.onclick = async () => { lb.disabled = true; lb.textContent = "Rašoma…"; try { const res = await lpLinkPhoto(p, { link: best.webUrl, name: best.name }); toast(res); if (MAP.lpSel === p) showLpPop(p); } catch (e) { toast(e.message, true); lb.disabled = false; lb.textContent = "Susieti su photo"; } };
  } catch (e) { el.textContent = "Nuotraukų sąrašo gauti nepavyko: " + e.message; }
}
function lpBindEditor(p) {
  if (p.idx == null || !$("lp-rv-save")) return;
  $("lp-rv-save").onclick = async () => {
    const btn = $("lp-rv-save"), v = $("lp-rv").value.trim();
    if (v === (p.at.RASTVIETOV || "")) { toast("Reikšmė nepakeista"); return; }
    if (v !== "" && !/^-?\d+$/.test(v)) { toast("RASTVIETOV turi būti sveikas skaičius", true); return; }
    if (!confirm(`Pakeisti ${p.at.PPAV || p.at.ID} RASTVIETOV į „${v || "(tuščia)"}“ ir įrašyti į SHP (OneDrive)?`)) return;
    btn.disabled = true; btn.textContent = "Saugoma…";
    try { await lpSaveAttr(p, "RASTVIETOV", v); toast("RASTVIETOV išsaugota"); if (MAP.lpSel === p) showLpPop(p); }
    catch (e) { toast(e.message, true); log("RASTVIETOV: " + e.message); btn.disabled = false; btn.textContent = "Saugoti"; }
  };
  $("lp-note-save").onclick = async () => {
    const btn = $("lp-note-save"), v = $("lp-note").value.replace(/\s*[\r\n]+\s*/g, " ").trim();
    if (v === (p.at.Pastaba || "")) { toast("Pastaba nepakeista"); return; }
    const first = !(LPD.meta && LPD.meta.note);
    if (!confirm((first ? `SHP dar nėra lauko „${LP_NOTE.field}“ – jis bus sukurtas (visiems punktams tuščias).\n\n` : "") + `Įrašyti pastabą punktui ${p.at.PPAV || p.at.ID}?`)) return;
    btn.disabled = true; btn.textContent = "Saugoma…";
    try { const r = await lpSaveAttr(p, LP_NOTE.field, v); toast(r.created ? "Laukas „Pastaba“ sukurtas, pastaba išsaugota" : "Pastaba išsaugota"); if (MAP.lpSel === p) showLpPop(p); }
    catch (e) { toast(e.message, true); log("Pastaba: " + e.message); btn.disabled = false; btn.textContent = "Saugoti pastabą"; }
  };
  $("lp-ph-add").onclick = () => $("lp-ph-file").click();
  $("lp-ph-file").onchange = async (ev) => {
    const f = ev.target.files && ev.target.files[0]; ev.target.value = ""; if (!f) return;
    const btn = $("lp-ph-add"); btn.disabled = true; btn.textContent = "Keliama…";
    try { const r = await lpAddPhoto(p, f); btn.textContent = "Rašoma nuoroda…"; const res = await lpLinkPhoto(p, r); toast(`Nuotrauka ${r.name} įkelta; ${res}`); if (MAP.lpSel === p) showLpPop(p); }
    catch (e) { toast(e.message, true); log("Nuotrauka: " + e.message); }
    if ($("lp-ph-add")) { $("lp-ph-add").disabled = false; $("lp-ph-add").textContent = "📷 Pridėti nuotrauką"; }
    lpShowPhotos(p);
  };
  lpShowPhotos(p);
}
async function lpReloadNow() {
  const info = $("set-lpinfo"); info.textContent = "Kraunama…";
  try {
    const pts = await lpFetchAll(true); info.textContent = `Gauta punktų: ${pts.length}`; toast("Vietiniai punktai atnaujinti");
    const n = await lpLogXY();
    if (n > 0) { info.textContent += `; X1/Y1 pakeitimų žurnale (Pastaba): ${n}`; toast(`X1/Y1 pakeitimai įrašyti į ${LP_NOTE.field}: ${n}`); }
    else if (n < 0) info.textContent += "; X1/Y1 pakeitimų žurnalas neįrašytas (paklaus kitą kartą)";
  } catch (e) { info.textContent = "Klaida: " + e.message; toast(e.message, true); }
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
  $("field-lpnear").onclick = lpNearMe;
  $("set-lpreload").onclick = lpReloadNow;
  $("set-lplink").onclick = lpLinkAll;
  $("field-here").onclick = fieldHere;
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
const AR = { on: false, stream: null, R: null, off: 0, fov: 62, h: 1.5, raf: 0, absolute: false, evAt: 0, W: 0, H: 0, dpr: 1, hudAt: 0, lis: [], ev: { abs: 0, rel: 0, nul: 0 }, sensor: null, via: "", perm: "", pitch: 0, fix: null, sm: null, smAt: 0, bh: 1 };
function arLoadPrefs() { if (typeof cfg.arFov === "number") AR.fov = cfg.arFov; if (typeof cfg.arH === "number") AR.h = cfg.arH; if (typeof cfg.arBh === "number") AR.bh = cfg.arBh; }
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
  const dep = -dz;
  if (!AR.pitch) return [dx, dy, dep];
  const q = AR.pitch * Math.PI / 180, cq = Math.cos(q), sq = Math.sin(q);   // kameros pasvirimo pataisa (aukštyn +)
  return [dx, dy * cq - dep * sq, dy * sq + dep * cq];
}
/* Telefono padėtis AR vaizdui: išlyginama (GPS šuoliai nešokdina taškų) arba užfiksuota ranka. Metrai kaip ir MAP. */
function arPos() {
  if (AR.fix) return AR.fix;
  if (!fld.pos) return null;
  const p = toM(fld.pos.lon, fld.pos.lat);
  if (!AR.sm) AR.sm = p.slice();
  else if (fld.lastAt !== AR.smAt) { const a = fld.pos.acc <= 1 ? 0.8 : 0.25; AR.sm = [AR.sm[0] + a * (p[0] - AR.sm[0]), AR.sm[1] + a * (p[1] - AR.sm[1])]; }
  AR.smAt = fld.lastAt;
  return AR.sm;
}
function arElev() {   // kameros ašies aukštis virš horizonto, laipsniais (be pataisos)
  if (!AR.R) return null;
  return Math.asin(Math.max(-1, Math.min(1, -AR.R[2][2]))) * 180 / Math.PI;
}
function arScreen(p) {   // p = kameros koordinatės, gylis > 0
  const f = (AR.H / 2) / Math.tan(AR.fov * Math.PI / 360);
  return [AR.W / 2 + f * p[0] / p[2], AR.H / 2 - f * p[1] / p[2]];
}
/* Žemės aukštis (m, LiDAR „Ground“) taške x, y (MAP metrai): vietinė plokštuma iš 8 artimiausių taškų, esant blogai – atvirkštinių atstumų svoris.
   Be reljefo paketo grąžina 0 (žemė laikoma lygia). */
function terrZ(x, y) {
  const T = MAP.terr; if (!T) return 0;
  const key = Math.round(x * 10) + "," + Math.round(y * 10), memo = MAP.terrMemo;
  const hit = memo.get(key); if (hit !== undefined) return hit;
  const best = [];
  for (const p of T) {
    const dx = p[0] - x, dy = p[1] - y, d2 = dx * dx + dy * dy;
    if (best.length < 8 || d2 < best[best.length - 1][0]) { best.push([d2, dx, dy, p[2]]); best.sort((a, b) => a[0] - b[0]); if (best.length > 8) best.pop(); }
  }
  let sw = 0, swz = 0; for (const b of best) { const w = 1 / (b[0] + 4); sw += w; swz += w * b[3]; }
  let v = swz / sw;
  if (best.length >= 4 && best[0][0] > 0.01) {   // svertinė plokštuma z = a + b·dx + c·dy; reikšmė taške = a
    let S = [[0, 0, 0], [0, 0, 0], [0, 0, 0]], R = [0, 0, 0];
    for (const [d2, dx, dy, z] of best) { const w = 1 / (d2 + 4), f = [1, dx, dy]; for (let i = 0; i < 3; i++) { R[i] += w * f[i] * z; for (let j = 0; j < 3; j++) S[i][j] += w * f[i] * f[j]; } }
    const det = (m) => m[0][0] * (m[1][1] * m[2][2] - m[1][2] * m[2][1]) - m[0][1] * (m[1][0] * m[2][2] - m[1][2] * m[2][0]) + m[0][2] * (m[1][0] * m[2][1] - m[1][1] * m[2][0]);
    const D = det(S);
    if (Math.abs(D) > 1e-9) {
      const Sa = S.map((r, i) => [R[i], r[1], r[2]]), a = det(Sa) / D;
      if (Math.abs(a - v) < 3) v = a;
    }
  }
  if (memo.size > 6000) memo.clear();
  memo.set(key, v); return v;
}
/* Pasaulio taškas -> kameros koordinatės. Linija rodoma AR.bh m virš žemės; kameros aukštis = žemė po telefonu + AR.h. */
function arPoint(wx, wy, ux, uy, ground = true) { return arCam(wx - ux, wy - uy, ground ? terrZ(wx, wy) + AR.bh - (terrZ(ux, uy) + AR.h) : 0); }
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
    $("ar-hud").innerHTML = $("field-info").innerHTML + `<div class="small">${st.join(" · ")} · FOV ${AR.fov}° · aukštis ${AR.h.toFixed(1).replace(".", ",")} m · pasvirimas ${AR.pitch > 0 ? "+" : ""}${AR.pitch.toFixed(1).replace(".", ",")}° · linija +${AR.bh.toFixed(1).replace(".", ",")} m · ${MAP.terr ? "reljefas: LiDAR (" + MAP.terr.length + ")" : "reljefo nėra (žemė lygi)"}${AR.fix ? " · vieta užfiksuota" : ""}</div>`;
  }
  c.strokeStyle = "rgba(255,255,255,.7)"; c.lineWidth = 1.5;
  c.beginPath(); c.moveTo(W / 2 - 12, H / 2); c.lineTo(W / 2 + 12, H / 2); c.moveTo(W / 2, H / 2 - 12); c.lineTo(W / 2, H / 2 + 12); c.stroke();
  if (!AR.R || !fld.pos) {
    c.fillStyle = "rgba(0,0,0,.55)"; c.fillRect(0, H / 2 - 28, W, 56);
    c.fillStyle = "#fff"; c.font = "bold 15px system-ui,sans-serif"; c.textAlign = "center";
    c.fillText(!fld.pos ? "Laukiama GPS padėties…" : "Laukiama kompaso / giroskopo…", W / 2, H / 2 + 5); return;
  }
  const [ux, uy] = arPos(), NEAR = 0.3;
  const seg = (a, b, col, lw) => {
    let pa = arPoint(a[0], a[1], ux, uy), pb = arPoint(b[0], b[1], ux, uy);
    if (pa[2] < NEAR && pb[2] < NEAR) return;
    if (pa[2] < NEAR) { const t = (NEAR - pa[2]) / (pb[2] - pa[2]); pa = pa.map((v, i) => v + t * (pb[i] - v)); }
    else if (pb[2] < NEAR) { const t = (NEAR - pb[2]) / (pa[2] - pb[2]); pb = pb.map((v, i) => v + t * (pa[i] - v)); }
    const s1 = arScreen(pa), s2 = arScreen(pb);
    c.strokeStyle = col; c.lineWidth = lw; c.beginPath(); c.moveTo(s1[0], s1[1]); c.lineTo(s2[0], s2[1]); c.stroke();
  };
  const line = (a, b, col, lw) => {   // su reljefu ilgi ruožai dalinami (~3 m), kad linija sektų žemės paviršių
    const len = Math.hypot(b[0] - a[0], b[1] - a[1]), n = MAP.terr ? Math.min(60, Math.max(1, Math.ceil(len / 3))) : 1;
    if (n === 1) { seg(a, b, col, lw); return; }
    let prev = a;
    for (let i = 1; i <= n; i++) { const t = i / n, cur = [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t]; seg(prev, cur, col, lw); prev = cur; }
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
  const [ux, uy] = arPos(), ul = $("ar-pick-list"); ul.innerHTML = "";
  arTargets(ux, uy).slice(0, 8).forEach((t) => {
    const li = document.createElement("li");
    li.innerHTML = `<div class="t">${esc(t.label)}</div><div class="s">${fmtM(t.d)} nuo jūsų</div>`;
    li.onclick = () => {
      const yaw = arYaw(); if (yaw == null) { toast("Nėra kompaso duomenų", true); return; }
      const brg = (Math.atan2(t.x - ux, t.y - uy) * 180 / Math.PI + 360) % 360;
      let off = brg - yaw; while (off > 180) off -= 360; while (off < -180) off += 360;
      AR.off = off;
      const m = arElev(), et = Math.atan2(terrZ(t.x, t.y) - terrZ(ux, uy) - AR.h, Math.max(t.d, 1)) * 180 / Math.PI;   // taškas ant žemės: kryžiukas turi būti į jį
      if (m != null) AR.pitch = Math.max(-20, Math.min(20, et - m));
      $("ar-pick").hidden = true; toast(`Sulygiuota: kryptis ${off.toFixed(0)}°, pasvirimas ${AR.pitch.toFixed(1).replace(".", ",")}°`);
    };
    ul.appendChild(li);
  });
  $("ar-pick").hidden = false;
}
async function openAR() {
  if (!fld.mapOpen) return;
  arLoadPrefs(); AR.R = null; AR.absolute = false; AR.off = 0; AR.pitch = 0; AR.fix = null; AR.sm = null; AR.ev = { abs: 0, rel: 0, nul: 0 }; AR.via = ""; AR.perm = "";
  $("ar-panel").hidden = true; $("ar-fix").classList.remove("on"); $("ar-fix").textContent = "📌 Fiksuoti vietą";
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
  $("ar-cfg").onclick = () => { $("ar-panel").hidden = !$("ar-panel").hidden; };
  $("ar-fix").onclick = () => {
    const b = $("ar-fix");
    if (AR.fix) { AR.fix = null; b.classList.remove("on"); b.textContent = "📌 Fiksuoti vietą"; toast("Vieta vėl seka GPS"); return; }
    const p = arPos(); if (!p) { toast("Padėtis dar negauta", true); return; }
    AR.fix = p.slice(); b.classList.add("on"); b.textContent = "📌 Vieta užfiksuota"; toast("Vieta užfiksuota – taškai nebešoks nuo GPS. Atrakinti – paspauskite dar kartą");
  };
  $("ar-up").onclick = () => { AR.pitch = Math.max(-20, AR.pitch - 0.5); };     // ▲ – taškai kyla ekrane
  $("ar-down").onclick = () => { AR.pitch = Math.min(20, AR.pitch + 0.5); };    // ▼ – taškai leidžiasi
  $("ar-bh").onclick = () => { const o = [0, 0.5, 1, 1.5, 2]; let i = 0; o.forEach((v, k) => { if (Math.abs(v - AR.bh) < Math.abs(o[i] - AR.bh)) i = k; }); AR.bh = o[(i + 1) % o.length]; cfg.arBh = AR.bh; saveCfg(cfg); toast(`Linija ${AR.bh.toFixed(1).replace(".", ",")} m virš žemės`); };
  $("ar-reset").onclick = () => { AR.off = 0; AR.pitch = 0; toast("Kryptis ir pasvirimas atstatyti"); };
  $("ar-left").onclick = () => { AR.off += 1; };
  $("ar-right").onclick = () => { AR.off -= 1; };
  $("ar-fov-").onclick = () => { AR.fov = Math.max(40, AR.fov - 2); cfg.arFov = AR.fov; saveCfg(cfg); };
  $("ar-fov+").onclick = () => { AR.fov = Math.min(100, AR.fov + 2); cfg.arFov = AR.fov; saveCfg(cfg); };
  $("ar-h").onclick = () => { const o = [1.2, 1.4, 1.5, 1.6, 1.8]; let i = 0; o.forEach((v, k) => { if (Math.abs(v - AR.h) < Math.abs(o[i] - AR.h)) i = k; }); AR.h = o[(i + 1) % o.length]; cfg.arH = AR.h; saveCfg(cfg); toast(`Telefono aukštis ${AR.h.toFixed(1).replace(".", ",")} m`); };
}
window.__ar = { terrZ, arPoint, MAP, rotFromEuler, arCam, arScreen, AR, arYaw, arElev, arPos, setR: (R) => { AR.R = R; } };

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
/* ------------------------------------------------------------------ LAUKO MATAVIMAS (ruletė / Leica DISTO per Bluetooth) */
const FM_SVC = "3ab10100-f831-4395-b29d-570977d5bf94";
const FM_KEY = "prj_fm_draft";
const fm = { project: "", items: [], dev: null, server: null, last: null, lastAt: 0, aux: "", on: false, names: [] };
function fmSaveDraft() {
  try { localStorage.setItem(FM_KEY, JSON.stringify({ project: fm.project, items: fm.items, from: $("fm-from").value, to: $("fm-to").value, sketch: $("fm-sketch").value, date: fm.date, ab: fmab.blds.length ? fmAbExport() : null })); } catch (_) { /* ignoruojam */ }
}
function fmLoadDraft() {
  try { const j = JSON.parse(localStorage.getItem(FM_KEY) || "null"); return j && typeof j === "object" ? j : null; } catch (_) { return null; }
}
const fmFmt = (v) => v.toFixed(3).replace(".", ",") + " m";
function fmStatus(txt, ok) { $("fm-status").textContent = txt; document.querySelector(".fm-box").classList.toggle("ok", !!ok); }
function fmNames() {
  const set = new Set(fm.names);
  fmab.sides.forEach((sd) => { set.add(sd.from); set.add(sd.to); });
  fm.items.forEach((i) => { set.add(i.from); set.add(i.to); });
  $("fm-names").innerHTML = Array.from(set).filter(Boolean).map((n) => `<option value="${esc(n)}">`).join("");
}
function fmRender() {
  const ul = $("fm-list"); ul.innerHTML = "";
  if (!fm.items.length) { const li = document.createElement("li"); li.className = "muted"; li.textContent = "Atkarpų dar nėra."; ul.appendChild(li); }
  fm.items.forEach((it, i) => {
    const li = document.createElement("li"); li.className = "fm-item";
    li.innerHTML = `<div><div class="t">${esc(it.from)} → ${esc(it.to)} · ${esc(fmFmt(it.dist_m))}</div>` +
      `<div class="s">${it.source === "disto" ? "DISTO" : "ranka"}${it.note ? " · " + esc(it.note) : ""}${it.plan_m ? " · brėžinyje " + esc(fmFmt(it.plan_m)) + " (Δ " + esc(fmDelta(it.dist_m - it.plan_m)) + ")" : ""}</div></div>` +
      `<div class="acts"><button type="button" class="btn small ghost" data-i="${i}" data-a="edit">Taisyti</button><button type="button" class="btn small ghost" data-i="${i}" data-a="del">✕</button></div>`;
    ul.appendChild(li);
  });
  ul.querySelectorAll("button").forEach((b) => {
    b.onclick = () => {
      const i = +b.dataset.i, it = fm.items[i]; if (!it) return;
      if (b.dataset.a === "edit") { $("fm-from").value = it.from; $("fm-to").value = it.to; $("fm-dist").value = it.dist_m.toFixed(3).replace(".", ","); $("fm-note").value = it.note || ""; }
      fm.items.splice(i, 1); fmRender(); fmNames(); fmSaveDraft();
    };
  });
  fmAbRender();
}
function fmAdd(source) {
  const from = $("fm-from").value.trim(), to = $("fm-to").value.trim(), dist = num($("fm-dist").value);
  if (!from || !to) return toast("Įrašykite „Nuo taško“ ir „Iki taško“", true);
  if (dist == null || !(dist > 0) || dist > 300) return toast("Įrašykite atstumą (0–300 m)", true);
  if (from === to) return toast("„Nuo“ ir „Iki“ negali būti vienodi", true);
  const sd = fmAbFind(from, to);
  const item = { from, to, dist_m: Math.round(dist * 1000) / 1000, note: $("fm-note").value.trim(), source: source || "manual", t: localISO(), aux: fm.aux || "" };
  if (sd) {
    if (sd.plan != null) item.plan_m = Math.round(sd.plan * 1000) / 1000;
    fm.items = fm.items.filter((i) => !fmAbSame(i, sd));   // ta pati kraštinė perrašoma
    if ($("fm-ab-apply").checked) { fmAbApplyLength(sd, item.dist_m); fmAbBuild(); }
  }
  fm.items.push(item);
  $("fm-dist").value = ""; $("fm-note").value = ""; fm.aux = "";
  if (sd && fmab.sel >= 0) fmAbNext(fmab.sides.find((x) => x.from === sd.from && x.to === sd.to) || sd);
  else if ($("fm-chain").checked) { $("fm-from").value = to; $("fm-to").value = ""; $("fm-to").focus(); }
  fmRender(); fmNames(); fmSaveDraft();
}
function fmOnData(ev) {
  const c = ev.target, dv = c.value, id = c.uuid;
  const hex = []; for (let i = 0; i < dv.byteLength; i++) hex.push(("0" + dv.getUint8(i).toString(16)).slice(-2));
  if (id.startsWith("3ab10101") && dv.byteLength >= 4) {
    const d = dv.getFloat32(0, true);
    if (!isFinite(d) || d <= 0.01 || d >= 300) return;
    const now = Date.now();
    if (fm.last != null && Math.abs(d - fm.last) < 1e-6 && now - fm.lastAt < 700) return;   // dubliuotas pranešimas
    fm.last = d; fm.lastAt = now;
    $("fm-big").textContent = fmFmt(d);
    $("fm-dist").value = d.toFixed(3).replace(".", ",");
    fmAbBar(fmab.sel >= 0 ? fmab.sides[fmab.sel] : null);
    try { if (navigator.vibrate) navigator.vibrate(70); } catch (_) { /* */ }
    log("DISTO: " + d.toFixed(4) + " m");
    if ($("fm-auto").checked || fmab.sel >= 0) setTimeout(() => fmAdd("disto"), 0);
  } else if (id.startsWith("3ab10102")) { fm.aux = hex.join(" "); }
}
async function fmSubscribe(server) {
  const svc = await server.getPrimaryService(FM_SVC);
  const chars = await svc.getCharacteristics();
  let n = 0;
  for (const c of chars) {
    if ((c.properties.notify || c.properties.indicate) && (c.uuid.startsWith("3ab10101") || c.uuid.startsWith("3ab10102"))) {
      c.removeEventListener("characteristicvaluechanged", fmOnData);
      c.addEventListener("characteristicvaluechanged", fmOnData);
      await c.startNotifications(); n++;
    }
  }
  if (!n) throw new Error("DISTO atstumo charakteristika nerasta");
}
async function fmConnect() {
  if (!navigator.bluetooth) { fmStatus("Ši naršyklė nepalaiko Bluetooth (naudokite „Chrome“ Android telefone)"); return; }
  try {
    fmStatus("Pasirinkite DISTO sąraše…");
    fm.dev = await navigator.bluetooth.requestDevice({ filters: [{ namePrefix: "DISTO" }, { namePrefix: "Leica" }, { services: [FM_SVC] }], optionalServices: [FM_SVC, "battery_service"] });
    fm.dev.removeEventListener("gattserverdisconnected", fmOnDisc);
    fm.dev.addEventListener("gattserverdisconnected", fmOnDisc);
    fmStatus("Jungiamasi…");
    fm.server = await fm.dev.gatt.connect();
    await fmSubscribe(fm.server);
    fm.on = true; $("fm-connect").textContent = "Atjungti";
    fmStatus("Prijungta: " + (fm.dev.name || "DISTO") + " – matuokite (ON/DIST)", true);
    log("DISTO prijungtas: " + (fm.dev.name || ""));
  } catch (e) {
    fm.on = false; $("fm-connect").textContent = "Prijungti DISTO";
    fmStatus(e.name === "NotFoundError" ? "Įrenginys nepasirinktas. Įjunkite DISTO Bluetooth ir bandykite dar kartą." : "Nepavyko prisijungti: " + e.message);
    log("DISTO klaida: " + e.name + ": " + e.message);
  }
}
async function fmOnDisc() {
  if (!fm.on) return;
  fmStatus("Ryšys nutrūko – jungiamasi iš naujo…");
  for (let i = 0; i < 3; i++) {
    try {
      await new Promise((r) => setTimeout(r, 1200));
      fm.server = await fm.dev.gatt.connect(); await fmSubscribe(fm.server);
      fmStatus("Prijungta: " + (fm.dev.name || "DISTO") + " – matuokite (ON/DIST)", true); return;
    } catch (_) { /* kartojam */ }
  }
  fm.on = false; $("fm-connect").textContent = "Prijungti DISTO";
  fmStatus("Ryšys nutrūko. Įjunkite DISTO Bluetooth ir spauskite „Prijungti DISTO“.");
}
function fmDisconnect() {
  fm.on = false;
  try { if (fm.dev && fm.dev.gatt.connected) fm.dev.gatt.disconnect(); } catch (_) { /* */ }
  $("fm-connect").textContent = "Prijungti DISTO"; fmStatus("DISTO atjungtas");
}
async function fmSetProject(prj) {
  fm.project = prj; if (prj) rememberProject(prj);
  setPick($("fm-project"), prj, "Pasirinkite projektą…");
  fm.names = [];
  if (prj) { const d = await cacheGet("points:" + prj); fm.names = ((d && d.points) || []).map((p) => String(p.nr)); }
  await fmAbLoad(prj);
  fmNames(); fmSaveDraft();
  if (prj) refreshBuildings(prj, true).then(async (j) => { if (j && fm.project === prj) { fmab.draft = fmab.blds.length ? fmAbExport() : null; await fmAbLoad(prj); fmNames(); } });
}

/* ------------------------------------------------------------------ Abriso juodraštis (pastatų kontūrai 2601/2602 iš DWG) */
const fmab = { blds: [], sides: [], sel: -1, tol: 0.10, move: false, draw: null, drag: null, tf: null, draft: null, full: false, view: { k: 1, cx: null, cy: null }, base: null, ptrs: new Map(), gest: null, noClick: false };
/* Mastelis ir stumdymas (viewBox): ＋/－ mygtukai, dviejų pirštų suspaudimas, vieno piršto stumdymas, pelės ratukas */
function fmAbGeom() { const r = $("fm-ab-svg").getBoundingClientRect(), b = fmab.base; return r.width && b ? { r, b } : null; }
function fmAbScale(g, k) { return Math.min(g.r.width / (g.b.vbw / k), g.r.height / (g.b.vbh / k)); }
function fmAbToSvg(g, mx, my, k, cx, cy) {
  const vw = g.b.vbw / k, vh = g.b.vbh / k, s = fmAbScale(g, k);
  return [cx - vw / 2 + (mx - g.r.left - (g.r.width - s * vw) / 2) / s, cy - vh / 2 + (my - g.r.top - (g.r.height - s * vh) / 2) / s];
}
function fmAbSetView(k, cx, cy) {
  const b = fmab.base; if (!b) return;
  k = Math.max(1, Math.min(40, k));
  if (k <= 1.0001) { k = 1; cx = b.bx + b.vbw / 2; cy = b.by + b.vbh / 2; }
  fmab.view = { k, cx: Math.max(b.bx, Math.min(b.bx + b.vbw, cx)), cy: Math.max(b.by, Math.min(b.by + b.vbh, cy)) };
}
function fmAbVBox() { const v = fmab.view, b = fmab.base, vw = b.vbw / v.k, vh = b.vbh / v.k; return `${v.cx - vw / 2} ${v.cy - vh / 2} ${vw} ${vh}`; }
/* k2 – naujas mastelis; (mx,my) – ekrano taškas, po kuriuo esantis brėžinio taškas `anchor` turi likti vietoje */
function fmAbZoomAt(k2, mx, my, anchor) {
  const g = fmAbGeom(); if (!g) return;
  k2 = Math.max(1, Math.min(40, k2));
  const a = anchor || fmAbToSvg(g, mx, my, fmab.view.k, fmab.view.cx, fmab.view.cy);
  const vw = g.b.vbw / k2, vh = g.b.vbh / k2, s = fmAbScale(g, k2);
  fmAbSetView(k2, a[0] - (mx - g.r.left - (g.r.width - s * vw) / 2) / s + vw / 2, a[1] - (my - g.r.top - (g.r.height - s * vh) / 2) / s + vh / 2);
}
function fmAbZoomBtn(f) { const g = fmAbGeom(); if (!g) return; fmAbZoomAt(fmab.view.k * f, g.r.left + g.r.width / 2, g.r.top + g.r.height / 2); fmAbRender(); }
function fmAbGestureBind(svg) {
  const info = () => { const a = [...fmab.ptrs.values()]; return { n: a.length, mx: a.reduce((t, q) => t + q.x, 0) / (a.length || 1), my: a.reduce((t, q) => t + q.y, 0) / (a.length || 1), d: a.length > 1 ? Math.hypot(a[0].x - a[1].x, a[0].y - a[1].y) : 0 }; };
  const begin = () => {
    const g = fmAbGeom(), i = info(); if (!g || !i.n) { fmab.gest = null; return; }
    fmab.gest = { k: fmab.view.k, cx: fmab.view.cx, cy: fmab.view.cy, mx: i.mx, my: i.my, d: i.d, n: i.n, moved: fmab.gest ? fmab.gest.moved : false, anchor: fmAbToSvg(g, i.mx, i.my, fmab.view.k, fmab.view.cx, fmab.view.cy) };
  };
  svg.addEventListener("pointerdown", (e) => {
    if (!fmab.full || fmab.drag || e.pointerType === "mouse" && e.button !== 0) return;
    if (!fmab.ptrs.size) fmab.noClick = false;
    fmab.ptrs.set(e.pointerId, { x: e.clientX, y: e.clientY }); begin();
  });
  svg.addEventListener("pointermove", (e) => {
    if (!fmab.ptrs.has(e.pointerId) || fmab.drag || !fmab.gest) return;
    fmab.ptrs.set(e.pointerId, { x: e.clientX, y: e.clientY });
    const i = info(), G = fmab.gest, g = fmAbGeom(); if (!g) return;
    if (i.n >= 2 && G.d > 0) {
      G.moved = true; fmab.noClick = true;
      fmAbZoomAt(G.k * (i.d / G.d), i.mx, i.my, G.anchor);
    } else if (i.n === 1) {
      const dx = i.mx - G.mx, dy = i.my - G.my;
      if (!G.moved && Math.hypot(dx, dy) < 8) return;
      if (!G.moved) { G.moved = true; fmab.noClick = true; try { svg.setPointerCapture(e.pointerId); } catch (_) { /* */ } }
      if (fmab.view.k > 1) { const s = fmAbScale(g, G.k); fmAbSetView(G.k, G.cx - dx / s, G.cy - dy / s); }
    } else return;
    svg.setAttribute("viewBox", fmAbVBox());
  });
  const up = (e) => {
    if (!fmab.ptrs.delete(e.pointerId)) return;
    const moved = fmab.gest && fmab.gest.moved;
    if (fmab.ptrs.size) begin(); else { fmab.gest = null; if (moved) { fmAbRender(); setTimeout(() => { fmab.noClick = false; }, 60); } }
  };
  svg.addEventListener("pointerup", up); svg.addEventListener("pointercancel", up);
  svg.addEventListener("click", (e) => { if (fmab.noClick) { e.stopPropagation(); e.preventDefault(); } }, true);
  svg.addEventListener("wheel", (e) => {
    if (!fmab.full) return;
    e.preventDefault(); fmAbZoomAt(fmab.view.k * Math.exp(-e.deltaY * 0.0018), e.clientX, e.clientY); fmAbRender();
  }, { passive: false });
}
const fmDelta = (v) => (v >= 0 ? "+" : "−") + Math.abs(v).toFixed(3).replace(".", ",") + " m";
const FM_KINDS = [["pastatas", "Pastatas", true], ["ukinis", "Ūkinis pastatas / priestatas", true], ["tvora", "Tvora (linija)", false], ["kita", "Kita linija", false], ["kontura", "Kitas kontūras", true]];
const cp = (pts) => pts.map((p) => [p[0], p[1]]);
function fmAbSame(it, sd) { return (it.from === sd.from && it.to === sd.to) || (it.from === sd.to && it.to === sd.from); }
function fmAbFind(from, to) { return fmab.sides.find((sd) => (sd.from === from && sd.to === to) || (sd.from === to && sd.to === from)) || null; }
function fmAbItem(sd) { for (let i = fm.items.length - 1; i >= 0; i--) if (fmAbSame(fm.items[i], sd)) return fm.items[i]; return null; }
function fmAbBuild() {
  fmab.sides = [];
  (fmab.blds || []).forEach((b, bi) => {
    const p = b.pts || [], n = p.length;
    if (n < 2) return;
    const m = b.closed ? n : n - 1;
    for (let i = 0; i < m; i++) {
      const j = (i + 1) % n, o = b.orig && b.orig.length === n ? b.orig : null;
      fmab.sides.push({ bi, i, from: `${b.id}-${i + 1}`, to: `${b.id}-${j + 1}`, a: p[i], b: p[j], len: Math.hypot(p[j][0] - p[i][0], p[j][1] - p[i][1]),
        plan: o ? Math.hypot(o[j][0] - o[i][0], o[j][1] - o[i][1]) : null });
    }
  });
  if (fmab.sel >= fmab.sides.length) fmab.sel = -1;
}
function fmAbExport() {
  return { buildings: fmab.blds.map((b) => ({ id: b.id, code: b.code || "", kind: b.kind || "", extra: !!b.extra, closed: !!b.closed, pts: b.pts.map((p) => [Math.round(p[0] * 1000) / 1000, Math.round(p[1] * 1000) / 1000]),
    orig_pts: b.orig ? b.orig.map((p) => [p[0], p[1]]) : null })) };
}
function fmAbApplyDraft() {
  const d = fmab.draft; fmab.draft = null;
  if (!d || !Array.isArray(d.buildings)) return;
  d.buildings.forEach((e) => {
    if (e.extra) { fmab.blds.push({ id: e.id, code: e.code || "", kind: e.kind || "", extra: true, closed: !!e.closed, pts: cp(e.pts), orig: null }); return; }
    const b = fmab.blds.find((x) => x.id === e.id);
    if (b && b.pts.length === (e.pts || []).length) b.pts = cp(e.pts);
  });
}
async function fmAbLoad(prj) {
  fmab.sel = -1; fmab.blds = []; fmab.draw = null; fmab.drag = null; fmab.view = { k: 1, cx: null, cy: null };
  if (prj) {
    try {
      const d = await cacheGet("buildings:" + prj);
      fmab.blds = ((d && d.buildings) || []).map((b) => ({ id: b.id, code: b.code || "", kind: "", extra: false, closed: !!b.closed, pts: cp(b.pts), orig: cp(b.pts) }));
    } catch (_) { /* */ }
  }
  fmAbApplyDraft(); fmAbBuild(); fmAbRender();
}
function fmAbSelect(i) {
  if (fmab.draw) return;
  if (i < 0 || i >= fmab.sides.length || fmab.sel === i) { fmab.sel = -1; fmAbRender(); return; }
  const sd = fmab.sides[i]; fmab.sel = i;
  $("fm-from").value = sd.from; $("fm-to").value = sd.to; $("fm-dist").value = "";
  const it = fmAbItem(sd); if (it) $("fm-dist").value = it.dist_m.toFixed(3).replace(".", ",");
  fmSaveDraft(); fmAbRender();
}
function fmAbNext(sd) {
  const start = fmab.sides.indexOf(sd);
  for (let k = 1; k <= fmab.sides.length; k++) {
    const i = (start + k) % fmab.sides.length, c = fmab.sides[i];
    if (!fmAbItem(c) && c !== sd) { fmAbSelect(i); return; }
  }
  fmab.sel = -1; toast("Visos kraštinės išmatuotos");
}
/* Kraštinės ilgis pakeičiamas į išmatuotą: pastumiamas galinis kampas; jei ties juo stačias kampas – ir kitas kampas (stačiakampis išlieka stačiakampiu). */
function fmAbApplyLength(sd, L) {
  const b = fmab.blds[sd.bi]; if (!b) return;
  const n = b.pts.length, i = sd.i, j = (i + 1) % n, A = b.pts[i], B = b.pts[j];
  const cur = Math.hypot(B[0] - A[0], B[1] - A[1]); if (cur < 1e-6 || !(L > 0)) return;
  const ux = (B[0] - A[0]) / cur, uy = (B[1] - A[1]) / cur, d = L - cur, mv = [j];
  const k = (j + 1) % n;
  if (k !== i && (b.closed || j + 1 < n)) {
    const C = b.pts[k], cx = C[0] - B[0], cy = C[1] - B[1], cl = Math.hypot(cx, cy);
    if (cl > 1e-6 && Math.abs((cx * ux + cy * uy) / cl) < 0.0872) mv.push(k);
  }
  mv.forEach((m) => { b.pts[m] = [b.pts[m][0] + ux * d, b.pts[m][1] + uy * d]; });
}
/* Kraštinės pakreipimas: pivot a = apie pradžią, b = apie pabaigą, m = apie vidurį; deg > 0 – prieš laikrodžio rodyklę */
function fmAbRotate(sd, deg, pivot) {
  const b = fmab.blds[sd.bi]; if (!b) return;
  const n = b.pts.length, i = sd.i, j = (i + 1) % n, A = b.pts[i], B = b.pts[j];
  const r = deg * Math.PI / 180, cs = Math.cos(r), sn = Math.sin(r);
  const P = pivot === "b" ? B : pivot === "m" ? [(A[0] + B[0]) / 2, (A[1] + B[1]) / 2] : A;
  const rot = (q) => [P[0] + (q[0] - P[0]) * cs - (q[1] - P[1]) * sn, P[1] + (q[0] - P[0]) * sn + (q[1] - P[1]) * cs];
  if (pivot !== "b") b.pts[j] = rot(B);
  if (pivot !== "a") b.pts[i] = rot(A);
}
function fmAbGeoChanged() { fmAbBuild(); fmSaveDraft(); fmAbRender(); }
function fmAbCenter() {
  const xs = [], ys = [];
  fmab.blds.forEach((b) => b.pts.forEach((p) => { xs.push(p[0]); ys.push(p[1]); }));
  return xs.length ? [(Math.min(...xs) + Math.max(...xs)) / 2, (Math.min(...ys) + Math.max(...ys)) / 2] : [0, 0];
}
function fmAbNewId() { let n = 1; const ids = new Set(fmab.blds.map((b) => b.id)); while (ids.has("X" + n)) n++; return "X" + n; }
function fmAbAddObject(kind, closed, pts) {
  const id = fmAbNewId();
  fmab.blds.push({ id, code: "", kind, extra: true, closed, pts: cp(pts), orig: null });
  fmAbBuild(); fmab.sel = -1; fmSaveDraft(); fmNames(); fmAbRender();
  toast(`Pridėtas objektas ${id}`);
  const first = fmab.sides.findIndex((sd) => fmab.blds[sd.bi].id === id); if (first >= 0) fmAbSelect(first);
}
function fmAbPtFromEvent(e) {
  const svg = $("fm-ab-svg"), m = svg.getScreenCTM(); if (!m || !fmab.tf) return null;
  const pt = svg.createSVGPoint(); pt.x = e.clientX; pt.y = e.clientY;
  const q = pt.matrixTransform(m.inverse());
  return [q.x + fmab.tf.minx, fmab.tf.maxy - q.y];
}
function fmAbRender() {
  const card = $("fm-abris"); if (!card) return;
  card.hidden = !fm.project;
  if (!fm.project) fmab.full = false;
  card.classList.toggle("full", fmab.full);
  $("fm-ab-bar").hidden = !fmab.full;
  $("fm-ab-full").textContent = fmab.full ? "✕ Uždaryti" : "⛶ Planas";
  try { document.body.style.overflow = fmab.full ? "hidden" : ""; } catch (_) { /* */ }
  const svg = $("fm-ab-svg"), info = $("fm-ab-info");
  svg.innerHTML = "";
  const sel = fmab.sel >= 0 ? fmab.sides[fmab.sel] : null;
  fmAbBar(sel);
  $("fm-ab-rot").hidden = !sel;
  $("fm-ab-move").classList.toggle("primary", fmab.move);
  $("fm-ab-draw").hidden = !fmab.draw;
  $("fm-ab-del").hidden = !(sel && fmab.blds[sel.bi] && fmab.blds[sel.bi].extra);
  const drawPts = fmab.draw ? fmab.draw.pts : [];
  if (!fmab.blds.length && !drawPts.length) {
    svg.setAttribute("hidden", "");
    info.textContent = fm.project ? "Šiam projektui pastatų (sluoksniai 2601–2699) nėra – kompiuteryje KADMAT taškų lange spauskite „Abriso pastatai…“. Objektus galima pridėti ranka: ＋ Objektas." : "";
    return;
  }
  svg.removeAttribute("hidden");
  const xs = [], ys = [];
  fmab.blds.forEach((b) => b.pts.forEach((p) => { xs.push(p[0]); ys.push(p[1]); }));
  drawPts.forEach((p) => { xs.push(p[0]); ys.push(p[1]); });
  const minx = Math.min(...xs), maxx = Math.max(...xs), miny = Math.min(...ys), maxy = Math.max(...ys);
  const w = Math.max(maxx - minx, 1), h = Math.max(maxy - miny, 1), pad = Math.max(w, h) * 0.16;
  const vbw = w + 2 * pad, vbh = h + 2 * pad, fz = Math.pow(fmab.view.k, 0.75), fs = Math.max(vbw, vbh) * 0.036 / fz, sw = fs * 0.16;
  fmab.tf = { minx, maxy };
  fmab.base = { bx: -pad, by: -pad, vbw, vbh };
  if (fmab.view.cx == null) fmAbSetView(fmab.view.k, -pad + vbw / 2, -pad + vbh / 2); else fmAbSetView(fmab.view.k, fmab.view.cx, fmab.view.cy);
  svg.setAttribute("viewBox", fmAbVBox());
  const X = (x) => (x - minx), Y = (y) => (maxy - y);
  const NS = "http://www.w3.org/2000/svg";
  const mk = (tag, attrs, txt) => { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (txt != null) e.textContent = txt; svg.appendChild(e); return e; };
  const cs = getComputedStyle(document.documentElement);
  const col = (v, d) => (cs.getPropertyValue(v) || d).trim() || d;
  const cPri = col("--primary", "#393187"), cMut = col("--muted", "#5d6079"), cAcc = col("--accent", "#e8201f"), cWarn = col("--warn", "#c4161c"), cSoft = col("--primary-soft", "#e8e7f5");
  const bg = mk("rect", { x: -pad - vbw * 10, y: -pad - vbh * 10, width: vbw * 21, height: vbh * 21, fill: "transparent" });
  bg.addEventListener("click", (e) => {
    if (!fmab.draw) return;
    const p = fmAbPtFromEvent(e); if (!p) return;
    fmab.draw.pts.push(p); fmAbRender();
  });
  const cen = fmab.blds.map((b) => { const p = b.pts; return [p.reduce((a, q) => a + q[0], 0) / p.length, p.reduce((a, q) => a + q[1], 0) / p.length]; });
  fmab.blds.forEach((b, bi) => {
    if (b.orig && fmab.blds[bi].pts.some((p, k) => Math.abs(p[0] - b.orig[k][0]) > 0.005 || Math.abs(p[1] - b.orig[k][1]) > 0.005))
      mk("polyline", { points: b.orig.concat(b.closed ? [b.orig[0]] : []).map((p) => `${X(p[0])},${Y(p[1])}`).join(" "), fill: "none", stroke: cMut, "stroke-width": sw * 0.6, "stroke-dasharray": `${sw * 1.5} ${sw * 1.5}`, opacity: "0.45" });
    if (b.closed) mk("polygon", { points: b.pts.map((p) => `${X(p[0])},${Y(p[1])}`).join(" "), fill: b.extra ? cWarn : cSoft, "fill-opacity": b.extra ? "0.12" : "0.55", stroke: "none" });
    mk("text", { x: X(cen[bi][0]), y: Y(cen[bi][1]), "text-anchor": "middle", "font-size": fs * 1.1, fill: cMut, "font-weight": "700", opacity: "0.6" }, b.id + (b.code && b.code !== "2601" && b.code !== "2602" ? " · " + b.code : ""));
  });
  let done = 0, bad = 0;
  fmab.sides.forEach((sd, i) => {
    const x1 = X(sd.a[0]), y1 = Y(sd.a[1]), x2 = X(sd.b[0]), y2 = Y(sd.b[1]);
    const it = fmAbItem(sd), isel = i === fmab.sel, ex = fmab.blds[sd.bi].extra;
    let stroke = ex ? cWarn : cMut, txt = sd.len.toFixed(2).replace(".", ","), tcol = stroke;
    if (it) {
      done++; txt = it.dist_m.toFixed(2).replace(".", ",");
      const ok = sd.plan == null || Math.abs(it.dist_m - sd.plan) <= fmab.tol; if (!ok) bad++;
      stroke = ok ? cPri : cWarn; tcol = stroke;
    }
    mk("line", { x1, y1, x2, y2, stroke: isel ? cAcc : stroke, "stroke-width": isel ? sw * 2.2 : sw * (it ? 1.6 : 1), "stroke-linecap": "round", "stroke-dasharray": it || isel ? "none" : `${sw * 3} ${sw * 2}` });
    const mx = (x1 + x2) / 2, my = (y1 + y2) / 2, cx = X(cen[sd.bi][0]), cy = Y(cen[sd.bi][1]);
    let nx = mx - cx, ny = my - cy; const nl = Math.hypot(nx, ny) || 1; nx /= nl; ny /= nl;
    mk("text", { x: mx + nx * fs * 1.0, y: my + ny * fs * 1.0 + fs * 0.35, "text-anchor": "middle", "font-size": fs, fill: isel ? cAcc : tcol, "font-weight": it || isel ? "700" : "400" }, txt);
    const hit = mk("line", { x1, y1, x2, y2, stroke: "transparent", "stroke-width": fs * 2.2, "stroke-linecap": "round", style: "cursor:pointer" });
    hit.addEventListener("click", () => fmAbSelect(i));
  });
  fmab.blds.forEach((b, bi) => b.pts.forEach((p, k) => {
    mk("circle", { cx: X(p[0]), cy: Y(p[1]), r: sw * 1.2, fill: fmab.move ? cAcc : cMut });
    let nx = X(p[0]) - X(cen[bi][0]), ny = Y(p[1]) - Y(cen[bi][1]); const nl = Math.hypot(nx, ny) || 1; nx /= nl; ny /= nl;
    mk("text", { x: X(p[0]) + nx * fs * 0.85, y: Y(p[1]) + ny * fs * 0.85 + fs * 0.3, "text-anchor": "middle", "font-size": fs * 0.78, fill: cMut, opacity: "0.8" }, String(k + 1));
    if (fmab.move && !fmab.draw) {
      const hd = mk("circle", { cx: X(p[0]), cy: Y(p[1]), r: fs * 1.15, fill: "transparent", stroke: cAcc, "stroke-width": sw * 0.5, "stroke-opacity": "0.5", style: "touch-action:none;cursor:grab" });
      hd.addEventListener("pointerdown", (e) => { e.preventDefault(); fmab.drag = { bi, k }; try { svg.setPointerCapture(e.pointerId); } catch (_) { /* */ } });
    }
  }));
  if (drawPts.length) {
    mk("polyline", { points: drawPts.map((p) => `${X(p[0])},${Y(p[1])}`).join(" "), fill: "none", stroke: cAcc, "stroke-width": sw * 1.4, "stroke-linejoin": "round" });
    drawPts.forEach((p) => mk("circle", { cx: X(p[0]), cy: Y(p[1]), r: sw * 1.6, fill: cAcc }));
  }
  info.textContent = `Objektų: ${fmab.blds.length}, kraštinių: ${fmab.sides.length}, išmatuota: ${done}` + (bad ? `, nukrypsta >${Math.round(fmab.tol * 100)} cm: ${bad}` : "") +
    (fmab.full && fmab.view.k === 1 && !sel ? " · Mastelis: dviem pirštais arba ＋/－" : "") + (fmab.draw ? " · Braižymas: palieskite brėžinį, kad pridėtumėte kampus" :
      sel ? ` · ${sel.from} → ${sel.to}: dabar ${fmFmt(sel.len)}${sel.plan != null ? `, DWG ${fmFmt(sel.plan)}` : ""} – matuokite D2 arba įveskite ilgį` : " · Palieskite kraštinę ir matuokite");
}
function fmAbBar(sel) {
  if (!fmab.full) return;
  $("fm-ab-sel").textContent = sel ? `Kraštinė ${sel.from} → ${sel.to} · dabar ${fmFmt(sel.len)}${sel.plan != null ? " · DWG " + fmFmt(sel.plan) : ""}` : "Palieskite kraštinę, kurios ilgį matuosite";
  const inp = $("fm-ab-dist");
  if (document.activeElement !== inp) inp.value = $("fm-dist").value;
  $("fm-ab-live").textContent = fm.last != null ? "D2: " + fmFmt(fm.last) : "D2: —";
  $("fm-ab-conn").textContent = fm.on ? "Atjungti DISTO" : "Prijungti DISTO";
}
function fmAbBind() {
  const svg = $("fm-ab-svg");
  fmAbGestureBind(svg);
  $("fm-ab-zin").onclick = () => fmAbZoomBtn(1.6); $("fm-ab-zout").onclick = () => fmAbZoomBtn(1 / 1.6);
  $("fm-ab-zfit").onclick = () => { fmab.view = { k: 1, cx: null, cy: null }; fmAbRender(); };
  $("fm-ab-full").onclick = () => { fmab.full = !fmab.full; fmAbRender(); try { $("fm-abris").scrollTop = 0; } catch (_) { /* */ } };
  $("fm-ab-dist").oninput = () => { $("fm-dist").value = $("fm-ab-dist").value; };
  $("fm-ab-save").onclick = () => { if (fmab.sel < 0) return toast("Pirmiausia palieskite kraštinę", true); $("fm-dist").value = $("fm-ab-dist").value; $("fm-add").click(); };
  $("fm-ab-conn").onclick = () => $("fm-connect").click();
  svg.addEventListener("pointermove", (e) => {
    if (!fmab.drag) return;
    const p = fmAbPtFromEvent(e); if (!p) return;
    const b = fmab.blds[fmab.drag.bi]; if (!b) return;
    b.pts[fmab.drag.k] = p; fmAbBuild(); fmAbRender();
  });
  const end = () => { if (fmab.drag) { fmab.drag = null; fmSaveDraft(); } };
  svg.addEventListener("pointerup", end); svg.addEventListener("pointercancel", end);
  $("fm-ab-len").onclick = () => { if (fmab.sel < 0) return toast("Pirmiausia palieskite kraštinę brėžinyje", true); $("fm-dist").focus(); $("fm-dist").scrollIntoView({ block: "center" }); };
  $("fm-ab-move").onclick = () => { fmab.move = !fmab.move; fmAbRender(); toast(fmab.move ? "Kampus galima tempti pirštu" : "Kampų tempimas išjungtas"); };
  const rot = (sign) => { const sd = fmab.sides[fmab.sel]; if (!sd) return; fmAbRotate(sd, sign * parseFloat($("fm-ab-step").value), $("fm-ab-pivot").value); fmAbGeoChanged(); };
  $("fm-ab-ccw").onclick = () => rot(1); $("fm-ab-cw").onclick = () => rot(-1);
  $("fm-ab-del").onclick = () => {
    const sd = fmab.sides[fmab.sel]; if (!sd) return; const b = fmab.blds[sd.bi]; if (!b || !b.extra) return;
    if (!confirm(`Šalinti objektą ${b.id}?`)) return;
    fm.items = fm.items.filter((it) => !(it.from.startsWith(b.id + "-") && it.to.startsWith(b.id + "-")));
    fmab.blds.splice(sd.bi, 1); fmab.sel = -1; fmAbBuild(); fmRender(); fmSaveDraft(); fmNames();
  };
  $("fm-ab-add").onclick = () => { $("fm-ab-new").hidden = !$("fm-ab-new").hidden; };
  $("fm-ab-kind").innerHTML = FM_KINDS.map((k) => `<option value="${k[0]}">${k[1]}</option>`).join("");
  $("fm-ab-how").onchange = () => { $("fm-ab-dims").hidden = $("fm-ab-how").value !== "rect"; };
  $("fm-ab-create").onclick = () => {
    const kind = $("fm-ab-kind").value, closed = FM_KINDS.find((k) => k[0] === kind)[2];
    if ($("fm-ab-how").value === "rect") {
      const L = num($("fm-ab-l").value), W = num($("fm-ab-w").value);
      if (!(L > 0) || (closed && !(W > 0))) return toast("Įrašykite ilgį" + (closed ? " ir plotį" : ""), true);
      const c = fmAbCenter();
      const pts = closed ? [[c[0] - L / 2, c[1] - W / 2], [c[0] + L / 2, c[1] - W / 2], [c[0] + L / 2, c[1] + W / 2], [c[0] - L / 2, c[1] + W / 2]] : [[c[0] - L / 2, c[1]], [c[0] + L / 2, c[1]]];
      $("fm-ab-new").hidden = true; fmAbAddObject(kind, closed, pts);
    } else {
      fmab.draw = { kind, closed, pts: [] }; $("fm-ab-new").hidden = true; fmab.sel = -1; fmAbRender();
    }
  };
  const fin = (closed) => {
    const d = fmab.draw; if (!d) return;
    if (d.pts.length < (closed ? 3 : 2)) return toast(closed ? "Reikia bent 3 kampų" : "Reikia bent 2 taškų", true);
    fmab.draw = null; fmAbAddObject(d.kind, closed, d.pts);
  };
  $("fm-ab-fin-c").onclick = () => fin(true); $("fm-ab-fin-o").onclick = () => fin(false);
  $("fm-ab-undo").onclick = () => { if (fmab.draw) { fmab.draw.pts.pop(); fmAbRender(); } };
  $("fm-ab-cancel").onclick = () => { fmab.draw = null; fmAbRender(); };
  $("fm-ab-reset").onclick = async () => {
    if (!confirm("Atstatyti pradinį brėžinį iš DWG ir pašalinti pridėtus objektus? Matavimų sąrašas nebus keičiamas.")) return;
    fmab.draft = null; await fmAbLoad(fm.project); fmSaveDraft();
  };
  $("fm-ab-refresh").onclick = async () => { if (!fm.project) return toast("Pasirinkite projektą", true); const keep = fmAbExport(); await refreshBuildings(fm.project, false); fmab.draft = keep; await fmAbLoad(fm.project); fmNames(); };
}
function initFm() {
  fm.items = []; fm.date = todayISO();
  $("fm-from").value = ""; $("fm-to").value = ""; $("fm-dist").value = ""; $("fm-note").value = ""; $("fm-sketch").value = "";
  setPick($("fm-project"), "", "Pasirinkite projektą…"); fm.project = ""; fm.names = [];
  const dr = fmLoadDraft();
  if (dr && ((dr.items && dr.items.length) || dr.project)) {
    fm.items = Array.isArray(dr.items) ? dr.items : []; fm.date = dr.date || todayISO();
    $("fm-from").value = dr.from || ""; $("fm-to").value = dr.to || ""; $("fm-sketch").value = dr.sketch || "";
    fmab.draft = dr.ab || null;
    if (dr.project) fmSetProject(dr.project);
    if (fm.items.length) toast("Atkurtas nebaigtas matavimas (" + fm.items.length + " atkarp.)");
  }
  if (!fm.on) { $("fm-connect").textContent = "Prijungti DISTO"; fmStatus(navigator.bluetooth ? "DISTO neprijungtas (galima ir įvesti ranka)" : "Bluetooth nepalaikomas – atstumus įveskite ranka"); }
  $("fm-big").textContent = fm.last != null ? fmFmt(fm.last) : "—";
  fmab.sel = -1; fmab.blds = []; fmab.sides = []; fmab.draw = null; fmab.move = false;
  fmRender(); fmNames();
}
function bindFm() {
  fmAbBind();
  $("fm-project").onclick = () => openPicker((p) => fmSetProject(p), { title: "Projektas (lauko matavimas)" });
  $("fm-connect").onclick = () => { if (fm.on) fmDisconnect(); else fmConnect(); };
  $("fm-add").onclick = () => fmAdd($("fm-dist").value && fm.last != null && Math.abs(num($("fm-dist").value) - fm.last) < 0.0006 ? "disto" : "manual");
  $("fm-swap").onclick = () => { const a = $("fm-from").value; $("fm-from").value = $("fm-to").value; $("fm-to").value = a; fmSaveDraft(); };
  ["fm-from", "fm-to", "fm-sketch"].forEach((id) => { $(id).oninput = fmSaveDraft; });
  $("fm-save").onclick = async () => {
    if (!fm.project) return toast("Pasirinkite projektą", true);
    if (!fm.items.length) return toast("Įrašykite bent vieną atkarpą", true);
    const total = fm.items.reduce((s, i) => s + i.dist_m, 0);
    await enqueue("field_measure", { project_no: fm.project, measure_date: fm.date || todayISO(), sketch_note: $("fm-sketch").value.trim(), items: fm.items, abris: fmab.blds.length ? fmAbExport() : null },
      [], `${fm.project}: lauko matavimas, ${fm.items.length} atkarp., ${total.toFixed(2).replace(".", ",")} m`);
    try { localStorage.removeItem(FM_KEY); } catch (_) { /* */ }
    fm.items = []; toast("Matavimas įrašytas į eilę"); stack.length = 0; show("home");
  };
}

async function boot() {
  $("btn-back").onclick = back;
  $("home-nav").onclick = () => openPicker((prj) => { if (prj) navigateToProject(prj); }, { title: "Naviguoti į projektą", nav: true });
  document.querySelectorAll("[data-go]").forEach((b) => {
    b.onclick = () => {
      const g = b.dataset.go;
      if (g === "note") initNote(); if (g === "marker") initMarker(); if (g === "expense") initExpense(); if (g === "files") initFiles(); if (g === "contact") initContact(); if (g === "field") initField(); if (g === "fm") initFm();
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
  bindNote(); bindMarker(); bindExpense(); bindFiles(); bindContact(); bindField(); bindFm(); bindAR(); bindSettings();
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
