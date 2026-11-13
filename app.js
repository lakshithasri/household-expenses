/* Household Expenses – phone app.
 * Signs in with a Microsoft account (OAuth 2.0 auth-code + PKCE, no backend) and appends rows to the
 * shared workbook's tables through the Microsoft Graph Excel API.
 * Add ?demo=1 to the URL to try the UI with fake data and no sign-in. */
"use strict";

const CFG = window.APP_CONFIG || {};
const DEMO = new URLSearchParams(location.search).has("demo");
const AUTH_BASE = `https://login.microsoftonline.com/${CFG.tenant || "consumers"}/oauth2/v2.0`;
const SCOPES = "openid profile offline_access User.Read Files.ReadWrite.All";
const GRAPH = "https://graph.microsoft.com/v1.0";
const REDIRECT = location.origin + location.pathname.replace(/index\.html$/, "");
const SETTINGS_RANGE = "A2:H41"; // Settings sheet: categories A:E, people F, pay methods G, income sources H

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const fmt = (n) => Math.round(n).toLocaleString("en-US");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

// ------------------------------------------------------------------ storage (per phone)
const PREFIX = DEMO ? "hxdemo." : "hx."; // demo data never touches the real app's storage
const store = {
  get(k, d = null) { try { const v = localStorage.getItem(PREFIX + k); return v === null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(PREFIX + k, JSON.stringify(v)); } catch { /* private mode */ } },
  del(k) { try { localStorage.removeItem(PREFIX + k); } catch {} },
};

// ------------------------------------------------------------------ dates
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const toSerial = (iso) => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - EXCEL_EPOCH) / 86400000; };
const fromSerial = (n) => new Date(EXCEL_EPOCH + Math.round(n) * 86400000);
const isoDay = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const monthKey = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
const thisMonth = () => isoDay(new Date()).slice(0, 7);
function toMonth(v) { // Settings Starts/Ends may be text "2027-01" or an Excel date serial
  if (v === "" || v == null) return null;
  if (typeof v === "number") return monthKey(fromSerial(v));
  const m = String(v).trim().match(/^(\d{4})-(\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, "0")}` : null;
}
function cellDate(v) {
  if (typeof v === "number" && v > 0) return fromSerial(v);
  const d = new Date(v);
  return isNaN(d) ? null : new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
}

// ------------------------------------------------------------------ auth (PKCE)
const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
const randStr = (n = 48) => b64url(crypto.getRandomValues(new Uint8Array(n)));

async function login(prompt) {
  const verifier = randStr(64);
  const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  const state = randStr(16);
  sessionStorage.setItem("hx.pkce", JSON.stringify({ verifier, state }));
  const p = new URLSearchParams({
    client_id: CFG.clientId, response_type: "code", redirect_uri: REDIRECT, response_mode: "query",
    scope: SCOPES, code_challenge: challenge, code_challenge_method: "S256", state,
  });
  const hint = store.get("username");
  if (prompt) p.set("prompt", prompt);
  else if (hint) p.set("login_hint", hint);
  location.assign(`${AUTH_BASE}/authorize?${p}`);
}

async function tokenRequest(body) {
  const r = await fetch(`${AUTH_BASE}/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CFG.clientId, scope: SCOPES, ...body }),
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error_description || j.error || "Sign-in failed");
  store.set("tok", { access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in - 120) * 1000 });
  if (j.id_token) {
    try {
      const claims = JSON.parse(atob(j.id_token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
      store.set("username", claims.preferred_username || claims.email || "");
      store.set("displayName", claims.name || "");
    } catch {}
  }
}

async function handleRedirect() {
  const q = new URLSearchParams(location.search);
  if (!q.has("code") && !q.has("error")) return;
  const saved = JSON.parse(sessionStorage.getItem("hx.pkce") || "null");
  history.replaceState(null, "", REDIRECT);
  sessionStorage.removeItem("hx.pkce");
  if (q.has("error")) throw new Error(q.get("error_description") || q.get("error"));
  if (!saved || saved.state !== q.get("state")) throw new Error("Sign-in response did not match; please try again.");
  await tokenRequest({ grant_type: "authorization_code", code: q.get("code"), redirect_uri: REDIRECT, code_verifier: saved.verifier });
}

async function getToken() {
  const t = store.get("tok");
  if (!t) return null;
  if (t.access && Date.now() < t.exp) return t.access;
  if (!t.refresh) return null;
  try { await tokenRequest({ grant_type: "refresh_token", refresh_token: t.refresh }); return store.get("tok").access; }
  catch (e) { if (navigator.onLine) store.del("tok"); return null; }
}

// ------------------------------------------------------------------ Graph / workbook
class NeedsLogin extends Error {}
class OfflineError extends Error {}

async function graph(path, { method = "GET", body, headers = {} } = {}, retry = true) {
  if (DEMO) return demoGraph(path, method, body);
  const token = await getToken();
  if (!token) throw new NeedsLogin("Please sign in again.");
  let r;
  try {
    r = await fetch(GRAPH + path, {
      method, body: body ? JSON.stringify(body) : undefined,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...headers },
    });
  } catch { throw new OfflineError("No connection"); }
  if (r.status === 401 && retry) { const t = store.get("tok"); if (t) { t.exp = 0; store.set("tok", t); } return graph(path, { method, body, headers }, false); }
  if (r.status === 429 || r.status === 503 || r.status === 504) throw new OfflineError("Excel is busy, will retry");
  if (r.status === 204) return null;
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error?.message || `Excel request failed (${r.status})`);
  return j;
}

function shareId(link) {
  const b = btoa(unescape(encodeURIComponent(link.trim())));
  return "u!" + b.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}

async function resolveWorkbook(link) {
  const item = await graph(`/shares/${shareId(link)}/driveItem?$select=id,name,parentReference`,
    { headers: { Prefer: "redeemSharingLinkIfNecessary" } });
  if (!/\.xlsx$/i.test(item.name)) throw new Error(`That link points to "${item.name}", not an .xlsx workbook.`);
  const wb = { base: `/drives/${item.parentReference.driveId}/items/${item.id}/workbook`, name: item.name };
  await graph(`${wb.base}/tables('tblExpenses')?$select=name`); // proves it is our tracker
  return wb;
}

async function loadSettings() {
  const r = await graph(`${state.wb.base}/worksheets('Settings')/range(address='${SETTINGS_RANGE}')?$select=values`);
  const rows = r.values;
  const col = (i) => rows.map((x) => String(x[i] ?? "").trim()).filter(Boolean);
  const s = {
    categories: rows.filter((x) => String(x[0] ?? "").trim()).map((x) => ({
      name: String(x[0]).trim(), group: String(x[1] || "Living").trim(), budget: Number(x[2]) || 0,
      starts: toMonth(x[3]), ends: toMonth(x[4]),
    })),
    people: col(5), methods: col(6), sources: col(7),
  };
  store.set("settings", s);
  return s;
}

async function loadEntries() {
  const [e, i] = await Promise.all([
    graph(`${state.wb.base}/tables('tblExpenses')/dataBodyRange?$select=values`),
    graph(`${state.wb.base}/tables('tblIncome')/dataBodyRange?$select=values`),
  ]);
  const parse = (vals, amountIdx) => vals.map((v, idx) => ({ idx, v, date: cellDate(v[0]), amount: Number(v[amountIdx]) }))
    .filter((x) => x.date && x.amount > 0);
  const data = { expenses: parse(e.values, 3), income: parse(i.values, 3), at: Date.now() };
  store.set("entries", { expenses: data.expenses.map(({ v }) => v).slice(-400), income: data.income.map(({ v }) => v).slice(-60), at: data.at });
  return data;
}

// ------------------------------------------------------------------ offline queue
function queue() { return store.get("queue", []); }
function setQueue(q) { store.set("queue", q); renderSyncState(); }

async function flushQueue() {
  let q = queue();
  if (!q.length || !state.wb || state.flushing) return;
  state.flushing = true;
  try {
    while (q.length) {
      const item = q[0];
      await graph(`${state.wb.base}/tables('${item.table}')/rows/add`, { method: "POST", body: { index: null, values: [item.values] } });
      q = queue().filter((x) => x.id !== item.id);
      setQueue(q);
    }
  } catch (e) {
    if (e instanceof NeedsLogin) showSetup("signin");
    else if (!(e instanceof OfflineError)) toast(`Couldn't save to Excel: ${e.message}`, "error");
  } finally {
    state.flushing = false;
    renderSyncState();
  }
}

// ------------------------------------------------------------------ state + UI
const state = { wb: store.get("wb"), settings: store.get("settings"), mode: "expense", form: {}, data: null, flushing: false };

const ICONS = {
  Groceries: "🛒", "Dining Out": "🍽️", "Electricity & Water": "💡", "Phone & Internet": "📱", "Fuel & Transport": "⛽",
  "Vehicle Maintenance": "🔧", Medical: "💊", "Personal Care": "🧴", Clothing: "👕", "Household Items": "🏠",
  "Gifts & Family": "🎁", Entertainment: "🎬", Education: "📚", Other: "📦", Rent: "🔑", "Loan Instalment": "🏦",
  "Vehicle Lease": "🚗", "Life Insurance": "🛡️", "Credit Card Repayment": "💳", "Friend Repayment": "🤝",
  "Cousin Repayment": "🤝", "Emergency Fund": "🆘", "House Fund": "🏡",
};
const GROUP_ORDER = ["Living", "Fixed", "Debt", "Savings"];
const GROUP_LABEL = { Living: "Everyday", Fixed: "Bills & commitments", Debt: "Debt repayments", Savings: "Savings" };

function toast(msg, kind = "ok", action) {
  const t = $("#toast");
  t.className = `toast show ${kind}`;
  t.innerHTML = `<span>${esc(msg)}</span>${action ? `<button type="button">${esc(action.label)}</button>` : ""}`;
  if (action) $("button", t).onclick = () => { t.className = "toast"; action.run(); };
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => (t.className = "toast"), action ? 7000 : 3200);
}

function show(id) {
  $$(".screen").forEach((s) => (s.hidden = s.id !== id));
}

function showSetup(step) {
  show("setup");
  $$(".setup-step").forEach((s) => (s.hidden = s.dataset.step !== step));
  if (step === "who") {
    const people = state.settings?.people || [];
    $("#whoList").innerHTML = people.map((p) => `<button type="button" class="chip big" data-v="${esc(p)}">${esc(p)}</button>`).join("");
    $$("#whoList .chip").forEach((b) => (b.onclick = () => { store.set("me", b.dataset.v); startMain(); }));
  }
  if (step === "link") $("#linkInput").value = "";
}

function renderSyncState() {
  const n = queue().length;
  const el = $("#sync");
  if (!el) return;
  el.className = "sync " + (n ? "pending" : navigator.onLine ? "ok" : "off");
  el.textContent = n ? `${n} waiting to sync` : navigator.onLine ? "Synced" : "Offline";
}

function chips(container, items, selected, onPick, cls = "chip") {
  container.innerHTML = items.map((v) => `<button type="button" class="${cls}${v === selected ? " on" : ""}" data-v="${esc(v)}">${esc(v)}</button>`).join("");
  $$("button", container).forEach((b) => (b.onclick = () => { $$("button", container).forEach((x) => x.classList.toggle("on", x === b)); onPick(b.dataset.v); }));
}

function renderForm() {
  const s = state.settings;
  const f = state.form;
  const isExp = state.mode === "expense";
  $$(".mode button").forEach((b) => b.classList.toggle("on", b.dataset.mode === state.mode));
  $("#expenseFields").hidden = !isExp;
  $("#incomeFields").hidden = isExp;
  $("#saveBtn").textContent = isExp ? "Save expense" : "Save income";

  if (isExp) {
    const month = thisMonth();
    const groups = GROUP_ORDER.map((g) => ({
      g, cats: s.categories.filter((c) => c.group === g && (!c.ends || c.ends >= month)),
    })).filter((x) => x.cats.length);
    $("#cats").innerHTML = groups.map(({ g, cats }) => `
      <details class="catgroup" ${g === "Living" || cats.some((c) => c.name === f.category) ? "open" : ""}>
        <summary>${esc(GROUP_LABEL[g] || g)}</summary>
        <div class="catgrid">${cats.map((c) => `
          <button type="button" class="cat${c.name === f.category ? " on" : ""}" data-v="${esc(c.name)}">
            <span class="ico">${ICONS[c.name] || "•"}</span><span>${esc(c.name)}</span></button>`).join("")}
        </div></details>`).join("");
    $$("#cats .cat").forEach((b) => (b.onclick = () => {
      $$("#cats .cat").forEach((x) => x.classList.toggle("on", x === b));
      f.category = b.dataset.v;
      suggestDescriptions();
      const c = s.categories.find((x) => x.name === f.category);
      if (c && (c.group === "Fixed" || c.group === "Debt") && !$("#amount").value && c.budget) $("#amount").value = c.budget;
    }));
    chips($("#paidBy"), s.people, f.paidBy, (v) => (f.paidBy = v));
    chips($("#method"), s.methods, f.method, (v) => { f.method = v; store.set("lastMethod", v); });
  } else {
    chips($("#person"), s.people, f.person, (v) => (f.person = v));
    chips($("#source"), s.sources, f.source, (v) => (f.source = v));
  }
}

function suggestDescriptions() {
  const rows = store.get("entries")?.expenses || [];
  const seen = new Map();
  for (let k = rows.length - 1; k >= 0 && seen.size < 12; k--) {
    const r = rows[k];
    if (r[1] === state.form.category && r[2] && !seen.has(r[2])) seen.set(r[2], 1);
  }
  $("#descList").innerHTML = [...seen.keys()].map((d) => `<option value="${esc(d)}">`).join("");
}

function resetForm(keepMode = true) {
  const me = store.get("me");
  state.form = { date: isoDay(new Date()), paidBy: me, person: me, method: store.get("lastMethod") || state.settings?.methods?.[0], source: "Salary" };
  if (!keepMode) state.mode = "expense";
  $("#amount").value = "";
  $("#desc").value = "";
  $("#notes").value = "";
  $("#incNotes").value = "";
  setDate(state.form.date);
  renderForm();
}

function setDate(iso) {
  state.form.date = iso;
  const today = isoDay(new Date());
  const y = new Date(); y.setDate(y.getDate() - 1);
  const yest = isoDay(y);
  $$("#dateChips button").forEach((b) => b.classList.toggle("on",
    (b.dataset.d === "today" && iso === today) || (b.dataset.d === "yesterday" && iso === yest) || (b.dataset.d === "pick" && iso !== today && iso !== yest)));
  $("#datePick").value = iso;
  $("#dateLabel").textContent = iso === today ? "" : new Date(iso + "T00:00").toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
}

async function save() {
  const amount = Number(String($("#amount").value).replace(/[, ]/g, ""));
  const f = state.form;
  if (!(amount > 0)) return toast("Enter an amount", "error"), $("#amount").focus();
  let item;
  if (state.mode === "expense") {
    if (!f.category) return toast("Pick a category", "error");
    if (!f.paidBy) return toast("Who paid?", "error");
    item = { table: "tblExpenses", values: [toSerial(f.date), f.category, $("#desc").value.trim(), amount, f.paidBy, f.method || "", $("#notes").value.trim()] };
  } else {
    if (!f.person) return toast("Whose income?", "error");
    item = { table: "tblIncome", values: [toSerial(f.date), f.person, f.source || "", amount, $("#incNotes").value.trim()] };
  }
  item.id = randStr(8);
  setQueue([...queue(), item]);
  const label = state.mode === "expense" ? `${f.category} · LKR ${fmt(amount)}` : `Income · LKR ${fmt(amount)}`;
  // optimistic local update so the summary reacts instantly
  const cache = store.get("entries") || { expenses: [], income: [] };
  (item.table === "tblExpenses" ? cache.expenses : cache.income).push(item.values);
  store.set("entries", cache);
  renderSummary();
  resetForm();
  window.scrollTo({ top: 0, behavior: "smooth" });
  await flushQueue();
  const pending = queue().some((x) => x.id === item.id);
  toast(pending ? `Saved on phone, will sync: ${label}` : `Saved to Excel ✓ ${label}`, "ok",
    pending ? null : { label: "Undo", run: () => undo(item) });
  if (!pending) refreshData();
}

async function undo(item) {
  try {
    const r = await graph(`${state.wb.base}/tables('${item.table}')/dataBodyRange?$select=values`);
    const same = (a, b) => a.length === b.length && a.every((v, k) => String(v) === String(b[k]) || Number(v) === Number(b[k]));
    let idx = -1;
    for (let k = r.values.length - 1; k >= 0; k--) if (same(r.values[k], item.values)) { idx = k; break; }
    if (idx < 0) return toast("Couldn't find that entry; delete it in Excel instead.", "error");
    await graph(`${state.wb.base}/tables('${item.table}')/rows/itemAt(index=${idx})`, { method: "DELETE" });
    toast("Entry removed");
    refreshData();
  } catch (e) { toast(`Undo failed: ${e.message}`, "error"); }
}

// ------------------------------------------------------------------ summary
function renderSummary() {
  const s = state.settings;
  const cache = store.get("entries");
  if (!s || !cache) return;
  const month = thisMonth();
  const groupOf = Object.fromEntries(s.categories.map((c) => [c.name, c.group]));
  const inMonth = (v) => { const d = cellDate(v[0]); return d && monthKey(d) === month && Number(v[3]) > 0; };
  const exp = cache.expenses.filter(inMonth);
  const byCat = {};
  let spent = 0, saved = 0;
  for (const v of exp) {
    const a = Number(v[3]);
    if (groupOf[v[1]] === "Savings") saved += a; else spent += a;
    byCat[v[1]] = (byCat[v[1]] || 0) + a;
  }
  const income = cache.income.filter(inMonth).reduce((t, v) => t + Number(v[3]), 0);
  const active = (c) => (!c.starts || c.starts <= month) && (!c.ends || c.ends >= month);
  const budget = s.categories.filter((c) => c.group !== "Savings" && active(c)).reduce((t, c) => t + c.budget, 0);
  const pct = budget ? spent / budget : 0;
  const day = new Date().getDate(), dim = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0).getDate();
  $("#sumMonth").textContent = new Date().toLocaleDateString("en-GB", { month: "long", year: "numeric" });
  $("#sumSpent").textContent = fmt(spent);
  $("#sumBudget").textContent = fmt(budget);
  $("#sumBar").style.width = Math.min(100, pct * 100) + "%";
  $("#sumBar").className = "bar " + (pct > 1 ? "over" : pct > day / dim + 0.1 ? "warn" : "");
  $("#sumMarker").style.left = (day / dim) * 100 + "%";
  $("#sumIncome").textContent = fmt(income);
  $("#sumSaved").textContent = fmt(saved);
  $("#sumLeft").textContent = fmt(income - spent - saved);

  const living = s.categories.filter((c) => c.group === "Living" && (c.budget > 0 || byCat[c.name]));
  $("#catBudgets").innerHTML = living.map((c) => {
    const a = byCat[c.name] || 0, p = c.budget ? a / c.budget : a > 0 ? 2 : 0;
    return `<div class="cb"><div class="cb-top"><span>${ICONS[c.name] || "•"} ${esc(c.name)}</span>
      <span class="${p > 1 ? "neg" : ""}">${fmt(a)} <small>/ ${fmt(c.budget)}</small></span></div>
      <div class="track"><div class="bar ${p > 1 ? "over" : p > 0.85 ? "warn" : ""}" style="width:${Math.min(100, p * 100)}%"></div></div></div>`;
  }).join("");

  const recent = cache.expenses.map((v) => ({ v, d: cellDate(v[0]) })).filter((x) => x.d && Number(x.v[3]) > 0).slice(-12).reverse();
  $("#recent").innerHTML = recent.length ? recent.map(({ v, d }) => `
    <li><span class="ico">${ICONS[v[1]] || "•"}</span>
      <div class="r-main"><b>${esc(v[2] || v[1])}</b><small>${esc(v[1])} · ${esc(v[4] || "")} · ${d.toLocaleDateString("en-GB", { day: "numeric", month: "short", timeZone: "UTC" })}</small></div>
      <span class="r-amt">${fmt(Number(v[3]))}</span></li>`).join("") : `<li class="empty">No expenses yet.</li>`;
  const at = cache.at ? new Date(cache.at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "–";
  $("#updated").textContent = `Updated ${at}`;
}

async function refreshData() {
  try {
    await loadEntries();
    renderSummary();
  } catch (e) {
    if (e instanceof NeedsLogin) return showSetup("signin");
    if (!(e instanceof OfflineError)) toast(e.message, "error");
  }
}

// ------------------------------------------------------------------ startup
async function startMain() {
  show("main");
  $("#who").textContent = store.get("me") || "";
  resetForm(false);
  renderSummary();
  renderSyncState();
  flushQueue();
  refreshData();
  try { state.settings = await loadSettings(); renderForm(); } catch (e) { if (e instanceof NeedsLogin) showSetup("signin"); }
}

async function boot() {
  wireEvents();
  if (DEMO) { initDemo(); }
  else if (!CFG.clientId || CFG.clientId.startsWith("PASTE")) { showSetup("config"); return; }
  try { await handleRedirect(); } catch (e) { toast(e.message, "error"); }
  if (!DEMO && !store.get("tok")) return showSetup("signin");
  if (!state.wb) return showSetup("link");
  if (!state.settings) {
    try { state.settings = await loadSettings(); } catch (e) {
      if (e instanceof NeedsLogin) return showSetup("signin");
      toast(e.message, "error"); return showSetup("link");
    }
  }
  if (!store.get("me") || !state.settings.people.includes(store.get("me"))) return showSetup("who");
  startMain();
}

function wireEvents() {
  $("#signinBtn").onclick = () => login("select_account");
  $("#linkBtn").onclick = async () => {
    const link = $("#linkInput").value.trim();
    if (!/^https:\/\//.test(link)) return toast("Paste the full https:// link to the Excel file", "error");
    $("#linkBtn").disabled = true; $("#linkBtn").textContent = "Checking…";
    try {
      state.wb = DEMO ? { base: "demo", name: "Household_Expenses.xlsx" } : await resolveWorkbook(link);
      store.set("wb", state.wb);
      state.settings = await loadSettings();
      toast(`Connected to ${state.wb.name}`);
      showSetup("who");
    } catch (e) {
      if (e instanceof NeedsLogin) return showSetup("signin");
      toast(`Couldn't open that file: ${e.message}`, "error");
    } finally { $("#linkBtn").disabled = false; $("#linkBtn").textContent = "Connect"; }
  };
  $$(".mode button").forEach((b) => (b.onclick = () => { state.mode = b.dataset.mode; renderForm(); }));
  $$("#dateChips button").forEach((b) => (b.onclick = () => {
    if (b.dataset.d === "today") setDate(isoDay(new Date()));
    else if (b.dataset.d === "yesterday") { const d = new Date(); d.setDate(d.getDate() - 1); setDate(isoDay(d)); }
    else { const p = $("#datePick"); p.showPicker ? p.showPicker() : p.click(); }
  }));
  $("#datePick").onchange = (e) => e.target.value && setDate(e.target.value);
  $("#entryForm").onsubmit = (e) => { e.preventDefault(); save(); };
  $("#refreshBtn").onclick = async () => { $("#refreshBtn").classList.add("spin"); await flushQueue(); await refreshData(); $("#refreshBtn").classList.remove("spin"); };
  $("#menuBtn").onclick = () => $("#menu").showModal();
  $("#menu").onclick = (e) => { if (e.target.id === "menu") $("#menu").close(); };
  $("#mWho").onclick = () => { $("#menu").close(); showSetup("who"); };
  $("#mLink").onclick = () => { $("#menu").close(); store.del("wb"); state.wb = null; showSetup("link"); };
  $("#mOut").onclick = () => {
    if (queue().length && !confirm("Some entries haven't synced yet and will be lost. Sign out anyway?")) return;
    ["tok", "wb", "settings", "entries", "queue", "me", "username", "displayName"].forEach(store.del);
    location.replace(REDIRECT + (DEMO ? "?demo=1" : ""));
  };
  window.addEventListener("online", () => { renderSyncState(); flushQueue(); });
  window.addEventListener("offline", renderSyncState);
  document.addEventListener("visibilitychange", () => { if (!document.hidden && state.wb && $("#main") && !$("#main").hidden) { flushQueue(); refreshData(); } });
  $("#amount").addEventListener("input", (e) => {
    const raw = e.target.value.replace(/[^\d.]/g, "");
    const [i, d] = raw.split(".");
    e.target.value = (i ? Number(i).toLocaleString("en-US") : "") + (d !== undefined ? "." + d.slice(0, 2) : "");
  });
}

// ------------------------------------------------------------------ demo backend (no sign-in, fake data)
function initDemo() {
  document.body.classList.add("demo");
  store.set("tok", { access: "demo", exp: Date.now() + 1e12 });
  if (!store.get("wb")) store.set("wb", { base: "demo", name: "Household_Expenses.xlsx" });
  if (!store.get("me")) store.set("me", "Alex");
  state.wb = store.get("wb");
}
const demoDb = (() => {
  const t = new Date(), y = t.getFullYear(), m = t.getMonth();
  const day = (d) => (Date.UTC(y, m, d) - EXCEL_EPOCH) / 86400000;
  return {
    settings: [
      ["Rent", "Fixed", 60000, "", "", "Alex", "Cash", "Salary"],
      ["Loan Instalment", "Fixed", 70000, "", "2029-06", "Sam", "Debit Card", "Bonus"],
      ["Vehicle Lease", "Fixed", 100000, "", "2030-06", "", "Credit Card", "Other"],
      ["Life Insurance", "Fixed", 20000, "", "", "", "Bank Transfer", ""],
      ["Credit Card Repayment", "Debt", 50000, "2027-01", "", "", "Online / App", ""],
      ["Groceries", "Living", 40000, "", "", "", "", ""], ["Dining Out", "Living", 10000, "", "", "", "", ""],
      ["Electricity & Water", "Living", 12000, "", "", "", "", ""], ["Phone & Internet", "Living", 6000, "", "", "", "", ""],
      ["Fuel & Transport", "Living", 20000, "", "", "", "", ""], ["Medical", "Living", 5000, "", "", "", "", ""],
      ["Clothing", "Living", 5000, "", "", "", "", ""], ["Other", "Living", 4000, "", "", "", "", ""],
      ["House Fund", "Savings", 0, "", "", "", "", ""],
    ],
    tblExpenses: [["", "", "", "", "", "", ""], [day(1), "Rent", "Monthly rent", 60000, "Alex", "Bank Transfer", ""],
      [day(2), "Groceries", "Keells weekly shop", 14250, "Sam", "Debit Card", ""], [day(3), "Fuel & Transport", "Petrol", 9000, "Alex", "Cash", ""],
      [day(4), "Dining Out", "Dinner", 6800, "Sam", "Credit Card", ""], [day(5), "Electricity & Water", "CEB bill", 8900, "Alex", "Online / App", ""]],
    tblIncome: [["", "", "", "", ""]],
  };
})();
async function demoGraph(path, method, body) {
  await new Promise((r) => setTimeout(r, 250));
  if (path.includes("/shares/")) return { id: "x", name: "Household_Expenses.xlsx", parentReference: { driveId: "d" } };
  if (path.includes("worksheets('Settings')")) return { values: demoDb.settings };
  const t = path.match(/tables\('(\w+)'\)/)?.[1];
  if (t && path.endsWith("/rows/add")) { demoDb[t].push(...body.values); return {}; }
  if (t && method === "DELETE") { const i = Number(path.match(/index=(\d+)/)[1]); demoDb[t].splice(i, 1); return null; }
  if (t && path.includes("dataBodyRange")) return { values: demoDb[t] };
  return {};
}

if ("serviceWorker" in navigator && !DEMO) navigator.serviceWorker.register("sw.js").catch(() => {});
boot();
