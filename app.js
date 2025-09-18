/***** GPM Jobs — Single-tenant + Excel *****/
const CLIENT_ID     = "c82303ea-e3b5-42da-84f5-f400a18148f8";
const TENANT        = "d643c7ef-29cf-4989-9503-928ec40af67c"; // your org only
const REDIRECT_URI  = "https://therealfulang.github.io/gpm-jobs/";
const SCOPES_LOGIN  = ["User.Read"]; // minimal on login to avoid consent blockers
const SCOPES_EXCEL  = ["User.Read","Files.ReadWrite","offline_access"]; // used only when calling Excel

// Excel location (OneDrive)
const WORKBOOK_PATH = "/me/drive/root:/Documents/Gpm General Contracting/App/GPM Jobs.xlsx:/workbook";
const TABLE_NAME    = "Jobs";

const $ = s => document.querySelector(s);
const toast = m => { const t=$("#toast"); if(!t) return; t.textContent=m; t.classList.add("show"); setTimeout(()=>t.classList.remove("show"),1500); };
const isIOS = () => /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints>1);

// Guard MSAL script
if (typeof window.msal === "undefined") {
  document.addEventListener("DOMContentLoaded", ()=>{
    const el = document.getElementById("cards");
    if (el) el.innerHTML = '<div class="hint">MSAL failed to load. Check network/script tag.</div>';
  });
  throw new Error("MSAL not found");
}

// MSAL init (single-tenant authority)
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT}`,
    redirectUri: REDIRECT_URI
  },
  cache: { cacheLocation:"localStorage", storeAuthStateInCookie:false },
  system: { loggerOptions: { loggerCallback: (_l, msg) => console.log("[MSAL]", msg) } }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;

// helper to show detailed msal/aad errors in the UI
function showMsalError(e, label="MSAL error") {
  const details = [
    `${label}`,
    `errorCode: ${e.errorCode || e.name || "n/a"}`,
    `message: ${e.errorMessage || e.message || "n/a"}`,
    `subError: ${e.subError || "n/a"}`,
    `correlationId: ${e.correlationId || "n/a"}`
  ].join("\n");
  console.error(e);
  const el = document.getElementById("cards");
  if (el) el.innerHTML = `<div class="hint" style="white-space:pre-wrap">${details}</div>`;
}

// handle login redirects (esp. iOS)
msalInstance.handleRedirectPromise()
  .then(resp=>{
    if (resp?.account) account = resp.account;
    else {
      const accs = msalInstance.getAllAccounts();
      if (accs.length) account = accs[0];
    }
    updateAuthUI();
  })
  .catch(e=> showMsalError(e, "Redirect processing failed"));

function updateAuthUI() {
  const signedIn = !!account;
  $("#signin").style.display  = signedIn ? "none" : "inline-block";
  $("#signout").style.display = signedIn ? "inline-block" : "none";
  $("#addForm").style.display = signedIn ? "block" : "none";
  if (signedIn) loadJobs();
}

async function signIn() {
  const req = { scopes: SCOPES_LOGIN, prompt: "select_account" };
  try {
    if (isIOS()) { await msalInstance.loginRedirect(req); return; }
    const resp = await msalInstance.loginPopup(req);
    account = resp.account; updateAuthUI(); toast("Signed in");
  } catch (e) {
    if (e.errorCode === "interaction_in_progress" || e.errorCode === "popup_window_error") {
      try { await msalInstance.loginRedirect(req); return; }
      catch (e2) { showMsalError(e2, "Redirect login failed"); return; }
    }
    showMsalError(e, "Popup login failed");
  }
}
async function signOut() {
  try { await (msalInstance.logoutPopup?.() ?? msalInstance.logoutRedirect?.()); }
  catch (e) { showMsalError(e, "Logout failed"); }
  finally { account=null; updateAuthUI(); toast("Signed out"); }
}
window.signIn = signIn; window.signOut = signOut;

async function getToken(scopes) {
  const req = { scopes, account: account || msalInstance.getAllAccounts()[0] };
  try { const s = await msalInstance.acquireTokenSilent(req); return s.accessToken; }
  catch (e) {
    try {
      if (isIOS()) { await msalInstance.acquireTokenRedirect(req); return null; }
      const p = await msalInstance.acquireTokenPopup(req); return p.accessToken;
    } catch (e2) { showMsalError(e2, "Token acquisition failed"); return null; }
  }
}

async function graph(path, method="GET", body=null){
  const token = await getToken(SCOPES_EXCEL); if(!token) return null;
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method,
    headers: { "Authorization": `Bearer ${token}`, "Content-Type":"application/json" },
    body: body ? JSON.stringify(body) : undefined
  });
  if (!res.ok) {
    const txt = await res.text().catch(()=>String(res.status));
    showMsalError({ errorCode: res.status, errorMessage: txt }, `Graph ${method} ${path} failed`);
    throw new Error(`${method} ${path} → ${res.status} ${txt}`);
  }
  return res.json();
}

async function workbookBase(){ return WORKBOOK_PATH; }

let ROWS=[];
async function loadJobs(){
  $("#cards").innerHTML = '<div class="hint">Loading…</div>';
  try {
    const base = await workbookBase();
    const data = await graph(`${base}/tables('${TABLE_NAME}')/rows`);
    ROWS = (data?.value || []).map(r => r.values?.[0]).filter(Boolean);
    renderCards(ROWS, $("#stageFilter").value);
  } catch(e) {
    console.error(e);
    // message already shown by showMsalError
  }
}
async function addJob(job){
  const base = await workbookBase();
  await graph(`${base}/tables('${TABLE_NAME}')/rows/add`, "POST", { values: [[job.number, job.name, job.status]] });
}

function renderCards(rows, stage="All"){
  const cards = $("#cards"); cards.innerHTML="";
  let count=0;
  rows.forEach((r,i)=>{
    const [number,name,status]=r;
    if(i===0 && String(number).toLowerCase().includes("number")) return; // skip header if present
    if(stage!=="All" && String(status)!==stage) return;
    const div=document.createElement("div");
    div.className=`card ${(status||"").toLowerCase()}`;
    div.innerHTML=`<div class="title">#${number} — ${name}</div><div class="status">STATUS: ${String(status||"").toUpperCase()}</div>`;
    cards.appendChild(div); count++;
  });
  if(!count) cards.innerHTML='<div class="hint">No jobs in this view.</div>';
}

document.addEventListener("DOMContentLoaded", ()=>{
  $("#refresh")?.addEventListener("click", loadJobs);
  $("#stageFilter")?.addEventListener("change", ()=> renderCards(ROWS, $("#stageFilter").value));
  $("#jobForm")?.addEventListener("submit", async e=>{
    e.preventDefault();
    const job={ number:$("#number").value.trim(), name:$("#name").value.trim(), status:$("#status").value.trim() };
    if(!job.number || !job.name){ $("#formMsg").className="msg err"; $("#formMsg").textContent="Enter Job # and Name"; return; }
    try{
      $("#formMsg").className="msg"; $("#formMsg").textContent="Saving…";
      await addJob(job);
      $("#formMsg").className="msg ok"; $("#formMsg").textContent="Saved ✔";
      e.target.reset(); await loadJobs();
    }catch(e2){
      $("#formMsg").className="msg err"; $("#formMsg").textContent="Save failed";
    }
  });

  // Optional PWA
  if ("serviceWorker" in navigator) window.addEventListener("load", ()=> navigator.serviceWorker.register("./sw.js"));
});
