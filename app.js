/***** GPM Jobs — Single-tenant + Excel *****/
const CLIENT_ID     = "c82303ea-e3b5-42da-84f5-f400a18148f8";
const TENANT        = "d643c7ef-29cf-4989-9503-928ec40af67c"; // your org only
const REDIRECT_URI  = "https://therealfulang.github.io/gpm-jobs/";
const SCOPES_LOGIN  = ["User.Read"]; // minimal on login
const SCOPES_EXCEL  = ["User.Read","Files.ReadWrite","offline_access"];

// Excel location (OneDrive)
const WORKBOOK_PATH = "/me/drive/root:/Documents/Gpm General Contracting/App/GPM Jobs.xlsx:/workbook";
const TABLE_NAME    = "Jobs";

const $ = s => document.querySelector(s);
const toast = m => { const t=$("#toast"); if(!t) return; t.textContent=m; t.classList.add("show"); setTimeout(()=>t.classList.remove("show"),1500); };
const isIOS = () => /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints>1);

// MSAL init (hard-coded tenant authority)
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT}`,
    redirectUri: REDIRECT_URI
  },
  cache: { cacheLocation:"localStorage", storeAuthStateInCookie:false }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;

msalInstance.handleRedirectPromise()
  .then(resp=>{
    if (resp?.account) account = resp.account;
    else {
      const accs = msalInstance.getAllAccounts();
      if (accs.length) account = accs[0];
    }
    updateAuthUI();
  })
  .catch(e=> showError(e, "Redirect error"));

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
    const resp = await msalInstance.loginPopup(req);
    account = resp.account; updateAuthUI(); toast("Signed in");
  } catch (e) {
    showError(e, "Sign-in failed");
  }
}
async function signOut() {
  try { await msalInstance.logoutPopup(); }
  catch (e) { showError(e, "Sign-out failed"); }
  finally { account=null; updateAuthUI(); }
}
window.signIn = signIn; window.signOut = signOut;

async function getToken(scopes) {
  const req = { scopes, account: account || msalInstance.getAllAccounts()[0] };
  try { const s = await msalInstance.acquireTokenSilent(req); return s.accessToken; }
  catch (e) {
    try { const p = await msalInstance.acquireTokenPopup(req); return p.accessToken; }
    catch (e2) { showError(e2, "Token error"); return null; }
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
    showError({ errorCode: res.status, errorMessage: txt }, `Graph ${method} ${path} failed`);
    throw new Error(`${method} ${path} → ${res.status}`);
  }
  return res.json();
}

async function loadJobs(){
  $("#cards").innerHTML = '<div class="hint">Loading…</div>';
  try {
    const data = await graph(`${WORKBOOK_PATH}/tables('${TABLE_NAME}')/rows`);
    const rows = (data?.value || []).map(r => r.values?.[0]).filter(Boolean);
    renderCards(rows, $("#stageFilter").value);
  } catch(e) { console.error(e); }
}

async function addJob(job){
  await graph(`${WORKBOOK_PATH}/tables('${TABLE_NAME}')/rows/add`, "POST", { values: [[job.number, job.name, job.status]] });
}

function renderCards(rows, stage="All"){
  const cards = $("#cards"); cards.innerHTML="";
  let count=0;
  rows.forEach((r,i)=>{
    const [number,name,status]=r;
    if(i===0 && String(number).toLowerCase().includes("number")) return;
    if(stage!=="All" && String(status)!==stage) return;
    const div=document.createElement("div");
    div.className=`card ${(status||"").toLowerCase()}`;
    div.innerHTML=`<div class="title">#${number} — ${name}</div><div class="status">STATUS: ${String(status||"").toUpperCase()}</div>`;
    cards.appendChild(div); count++;
  });
  if(!count) cards.innerHTML='<div class="hint">No jobs in this view.</div>';
}

function showError(e, label="Error") {
  const el = document.getElementById("cards");
  const msg = `${label}\n${e.errorCode||e.name||""}\n${e.errorMessage||e.message||""}`;
  el.innerHTML = `<div class="hint" style="white-space:pre-wrap">${msg}</div>`;
  console.error(e);
}

document.addEventListener("DOMContentLoaded", ()=>{
  $("#refresh")?.addEventListener("click", loadJobs);
  $("#stageFilter")?.addEventListener("change", ()=> loadJobs());
  $("#jobForm")?.addEventListener("submit", async e=>{
    e.preventDefault();
    const job={ number:$("#number").value.trim(), name:$("#name").value.trim(), status:$("#status").value.trim() };
    try{
      await addJob(job);
      toast("Job saved");
      e.target.reset(); await loadJobs();
    }catch(e2){ showError(e2, "Save failed"); }
  });
});
