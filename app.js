// ══════════════════════════════════════════════════════════════
// CONFIG
// ══════════════════════════════════════════════════════════════
const CFG = {
  month:6, year:2026, monthName:"June 2026",
  cutoffHour:15, workStart:8, slotMins:30,

  // SharePoint — Library: "ARTC CRM Operations", Subfolder: "Service Appointments"
  // Files inside branch folders: e.g. "AD - June 2026.xlsx"
  spDefaults: {
    site:   "https://alrostamanigroupae.sharepoint.com/sites/DataExchange",
    lib:    "ARTC CRM Operations",
    folder: "Service Appointments"
  },

  // spFolder: folder name on SharePoint (matches screenshots)
  // filePrefix: prefix used in xlsx filenames e.g. "AD" → "AD - June 2026.xlsx"
  branches:{
    RAK:      {label:"RAK",      spFolder:"RAK",           filePrefix:"RAK", dayOff:"Sunday",  advisors:["Arun","Hanan"],             cap:{wday:15,fri:10,sat:15,sun:0}, maps:"https://maps.google.com/?q=RAK+Suzuki+Service+Center,+Ras+Al+Khaimah"},
    SZR:      {label:"SZR",      spFolder:"SZR",           filePrefix:"SZR", dayOff:"Sunday",  advisors:["Khalid","Ehsan","Michille"], cap:{wday:42,fri:30,sat:42,sun:0}, maps:"https://maps.google.com/?q=SZR+Suzuki+Service+Center,+Sheikh+Zayed+Road+Dubai"},
    Sharjah:  {label:"Sharjah",  spFolder:"Sharjah",       filePrefix:"SHJ", dayOff:"Sunday",  advisors:["Omar","Hamdan"],             cap:{wday:28,fri:20,sat:35,sun:0}, maps:"https://maps.google.com/?q=Sharjah+Suzuki+Service+Center,+Sharjah"},
    Deira:    {label:"Deira",    spFolder:"Deira",          filePrefix:"DXB", dayOff:"Sunday",  advisors:["Ramy","Chona","Churchill"],  cap:{wday:28,fri:20,sat:35,sun:0}, maps:"https://maps.google.com/?q=Deira+Suzuki+Service+Center,+Dubai"},
    Musaffah: {label:"Musaffah", spFolder:"Musaffah (AD)",  filePrefix:"AD",  dayOff:"Sunday",  advisors:["Hossam","Dayanand"],         cap:{wday:28,fri:18,sat:40,sun:0}, maps:"https://maps.google.com/?q=Musaffah+Suzuki+Service+Center,+Abu+Dhabi"},
    AlAin:    {label:"Al Ain",   spFolder:"Al-Ain",         filePrefix:"AIN", dayOff:"Sunday",  advisors:["Hammad","Fazil"],            cap:{wday:15,fri:10,sat:15,sun:0}, maps:"https://maps.google.com/?q=Al+Ain+Suzuki+Service+Center,+Al+Ain"},
    Fujairah: {label:"Fujairah", spFolder:"Fujairah",       filePrefix:"FUJ", dayOff:"Friday",  advisors:["Sharif"],                   cap:{wday:5, fri:0, sat:3, sun:5},  maps:"https://maps.google.com/?q=Fujairah+Suzuki+Service+Center,+Fujairah"}
  },

  twixorBase:    "https://twixor.karix.ae",
  twixorChannel: "699c4aa677e05867ae76b322",
  twixorTplId:   "appointment_confirmation",
  twixorTplUId:  "4226955467558091"
};

// ══════════════════════════════════════════════════════════════
// UAE PLATE NUMBER DEFINITIONS
// Format rules:
//   Dubai:   DXB + code(letters) + number  →  DXBDD78387
//   AbuDhabi:AUH + code(digits)  + number  →  AUH1387209  (no code → AUH387209)
//   Sharjah: SHJ + code(digits)  + number  →  SHJ318273   (no code → SHJ18263)
//   Others:  PREFIX + code(letters) + number
// ══════════════════════════════════════════════════════════════
const PLATE_EMIRATES = {
  DXB:{
    name:"Dubai", prefix:"DXB",
    codes:["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
           "AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
           "BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ",
           "CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ",
           "DA","DB","DC","DD","DE","DF","DG","DH","DI","DJ","DK","DL","DM","DN","DO","DP","DQ","DR","DS","DT","DU","DV","DW","DX","DY","DZ",
           "EA","EB","EC","ED","EE","EF","EG","EH","EI","EJ","EK","EL","EM","EN","EO","EP","EQ","ER","ES","ET","EU","EV","EW","EX","EY","EZ",
           "FA","FB","FC","FD","FE","FF","FG","FH","FI","FJ","FK","FL","FM","FN","FO","FP","FQ","FR","FS","FT","FU","FV","FW","FX","FY","FZ"]
  },
  AUH:{
    name:"Abu Dhabi", prefix:"AUH",
    codes:["","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20",
           "21","22","23","24","25","26","27","28","29","30","31","32","33","34","35","36","37","38","39","40",
           "41","42","43","44","45","46","47","48","49","50","51","52","53","54","55","56","57","58","59","60",
           "61","62","63","64","65","66","67","68","69","70","71","72","73","74","75","76","77","78","79","80",
           "81","82","83","84","85","86","87","88","89","90","91","92","93","94","95","96","97","98","99","100"]
  },
  SHJ:{
    name:"Sharjah", prefix:"SHJ",
    codes:["","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20",
           "21","22","23","24","25","26","27","28","29","30","31","32","33","34","35"]
  },
  AJM:{
    name:"Ajman", prefix:"AJM",
    codes:["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
  },
  UAQ:{
    name:"Umm Al Quwain", prefix:"UAQ",
    codes:["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
  },
  RAK:{
    name:"Ras Al Khaimah", prefix:"RAK",
    codes:["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
  },
  FUJ:{
    name:"Fujairah", prefix:"FUJ",
    codes:["","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
  }
};

function buildPlate(emirateKey, code, num) {
  if (!emirateKey || !num) return "";
  const em = PLATE_EMIRATES[emirateKey];
  if (!em) return "";
  return em.prefix + code + num.replace(/\s/g,"").toUpperCase();
}

// ══════════════════════════════════════════════════════════════
// CONSTANTS — month-aware
// ══════════════════════════════════════════════════════════════
function getDaysInMonth(m, y) { return new Date(y, m, 0).getDate(); }
function getMonthDays(m, y) {
  const n=getDaysInMonth(m,y);
  return Array.from({length:n},(_,i)=>{const d=i+1;return{day:d,dow:new Date(y,m-1,d).getDay()};});
}
const DOW   = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
const DOW_S = ["Su","Mo","Tu","We","Th","Fr","Sa"];
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const TODAY_DAY = (new Date().getMonth()+1===CFG.month && new Date().getFullYear()===CFG.year) ? new Date().getDate() : 1;

// Mutable — updated when user switches month
let DAYS_IN_MONTH = getDaysInMonth(CFG.month, CFG.year);
let JUNE = getMonthDays(CFG.month, CFG.year); // "JUNE" = active month's day array (legacy alias kept)

// ══════════════════════════════════════════════════════════════
// STATE
// ══════════════════════════════════════════════════════════════
const appts={}, spLoaded={}, spFileExists={};
let activeB=null, activeDay=TODAY_DAY, activeAdv=null, selSlot=null;
let demoMode=false, spCfg=null, currentAccount=null;
let rschTarget=null, waTarget=null, twxToken=null, twxUser=null;
let activeMonth=CFG.month, activeYear=CFG.year;
const agentLog=[];

function reinitBranchState() {
  Object.keys(CFG.branches).forEach(b=>{
    appts[b]={}; spLoaded[b]={}; spFileExists[b]=null;
    for(let d=1;d<=DAYS_IN_MONTH;d++) appts[b][d]=[];
  });
}
reinitBranchState();

// ══════════════════════════════════════════════════════════════
// CAPACITY / SLOTS
// ══════════════════════════════════════════════════════════════
function getDayCap(branch, dow) {
  const b=CFG.branches[branch], day=DOW[dow];
  if(day===b.dayOff) return 0;
  if(day==="Friday")   return b.cap.fri;
  if(day==="Saturday") return b.cap.sat;
  if(day==="Sunday")   return b.cap.sun??b.cap.wday;
  return b.cap.wday;
}

/** All half-hour slots 08:00–14:30 (last slot starts before 15:00) */
function allSlots() {
  const s=[];
  for(let m=CFG.workStart*60; m<CFG.cutoffHour*60; m+=CFG.slotMins)
    s.push(`${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`);
  return s;
}

/** Returns all time slots for a given advisor on a day.
 *  Slots 08:00–14:30 are all available; per-advisor capacity
 *  is enforced separately (cap ÷ number of advisors). */
function advisorSlots(branch, dow, advisor) {
  const cap=getDayCap(branch,dow);
  if(!cap) return [];
  return allSlots();
}

/** Per-advisor daily capacity ceiling */
function advisorCap(branch, dow) {
  const bc=getDayCap(branch,dow);
  if(!bc) return 0;
  return Math.ceil(bc/CFG.branches[branch].advisors.length);
}

// ══════════════════════════════════════════════════════════════
// TOKEN-PASTE AUTHENTICATION
// ══════════════════════════════════════════════════════════════
let _accessToken=null, _tokenExp=null, _tokenTimer=null;

function _jwtExp(token) {
  try { return JSON.parse(atob(token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/'))).exp; } catch { return null; }
}
function _jwtName(token) {
  try {
    const p=JSON.parse(atob(token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/')));
    return p.name||p.unique_name?.split('@')[0]||p.upn?.split('@')[0]||'Agent';
  } catch { return 'Agent'; }
}
function _storeToken(token) {
  _accessToken=token.trim(); _tokenExp=_jwtExp(_accessToken);
  sessionStorage.setItem('svc_token',_accessToken); _startExpiryWatch();
}
function _loadStoredToken() {
  const t=sessionStorage.getItem('svc_token'); if(!t) return false;
  const exp=_jwtExp(t);
  if(exp&&Date.now()/1000>exp){sessionStorage.removeItem('svc_token');return false;}
  _accessToken=t; _tokenExp=exp; _startExpiryWatch(); return true;
}
function _startExpiryWatch() {
  clearInterval(_tokenTimer); if(!_tokenExp) return;
  _tokenTimer=setInterval(()=>{
    const secsLeft=_tokenExp-Date.now()/1000;
    const warn=document.getElementById('token-warn');
    if(secsLeft<=0){
      clearInterval(_tokenTimer); _accessToken=null; sessionStorage.removeItem('svc_token');
      if(warn) warn.style.display='none';
      toast('Access token expired — please paste a new one','t-err'); openTokenRefresh();
    } else if(secsLeft<600){if(warn) warn.style.display='flex';}
  },15000);
}
function getToken() {
  if(demoMode) return null;
  if(_accessToken){
    if(_tokenExp&&Date.now()/1000>_tokenExp){_accessToken=null;sessionStorage.removeItem('svc_token');openTokenRefresh();return null;}
    return _accessToken;
  }
  return null;
}

function doTokenConnect() {
  const raw=document.getElementById('token-input').value.trim();
  if(!raw||!raw.startsWith('ey')){toast('Paste a valid JWT access token (starts with "ey…")','t-err');return;}
  const exp=_jwtExp(raw);
  if(exp&&Date.now()/1000>exp){toast('This token has already expired — copy a fresh one from Graph Explorer','t-err');return;}
  _storeToken(raw); currentAccount={name:_jwtName(raw)};
  logAction('login',`Connected via Graph Explorer token as ${currentAccount.name}`);
  afterAuth();
}
function openTokenRefresh()  {document.getElementById('token-refresh-modal').classList.add('show');}
function closeTokenRefresh() {document.getElementById('token-refresh-modal').classList.remove('show');}
function applyRefreshedToken() {
  const raw=document.getElementById('token-refresh-input').value.trim();
  if(!raw||!raw.startsWith('ey')){toast('Paste a valid token','t-err');return;}
  const exp=_jwtExp(raw);
  if(exp&&Date.now()/1000>exp){toast('This token is already expired','t-err');return;}
  _storeToken(raw); currentAccount={name:_jwtName(raw)};
  document.getElementById('tb-user').textContent=currentAccount.name.split(' ')[0];
  document.getElementById('token-warn').style.display='none';
  closeTokenRefresh(); toast('✓ Token refreshed — connected','t-ok'); setSP('pill-live','Connected');
}
function doSignOut(){_accessToken=null;sessionStorage.removeItem('svc_token');clearInterval(_tokenTimer);location.reload();}

function afterAuth() {
  const saved=localStorage.getItem('svc_sp_cfg');
  if(saved){spCfg=JSON.parse(saved);show('app');hide('authScreen');launch();}
  else{
    document.getElementById('cfg-site').value  =CFG.spDefaults.site;
    document.getElementById('cfg-lib').value   =CFG.spDefaults.lib;
    document.getElementById('cfg-folder').value=CFG.spDefaults.folder;
    show('setupScreen');hide('authScreen');
  }
}
function enterDemo(){demoMode=true;show('app');hide('authScreen');hide('setupScreen');launch();}
function saveSetup() {
  const site=document.getElementById('cfg-site').value.trim().replace(/\/$/,'');
  if(!site){toast('SharePoint site URL required','t-err');return;}
  spCfg={
    site,
    lib:   document.getElementById('cfg-lib').value.trim()   ||CFG.spDefaults.lib,
    folder:document.getElementById('cfg-folder').value.trim().replace(/^\/|\/$/g,'')
  };
  localStorage.setItem('svc_sp_cfg',JSON.stringify(spCfg));
  show('app');hide('setupScreen');launch();
}
function showSetup() {
  if(spCfg){
    document.getElementById('cfg-site').value  =spCfg.site  ||CFG.spDefaults.site;
    document.getElementById('cfg-lib').value   =spCfg.lib   ||CFG.spDefaults.lib;
    document.getElementById('cfg-folder').value=spCfg.folder||CFG.spDefaults.folder;
  }
  show('setupScreen');hide('app');
}

// ══════════════════════════════════════════════════════════════
// SHAREPOINT — GRAPH API
// ══════════════════════════════════════════════════════════════
let _di=null;
async function getDI() {
  if(_di) return _di;
  const token=getToken(); if(!token||!spCfg) return null;
  try {
    const hn=new URL(spCfg.site).hostname, sp=new URL(spCfg.site).pathname;
    const sr=await fetch(`https://graph.microsoft.com/v1.0/sites/${hn}:${sp}`,{headers:{Authorization:`Bearer ${token}`}});
    if(!sr.ok) return null;
    const site=await sr.json();
    const dr=await fetch(`https://graph.microsoft.com/v1.0/sites/${site.id}/drives`,{headers:{Authorization:`Bearer ${token}`}});
    if(!dr.ok) return null;
    const drives=await dr.json();
    const match=drives.value.find(d=>d.name.toLowerCase()===spCfg.lib.toLowerCase())||drives.value[0];
    _di={siteId:site.id,driveId:match.id}; return _di;
  } catch { return null; }
}

/** Build SharePoint path for a file.
 *  Structure: [folder]/[branchSpFolder]/[filePrefix] - [Month] [Year].xlsx
 *  e.g.  Service Appointments/Musaffah (AD)/AD - June 2026.xlsx  */
function spFileName(b, mo, yr) {
  mo=mo||activeMonth; yr=yr||activeYear;
  return `${CFG.branches[b].filePrefix} - ${MONTHS[mo-1]} ${yr}.xlsx`;
}
function spPath(b, mo, yr) {
  mo=mo||activeMonth; yr=yr||activeYear;
  const base=spCfg?.folder?`${spCfg.folder}/`:"";
  return `${base}${CFG.branches[b].spFolder}/${spFileName(b,mo,yr)}`;
}

async function checkFile(b) {
  if(demoMode||spFileExists[b]!==null) return spFileExists[b]??true;
  const token=getToken(), di=await getDI();
  if(!token||!di){spFileExists[b]=false;return false;}
  const r=await fetch(`https://graph.microsoft.com/v1.0/sites/${di.siteId}/drives/${di.driveId}/root:/${encodeURIComponent(spPath(b))}`,{headers:{Authorization:`Bearer ${token}`}});
  spFileExists[b]=r.ok; return r.ok;
}

/** List xlsx files in a branch's SharePoint folder to discover available months.
 *  Parses filenames like "AD - June 2026.xlsx", "AD - MAY 2026.xlsx", "RAK - July 2026.xlsx" */
async function listBranchFiles(b) {
  if(demoMode) return [];
  const token=getToken(), di=await getDI();
  if(!token||!di) return [];
  const br=CFG.branches[b];
  const folderPath=(spCfg?.folder?`${spCfg.folder}/`:"")+br.spFolder;
  try {
    const r=await fetch(
      `https://graph.microsoft.com/v1.0/sites/${di.siteId}/drives/${di.driveId}/root:/${encodeURIComponent(folderPath)}:/children?$select=name,lastModifiedDateTime`,
      {headers:{Authorization:`Bearer ${token}`}}
    );
    if(!r.ok) return [];
    const data=await r.json();
    const files=[];
    (data.value||[]).forEach(item=>{
      const name=item.name||"";
      if(!name.match(/\.xlsx?$/i)) return;
      // Match: anything - Month Year.xlsx  or  anything Month Year.xlsx
      const base=name.replace(/\.xlsx?$/i,"");
      const match=base.match(/[-\s]([A-Za-z]+)\s+(\d{4})$/);
      if(match){
        const mi=MONTHS.findIndex(x=>x.toLowerCase()===match[1].toLowerCase());
        const yr=parseInt(match[2]);
        if(mi>=0&&yr>=2024) files.push({month:mi+1,year:yr,filename:name});
      }
    });
    files.sort((a,b)=>a.year===b.year?a.month-b.month:a.year-b.year);
    return files;
  } catch { return []; }
}

async function spRange(b, day, addr, values) {
  if(demoMode) return {ok:true};
  const token=getToken(), di=await getDI(); if(!token||!di) return {ok:false};
  const r=await fetch(
    `https://graph.microsoft.com/v1.0/sites/${di.siteId}/drives/${di.driveId}/root:/${encodeURIComponent(spPath(b))}:/workbook/worksheets('${day}')/range(address='${addr}')`,
    {method:"PATCH",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify({values})}
  );
  return {ok:r.ok};
}
async function spUsedRange(b, day) {
  if(demoMode) return null;
  const token=getToken(), di=await getDI(); if(!token||!di) return null;
  const r=await fetch(
    `https://graph.microsoft.com/v1.0/sites/${di.siteId}/drives/${di.driveId}/root:/${encodeURIComponent(spPath(b))}:/workbook/worksheets('${day}')/usedRange`,
    {headers:{Authorization:`Bearer ${token}`}}
  );
  if(!r.ok) return null; return await r.json();
}
async function scanSheet(b, day) {
  const data=await spUsedRange(b,day);
  if(!data||!data.values) return {lastDataRow:4,summaryRow:null,rschLogRow:null,totalRows:4};
  const rows=data.values; let lastDataRow=4, summaryRow=null, rschLogRow=null;
  for(let i=0;i<rows.length;i++){
    const txt=String(rows[i][0]||rows[i][1]||"").toLowerCase();
    if(txt.includes("daily summary"))       summaryRow=i+1;
    else if(txt.includes("reschedule log")) rschLogRow=i+1;
    else if(!summaryRow&&(rows[i][1]||rows[i][2])) lastDataRow=i+1;
  }
  return {lastDataRow,summaryRow,rschLogRow,totalRows:rows.length};
}
async function writeBookingRow(b, day, appt) {
  if(demoMode) return {ok:true};
  if(!await checkFile(b)) return {ok:false,reason:"file_not_found"};
  const {lastDataRow}=await scanSheet(b,day);
  const row=lastDataRow+1;
  const date=new Date(activeYear,activeMonth-1,day).toLocaleDateString("en-GB");
  const now=new Date().toLocaleDateString("en-GB");
  const vals=[[date,appt.name,appt.phone,appt.plate,appt.model,appt.service,appt.advisor,appt.time,null,null,now,appt.agent,appt.remarks||""]];
  const res=await spRange(b,day,`B${row}:N${row}`,vals);
  if(res.ok) appt.spRow=row;
  return res;
}
async function writeRescheduleLog(b, day, entry) {
  if(demoMode) return {ok:true};
  let {rschLogRow,totalRows}=await scanSheet(b,day);
  if(!rschLogRow){
    const hdrRow=totalRows+2;
    await spRange(b,day,`A${hdrRow}:A${hdrRow}`,[["RESCHEDULE LOG"]]);
    await spRange(b,day,`A${hdrRow+1}:J${hdrRow+1}`,[["Old Date","Old Time","New Date","New Time","Customer","Plate","Advisor","Rescheduled By","Timestamp","Reason"]]);
    totalRows=hdrRow+1;
  }
  const nextRow=totalRows+1;
  const vals=[[entry.oldDate,entry.oldTime,entry.newDate,entry.newTime,entry.customer,entry.plate,entry.advisor,entry.by,entry.ts,entry.reason]];
  return await spRange(b,day,`A${nextRow}:J${nextRow}`,vals);
}

// ══════════════════════════════════════════════════════════════
// MONTH BROWSER
// ══════════════════════════════════════════════════════════════
async function openMonthPicker() {
  const modal=document.getElementById("month-picker-modal");
  modal.classList.add("show");
  const body=document.getElementById("month-picker-body");
  body.innerHTML=`<p style="font-size:12px;color:var(--gray-600);padding:8px 0">Scanning SharePoint for available files…</p>`;

  if(demoMode){
    const demo=[
      {month:4,year:2026,filename:"Demo - April 2026.xlsx"},
      {month:5,year:2026,filename:"Demo - May 2026.xlsx"},
      {month:6,year:2026,filename:"Demo - June 2026.xlsx"},
      {month:7,year:2026,filename:"Demo - July 2026.xlsx"}
    ];
    renderMonthPickerOptions(demo,body); return;
  }
  if(!activeB){body.innerHTML=`<p style="font-size:12px;color:var(--gray-600)">Select a branch first.</p>`;return;}
  const files=await listBranchFiles(activeB);
  if(!files.length){
    body.innerHTML=`<p style="font-size:12px;color:var(--gray-600)">No Excel files found for <strong>${CFG.branches[activeB].label}</strong>.<br>Folder searched: <code>${(spCfg?.folder||"")+"/"+CFG.branches[activeB].spFolder}</code></p>`;
    return;
  }
  renderMonthPickerOptions(files,body);
}
function renderMonthPickerOptions(files,body){
  body.innerHTML=files.map(f=>{
    const isCur=f.month===activeMonth&&f.year===activeYear;
    return`<div class="month-opt${isCur?" month-opt-cur":""}" onclick="selectMonth(${f.month},${f.year})">
      <div class="month-opt-name">${MONTHS[f.month-1]} ${f.year}${isCur?` <span class="badge" style="background:var(--blue-light);color:var(--blue-dark);margin-left:4px">Current</span>`:""}</div>
      <div class="month-opt-file">${f.filename}</div>
    </div>`;
  }).join("");
}
async function selectMonth(mo,yr){
  closeMonthPicker();
  if(mo===activeMonth&&yr===activeYear) return;
  activeMonth=mo; activeYear=yr;
  DAYS_IN_MONTH=getDaysInMonth(mo,yr);
  JUNE=getMonthDays(mo,yr);
  reinitBranchState();
  activeDay=1; selSlot=null;
  document.getElementById("tb-month").textContent=`${MONTHS[mo-1]} ${yr}`;
  const sblabels=document.querySelectorAll(".sb-label");
  if(sblabels[1]) sblabels[1].textContent=`${MONTHS[mo-1]} ${yr}`;
  buildCalDows();
  if(activeB){
    showLoader("Loading…"); await loadDay(activeB,activeDay); hideLoader(); renderMain();
  }
  toast(`Switched to ${MONTHS[mo-1]} ${yr}`,"t-ok");
}
function closeMonthPicker(){document.getElementById("month-picker-modal").classList.remove("show");}

// ══════════════════════════════════════════════════════════════
// LOAD FROM SHAREPOINT
// ══════════════════════════════════════════════════════════════
async function loadDay(b, day, force=false) {
  if((spLoaded[b][day]&&!force)||demoMode) return;
  if(!await checkFile(b)){spLoaded[b][day]=true;return;}
  setSP("pill-sync","Reading…");
  const data=await spUsedRange(b,day);
  if(data?.values){
    const parsed=[];
    for(let i=4;i<data.values.length;i++){
      const row=data.values[i]; if(!row) continue;
      const txt=String(row[0]||row[1]||"").toLowerCase();
      if(txt.includes("daily summary")||txt.includes("reschedule log")) break;
      if(!row[1]&&!row[2]) continue;
      const raw=row[8]; let t="";
      if(typeof raw==="number"){const m=Math.round(raw*24*60);t=`${String(Math.floor(m/60)).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;}
      else if(typeof raw==="string") t=raw.substring(0,5);
      if(!t||t.length<4) continue;
      const isReschd=String(row[13]||"").toUpperCase().startsWith("RESCHEDULED");
      parsed.push({time:t,name:String(row[2]||""),phone:String(row[3]||""),plate:String(row[4]||""),model:String(row[5]||""),service:String(row[6]||""),advisor:String(row[7]||""),agent:String(row[12]||""),remarks:String(row[13]||""),fromSP:true,spRow:i+1,reschd:isReschd});
    }
    appts[b][day]=parsed.filter(a=>a.time);
  }
  spLoaded[b][day]=true;
  setSP(demoMode?"pill-demo":"pill-live",demoMode?"Demo":"Connected");
}

// ══════════════════════════════════════════════════════════════
// LAUNCH
// ══════════════════════════════════════════════════════════════
function launch() {
  document.getElementById("tb-month").textContent=`${MONTHS[activeMonth-1]} ${activeYear}`;
  if(currentAccount){document.getElementById("tb-user").textContent=currentAccount.name?.split(" ")[0]||"Agent";setSP("pill-live","Connected");}
  else setSP("pill-demo","Demo mode");
  buildNav(); buildCalDows();
  selectBranch(Object.keys(CFG.branches)[0]);
}
function setSP(cls,lbl){document.getElementById("sync-pill").className="pill "+cls;document.getElementById("sync-lbl").textContent=lbl;}

// ══════════════════════════════════════════════════════════════
// SIDEBAR / NAV
// ══════════════════════════════════════════════════════════════
function buildNav() {
  document.getElementById("branch-nav").innerHTML=Object.entries(CFG.branches).map(([id,br])=>
    `<button class="nav-btn" id="nb_${id}" onclick="selectBranch('${id}')"><span class="nav-dot" id="nd_${id}"></span>${br.label}<span class="nb-count" id="nc_${id}">—</span></button>`
  ).join("");
}
function updateCounts() {
  Object.keys(CFG.branches).forEach(b=>{
    const total=(appts[b][activeDay]||[]).filter(a=>!a.reschd).length;
    const cap=getDayCap(b,JUNE[activeDay-1]?.dow??0);
    const el=document.getElementById(`nc_${b}`); if(!el) return;
    el.textContent=`${total}/${cap||"—"}`;
    const pct=cap>0?total/cap:0;
    el.className=pct>=1?"nb-count full":pct>=0.8?"nb-count warn":"nb-count";
    const dot=document.getElementById(`nd_${b}`);
    if(dot) dot.style.background=pct>=1?"#E24B4A":pct>=0.8?"#EF9F27":"#C0DD97";
  });
}
async function selectBranch(b) {
  activeB=b; activeAdv=CFG.branches[b].advisors[0]; selSlot=null;
  document.querySelectorAll(".nav-btn").forEach(el=>el.classList.remove("active"));
  document.getElementById(`nb_${b}`)?.classList.add("active");
  showLoader("Loading schedule…"); await loadDay(b,activeDay); hideLoader(); renderMain();
}

// ══════════════════════════════════════════════════════════════
// MINI CALENDAR
// ══════════════════════════════════════════════════════════════
function buildCalDows() {
  document.getElementById("mc-dow").innerHTML=DOW_S.map(d=>`<div class="mc-dow">${d}</div>`).join("");
  const labels=document.querySelectorAll(".sb-label");
  if(labels[1]) labels[1].textContent=`${MONTHS[activeMonth-1]} ${activeYear}`;
}
function buildCal() {
  const firstDow=JUNE[0]?.dow??0;
  let html="";
  for(let i=0;i<firstDow;i++) html+=`<div class="mcd mcd-x"></div>`;
  JUNE.forEach(({day,dow})=>{
    const closed=activeB&&DOW[dow]===CFG.branches[activeB]?.dayOff;
    const total=activeB?(appts[activeB][day]||[]).filter(a=>!a.reschd).length:0;
    const cap=activeB?getDayCap(activeB,dow):0;
    const pct=cap>0?total/cap:0;
    let cls="mcd";
    if(closed) cls+=" mcd-cl";
    else if(day===activeDay) cls+=" mcd-sel";
    else if(pct>=1) cls+=" mcd-full";
    else if(pct>=0.75) cls+=" mcd-near";
    else if(total>0) cls+=" mcd-has";
    html+=`<div class="${cls}" ${closed?"":"onclick=\"jd("+day+")\""} title="${DOW[dow]} ${day}">${day}</div>`;
  });
  document.getElementById("mc-grid").innerHTML=html;
}
function jd(day){
  activeDay=day; selSlot=null;
  if(activeB){showLoader("Loading…");loadDay(activeB,day).then(()=>{hideLoader();renderMain();});}
  updateCounts();
}

// ══════════════════════════════════════════════════════════════
// PLATE BUILDER UI
// ══════════════════════════════════════════════════════════════
function plateEmirateChanged(){
  const ek=document.getElementById("f_plate_emirate")?.value;
  const codeSelect=document.getElementById("f_plate_code");
  if(!ek||!codeSelect) return;
  const em=PLATE_EMIRATES[ek]; if(!em) return;
  codeSelect.innerHTML=em.codes.map(c=>`<option value="${c}">${c===""?"(no code)":c}</option>`).join("");
  updatePlatePreview();
}
function updatePlatePreview(){
  const ek  =(document.getElementById("f_plate_emirate")?.value||"");
  const code=(document.getElementById("f_plate_code")?.value||"");
  const num =(document.getElementById("f_plate_num")?.value||"").replace(/\s/g,"").toUpperCase();
  const plate=buildPlate(ek,code,num);
  const prev=document.getElementById("f_plate_preview");
  if(prev) prev.textContent=plate||"—";
  const hidden=document.getElementById("f_plate");
  if(hidden) hidden.value=plate;
}

// ══════════════════════════════════════════════════════════════
// MAIN RENDER
// ══════════════════════════════════════════════════════════════
function renderMain() {
  if(!activeB) return;
  const b=CFG.branches[activeB];
  const dayObj=JUNE[activeDay-1]; if(!dayObj) return;
  const {day,dow}=dayObj;
  const dayName=DOW[dow], isClosed=dayName===b.dayOff;
  const allList=appts[activeB][day]||[];
  const activeList=allList.filter(a=>!a.reschd);
  const cap=getDayCap(activeB,dow);
  const total=activeList.length, remaining=cap-total, pct=cap>0?Math.round(total/cap*100):0;
  const fc=pct>=90?"#E24B4A":pct>=70?"#EF9F27":"#639922";
  const adv=activeAdv||b.advisors[0];
  const slots=advisorSlots(activeB,dow,adv);
  const advCapLimit=advisorCap(activeB,dow);
  const fileMissing=!demoMode&&spFileExists[activeB]===false;
  const mc=document.getElementById("main-content");

  const hdr=`<div class="day-hdr">
    <button class="dh-nav" onclick="jd(${day>1?day-1:1})" ${day<=1?"disabled":""}>&#8249;</button>
    <div>
      <div class="dh-title">${dayName}, ${day} ${MONTHS[activeMonth-1]} ${activeYear}</div>
      <div class="dh-sub">${b.label} Service Centre</div>
    </div>
    <button class="dh-nav" onclick="jd(${day<DAYS_IN_MONTH?day+1:DAYS_IN_MONTH})" ${day>=DAYS_IN_MONTH?"disabled":""}>&#8250;</button>
    <div class="dh-right">
      <button class="btn" style="font-size:11px;padding:4px 9px" onclick="openMonthPicker()">&#128197; ${MONTHS[activeMonth-1].slice(0,3)} ${activeYear} &#9660;</button>
      <button class="btn" style="font-size:11px;padding:4px 9px" onclick="syncDay()">&#8635; Sync</button>
    </div>
  </div>`;

  if(isClosed){mc.innerHTML=hdr+`<div class="closed-banner">&#128197; ${b.label} is closed on ${b.dayOff}s</div>`;buildCal();updateCounts();return;}

  const fBanner=fileMissing?`<div class="no-file-banner"><strong>SharePoint file not found</strong>&nbsp;Expected: <code>${spPath(activeB)}</code><br>Working in local mode.</div>`:"";

  // Build emirate dropdown options
  const emOptions=Object.entries(PLATE_EMIRATES).map(([k,v])=>`<option value="${k}">${v.name}</option>`).join("");
  const firstEm=Object.keys(PLATE_EMIRATES)[0];
  const codeOptions=PLATE_EMIRATES[firstEm].codes.map(c=>`<option value="${c}">${c===""?"(no code)":c}</option>`).join("");

  mc.innerHTML=hdr+fBanner+`
  <div class="stats">
    <div class="stat"><div class="stat-v">${total}</div><div class="stat-k">booked</div></div>
    <div class="stat ${remaining<=2?"sr":remaining<=5?"sw":""}"><div class="stat-v">${remaining}</div><div class="stat-k">remaining</div></div>
    <div class="stat"><div class="stat-v">${cap}</div><div class="stat-k">capacity</div></div>
    <div class="stat ${pct>=90?"sr":pct>=70?"sw":""}"><div class="stat-v">${pct}%</div><div class="stat-k">fill rate</div></div>
  </div>
  <div class="cap-wrap">
    <div class="cap-hdr"><span>${b.label} &mdash; ${dayName}</span><span>${total}/${cap} &middot; 08:00–14:30</span></div>
    <div class="cap-bar"><div class="cap-fill" style="width:${pct}%;background:${fc}"></div></div>
  </div>
  <div class="cgrid">
    <div>
      <div class="card">
        <div class="card-ttl">New appointment</div>
        <div class="sec-div">Advisor</div>
        <div class="adv-tabs">${b.advisors.map(a=>{
          const cnt=activeList.filter(x=>x.advisor===a).length;
          const full=advCapLimit>0&&cnt>=advCapLimit;
          return`<button class="adv-tab ${a===adv?"active":""} ${full?"at-full":""}" onclick="pickAdv('${a}')">${a} <span style="opacity:.65">${cnt}/${advCapLimit||"∞"}</span></button>`;
        }).join("")}</div>
        <div class="sec-div">Time slot <span style="font-weight:400;font-size:10px;color:var(--gray-400)">&nbsp;08:00–14:30</span></div>
        ${slots.length?`<div class="slot-grid">${slots.map(t=>{
          const taken=activeList.some(x=>x.advisor===adv&&x.time===t), sel=selSlot===t;
          return`<div class="slot ${taken?"sl-taken":sel?"sl-sel":"sl-free"}" onclick="${taken?"":"pickSlot('"+t+"')"}">${t}${taken?`<div style="font-size:9px;opacity:.6">taken</div>`:""}</div>`;
        }).join("")}</div>`:`<p style="font-size:11px;color:var(--gray-600);margin-bottom:8px">No slots available.</p>`}
        <div class="sec-div">Customer details</div>
        <div class="fgrid">
          <div class="fg"><label>Customer name *</label><input id="f_name" placeholder="Full name"></div>
          <div class="fg">
            <label>Contact number</label>
            <input id="f_phone" placeholder="971XXXXXXXXX" maxlength="15"
              oninput="this.value=this.value.replace(/[^0-9]/g,'')"
              onpaste="setTimeout(()=>{this.value=this.value.replace(/[^0-9]/g,'')},0)"
              title="Digits only — no spaces. Format: 971XXXXXXXXX or 05XXXXXXXX">
          </div>

          <div class="fg full">
            <label>Licence plate</label>
            <div class="plate-builder">
              <select id="f_plate_emirate" onchange="plateEmirateChanged()" title="Emirate">${emOptions}</select>
              <select id="f_plate_code" onchange="updatePlatePreview()" title="Plate code"></select>
              <input id="f_plate_num" placeholder="Number" maxlength="8" inputmode="numeric"
                oninput="this.value=this.value.replace(/[^0-9]/g,'');updatePlatePreview()"
                onpaste="setTimeout(()=>{this.value=this.value.replace(/[^0-9]/g,'');updatePlatePreview()},0)"
                title="Plate number digits only">
              <input type="hidden" id="f_plate">
            </div>
            <div class="plate-preview">Formatted: <strong id="f_plate_preview">—</strong></div>
          </div>

          <div class="fg"><label>Vehicle model</label><input id="f_model" placeholder="Grand Vitara"></div>
          <div class="fg full"><label>Service / repair type</label><input id="f_service" placeholder="e.g. 20K service, brake check…"></div>
          <div class="fg"><label>Booked by (agent) *</label><input id="f_agent" placeholder="Your name"></div>
          <div class="fg"><label>Channel</label>
            <select id="f_channel"><option>Inbound call - IBT</option><option>WhatsApp</option><option>Walk-in</option><option>Online</option><option>Outbound call</option><option>Other</option></select>
          </div>
          <div class="fg full"><label>Remarks</label><input id="f_remarks" placeholder="Notes…"></div>
        </div>
        <div class="btn-row">
          <button class="btn btn-primary" id="book-btn" onclick="bookAppt()">+ Book</button>
          <button class="btn" onclick="clearForm()">Clear</button>
        </div>
      </div>
    </div>
    <div>
      <div class="card">
        <div class="card-ttl">Today's bookings (${total})</div>
        ${allList.length===0?`<div class="empty-state" style="padding:1rem 0"><div style="font-size:22px;opacity:.3">&#128203;</div><p style="font-size:11px;margin-top:5px">No appointments yet</p></div>`:`
        <div class="appt-list">${allList.map((a,i)=>`
          <div class="appt-row ${a.reschd?"reschd":""}">
            <div class="at">${a.time}</div>
            <div class="ai">
              <div class="an">${a.name||"—"}<span class="badge b-adv">${a.advisor}</span>${a.reschd?`<span class="badge b-rsch">rescheduled</span>`:""}${a.waSent?`<span class="badge b-wa">&#10003; WA</span>`:""}</div>
              <div class="ad">${[a.plate,a.model].filter(Boolean).join(" &middot; ")}${a.service?" &mdash; "+a.service:""}</div>
              ${a.agent?`<div class="ad">By: ${a.agent}${a.remarks?" &middot; "+a.remarks:""}</div>`:""}
            </div>
            <div class="appt-actions">
              ${!a.reschd?`<button class="icon-btn ib-wa ${a.waSent?"wa-done":""}" title="Send WhatsApp" onclick="openWaModal(${i})">&#128172;</button>`:""}
              ${!a.reschd?`<button class="icon-btn ib-warn" title="Reschedule" onclick="openRschedule(${i})">&#8635;</button>`:""}
              ${!a.fromSP&&!a.reschd?`<button class="icon-btn ib-red" title="Remove" onclick="cancelAppt(${i})">&#10005;</button>`:""}
            </div>
          </div>`).join("")}
        </div>`}
      </div>
      <div class="card">
        <div class="card-ttl">Advisor load</div>
        ${b.advisors.map(a=>{
          const cnt=activeList.filter(x=>x.advisor===a).length;
          const max=advCapLimit||1;
          const p=Math.round(cnt/max*100), col=p>=90?"#E24B4A":p>=70?"#EF9F27":"#639922";
          return`<div class="adv-load"><div class="adv-load-hdr"><span>${a}</span><span style="color:var(--gray-600)">${cnt}/${advCapLimit||"∞"} (${p}%)</span></div><div class="adv-load-bar"><div class="adv-load-fill" style="width:${Math.min(p,100)}%;background:${col}"></div></div></div>`;
        }).join("")}
      </div>
    </div>
  </div>`;
  buildCal(); updateCounts();
  // Init plate builder dropdowns after render
  plateEmirateChanged();
}

// ══════════════════════════════════════════════════════════════
// BOOKING
// ══════════════════════════════════════════════════════════════
function pickAdv(a){activeAdv=a;selSlot=null;renderMain();}
function pickSlot(t){selSlot=t;renderMain();}

async function bookAppt() {
  const name=g("f_name"), phone=g("f_phone"), plate=g("f_plate"), model=g("f_model"),
        service=g("f_service"), agent=g("f_agent"), channel=g("f_channel"), remarks=g("f_remarks");
  if(!name){mk("f_name");toast("Customer name required","t-err");return;}
  if(!agent){mk("f_agent");toast("Agent name required","t-err");return;}
  if(!selSlot){toast("Select a time slot","t-err");return;}
  // Phone validation — no spaces allowed
  if(phone&&/\s/.test(phone)){toast("Phone number must not contain spaces","t-err");return;}
  const activeList=appts[activeB][activeDay].filter(a=>!a.reschd);
  if(activeList.some(a=>a.advisor===activeAdv&&a.time===selSlot)){toast("Slot already taken","t-err");return;}

  const appt={time:selSlot,advisor:activeAdv,name,phone,plate,model,service,agent,remarks:`${channel}${remarks?" — "+remarks:""}`};
  appts[activeB][activeDay].push(appt);
  appts[activeB][activeDay].sort((a,b)=>a.time.localeCompare(b.time));
  selSlot=null;
  logAction("book",`${agent} booked ${appt.time} for ${name} (${plate||"no plate"}) @ ${CFG.branches[activeB].label} with ${activeAdv}`);

  document.getElementById("book-btn").disabled=true;
  if(!demoMode&&spFileExists[activeB]!==false){
    setSP("pill-sync","Saving…");showLoader("Writing to SharePoint…");
    const res=await writeBookingRow(activeB,activeDay,appt);hideLoader();
    if(res.ok){appt.fromSP=true;setSP("pill-live","Saved ✓");toast(`✓ ${appt.time} booked for ${name}`,"t-ok");}
    else{setSP("pill-err","Save failed");toast("Saved locally — SP write failed","t-warn");}
  } else {toast(`✓ ${appt.time} booked for ${name} with ${activeAdv}`,"t-ok");}
  clearForm(); renderMain();
}

function cancelAppt(i){
  if(!confirm("Remove this appointment?")) return;
  const a=appts[activeB][activeDay][i];
  logAction("cancel",`Removed ${a.time} — ${a.name||"unknown"} @ ${CFG.branches[activeB].label}`);
  appts[activeB][activeDay].splice(i,1);
  toast("Removed","t-info"); renderMain();
}
async function syncDay(){
  if(!activeB) return;
  spLoaded[activeB][activeDay]=false; spFileExists[activeB]=null;
  showLoader("Syncing…"); await loadDay(activeB,activeDay,true); hideLoader(); renderMain();
  toast("Refreshed","t-info");
}
function clearForm(){
  ["f_name","f_phone","f_plate","f_model","f_service","f_agent","f_remarks"].forEach(id=>{
    const el=document.getElementById(id); if(el){el.value="";el.classList.remove("err");}
  });
  const pnum=document.getElementById("f_plate_num"); if(pnum) pnum.value="";
  const prev=document.getElementById("f_plate_preview"); if(prev) prev.textContent="—";
  selSlot=null; renderMain();
}

// ══════════════════════════════════════════════════════════════
// RESCHEDULE
// ══════════════════════════════════════════════════════════════
function openRschedule(idx){
  const a=appts[activeB][activeDay][idx]; if(!a||a.reschd) return;
  rschTarget={branch:activeB,day:activeDay,idx};
  document.getElementById("rsch-preview").innerHTML=`
    <div class="rp-name">${a.name||"Unknown"} — <span style="color:var(--blue-dark)">${a.time}</span> with ${a.advisor}</div>
    <div class="rp-detail">${[a.plate,a.model,a.service].filter(Boolean).join(" · ")||"No vehicle details"}</div>`;
  const dayEl=document.getElementById("rsch-day");
  dayEl.innerHTML=JUNE.filter(({dow})=>DOW[dow]!==CFG.branches[activeB].dayOff)
    .map(({day,dow})=>`<option value="${day}" ${day===activeDay?"selected":""}>${day} ${MONTHS[activeMonth-1].slice(0,3)} (${DOW[dow]})</option>`).join("");
  const advEl=document.getElementById("rsch-adv");
  advEl.innerHTML=CFG.branches[activeB].advisors.map(av=>`<option value="${av}" ${av===a.advisor?"selected":""}>${av}</option>`).join("");
  const nd=dayEl.cloneNode(true), na=advEl.cloneNode(true);
  dayEl.replaceWith(nd); advEl.replaceWith(na);
  nd.value=activeDay; na.value=a.advisor;
  nd.addEventListener("change",fillRschSlots); na.addEventListener("change",fillRschSlots);
  fillRschSlots();
  document.getElementById("rsch-modal").classList.add("show");
}
function fillRschSlots(){
  const day=parseInt(document.getElementById("rsch-day").value);
  const adv=document.getElementById("rsch-adv").value;
  const dayObj=JUNE[day-1]; if(!dayObj) return;
  const {dow}=dayObj;
  const taken=(appts[activeB][day]||[]).filter(a=>a.advisor===adv&&!a.reschd).map(a=>a.time);
  const el=document.getElementById("rsch-slot"); el.innerHTML="";
  let has=false;
  allSlots().forEach(t=>{
    if(getDayCap(activeB,dow)===0||taken.includes(t)) return;
    const o=document.createElement("option"); o.value=t; o.textContent=t; el.appendChild(o); has=true;
  });
  if(!has){const o=document.createElement("option");o.textContent="No slots available";o.disabled=true;el.appendChild(o);}
}
function closeRschModal(){document.getElementById("rsch-modal").classList.remove("show");rschTarget=null;}
async function confirmReschedule(){
  if(!rschTarget) return;
  const newDay  =parseInt(document.getElementById("rsch-day").value);
  const newAdv  =document.getElementById("rsch-adv").value;
  const newSlot =document.getElementById("rsch-slot").value;
  const reason  =document.getElementById("rsch-reason").value.trim();
  if(!newSlot||newSlot.includes("No slot")){toast("Select a valid slot","t-err");return;}
  const orig=appts[rschTarget.branch][rschTarget.day][rschTarget.idx];
  if(!orig){closeRschModal();return;}
  const agent=orig.agent||"System";
  const now=new Date().toLocaleString("en-GB");
  const mo=MONTHS[activeMonth-1].slice(0,3);
  const origDate=`${rschTarget.day} ${mo} ${activeYear}`;
  const newDate=`${newDay} ${mo} ${activeYear}`;
  const origSnapshot={...orig};
  appts[rschTarget.branch][rschTarget.day].splice(rschTarget.idx,1);
  const newAppt={...origSnapshot,time:newSlot,advisor:newAdv,day:newDay,reschd:false,
    remarks:`Rescheduled from ${rschTarget.day} ${mo} ${origSnapshot.time}${reason?" ("+reason+")":""}`,fromSP:false,waSent:false};
  if(!appts[rschTarget.branch][newDay]) appts[rschTarget.branch][newDay]=[];
  appts[rschTarget.branch][newDay].push(newAppt);
  appts[rschTarget.branch][newDay].sort((a,b)=>a.time.localeCompare(b.time));
  logAction("rsch",`${agent} rescheduled ${origSnapshot.name||"unknown"}: ${rschTarget.day} ${mo} ${origSnapshot.time} → ${newDay} ${mo} ${newSlot}${reason?" | "+reason:""}`);
  closeRschModal();
  if(!demoMode&&spFileExists[activeB]!==false){
    setSP("pill-sync","Saving…");showLoader("Writing reschedule…");
    if(origSnapshot.spRow){
      const rschdRemark=`RESCHEDULED → ${newDay} ${mo} ${newSlot}${reason?" ("+reason+")":""}`;
      await spRange(rschTarget.branch,rschTarget.day,`N${origSnapshot.spRow}:N${origSnapshot.spRow}`,[[rschdRemark]]);
    }
    if(!spLoaded[rschTarget.branch][newDay]) await loadDay(rschTarget.branch,newDay);
    const res=await writeBookingRow(rschTarget.branch,newDay,newAppt);
    if(res.ok) newAppt.fromSP=true;
    await writeRescheduleLog(rschTarget.branch,rschTarget.day,{
      oldDate:origDate,oldTime:origSnapshot.time,newDate,newTime:newSlot,
      customer:origSnapshot.name||"",plate:origSnapshot.plate||"",advisor:origSnapshot.advisor,
      by:agent,ts:now,reason:reason||""
    });
    hideLoader();setSP("pill-live","Saved ✓");
    toast(`✓ Rescheduled to ${newDay} ${mo} ${newSlot} — log updated`,"t-ok");
  } else {toast(`✓ Rescheduled to ${newDay} ${mo} ${newSlot}`,"t-ok");}
  renderMain();
}

// ══════════════════════════════════════════════════════════════
// WHATSAPP — TWIXOR
// ══════════════════════════════════════════════════════════════
function openTwxModal(){document.getElementById("wa-login-modal").classList.add("show");}
function closeTwxModal(){document.getElementById("wa-login-modal").classList.remove("show");}
async function doTwxLogin(){
  const user=document.getElementById("twx-user").value.trim();
  const pass=document.getElementById("twx-pass").value;
  if(!user||!pass){toast("Enter Twixor credentials","t-err");return;}
  showLoader("Signing in to Twixor…");
  try{
    const r=await fetch(`${CFG.twixorBase}/chatbird/api/auth/login`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({username:user,password:pass})});
    hideLoader();
    if(r.ok){
      const d=await r.json();
      twxToken=d.token||d.authToken||d.authentication_token||d.access_token||d.data?.token;
      if(!twxToken){toast("Login OK but no token in response","t-warn");return;}
    } else {
      twxToken=prompt("Twixor login API returned an error.\nPaste your authentication-token from browser devtools:");
      if(!twxToken) return;
    }
  }catch(e){
    hideLoader();
    twxToken=prompt("Could not reach Twixor.\nPaste your authentication-token manually:");
    if(!twxToken) return;
  }
  twxUser=user;
  document.getElementById("wa-chip").className="wa-chip";
  document.getElementById("wa-lbl").textContent=`WA: ${user.split("@")[0]}`;
  logAction("login",`Twixor WA connected as ${user}`);
  closeTwxModal(); toast(`✓ WhatsApp connected`,"t-ok");
}
function openWaModal(idx){
  if(!twxToken){toast("Connect to Twixor first — click the WA chip","t-warn");openTwxModal();return;}
  const a=appts[activeB][activeDay][idx]; if(!a) return;
  waTarget={branch:activeB,day:activeDay,idx};
  const date=new Date(activeYear,activeMonth-1,activeDay).toLocaleDateString("en-GB");
  document.getElementById("wa-preview").innerHTML=`
    <strong>${a.name||"Customer"}</strong><br>
    <span style="color:var(--gray-600)">Plate: ${a.plate||"—"} &nbsp;·&nbsp; ${date} at ${a.time}<br>
    ${CFG.branches[activeB].label} Service Centre</span>`;
  let ph=(a.phone||"").replace(/\s+/g,"").replace(/^00/,"").replace(/^\+/,"");
  if(ph.startsWith("0")) ph="971"+ph.substring(1);
  if(!ph.startsWith("971")&&ph.length>=9) ph="971"+ph;
  document.getElementById("wa-phone").value=ph||"971";
  document.getElementById("wa-confirm-modal").classList.add("show");
}
function closeWaModal(){document.getElementById("wa-confirm-modal").classList.remove("show");waTarget=null;}
async function sendWa(){
  if(!waTarget||!twxToken) return;
  const phone=document.getElementById("wa-phone").value.trim().replace(/\s+/g,"");
  if(!phone||phone.length<11){toast("Enter valid UAE number (971XXXXXXXXX)","t-err");return;}
  const a=appts[waTarget.branch][waTarget.day][waTarget.idx]; if(!a) return;
  const date=new Date(activeYear,activeMonth-1,waTarget.day).toLocaleDateString("en-GB");
  const br=CFG.branches[waTarget.branch];
  const payload={
    channelId:CFG.twixorChannel,phoneNumber:[phone],
    message:{template:{templateId:CFG.twixorTplId,templateUId:CFG.twixorTplUId}},
    columnMapping:{
      body:[{variable:"1",value:a.plate||"—"},{variable:"2",value:date},
            {variable:"3",value:a.time},{variable:"4",value:br.label+" Service Centre"},{variable:"5",value:br.maps}],
      messageTags:[{variable:"messageTag2",value:null},{variable:"messageTag3",value:null},{variable:"messageTag4",value:null},
                   {variable:"messageTag5",value:null},{variable:"messageTag6",value:null},{variable:"messageTag7",value:null},
                   {variable:"messageTag8",value:null},{variable:"messageTag9",value:null},{variable:"messageTag10",value:null}]
    }
  };
  showLoader("Sending WhatsApp…");
  try{
    const r=await fetch(`${CFG.twixorBase}/chatbird/api/broadcast`,{method:"POST",
      headers:{"Content-Type":"application/json; charset=UTF-8","authentication-token":twxToken},body:JSON.stringify(payload)});
    hideLoader();
    if(r.ok){
      a.waSent=true;a.waPhone=phone;
      logAction("wa",`WA sent for ${a.name||"unknown"} → ${phone} (${waTarget.branch} ${date} ${a.time})`);
      closeWaModal();toast(`✓ WhatsApp sent to ${phone}`,"t-ok");renderMain();
    } else {
      const err=await r.json().catch(()=>({}));toast(`Send failed: ${err.message||r.status}`,"t-err");
    }
  }catch(e){hideLoader();toast("Network error sending WhatsApp","t-err");}
}

// ══════════════════════════════════════════════════════════════
// AGENT LOG
// ══════════════════════════════════════════════════════════════
function logAction(type,msg){
  const ts=new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit",second:"2-digit"});
  agentLog.unshift({type,msg,ts,agent:twxUser||currentAccount?.name||"Agent"});
  renderLog();
}
function renderLog(){
  const el=document.getElementById("log-body"); if(!el) return;
  const cls={book:"l-book",rsch:"l-rsch",wa:"l-wa",cancel:"l-cancel",login:"l-login"};
  el.innerHTML=agentLog.length
    ?agentLog.map(e=>`<div class="log-entry ${cls[e.type]||"l-login"}">${e.msg}<div class="log-ts">${e.ts} · ${e.agent}</div></div>`).join("")
    :`<p style="font-size:12px;color:var(--gray-400);padding:8px">No actions yet.</p>`;
}
function toggleLog(){document.getElementById("log-panel").classList.toggle("open");}
function exportLog(){
  const lines=agentLog.map(e=>`[${e.ts}] [${e.type.toUpperCase()}] ${e.msg} (${e.agent})`).join("\n");
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([lines],{type:"text/plain"}));
  a.download=`agent_log_${new Date().toISOString().split("T")[0]}.txt`;a.click();
}

// ══════════════════════════════════════════════════════════════
// UTILS
// ══════════════════════════════════════════════════════════════
function g(id){return document.getElementById(id)?.value.trim()||"";}
function mk(id){document.getElementById(id)?.classList.add("err");}
function show(id){document.getElementById(id).classList.add("show");}
function hide(id){document.getElementById(id).classList.remove("show");}
function showLoader(t){document.getElementById("loader-txt").textContent=t||"Loading…";document.getElementById("loader").classList.add("show");}
function hideLoader(){document.getElementById("loader").classList.remove("show");}
let _tt;
function toast(msg,cls="t-ok"){const el=document.getElementById("toast");el.textContent=msg;el.className="toast show "+cls;clearTimeout(_tt);_tt=setTimeout(()=>el.className="toast",4000);}

// ── Boot ──
window.addEventListener("load",()=>{
  if(_loadStoredToken()){currentAccount={name:_jwtName(_accessToken)};afterAuth();}
});