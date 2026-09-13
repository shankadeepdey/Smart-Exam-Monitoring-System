import { FaceLandmarker, FilesetResolver } from "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.14";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
         HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType, PageBreak }
         from "https://cdn.jsdelivr.net/npm/docx@8.5.0/+esm";



const SUSPICIOUS_TURNS = 5, MAX_TURNS = 15;
const LEVEL_WARNING = 30, LEVEL_ALERT = 60;
const YAW_THRESHOLD = 22;
const GAZE_THRESHOLD = 0.17, GAZE_TIME_LIMIT = 2, GAZE_COOLDOWN = 5, GAZE_MAX = 7;
const EAR_THRESHOLD = 0.21, EAR_TIME_LIMIT = 1.5, DROWSY_COOLDOWN = 5, DROWSY_MAX = 10;
const COVER_TIME_LIMIT = 1.0, COVER_TIMEOUT = 15.0, COVER_COOLDOWN = 5, COVER_MAX = 10;
const PHONE_TIME_LIMIT = 1.0, PHONE_COOLDOWN = 6, PHONE_MAX = 15, PHONE_CONF = 0.5;
const HEAD_MAX = 7;
const YOLO_EVERY = 3; // run object detector every N frames

// Eye-closed / drowsiness detection relies on iris landmark precision that mobile
// front cameras (lower resolution, harder autofocus, more motion) can't reliably
// deliver — it was firing false "eyes closed" alerts on phones. Disable the whole
// drowsiness feature on mobile/touch devices while keeping every other check.
// Eye-closed / drowsiness detection relies on iris landmark precision that mobile
// front cameras (lower resolution, harder autofocus, more motion) can't reliably
// deliver — it was firing false "eyes closed" alerts on phones. Disable the whole
// drowsiness feature on mobile/touch devices while keeping every other check.
// NOTE: detected by device type (user agent / touch), not window width, so
// resizing a laptop browser window narrower does not count as "mobile".
const isMobile = /Android|iPhone|iPad|iPod|Mobi/i.test(navigator.userAgent) ||
  (navigator.maxTouchPoints > 1 && window.matchMedia('(pointer: coarse)').matches);
document.body.classList.toggle('no-drowsy', isMobile);

const L_IRIS=473, L_INNER=362, L_OUTER=263, R_IRIS=468, R_INNER=33, R_OUTER=133;
const R_EAR_PTS=[33,160,158,133,153,144], L_EAR_PTS=[362,385,387,263,373,380];


const SOUND_FILES = {
  head:     'too_much_moving_your_head.mp3',
  lookaway: 'stop_seeing_here_and_there.mp3',
  phone:    'put_your_phone_away.mp3',
  covered:  'don_t_cover_your_face.mp3'
  // drowsy: 'your_file_here.mp3'
};


const S = {
  name:'', turnCount:0, turnLeft:0, turnRight:0, score:0, bonusScore:0, wasLookingAway:false, direction:'center',
  gazeDir:'center', gazeRatio:0.5, gazeOutSince:null, gazeAlerted:false,
  ear:0.30, eyesClosedSince:null, drowsyActive:false, drowsyAlerted:false,
  lastSeenTime:null, coverSince:null, coverActive:false, coverAlerted:false,
  phoneActive:false, phoneSeenSince:null, phoneAlerted:false, phoneInstant:false,
  sessionStart:null, sessionEnd:null, running:false, muted:false,
  headEvidence:[], gazeEvidence:[], drowsyEvidence:[], coverEvidence:[], phoneEvidence:[],
  lastSoundTime:{}, lastPhotoTime:{}
};

const $ = id => document.getElementById(id);
const loadScreen=$('loadScreen'), loadBar=$('loadBar'), loadText=$('loadText'), loadPct=$('loadPct');
const nameModal=$('nameModal'), nameInput=$('studentName'), nameHint=$('nameHint'), nameSubmit=$('nameSubmit');
const guideModal=$('guideModal'), guideName=$('guideName'), guideContinue=$('guideContinue'), demoVideo=$('demoVideo'), replayVideoBtn=$('replayVideoBtn');
const app=$('app'), video=$('video'), overlay=$('overlay'), octx=overlay.getContext('2d');
const stage=$('stage'), stageIdle=$('stageIdle'), flashRing=$('flashRing'), alertStack=$('alertStack');
const startBtn=$('startBtn'), endBtn=$('endBtn'), muteBtn=$('muteBtn'), muteIcon=$('muteIcon');
const statusPill=$('statusPill'), statusText=$('statusText'), whoLabel=$('whoLabel');
const scoreNum=$('scoreNum'), dialFill=$('dialFill'), riskLabel=$('riskLabel'), turnsLabel=$('turnsLabel');
const sessionTimer=$('sessionTimer'), sessionState=$('sessionState');
const mFace=$('mFace'), mHead=$('mHead'), mGaze=$('mGaze'), mEar=$('mEar'), mPhone=$('mPhone'), mEarChip=$('mEarChip');
const logList=$('logList'), logCount=$('logCount');
const completionModal=$('completionModal'), doneName=$('doneName'), summaryGrid=$('summaryGrid');
const dlReport=$('dlReport'), dlZip=$('dlZip'), reportSub=$('reportSub'), zipSub=$('zipSub'), reportGo=$('reportGo'), zipGo=$('zipGo');
const closeCompletion=$('closeCompletion'), toast=$('toast'), fsBtn=$('fsBtn'), serialBtn=$('serialBtn');
const topToast=$('topToast'), topToastMsg=$('topToastMsg');
const fpsHint=$('fpsHint');
const clockNow=$('clockNow'), sessionChip=$('sessionChip'), viewTitle=$('viewTitle');
const sideStatusDot=$('sideStatusDot'), sideStatusValue=$('sideStatusValue');
const statTotalAlerts=$('statTotalAlerts'), statHeadTurns=$('statHeadTurns'), statPhoneAlerts=$('statPhoneAlerts'), statOtherAlerts=$('statOtherAlerts');
const reportsList=$('reportsList'), reportsEmpty=$('reportsEmpty'), clearReportsBtn=$('clearReportsBtn');
$('thYaw').textContent = YAW_THRESHOLD + '\u00b0';
$('thGaze').textContent = GAZE_THRESHOLD;

const DIAL_CIRC = 276.5; // matches r=44 ring in the current index.html dial svg

function showToast(msg){ toast.textContent=msg; toast.classList.add('show'); clearTimeout(showToast._t); showToast._t=setTimeout(()=>toast.classList.remove('show'),2600); }
function showTopToast(msg, durationMs=10000){
  if(!topToast || !topToastMsg) return;
  topToastMsg.textContent = msg;
  topToast.classList.add('show');
  clearTimeout(showTopToast._t);
  showTopToast._t = setTimeout(()=>topToast.classList.remove('show'), durationMs);
}
function fmtClock(sec){ const h=String(Math.floor(sec/3600)).padStart(2,'0'), m=String(Math.floor(sec/60)%60).padStart(2,'0'), s=String(Math.floor(sec)%60).padStart(2,'0'); return `${h}:${m}:${s}`; }
function nowStamp(){ const d=new Date(); return d.toLocaleTimeString('en-GB',{hour12:false}); }
function scoreOf(turns){ if(turns<=SUSPICIOUS_TURNS) return 0; return Math.min(Math.round(((turns-SUSPICIOUS_TURNS)/(MAX_TURNS-SUSPICIOUS_TURNS))*100),100); }
function levelOf(score){ if(score>=LEVEL_ALERT) return 2; if(score>=LEVEL_WARNING) return 1; return 0; }

function computeScore(){ return Math.min(100, scoreOf(S.turnCount) + S.bonusScore); }
function addScoreBonus(amount){ S.bonusScore = Math.min(100, S.bonusScore + amount); S.score = computeScore(); }

//NAVIGATION (Dashboard / Reports)
const navBtns = document.querySelectorAll('.nav-btn');
const viewEls = { dashboard: $('viewDashboard'), reports: $('viewReports') };
function showView(name){
  Object.entries(viewEls).forEach(([k,el])=>{ if(el) el.hidden = k!==name; });
  navBtns.forEach(b=> b.classList.toggle('active', b.dataset.view===name));
  if(viewTitle) viewTitle.textContent = name==='reports' ? 'Reports' : 'Dashboard';
  if(name==='reports') renderReports();
}
navBtns.forEach(b=> b.addEventListener('click', ()=> showView(b.dataset.view)));

//REPORTS HISTORY 
const REPORTS_KEY = 'sentinelReports';
function loadReports(){ try{ return JSON.parse(sessionStorage.getItem(REPORTS_KEY) || '[]'); }catch(e){ return []; } }
function saveReportRecord(rec){
  const list = loadReports();
  list.unshift(rec);
  try{ sessionStorage.setItem(REPORTS_KEY, JSON.stringify(list.slice(0,50))); }catch(e){ console.warn('Could not persist report', e); }
}
function renderReports(){
  if(!reportsList) return;
  const list = loadReports();
  if(!list.length){
    reportsList.innerHTML = '';
    if(reportsEmpty) reportsEmpty.style.display = 'block';
    return;
  }
  if(reportsEmpty) reportsEmpty.style.display = 'none';
  reportsList.innerHTML = list.map(r => `
    <div class="report-row">
      <div class="report-main">
        <div class="report-name">${r.name}</div>
        <div class="report-date">${r.date}</div>
      </div>
      <div class="report-stat"><span>${r.score}%</span><small>${r.risk}</small></div>
      <div class="report-stat"><span>${fmtClock(r.durationSec)}</span><small>duration</small></div>
      <div class="report-stat"><span>${r.headTurns}</span><small>head turns</small></div>
      <div class="report-stat"><span>${r.phoneDetections}</span><small>phone</small></div>
      <div class="report-stat"><span>${r.evidenceCount}</span><small>evidence</small></div>
    </div>
  `).join('');
}
if(clearReportsBtn) clearReportsBtn.addEventListener('click', ()=>{
  sessionStorage.removeItem(REPORTS_KEY);
  renderReports();
  showToast('Report history cleared');
});


//LIVE CLOCK 
setInterval(()=>{ if(clockNow) clockNow.textContent = nowStamp(); }, 1000);
if(clockNow) clockNow.textContent = nowStamp();

//eyechart
let eyeChart=null, alertFreqChart=null, headDirChart=null;
const earSeries=[];
const sparklineOpts = {
  responsive:true, maintainAspectRatio:false,
  animation:false,
  plugins:{legend:{display:false}, tooltip:{enabled:false}},
  scales:{ x:{display:false}, y:{display:false} },
  elements:{ point:{radius:0} }
};
function initLiveCharts(){
  if(typeof Chart==='undefined') return;
  if(isMobile) return; // eye-ratio sparkline is hidden on mobile, skip building it
  const eyeCanvas = $('eyeRatioChart');
  if(eyeCanvas && !eyeChart){
    eyeChart = new Chart(eyeCanvas, { type:'line', data:{ labels:[], datasets:[{ data:[], borderColor:'#2FE0C4', borderWidth:2, tension:.35, fill:true, backgroundColor:'rgba(47,224,196,.12)' }] }, options: sparklineOpts });
  }
}
function updateLiveCharts(){
  if(!eyeChart) return;
  earSeries.push(S.ear); if(earSeries.length>40) earSeries.shift();
  eyeChart.data.labels = earSeries.map((_,i)=>i);
  eyeChart.data.datasets[0].data = earSeries;
  eyeChart.update('none');
}
function renderCompletionCharts(){
  if(typeof Chart==='undefined') return;
  const freqCanvas = $('alertFreqChart'), dirCanvas = $('headTurnDirChart');
  const labels = ['Head','Gaze','Drowsy','Cover','Phone'];
  const data = [S.headEvidence.length, S.gazeEvidence.length, S.drowsyEvidence.length, S.coverEvidence.length, S.phoneEvidence.length];
  if(freqCanvas){
    if(alertFreqChart) alertFreqChart.destroy();
    alertFreqChart = new Chart(freqCanvas, {
      type:'bar',
      data:{ labels, datasets:[{ data, backgroundColor:['#5B9BD5','#F5A623','#A78BFA','#A78BFA','#EF4460'], borderRadius:4 }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
        scales:{ y:{ beginAtZero:true, ticks:{color:'#5A6873', precision:0}, grid:{color:'rgba(255,255,255,.05)'} },
                 x:{ ticks:{color:'#93A4AE', font:{size:10}}, grid:{display:false} } } }
    });
  }
  if(dirCanvas){
    if(headDirChart) headDirChart.destroy();
    const dirData = [S.turnLeft, S.turnRight];
    headDirChart = new Chart(dirCanvas, {
      type:'doughnut',
      data:{ labels:['Left','Right'], datasets:[{ data: dirData.some(v=>v>0) ? dirData : [1,1], backgroundColor:['#EF4460','#2FE0C4'], borderWidth:0 }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ position:'bottom', labels:{ color:'#93A4AE', boxWidth:10, font:{size:11} } } } }
    });
  }
}

//audio alerts
let actx=null;
const soundBuffers = {};
function ensureAudio(){
  if(!actx){ try{ actx=new (window.AudioContext||window.webkitAudioContext)(); }catch(e){ console.warn('Web Audio unavailable', e); } }
  if(actx && actx.state==='suspended'){ actx.resume().catch(()=>{}); }
  return actx;
}
async function preloadSounds(){
  ensureAudio();
  if(!actx) return;
  await Promise.all(Object.entries(SOUND_FILES).map(async ([kind, path])=>{
    try{
      const res = await fetch(path);
      if(!res.ok) throw new Error('HTTP '+res.status);
      const arr = await res.arrayBuffer();
      soundBuffers[kind] = await actx.decodeAudioData(arr);
    }catch(e){ console.warn(`[sound] Could not load "${path}" for "${kind}" — falling back to a tone.`, e); }
  }));
}
function playBuffer(kind){
  if(!actx || !soundBuffers[kind]) return false;
  try{
    const src = actx.createBufferSource();
    src.buffer = soundBuffers[kind];
    const gain = actx.createGain(); gain.gain.value = 0.9;
    src.connect(gain); gain.connect(actx.destination);
    src.start(0);
    return true;
  }catch(e){ console.warn('[sound] playback failed', e); return false; }
}
function beepFallback(kind){
  if(!actx) return;
  const now=actx.currentTime;
  const patterns = { drowsy:[[220,0.35]] };
  const pat = patterns[kind] || [[440,0.12]];
  let t=now;
  pat.forEach(([freq,dur])=>{
    const osc=actx.createOscillator(), gain=actx.createGain();
    osc.frequency.value=freq; osc.type='sine';
    gain.gain.setValueAtTime(0.0001,t);
    gain.gain.exponentialRampToValueAtTime(0.22,t+0.02);
    gain.gain.exponentialRampToValueAtTime(0.0001,t+dur);
    osc.connect(gain); gain.connect(actx.destination);
    osc.start(t); osc.stop(t+dur+0.02);
    t+=dur+0.06;
  });
}
function chime(){
  if(!actx) return;
  const now=actx.currentTime;
  [[880,0],[1180,0.09]].forEach(([freq,delay])=>{
    const osc=actx.createOscillator(), gain=actx.createGain();
    osc.frequency.value=freq; osc.type='sine';
    const t=now+delay;
    gain.gain.setValueAtTime(0.0001,t);
    gain.gain.exponentialRampToValueAtTime(0.18,t+0.015);
    gain.gain.exponentialRampToValueAtTime(0.0001,t+0.09);
    osc.connect(gain); gain.connect(actx.destination);
    osc.start(t); osc.stop(t+0.11);
  });
}
function playSound(kind, cooldown=4.0){
  const t=performance.now()/1000;
  if((t-(S.lastSoundTime[kind]||0))<cooldown) return;
  S.lastSoundTime[kind]=t;
  if(S.muted) return;
  ensureAudio();
  chime();
  setTimeout(()=>{ if(S.muted) return; if(!playBuffer(kind)) beepFallback(kind); }, 260);
}

const screenFlash=$('screenFlash'), screenFlashMsg=$('screenFlashMsg');
let flashHideTimer=null;
function triggerFullFlash(kind, message, durationMs=2000){
  const bg = {
    phone:'rgba(200,0,20,.55)', drowsy:'rgba(20,60,180,.5)',
    covered:'rgba(120,10,140,.5)', crit:'rgba(200,0,20,.55)'
  }[kind] || 'rgba(200,0,20,.5)';
  screenFlash.style.setProperty('--flash-bg', bg);
  screenFlashMsg.textContent = message;
  screenFlash.classList.add('show');
  clearTimeout(flashHideTimer);
  flashHideTimer = setTimeout(()=>screenFlash.classList.remove('show'), durationMs);
}

//event log & alert banners
const LEVEL_META = {
  head:   {label:'HEAD TURN',     cls:'level-warn',  color:'var(--amber)'},
  gaze:   {label:'EYE ALERT',     cls:'level-warn',  color:'var(--amber)'},
  drowsy: {label:'DROWSINESS',    cls:'level-drowsy',color:'var(--blue)'},
  cover:  {label:'FACE COVERED',  cls:'level-cover', color:'var(--violet)'},
  phone:  {label:'PHONE DETECTED',cls:'level-phone', color:'var(--red)'},
  crit:   {label:'HIGH RISK',     cls:'level-crit',  color:'var(--red)'}
};
let logN=0;
function addLog(kind, msg){
  logN++; logCount.textContent = logN+' event'+(logN===1?'':'s');
  const empty = logList.querySelector('.log-empty'); if(empty) empty.remove();
  const item=document.createElement('div'); item.className='log-item';
  item.innerHTML = `<div class="t">${nowStamp()}</div><div class="m"><b>${LEVEL_META[kind]?.label||kind}</b> &mdash; ${msg}</div>`;
  logList.appendChild(item);
  while(logList.children.length>60) logList.removeChild(logList.firstChild);
}
const bannerEls = {};
function showBanner(kind, msg, sticky){
  const meta=LEVEL_META[kind]; if(!meta) return;
  let el=bannerEls[kind];
  if(!el){ el=document.createElement('div'); el.className='alert-banner '+meta.cls; alertStack.appendChild(el); bannerEls[kind]=el; }
  el.innerHTML = `<b>${meta.label}</b> &nbsp;${msg}`;
  el.classList.add('show');
  flashRing.style.setProperty('--flash-color', meta.color);
  flashRing.classList.add('on');
  clearTimeout(el._t);
  if(!sticky){ el._t=setTimeout(()=>{ el.classList.remove('show'); maybeClearFlash(); },1600); }
}
function hideBanner(kind){ const el=bannerEls[kind]; if(el){ el.classList.remove('show'); } maybeClearFlash(); }
function maybeClearFlash(){ const anyOn=Object.values(bannerEls).some(el=>el.classList.contains('show')); if(!anyOn) flashRing.classList.remove('on'); }

//evidence capture
const snapCanvas=document.createElement('canvas'), sctx=snapCanvas.getContext('2d');
function captureEvidence(stampTitle, stampColor){
  const w=video.videoWidth||1280, h=video.videoHeight||720;
  snapCanvas.width=w; snapCanvas.height=h;
  sctx.save(); sctx.translate(w,0); sctx.scale(-1,1); sctx.drawImage(video,0,0,w,h); sctx.restore();
  if(stampTitle){
    sctx.fillStyle=stampColor||'rgba(0,0,0,.7)';
    sctx.fillRect(0,0,w,64);
    sctx.fillStyle='#fff'; sctx.font='600 26px Arial';
    sctx.fillText(stampTitle, 16, 32);
    sctx.font='14px Arial'; sctx.fillStyle='rgba(255,255,255,.85)';
    sctx.fillText(new Date().toLocaleString(), 16, 54);
  }
  return { ts: nowStamp(), dataUrl: snapCanvas.toDataURL('image/jpeg',0.85), w, h };
}
function pushEvidence(arr, entry, cap){ if(arr.length<cap) arr.push(entry); }


let faceLandmarker=null, cocoModel=null;


let progRAF = null, progCurrent = 0;
function setLoadBarUI(pct){
  loadBar.style.width = pct+'%';
  if(loadPct) loadPct.textContent = Math.round(pct)+'%';
}
function animateLoadProgress(target, durationMs=900){
  cancelAnimationFrame(progRAF);
  const start = progCurrent;
  target = Math.max(target, start); 
  const t0 = performance.now();
  const step = (now)=>{
    const t = Math.min(1, (now-t0)/durationMs);
    const eased = 1 - Math.pow(1-t, 3); // ease-out
    progCurrent = start + (target-start)*eased;
    setLoadBarUI(progCurrent);
    if(t<1){ progRAF = requestAnimationFrame(step); }
    else{ progCurrent = target; setLoadBarUI(progCurrent); }
  };
  progRAF = requestAnimationFrame(step);
}
async function loadModels(){
  const slowTimer = setTimeout(()=>{
    loadText.textContent = 'still downloading models\u2026 first run pulls ~15\u201320\u2009MB, hang tight on slower connections';
  }, 6000);
  try{
    loadText.textContent='connecting to neural network\u2026';
    animateLoadProgress(5, 300);
    animateLoadProgress(32, 3000);
    const fileset = await FilesetResolver.forVisionTasks("https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.14/wasm");
    animateLoadProgress(38, 400);
    loadText.textContent='loading face landmark model\u2026';
    animateLoadProgress(68, 5500);
    try{
      faceLandmarker = await FaceLandmarker.createFromOptions(fileset, {
        baseOptions:{ modelAssetPath:"https://storage.googleapis.com/mediapipe-models/face_landmarker/face_landmarker/float16/1/face_landmarker.task", delegate:"GPU" },
        outputFacialTransformationMatrixes:true, outputFaceBlendshapes:false, runningMode:"VIDEO", numFaces:1
      });
    }catch(e){
      faceLandmarker = await FaceLandmarker.createFromOptions(fileset, {
        baseOptions:{ modelAssetPath:"https://storage.googleapis.com/mediapipe-models/face_landmarker/face_landmarker/float16/1/face_landmarker.task", delegate:"CPU" },
        outputFacialTransformationMatrixes:true, outputFaceBlendshapes:false, runningMode:"VIDEO", numFaces:1
      });
    }
    animateLoadProgress(70, 400);
    loadText.textContent='AI phone-detection model is loading\u2026';
    animateLoadProgress(96, 4500);
    cocoModel = await cocoSsd.load({base:'lite_mobilenet_v2'});
    animateLoadProgress(100, 400);
    loadText.textContent='validating environmental baseline\u2026';
    clearTimeout(slowTimer);
    initLiveCharts();
    await new Promise(r=>setTimeout(r,350));
    loadScreen.classList.add('hidden');
    nameModal.classList.add('show');
    if(isMobile){
      showTopToast('For the best experience, open Sentinel on a laptop — mobile support is in beta.', 10000);
    }
  }catch(err){
    cancelAnimationFrame(progRAF);
    clearTimeout(slowTimer);
    throw err;
  }
}


nameInput.addEventListener('input', ()=>{
  const v=nameInput.value.trim();
  nameSubmit.disabled = v.length<2;
  nameHint.textContent = v.length>0 && v.length<2 ? 'Enter at least 2 characters' : '\u00a0';
  nameHint.classList.toggle('err', v.length>0 && v.length<2);
});
nameInput.addEventListener('keydown', e=>{ if(e.key==='Enter' && !nameSubmit.disabled) nameSubmit.click(); });

nameSubmit.addEventListener('click', async ()=>{
  S.name = nameInput.value.trim();
  nameSubmit.disabled=true; nameSubmit.textContent='Requesting camera\u2026';
  try{
    const stream = await navigator.mediaDevices.getUserMedia({ video:{ facingMode:'user', width:{ideal:1280}, height:{ideal:720} }, audio:false });
    video.srcObject = stream;
    await new Promise(res=>{ video.onloadedmetadata=()=>{ video.play(); res(); }; });
    overlay.width = video.videoWidth; overlay.height = video.videoHeight;
    nameModal.classList.remove('show');
    guideName.textContent = `A few things to avoid, ${S.name.split(' ')[0]}`;
    guideModal.classList.add('show');
    replayVideoBtn.style.display='none';
    demoVideo.muted = true;
    demoVideo.currentTime = 0;
    demoVideo.play().then(()=>{
      demoVideo.muted = false;
    }).catch(()=>{ });
  }catch(err){
    nameSubmit.disabled=false; nameSubmit.textContent='Continue & enable camera';
    nameHint.textContent = 'Camera access was blocked. Please allow camera permission and try again.';
    nameHint.classList.add('err');
  }
});

demoVideo.addEventListener('ended', ()=>{ replayVideoBtn.style.display='flex'; });
replayVideoBtn.addEventListener('click', ()=>{
  replayVideoBtn.style.display='none';
  demoVideo.currentTime = 0;
  demoVideo.muted = false;
  demoVideo.play().catch(()=>{});
});

guideContinue.addEventListener('click', ()=>{
  demoVideo.pause();
  guideModal.classList.remove('show');
  app.classList.add('active');
  whoLabel.textContent = S.name;
  stageIdle.style.display='none';
  video.style.display='block';
  if('serial' in navigator) serialBtn.style.display='flex';
});

function headYaw(matrixData){
  const yaw = Math.atan2(matrixData[2], matrixData[10]) * 180/Math.PI;
  if(yaw > YAW_THRESHOLD) return 'right';
  if(yaw < -YAW_THRESHOLD) return 'left';
  return 'center';
}
function gazeDirection(lm){
  const spanL = lm[L_OUTER].x - lm[L_INNER].x;
  const rl = Math.abs(spanL)<1e-5 ? 0.5 : (lm[L_IRIS].x - lm[L_INNER].x)/spanL;
  const spanR = lm[R_OUTER].x - lm[R_INNER].x;
  const rr = Math.abs(spanR)<1e-5 ? 0.5 : (lm[R_IRIS].x - lm[R_INNER].x)/spanR;
  const avg = Math.min(1,Math.max(0,(rl+rr)/2));
  let d='center';
  if(avg < 0.5-GAZE_THRESHOLD) d='left'; else if(avg > 0.5+GAZE_THRESHOLD) d='right';
  return [d, avg];
}
function eyeAspectRatio(lm){
  const dist=(i,j)=>Math.hypot(lm[i].x-lm[j].x, lm[i].y-lm[j].y);
  const earOf = pts=>{ const [p1,p2,p3,p4,p5,p6]=pts; const v=dist(p2,p6)+dist(p3,p5); const h=dist(p1,p4); return h<1e-6?0.30:v/(2*h); };
  return (earOf(L_EAR_PTS)+earOf(R_EAR_PTS))/2;
}

let rafId=null, frameNo=0, lastPhoneBoxes=[], lastFrameTime=performance.now(), fpsSmooth=0;

function drawFaceMesh(lm){
  octx.strokeStyle='rgba(47,224,196,0.55)'; octx.lineWidth=1;
  const w=overlay.width,h=overlay.height;
  octx.fillStyle='rgba(47,224,196,0.55)';
  for(let i=0;i<lm.length;i+=3){
    octx.beginPath(); octx.arc(lm[i].x*w, lm[i].y*h, 1, 0, 7); octx.fill();
  }
}
function drawEyeRings(lm){
  const w=overlay.width,h=overlay.height;
  octx.strokeStyle='#B07CFF'; octx.lineWidth=1.5;
  [[L_IRIS,L_INNER,L_OUTER],[R_IRIS,R_INNER,R_OUTER]].forEach(([iris,inner,outer])=>{
    const cx=lm[iris].x*w, cy=lm[iris].y*h;
    const r=Math.max(4, Math.abs(lm[outer].x-lm[inner].x)*w*0.42);
    octx.beginPath(); octx.arc(cx,cy,r,0,7); octx.stroke();
  });
}
function drawPhoneBoxes(boxes){
  octx.lineWidth=3; octx.strokeStyle='#FF4462'; octx.font='700 15px "IBM Plex Sans"';
  boxes.forEach(b=>{
    octx.strokeRect(b.x,b.y,b.w,b.h);
    octx.fillStyle='#FF4462'; octx.fillRect(b.x,Math.max(0,b.y-26),150,26);
    octx.fillStyle='#fff'; octx.fillText('PHONE', b.x+8, Math.max(18,b.y-7));
  });
}

async function loop(){
  if(!S.running) return;
  rafId = requestAnimationFrame(loop);
  if(video.readyState < 2) return;
  frameNo++;
  const now = performance.now();
  const dt = now-lastFrameTime; lastFrameTime=now;
  fpsSmooth = fpsSmooth ? fpsSmooth*0.9+ (1000/dt)*0.1 : (1000/dt);
  if(frameNo % 15 === 0) fpsHint.textContent = Math.round(fpsSmooth)+' fps';

  octx.clearRect(0,0,overlay.width,overlay.height);
  document.querySelectorAll('.eye-ratio-tag').forEach(e=>e.remove());

  const tSec = performance.now()/1000;
  const result = faceLandmarker.detectForVideo(video, now);

  if(frameNo % YOLO_EVERY === 0 && cocoModel){
    cocoModel.detect(video).then(preds=>{
      lastPhoneBoxes = preds.filter(p=>p.class==='cell phone' && p.score>=PHONE_CONF)
        .map(p=>({x:p.bbox[0],y:p.bbox[1],w:p.bbox[2],h:p.bbox[3]}));
    }).catch(()=>{});
  }

  const hasFace = result.faceLandmarks && result.faceLandmarks.length>0;
  mFace.textContent = hasFace ? 'yes' : 'no';
  mFace.className = 'val ' + (hasFace?'ok':'bad');

  if(hasFace){
    const lm = result.faceLandmarks[0];
    S.lastSeenTime = tSec;
    S.coverSince = null; S.coverAlerted=false;
    if(S.coverActive){ S.coverActive=false; hideBanner('cover'); }

    drawFaceMesh(lm); drawEyeRings(lm);

    let direction='center';
    if(result.facialTransformationMatrixes && result.facialTransformationMatrixes[0]){
      direction = headYaw(result.facialTransformationMatrixes[0].data);
    }
    S.direction = direction;
    mHead.textContent = direction;

    if(direction!=='center' && !S.wasLookingAway){
      S.turnCount++; S.wasLookingAway=true;
      if(direction==='left') S.turnLeft++; else if(direction==='right') S.turnRight++;
      addLog('head', `Head turned ${direction} (turn #${S.turnCount})`);
      if(levelOf(scoreOf(S.turnCount))>=1) playSound('head');
    } else if(direction==='center'){
      S.wasLookingAway=false;
    }
    S.score = computeScore();
    const lvl = levelOf(S.score);
    if(lvl===2 && S.wasLookingAway && S.headEvidence.length<HEAD_MAX){
      pushEvidence(S.headEvidence, captureEvidence('HIGH RISK — Repeated head turns','rgba(180,10,20,.85)'), HEAD_MAX);
      addLog('crit', `Evidence photo captured (${S.headEvidence.length}/${HEAD_MAX})`);
      showBanner('crit', `${S.name || 'Student'} — repeated head turns`, true);
      triggerFullFlash('crit', 'CHEATING ALERT', 2200);
    }

    const [gd, gratio] = gazeDirection(lm);
    S.gazeDir=gd; S.gazeRatio=gratio;
    mGaze.textContent = gd;
    if(gd!=='center'){
      if(S.gazeOutSince===null){ S.gazeOutSince=tSec; S.gazeAlerted=false; }
      const elapsed = tSec-S.gazeOutSince;
      if(elapsed>=GAZE_TIME_LIMIT){
        showBanner('gaze', `looking ${gd} for ${elapsed.toFixed(1)}s`, true);
        playSound('lookaway',4);
        if(!S.gazeAlerted && S.gazeEvidence.length<GAZE_MAX && (tSec-(S.lastPhotoTime.gaze||0))>=GAZE_COOLDOWN){
          S.gazeAlerted=true; S.lastPhotoTime.gaze=tSec;
          pushEvidence(S.gazeEvidence, captureEvidence(), GAZE_MAX);
          addLog('gaze', `Sustained gaze away (${gd}), photo ${S.gazeEvidence.length}/${GAZE_MAX}`);
          triggerFullFlash('crit', `CHEATING ALERT — eyes looking ${gd}`, 1800);
          addScoreBonus(5);
        }
      }
    } else {
      S.gazeOutSince=null; S.gazeAlerted=false; hideBanner('gaze');
    }

    S.ear = eyeAspectRatio(lm);
    mEar.textContent = S.ear.toFixed(2);
    if(mEarChip) mEarChip.textContent = S.ear.toFixed(2);
    // Drowsiness / eyes-closed detection is disabled on mobile — see isMobile above.
    if(!isMobile){
      if(S.ear < EAR_THRESHOLD){
        if(S.eyesClosedSince===null){ S.eyesClosedSince=tSec; S.drowsyAlerted=false; }
        const ed = tSec-S.eyesClosedSince;
        S.drowsyActive = ed>=EAR_TIME_LIMIT;
        if(S.drowsyActive){
          showBanner('drowsy','eyes closed — open your eyes!', true);
          playSound('drowsy',4);
          if(!S.drowsyAlerted && S.drowsyEvidence.length<DROWSY_MAX && (tSec-(S.lastPhotoTime.drowsy||0))>=DROWSY_COOLDOWN){
            S.drowsyAlerted=true; S.lastPhotoTime.drowsy=tSec;
            pushEvidence(S.drowsyEvidence, captureEvidence('OPEN YOUR EYES!','rgba(20,60,180,.85)'), DROWSY_MAX);
            addLog('drowsy', `Eyes closed, photo ${S.drowsyEvidence.length}/${DROWSY_MAX}`);
            triggerFullFlash('drowsy', 'OPEN YOUR EYES!', 2000);
          }
        }
      } else {
        S.eyesClosedSince=null; S.drowsyAlerted=false;
        if(S.drowsyActive){ S.drowsyActive=false; hideBanner('drowsy'); }
      }
    } else if(S.drowsyActive){
      S.eyesClosedSince=null; S.drowsyAlerted=false; S.drowsyActive=false; hideBanner('drowsy');
    }

    if(!isMobile){
      const rect = overlay.getBoundingClientRect();
      const tagX = rect.left + (1-lm[10].x)*rect.width;
      const tagY = rect.top + lm[10].y*rect.height;
      const tag=document.createElement('div'); tag.className='eye-ratio-tag';
      tag.style.left=tagX+'px'; tag.style.top=tagY+'px';
      tag.textContent = 'Eye Ratio '+Math.round(S.ear*100);
      document.body.appendChild(tag);
      setTimeout(()=>tag.remove(), 120);
    }
  } else {
    mHead.textContent='—'; mGaze.textContent='—'; mEar.textContent='—';
    if(mEarChip) mEarChip.textContent='—';
    if(S.lastSeenTime!==null){
      const gap = tSec-S.lastSeenTime;
      if(gap<=COVER_TIMEOUT){
        if(S.coverSince===null){ S.coverSince=tSec; S.coverAlerted=false; }
        const ec=tSec-S.coverSince;
        S.coverActive = ec>=COVER_TIME_LIMIT;
        if(S.coverActive){
          showBanner('cover',"don't cover your face!", true);
          playSound('covered',4);
          if(!S.coverAlerted && S.coverEvidence.length<COVER_MAX && (tSec-(S.lastPhotoTime.cover||0))>=COVER_COOLDOWN){
            S.coverAlerted=true; S.lastPhotoTime.cover=tSec;
            pushEvidence(S.coverEvidence, captureEvidence("DON'T COVER YOUR FACE!",'rgba(120,10,140,.85)'), COVER_MAX);
            addLog('cover', `Face hidden, photo ${S.coverEvidence.length}/${COVER_MAX}`);
            triggerFullFlash('covered', "DON'T COVER YOUR FACE!", 2000);
          }
        }
      } else {
        S.coverSince=null; S.coverActive=false; hideBanner('cover');
      }
    }
  }


  const phoneBoxesThisFrame = lastPhoneBoxes;
  drawPhoneBoxes(phoneBoxesThisFrame);
  S.phoneActive = phoneBoxesThisFrame.length>0;
  mPhone.textContent = S.phoneActive ? 'yes' : 'no';
  mPhone.className = 'val ' + (S.phoneActive?'bad':'ok');
  if(S.phoneActive){
    if(S.phoneSeenSince===null){ S.phoneSeenSince=tSec; S.phoneAlerted=false; S.phoneInstant=false; }
    showBanner('phone','put the phone away!', true);
    if(!S.phoneInstant && S.phoneEvidence.length<PHONE_MAX){
      S.phoneInstant=true;
      pushEvidence(S.phoneEvidence, captureEvidence('PUT THE PHONE AWAY!','rgba(160,0,20,.88)'), PHONE_MAX);
      addLog('phone', `Phone detected (instant), photo ${S.phoneEvidence.length}/${PHONE_MAX}`);
      playSound('phone',6);
      triggerFullFlash('phone', 'PUT THE PHONE AWAY!', 2200);
      addScoreBonus(20);
    }
    const ep = tSec-S.phoneSeenSince;
    if(ep>=PHONE_TIME_LIMIT && !S.phoneAlerted && S.phoneEvidence.length<PHONE_MAX && (tSec-(S.lastPhotoTime.phone||0))>=PHONE_COOLDOWN){
      S.phoneAlerted=true; S.lastPhotoTime.phone=tSec;
      pushEvidence(S.phoneEvidence, captureEvidence('PUT THE PHONE AWAY!','rgba(160,0,20,.88)'), PHONE_MAX);
      addLog('phone', `Phone confirmed, photo ${S.phoneEvidence.length}/${PHONE_MAX}`);
      playSound('phone',6);
      triggerFullFlash('phone', 'PUT THE PHONE AWAY!', 2200);
      addScoreBonus(20);
    }
  } else {
    S.phoneSeenSince=null; S.phoneAlerted=false; S.phoneInstant=false; hideBanner('phone');
  }

  updateHUD();
  sendArduinoScore();
}


function updateHUD(){
  scoreNum.textContent = S.score;
  const off = DIAL_CIRC * (1 - S.score/100);
  dialFill.style.strokeDashoffset = off;
  const lvl=levelOf(S.score);
  const color = lvl===2?'var(--red)':lvl===1?'var(--amber)':'var(--cyan)';
  dialFill.style.stroke=color; riskLabel.style.color=color;
  riskLabel.textContent = lvl===2?'High risk':lvl===1?'Elevated risk':'Low risk';
  turnsLabel.textContent = `${S.turnCount} head turn${S.turnCount===1?'':'s'}`;

  if(S.sessionStart){
    const elapsed = (performance.now()-S.sessionStart)/1000;
    sessionTimer.textContent = fmtClock(elapsed);
    if(sessionChip) sessionChip.textContent = 'Session ' + fmtClock(elapsed);
  }

  if(statTotalAlerts) statTotalAlerts.textContent = logN;
  if(statHeadTurns) statHeadTurns.textContent = S.turnCount;
  if(statPhoneAlerts) statPhoneAlerts.textContent = S.phoneEvidence.length;
  if(statOtherAlerts) statOtherAlerts.textContent = S.gazeEvidence.length + S.drowsyEvidence.length + S.coverEvidence.length;
}
let timerInt=null;


let serialPort=null, serialWriter=null, lastSerialScore=-1, lastSerialTime=0;
serialBtn.addEventListener('click', async ()=>{
  try{
    serialPort = await navigator.serial.requestPort();
    await serialPort.open({ baudRate:9600 });
    serialWriter = serialPort.writable.getWriter();
    serialBtn.style.color='var(--cyan)';
    showToast('Hardware indicator connected');
  }catch(e){ showToast('Could not connect to device'); }
});
function sendArduinoScore(){
  if(!serialWriter) return;
  const anyPhone = S.phoneActive;
  const score = anyPhone?101 : S.coverActive?95 : S.drowsyActive?90 : (S.gazeOutSince && S.gazeDir!=='center' && (performance.now()/1000-S.gazeOutSince)>=GAZE_TIME_LIMIT)?75 : S.score;
  const t=performance.now()/1000;
  if(score!==lastSerialScore || (t-lastSerialTime)>=1){
    lastSerialScore=score; lastSerialTime=t;
    serialWriter.write(new TextEncoder().encode(score+"\n")).catch(()=>{});
  }
}



fsBtn.addEventListener('click', ()=>{
  if(!document.fullscreenElement){
    (stage.requestFullscreen ? stage.requestFullscreen() : Promise.reject()).catch(()=>showToast('Fullscreen not supported'));
  } else {
    document.exitFullscreen();
  }
});
document.addEventListener('fullscreenchange', ()=>{
  stage.classList.toggle('is-fullscreen', document.fullscreenElement===stage);
});



startBtn.addEventListener('click', ()=>{
  ensureAudio();
  S.running=true; S.sessionStart=performance.now(); S.bonusScore=0;
  startBtn.style.display='none'; endBtn.style.display='inline-flex';
  statusPill.classList.add('live'); statusText.textContent='Monitoring';
  sessionState.textContent='in progress';
  stage.classList.add('monitoring');
  if(sideStatusDot) sideStatusDot.classList.add('on');
  if(sideStatusValue) sideStatusValue.textContent='Active';
  addLog('head','Session started');
  timerInt = setInterval(()=>{ updateHUD(); updateLiveCharts(); }, 1000);
  loop();
});

muteBtn.addEventListener('click', ()=>{
  S.muted=!S.muted;
  muteBtn.classList.toggle('btn-outline-cyan', !S.muted);
  muteIcon.style.opacity = S.muted?0.4:1;
  showToast(S.muted?'Alert sounds muted':'Alert sounds on');
});

endBtn.addEventListener('click', endSession);

async function endSession(){
  S.running=false; S.sessionEnd=performance.now();
  cancelAnimationFrame(rafId); clearInterval(timerInt);
  statusPill.classList.remove('live'); statusText.textContent='Ended';
  sessionState.textContent='completed';
  stage.classList.remove('monitoring');
  if(sideStatusDot) sideStatusDot.classList.remove('on');
  if(sideStatusValue) sideStatusValue.textContent='Standby';
  endBtn.style.display='none';
  video.srcObject?.getTracks().forEach(t=>t.stop());
  sendFinalZero();
  addLog('head','Session ended');
  openCompletion();
}
function sendFinalZero(){ if(serialWriter){ serialWriter.write(new TextEncoder().encode("0\n")).catch(()=>{}); } }



function openCompletion(){
  const isSuspicious = S.score >= 50;
  if (S.name) {
    doneName.textContent = isSuspicious
      ? `Suspicious activity detected — ${S.name.split(' ')[0]}`
      : `Great work, ${S.name.split(' ')[0]}`;
  } else {
    doneName.textContent = isSuspicious ? 'Suspicious activity detected' : 'Session summary';
  }
  const totalPhotos = S.headEvidence.length+S.gazeEvidence.length+S.drowsyEvidence.length+S.coverEvidence.length+S.phoneEvidence.length;
  const durSec = (S.sessionEnd-S.sessionStart)/1000;
  summaryGrid.innerHTML = `
    <div class="summary-cell"><div class="n" style="color:${S.score>=LEVEL_ALERT?'var(--red)':S.score>=LEVEL_WARNING?'var(--amber)':'var(--cyan)'}">${S.score}%</div><div class="l">Suspicion score</div></div>
    <div class="summary-cell"><div class="n">${fmtClock(durSec)}</div><div class="l">Session length</div></div>
    <div class="summary-cell"><div class="n">${S.turnCount}</div><div class="l">Head turns</div></div>
    <div class="summary-cell"><div class="n">${totalPhotos}</div><div class="l">Evidence photos</div></div>
  `;
  reportSub.textContent='Preparing\u2026'; zipSub.textContent='Preparing\u2026';
  dlReport.classList.add('pending'); dlZip.classList.add('pending');
  reportGo.innerHTML='<div class="spin"></div>'; zipGo.innerHTML='<div class="spin"></div>';



  const lvl = levelOf(S.score);
  const riskTxt = lvl===2 ? 'High risk' : lvl===1 ? 'Elevated risk' : 'Low risk';
  saveReportRecord({
    id: Date.now(),
    name: S.name || 'Student',
    date: new Date().toLocaleString(),
    durationSec: durSec,
    score: S.score,
    risk: riskTxt,
    headTurns: S.turnCount,
    phoneDetections: S.phoneEvidence.length,
    totalAlerts: logN,
    evidenceCount: totalPhotos
  });

  completionModal.classList.add('show');
  renderCompletionCharts();
  buildReport(); buildZip();
}
closeCompletion.addEventListener('click', ()=>{
  completionModal.classList.remove('show');
  showView('dashboard');
});

function dataUrlToUint8(dataUrl){
  const base64 = dataUrl.split(',')[1];
  const bin = atob(base64);
  const arr = new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
  return arr;
}
function safeName(n){ return (n||'Student').replace(/[^a-z0-9]+/gi,'_'); }

async function buildReport(){
  try{
    const score=S.score, lvl=levelOf(score);
    const risk = lvl===2?'HIGH':lvl===1?'MEDIUM':'LOW';
    const riskCol = lvl===2?'C0392B':lvl===1?'E67E22':'27AE60';
    const scoreCol = riskCol;
    const phoneN=S.phoneEvidence.length, drowsyN=S.drowsyEvidence.length, coverN=S.coverEvidence.length;

    let assessment;
    if(phoneN>0) assessment = `${S.name} was detected with a mobile phone during the exam. This represents a serious academic integrity violation.`;
    else if(lvl===2) assessment = `${S.name} showed a HIGH level of suspicious activity. The head-turn and gaze counts exceed the alert threshold. Immediate review is recommended.`;
    else if(lvl===1) assessment = `${S.name} showed a MEDIUM level of suspicious activity. Further review of the evidence photographs is recommended.`;
    else if(coverN>0) assessment = `${S.name}'s face was repeatedly hidden from the camera during the exam. This prevents proper monitoring and should be reviewed.`;
    else if(drowsyN>0) assessment = `${S.name} showed minimal suspicious head/eye movement, but repeated drowsiness was detected. A short break or attentiveness check is recommended.`;
    else assessment = `${S.name} showed minimal suspicious activity during the exam session. No immediate action required.`;

    const cellBorder = { top:{style:BorderStyle.SINGLE,size:1,color:"CCCCCC"}, bottom:{style:BorderStyle.SINGLE,size:1,color:"CCCCCC"}, left:{style:BorderStyle.SINGLE,size:1,color:"CCCCCC"}, right:{style:BorderStyle.SINGLE,size:1,color:"CCCCCC"} };
    const cellMargins = { top:80, bottom:80, left:120, right:120 };
    const noBorder = { top:{style:BorderStyle.NONE,size:0}, bottom:{style:BorderStyle.NONE,size:0}, left:{style:BorderStyle.NONE,size:0}, right:{style:BorderStyle.NONE,size:0} };

    const statCell = (label,val,color,fill)=> new TableCell({ borders:cellBorder, width:{size:4680,type:WidthType.DXA}, shading:{fill:fill||"F8F9FA",type:ShadingType.CLEAR}, margins:cellMargins,
      children:[ new Paragraph({children:[new TextRun({text:label,bold:true,size:20})]}), new Paragraph({children:[new TextRun({text:String(val),bold:true,size:30,color})]}) ] });

    const allEvidence = [
      ...S.headEvidence.map(e=>({...e,label:'Head Turn'})),
      ...S.gazeEvidence.map(e=>({...e,label:'Eye / Gaze'})),
      ...S.drowsyEvidence.map(e=>({...e,label:'Drowsiness'})),
      ...S.coverEvidence.map(e=>({...e,label:'Face Covered'})),
      ...S.phoneEvidence.map(e=>({...e,label:'Phone Detected'})),
    ];
    const ecol = l=> l==='Phone Detected'?'C0392B': l==='Face Covered'?'8E44AD': l==='Drowsiness'?'1F618D': l==='Eye / Gaze'?'E67E22':'1A5276';
    const ebg  = l=> l==='Phone Detected'?'FFE5E5': l==='Face Covered'?'F4E6F7': l==='Drowsiness'?'D6EAF8': l==='Eye / Gaze'?'FFF3CD':'FFFFFF';

    const tlRows = allEvidence.length ? allEvidence.map(e=> new TableRow({children:[
        new TableCell({borders:cellBorder,width:{size:2400,type:WidthType.DXA},margins:cellMargins,children:[new Paragraph({children:[new TextRun({text:e.ts,size:20})]})]}),
        new TableCell({borders:cellBorder,width:{size:2500,type:WidthType.DXA},shading:{fill:ebg(e.label),type:ShadingType.CLEAR},margins:cellMargins,children:[new Paragraph({children:[new TextRun({text:e.label,bold:true,size:20,color:ecol(e.label)})]})]}),
        new TableCell({borders:cellBorder,width:{size:4460,type:WidthType.DXA},margins:cellMargins,children:[new Paragraph({children:[new TextRun({text:'Evidence photo captured',size:20})]})]}),
      ]})) : [ new TableRow({children:[ new TableCell({columnSpan:3,borders:cellBorder,margins:cellMargins,children:[new Paragraph({children:[new TextRun({text:'No incidents recorded.',italics:true,color:'888888'})]})]}) ]}) ];

    function photoSection(entries, heading, color){
      const out = [ new Paragraph({heading:HeadingLevel.HEADING_2, children:[new TextRun({text:`${heading}  (${entries.length} photo${entries.length===1?'':'s'})`, color})]}) ];
      if(!entries.length){
        out.push(new Paragraph({children:[new TextRun({text:'No incidents recorded.',italics:true,color:'888888'})]}));
        out.push(new Paragraph({children:[]}));
        return out;
      }
      entries.forEach((e,i)=>{
        out.push(new Paragraph({children:[new TextRun({text:`Incident ${i+1} \u2014 ${e.ts}`,bold:true,size:20})]}));
        // Keep each photo's original aspect ratio (mobile shots are often portrait)
        // instead of forcing a fixed landscape box, which stretched them.
        const maxW=440, maxH=560;
        let dw=e.w||1280, dh=e.h||720;
        const scale=Math.min(maxW/dw, maxH/dh);
        dw=Math.round(dw*scale); dh=Math.round(dh*scale);
        out.push(new Paragraph({spacing:{after:160}, children:[new ImageRun({ data:dataUrlToUint8(e.dataUrl), transformation:{width:dw,height:dh}, type:"jpg" })]}));
      });
      return out;
    }

    const photosChildren = [
      ...photoSection(S.headEvidence,'Head Turn Evidence','1A5276'),
      ...photoSection(S.gazeEvidence,'Eye / Gaze Evidence','7D6608'),
      ...photoSection(S.drowsyEvidence,'Drowsiness Evidence','1F618D'),
      ...photoSection(S.coverEvidence,'Face Covered Evidence','8E44AD'),
      ...photoSection(S.phoneEvidence,'Phone Detection Evidence','C0392B'),
    ];

    const doc = new Document({
      styles:{ default:{ document:{ run:{ font:"Arial", size:22 } } },
        paragraphStyles:[
          {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:36,bold:true,font:"Arial",color:"1B2631"},paragraph:{spacing:{before:280,after:160},outlineLevel:0}},
          {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:26,bold:true,font:"Arial",color:"2C3E50"},paragraph:{spacing:{before:240,after:120},outlineLevel:1}},
        ] },
      sections:[{
        properties:{ page:{ size:{width:12240,height:15840}, margin:{top:1080,right:1080,bottom:1080,left:1080} } },
        children:[
          new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:80}, children:[new TextRun({text:"EXAM INTEGRITY REPORT",bold:true,size:44,color:"1B2631"})]}),
          new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:40}, children:[new TextRun({text:"Sentinel Exam Monitor — Web Edition",size:20,color:"888888"})]}),
          new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:320}, children:[new TextRun({text:`Generated: ${new Date().toLocaleString()}`,size:20,color:"AAAAAA"})]}),
          new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:6,color:"2E4057",space:1}}, children:[], spacing:{after:320}}),

          new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[4680,4680], rows:[
            new TableRow({children:[ new TableCell({columnSpan:2,borders:noBorder, shading:{fill:"2C3E50",type:ShadingType.CLEAR}, margins:{top:160,bottom:160,left:240,right:240},
              children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:S.name,bold:true,size:40,color:"FFFFFF"})]})] }) ]}),
            new TableRow({children:[ statCell("Suspicion Score", score+"%", scoreCol, "EBF5FB"), statCell("Risk Level", risk, riskCol, lvl===2?"FDEDEC":lvl===1?"FEF9E7":"EAFAF1") ]}),
            new TableRow({children:[ statCell("Head Turn Alerts", S.headEvidence.length, "1A5276"), statCell("Eye / Gaze Alerts", S.gazeEvidence.length, "7D6608") ]}),
            new TableRow({children:[ statCell("Drowsiness Alerts", drowsyN, drowsyN>0?"1F618D":"27AE60", drowsyN>0?"D6EAF8":"F8F9FA"), statCell("Face Covered Alerts", coverN, coverN>0?"8E44AD":"27AE60", coverN>0?"F4E6F7":"F8F9FA") ]}),
            new TableRow({children:[ new TableCell({columnSpan:2,borders:cellBorder, shading:{fill:phoneN>0?"FDEDEC":"F8F9FA",type:ShadingType.CLEAR}, margins:cellMargins,
              children:[ new Paragraph({children:[new TextRun({text:"Phone Detection Events",bold:true,size:20})]}), new Paragraph({children:[new TextRun({text:String(phoneN),bold:true,size:30,color:phoneN>0?"C0392B":"27AE60"})]}) ] }) ]}),
          ] }),

          new Paragraph({children:[], spacing:{after:240}}),
          new Paragraph({heading:HeadingLevel.HEADING_2, children:[new TextRun({text:"Assessment",color:"2C3E50"})]}),
          new Paragraph({spacing:{after:240}, children:[new TextRun({text:assessment,size:22})]}),
          new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:2,color:"CCCCCC",space:1}}, children:[], spacing:{after:240}}),

          new Paragraph({heading:HeadingLevel.HEADING_2, children:[new TextRun({text:"Incident Timeline",color:"2C3E50"})]}),
          new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[2400,2500,4460], rows:[
            new TableRow({tableHeader:true, children:[
              new TableCell({borders:cellBorder,width:{size:2400,type:WidthType.DXA}, shading:{fill:"2C3E50",type:ShadingType.CLEAR}, margins:cellMargins, children:[new Paragraph({children:[new TextRun({text:"Timestamp",bold:true,size:20,color:"FFFFFF"})]})]}),
              new TableCell({borders:cellBorder,width:{size:2500,type:WidthType.DXA}, shading:{fill:"2C3E50",type:ShadingType.CLEAR}, margins:cellMargins, children:[new Paragraph({children:[new TextRun({text:"Event Type",bold:true,size:20,color:"FFFFFF"})]})]}),
              new TableCell({borders:cellBorder,width:{size:4460,type:WidthType.DXA}, shading:{fill:"2C3E50",type:ShadingType.CLEAR}, margins:cellMargins, children:[new Paragraph({children:[new TextRun({text:"Details",bold:true,size:20,color:"FFFFFF"})]})]}),
            ]}), ...tlRows
          ] }),

          new Paragraph({children:[], spacing:{after:240}}),
          new Paragraph({children:[new PageBreak()]}),
          new Paragraph({heading:HeadingLevel.HEADING_1, children:[new TextRun("Evidence Photographs")]}),
          new Paragraph({spacing:{after:240}, children:[new TextRun({text:"The following screenshots were captured automatically when suspicious behaviour was detected.", size:20, color:"555555"})]}),
          ...photosChildren,
          new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:2,color:"CCCCCC",space:1}}, children:[], spacing:{before:320}}),
          new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"This report was generated automatically by Sentinel Exam Monitor. All timestamps are local device time.", size:18, color:"AAAAAA"})]}),
        ]
      }]
    });

    const blob = await Packer.toBlob(doc);
    const fname = `${safeName(S.name)}_Exam_Report.docx`;
    reportSub.textContent = `${(blob.size/1024).toFixed(0)} KB \u00b7 downloaded automatically`;
    reportGo.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke-width="1.8" style="width:16px;height:16px;stroke:currentColor;"><path d="M12 5v14M5 12l7 7 7-7"/></svg>';
    dlReport.classList.remove('pending');
    dlReport.onclick = ()=>{ saveAs(blob, fname); showToast('Report downloaded'); };
    saveAs(blob, fname);
    showToast('Report downloaded automatically — check your Downloads folder');
  }catch(err){
    console.error(err);
    reportSub.textContent='Failed to generate — try again';
    reportGo.innerHTML='';
  }
}

async function buildZip(){
  try{
    const zip = new JSZip();
    const add = (arr,folder)=> arr.forEach((e,i)=> zip.file(`${folder}/${folder}_${i+1}_${e.ts.replace(/:/g,'-')}.jpg`, e.dataUrl.split(',')[1], {base64:true}));
    add(S.headEvidence,'head_turn'); add(S.gazeEvidence,'gaze'); add(S.drowsyEvidence,'drowsiness'); add(S.coverEvidence,'face_covered'); add(S.phoneEvidence,'phone');
    const total = S.headEvidence.length+S.gazeEvidence.length+S.drowsyEvidence.length+S.coverEvidence.length+S.phoneEvidence.length;
    if(total===0) zip.file('README.txt','No incidents were recorded during this session.');
    const blob = await zip.generateAsync({type:'blob'});
    const fname = `${safeName(S.name)}_Evidence.zip`;
    zipSub.textContent = `${total} photo${total===1?'':'s'} \u00b7 ${(blob.size/1024).toFixed(0)} KB`;
    zipGo.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke-width="1.8" style="width:16px;height:16px;stroke:currentColor;"><path d="M12 5v14M5 12l7 7 7-7"/></svg>';
    dlZip.classList.remove('pending');
    dlZip.onclick = ()=>{ saveAs(blob, fname); showToast('Evidence archive downloaded'); };
  }catch(err){
    console.error(err);
    zipSub.textContent='Failed to generate — try again';
    zipGo.innerHTML='';
  }
}



preloadSounds();
loadModels().catch(err=>{
  console.error(err);
  cancelAnimationFrame(progRAF);
  loadBar.style.background = 'var(--red)';
  loadText.textContent='Failed to load models — check your connection and reload.';
});
window.addEventListener('resize', ()=>{ if(video.videoWidth){ overlay.width=video.videoWidth; overlay.height=video.videoHeight; } });