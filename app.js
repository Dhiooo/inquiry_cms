/* ===== Country map (derived from embedded phone prefixes) ===== */
const COUNTRIES = {
  '62':'Indonesia','60':'Malaysia','63':'Philippines','65':'Singapore',
  '66':'Thailand','84':'Vietnam','91':'India','998':'Uzbekistan',
  '1':'United States','44':'United Kingdom','61':'Australia'
};
function dialInfo(digits){
  for(const code of ['998','91','66','84','65','63','62','61','60','44','1']){
    if(digits.startsWith(code)) return {code:'+'+code, country:COUNTRIES[code]};
  }
  return null;
}
// normalize a raw phone from the spreadsheet into digits w/ country code
function normalizePhone(raw){
  if(raw==null) return {ok:false, digits:'', pretty:'', code:null, country:null, note:null};
  let s=String(raw).trim();
  let note=null;
  if(/[\u2010-\u2015\u2212]/.test(s)) note='format';        // unicode dashes
  if(/\.0$/.test(s)){ s=s.replace(/\.0$/,''); note='excel'; } // excel number artifact
  let d=s.replace(/[^0-9]/g,'');
  if(!d) return {ok:false, digits:'', pretty:s, code:null, country:null, note:null};
  if(d.startsWith('0')) d='62'+d.slice(1);                    // local 08.. -> 62..
  else if(d.startsWith('8')) d='62'+d;                        // bare 8.. -> 62..
  const info=dialInfo(d);
  const local = info ? d.slice(info.code.length-1) : d;
  return {ok:true, digits:d, pretty:local, code:info?info.code:null, country:info?info.country:null, note};
}

/* ===== Import dataset (real rows from “Inquiry Records Exsample.xlsx” + injected edge cases) ===== */
/* sheet columns: Date, Phone Number, Name(Optional), Source Ads, Age Group, Domisili, Chat(status), Notes */
const RAW = [
  {r:3,  date:'2026-08-01', phone:'85271947194.0',     name:null,        source:'Others Admin', chat:'C1'},
  {r:4,  date:null,         phone:'62 858-8190-5895',  name:'Andaleeb',  source:'Unknown',      chat:'C3'},
  {r:5,  date:'2026-08-02', phone:'+62 812-1346-1452', name:'Ika',       source:'Instagram',    chat:'C2'},
  {r:6,  date:null,         phone:'+62 812-1346-1452', name:'Rian',      source:'Facebook',     chat:'C1'}, // DUP of Ika
  {r:7,  date:null,         phone:'62 856-8561-118',   name:null,        source:'Instagram',    chat:'C3'},
  {r:8,  date:null,         phone:null,                name:'Budi',      source:'Facebook',     chat:'A'},  // EMPTY phone
  {r:9,  date:'2026-08-04', phone:'0812 9452 7929 ',   name:null,        source:'Others Admin', chat:'A'},
  {r:10, date:null,         phone:'62 852\u20111968\u20110024', name:'Sena', source:'Others Admin', chat:'D3'}, // unicode dash
  {r:11, date:null,         phone:'81234861881.0',     name:null,        source:'Others Admin', chat:'D2'},
  {r:12, date:null,         phone:'62 816-513-514',    name:'Linda',     source:'Web',          chat:'A'},
  {r:13, date:'2026-08-06', phone:'62 813-6611-2700',  name:null,        source:'Instagram',    chat:'C2'},
  {r:14, date:null,         phone:'62 812-1083-210',   name:'Sukirman',  source:'Walk in',      chat:'C1'},
  {r:15, date:'2026-08-07', phone:'+62 813-3173-4534', name:null,        source:'Instagram',    chat:'C1'},
  {r:16, date:'2026-08-08', phone:'62 813-3366-1588',  name:'Ayu Narita',source:'',             chat:'C4'}  // empty source
];

/* ===== Status system (from the Chat column funnel codes) ===== */
// Official legend from Mr Dewa's spreadsheet (Code / Category / Criteria)
const STATUS_LIST = [
  {code:'A',  cat:'Registered',          lbl:'Registered', cls:'reg',   desc:'Sudah daftar untuk trial class.'},
  {code:'B',  cat:'Waiting',             lbl:'Waiting',    cls:'wait',  desc:'Masih diskusi dulu dengan orang tua/keluarga.'},
  {code:'C1', cat:'No Response',         lbl:'No Response',cls:'prog',  desc:'Pesan belum dibaca setelah balasan pertama kita.'},
  {code:'C2', cat:'No Response',         lbl:'No Response',cls:'prog',  desc:'Sudah dibaca tapi tidak dibalas setelah reply pertama.'},
  {code:'C3', cat:'No Response',         lbl:'No Response',cls:'prog',  desc:'Tidak merespons setelah ditanya/diberi info harga.'},
  {code:'C4', cat:'No Response',         lbl:'No Response',cls:'prog',  desc:'Tidak merespons setelah ditawari jadwal trial.'},
  {code:'C5', cat:'No Response',         lbl:'No Response',cls:'prog',  desc:'Tidak merespons setelah diminta data untuk trial.'},
  {code:'D1', cat:'Issues',              lbl:'Issues',     cls:'issue', desc:'Terkendala harga.'},
  {code:'D2', cat:'Issues',              lbl:'Issues',     cls:'issue', desc:'Terkendala lokasi / jarak.'},
  {code:'D3', cat:'Issues',              lbl:'Issues',     cls:'issue', desc:'Terkendala jadwal.'},
  {code:'D4', cat:'Issues',              lbl:'Issues',     cls:'issue', desc:'Terkendala durasi belajar / kurikulum.'},
  {code:'E1', cat:'Low customer quality',lbl:'Low Quality',cls:'low',   desc:'Tidak tertarik / tidak paham coding.'},
  {code:'E2', cat:'Low customer quality',lbl:'Low Quality',cls:'low',   desc:'Inquiry iseng / prank (bercanda, dll).'},
  {code:'F1', cat:'Others',              lbl:'Others',     cls:'other', desc:'Lain-lain.'},
  {code:'F2', cat:'Others',              lbl:'Others',     cls:'other', desc:'Lain-lain (butuh penjelasan tambahan).'}
];
const STATUS = {}; STATUS_LIST.forEach(s=>STATUS[s.code]=s);
const CLS_COLOR = {reg:'#1c8f42',wait:'#c99411',prog:'#2586bd',issue:'#d5803b',low:'#c0563f',other:'#8a95a1'};
function statusMeta(code){return STATUS[code]||{code,cat:'',lbl:(code||'—'),cls:'other',desc:''};}
function statusPill(code){const s=statusMeta(code);return `<span class="pill ${s.cls}" title="${esc(s.code)} · ${esc(s.cat)} — ${esc(s.desc)}">${esc(s.lbl)}<span class="scode">${esc(s.code)}</span></span>`;}
function statusCats(){const c=[];STATUS_LIST.forEach(s=>{if(!c.includes(s.cat))c.push(s.cat);});return c;}
function statusOptions(){return statusCats().map(cat=>`<optgroup label="${cat}">`+STATUS_LIST.filter(s=>s.cat===cat).map(s=>`<option value="${s.code}">${s.code} — ${esc(s.desc)}</option>`).join('')+`</optgroup>`).join('');}

/* ===== Existing dashboard rows (mirrors the current CMS screenshots) ===== */
let dashRows = [
  {branch:'HQ Training',student:'Jim Low Lap Hong',parent:'MY ON',source:'OTHER ADMIN',country:'Malaysia',code:'+60',phone:'0000000000',date:'01 Aug 2026',chat:'A'},
  {branch:'HQ Training',student:'Isaac Hsu Li-Hang',parent:'MY ON',source:'OTHER ADMIN',country:'Malaysia',code:'+60',phone:'0000000000',date:'01 Aug 2026',chat:'A'},
  {branch:'HQ Training',student:'Ilhan Kamil',parent:'MY ON',source:'OTHER ADMIN',country:'Malaysia',code:'+60',phone:'0000000000',date:'01 Aug 2026',chat:'A'},
  {branch:'HQ Training',student:'Fathulloh',parent:'UZ FC',source:'OTHER ADMIN',country:'Uzbekistan',code:'+998',phone:'0000000000',date:'01 Aug 2026',chat:'A'},
  {branch:'HQ Training',student:'Sasuke',parent:'Uchiha',source:'WALK IN',country:'Philippines',code:'+63',phone:'099712321123',date:'18 Aug 2026',chat:'C2'},
  {branch:'HQ Training',student:'Adam',parent:'Zeny',source:'WALK IN',country:'Philippines',code:'+63',phone:'9060182075',date:'18 Aug 2026',chat:'C2'},
  {branch:'HQ Training',student:'Adhwa',parent:'Tammy',source:'WA',country:'Indonesia',code:'+62',phone:'81213694239',date:'18 Aug 2026',chat:'A'},
  {branch:'HQ Training',student:'Dewa',parent:"Dewa's Dad",source:'WA',country:'Indonesia',code:'+62',phone:'81231241441',date:'—',chat:'C1'}
];

function esc(s){return (s==null?'':String(s)).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));}

function renderDash(){
  const tb=document.getElementById('tbody');
  tb.innerHTML = dashRows.map((x,i)=>`<tr class="${x._new?'newrow':''}">
    <td>${esc(x.branch)}</td>
    <td>${esc(x.student)||'<span class=miss>—</span>'}</td>
    <td>${esc(x.parent)||'<span class=miss>—</span>'}</td>
    <td>${esc(x.source)||'<span class=miss>—</span>'}</td>
    <td>${esc(x.country)||'<span class=miss>—</span>'}</td>
    <td>${esc(x.code)||'<span class=miss>—</span>'}</td>
    <td>${x.phone?esc(x.phone):'<span class=miss>belum diisi</span>'}</td>
    <td>${esc(x.date)}</td>
    <td class="sm-cell">${x.sosmed?esc(x.sosmed):'<span class=miss>—</span>'}</td>
    <td class="note-cell"${x.note?' title="'+esc(x.note)+'"':''}>${x.note?esc(x.note):'<span class=miss>—</span>'}</td>
    <td>${statusPill(x.chat)}${x._incomplete?' <span class="pill warn" title="Perlu dilengkapi di CMS">⚠ Perlu dilengkapi</span>':''}</td>
    <td style="position:relative"><button class="dotsbtn" onclick="openActionMenu(event,${i})">⋯</button></td>
  </tr>`).join('');
  document.getElementById('rangelbl').textContent = `1-${Math.min(10,dashRows.length)} of ${73+dashRows.length}`;
}
renderDash();

/* ===== Validation engine ===== */
// field importance: phone = FATAL identifier; date/source/status = recommended (skippable); name = optional
let VROWS = [];
function validate(){
  const seen = {}; // normalized digits -> {row, who}
  // seed with existing dashboard phones so cross-checks work too
  VROWS = RAW.map(x=>{
    const p = normalizePhone(x.phone);
    const who = x.name || (p.ok? (p.code+' '+p.pretty) : 'Tanpa nama');
    const issues = [];
    let sev = 'ok';
    // --- FATAL: phone ---
    if(!p.ok){
      sev='er';
      issues.push({s:'e',short:'Nomor telepon kosong',fix:'wajib diisi (identitas unik) — atau Force Add'});
    } else {
      if(seen[p.digits]){
        sev='er';
        const o=seen[p.digits];
        issues.push({s:'e',short:`Nomor duplikat dengan Baris ${o.row} (${esc(o.who)})`,fix:'perbaiki salah satu, atau Force Add'});
      } else {
        seen[p.digits]={row:x.r,who};
      }
      // derived from phone → shown as small auto-chips, not errors
      issues.push({s:'i',chip:`🌐 ${p.country||'?'} (${p.code||'?'})`});
      if(p.note==='format') issues.push({s:'i',chip:'format dirapikan'});
      if(p.note==='excel') issues.push({s:'i',chip:'.0 dibersihkan'});
    }
    // --- recommended (skippable) ---
    if(!x.date){ if(sev!=='er')sev=sev==='ok'?'wr':sev; issues.push({s:'w',short:'Inquiry Date kosong',fix:'bisa dilengkapi nanti'}); }
    if(!x.source){ if(sev!=='er')sev=sev==='ok'?'wr':sev; issues.push({s:'w',short:'Source kosong',fix:'bisa dilengkapi nanti'}); }
    if(!x.chat){ if(sev!=='er')sev=sev==='ok'?'wr':sev; issues.push({s:'w',short:'Status kosong',fix:'default akan dipakai'}); }
    // --- optional (info only) ---
    if(!x.name) issues.push({s:'i',chip:'nama kosong · opsional'});
    return {raw:x, p, who, issues, sev, force:false};
  });
  paint();
}

function counts(){
  let ok=0,wr=0,er=0;
  VROWS.forEach(v=>{ if(v.sev==='er')er++; else if(v.sev==='wr')wr++; else ok++; });
  return {ok,wr,er};
}

function paint(){
  const c=counts();
  document.getElementById('cOk').textContent=c.ok;
  document.getElementById('cWr').textContent=c.wr;
  document.getElementById('cEr').textContent=c.er;
  const skip=document.getElementById('skipToggle').checked;
  const filt=document.querySelector('.chip.on').dataset.f;

  const list=document.getElementById('rowlist');
  list.innerHTML = VROWS.filter(v=>filt==='all'|| (filt==='wr'&&v.sev==='wr') || (filt==='er'&&v.sev==='er')).map((v,i)=>{
    const idx=VROWS.indexOf(v);
    const stateLbl = v.sev==='er'?(v.force?'Force Add aktif':'Fatal'):(v.sev==='wr'?(skip?'Akan di-skip':'Perlu dilengkapi'):'Siap');
    const stateCls = v.sev==='er'?(v.force?'wr':'er'):(v.sev==='wr'?(skip?'ok':'wr'):'ok');
    const rowCls = v.sev==='er'&&!v.force?'er':(v.sev==='wr'&&!skip?'wr':(v.sev==='er'&&v.force?'wr':'ok'));
    const chips=v.issues.filter(is=>is.s==='i');
    const probs=v.issues.filter(is=>is.s!=='i');
    const chipsHtml = chips.length?`<div class="rchips">`+chips.map(c=>`<span class="ac">${c.chip}</span>`).join('')+`</div>`:'';
    const probsHtml = probs.length?`<div class="rprobs">`+probs.map(pb=>`<div class="pr ${pb.s}"><span class="pi">${pb.s==='e'?'✖':'!'}</span><span class="pt">${pb.short}${pb.fix?` <span class="fix">· ${esc(pb.fix)}</span>`:''}</span></div>`).join('')+`</div>`:'';
    const force = v.sev==='er'?`<div class="forcebar"><span class="warnico">⛔ Fatal</span><label class="switch"><input type="checkbox" ${v.force?'checked':''} onchange="toggleForce(${idx})"><span class="track"></span> Force Add — tambah tetap, perbaiki di CMS</label></div>`:'';
    return `<div class="rowc ${rowCls}">
      <div class="rhead">
        <span class="ricon">${v.sev==='er'?'✕':(v.sev==='wr'?'!':'✓')}</span>
        <div class="rid">
          <div class="rname${v.raw.name?'':' muted'}">${v.raw.name?esc(v.raw.name):'Tanpa nama'}</div>
          <div class="rsub"><span class="rn">Baris ${v.raw.r}</span><span class="rphone">${v.p.ok?esc(v.p.code+' '+v.p.pretty):'Nomor kosong'}</span></div>
        </div>
        <span class="rstate ${stateCls}">${stateLbl}</span>
      </div>
      ${chipsHtml?`<div class="rmeta">Terdeteksi otomatis</div>${chipsHtml}`:''}
      ${probsHtml?`<div class="rmeta">Perlu diperhatikan</div>${probsHtml}`:''}
      ${force}
    </div>`;
  }).join('');

  document.querySelectorAll('.filterchips .chip').forEach(ch=>{ const f=ch.dataset.f; const n=f==='all'?VROWS.length:(f==='wr'?c.wr:c.er); const base=f==='all'?'Semua':(f==='wr'?'Perlu dilengkapi':'Fatal'); ch.textContent=base+' ('+n+')'; });

  // footer summary + Add button
  const willAdd = VROWS.filter(v=> v.sev==='ok' || (v.sev==='wr'&&skip) || (v.sev==='er'&&v.force)).length;
  const blocked = VROWS.filter(v=> v.sev==='er'&&!v.force).length;
  const wskip = VROWS.filter(v=> v.sev==='wr').length;
  document.getElementById('fsummary').innerHTML = `<b>${willAdd}</b> akan ditambahkan${skip&&wskip?` · <b>${wskip}</b> dilengkapi nanti`:''}${blocked?` · <b style="color:var(--red)">${blocked}</b> tertahan (butuh Force Add)`:''}`;
  document.getElementById('addBtn').disabled = willAdd===0;
  document.getElementById('addBtn').textContent = `Add Inquiry${willAdd?` (${willAdd})`:''}`;
}

window.toggleForce=function(i){ VROWS[i].force=!VROWS[i].force; paint(); };

/* ===== Add to dashboard (new/incomplete rows go to top) ===== */
function addInquiries(){
  const skip=document.getElementById('skipToggle').checked;
  const toAdd = VROWS.filter(v=> v.sev==='ok' || (v.sev==='wr'&&skip) || (v.sev==='er'&&v.force));
  dashRows.forEach(r=>{r._new=false;});
  const mapped = toAdd.map(v=>({
    branch:'HQ Training',
    student:v.raw.name||'',
    parent:'',
    source:(v.raw.source||'').toUpperCase(),
    country:v.p.country||'',
    code:v.p.code||'',
    phone:v.p.ok?v.p.pretty:'',
    date:v.raw.date?fmtDate(v.raw.date):'—',
    chat:v.raw.chat||'A',
    _new:true,
    _incomplete: (v.sev!=='ok')
  }));
  dashRows = mapped.concat(dashRows);
  renderDash();
  closeModal();
}
function fmtDate(iso){const m=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];const d=new Date(iso+'T00:00:00');return `${String(d.getDate()).padStart(2,'0')} ${m[d.getMonth()]} ${d.getFullYear()}`;}

/* ===== Spreadsheet issue filtering, severity and sorting ===== */
const VALIDATION_ISSUES = [
  {row:8,name:'Budi',meta:'Source · Facebook',severity:'fatal',cats:['phone'],issues:[['Phone Number - belum diisi','fatal']]},
  {row:9,name:'Ika',meta:'+62 812 4000 7788',severity:'fatal',cats:['phone'],issues:[['Phone Number - sama dengan data di baris 10','fatal']]},
  {row:6,name:'Rian',meta:'+62 512 1346 1452',severity:'warning',cats:['date','source','chat'],issues:[['Inquiry Date - tidak sesuai format','warning'],['Source Ads - tidak sesuai pilihan','warning'],['Chat - “C9” tidak ditemukan','warning']]},
  {row:4,name:'Andaleeb',meta:'+62 555 5190 5595',severity:'warning',cats:['date'],issues:[['Inquiry Date - belum diisi','warning']]},
  {row:7,name:'Tanpa nama',italic:true,meta:'+62 556 5551 115',severity:'warning',cats:['date','source'],issues:[['Inquiry Date - belum diisi','warning'],['Source Ads - pilihan tidak tersedia','warning'],['Name - belum diisi (opsional)','more']]},
  {row:10,name:'Sena',meta:'+62 852 1968 0024',severity:'warning',cats:['date'],issues:[['Inquiry Date - belum diisi','warning']]},
  {row:11,name:'Tanpa nama',italic:true,meta:'+62 812 3486 1881',severity:'warning',cats:[],issues:[['Age Group - belum diisi','warning']]},
  {row:12,name:'Linda',meta:'+62 816 513 514',severity:'warning',cats:[],issues:[['Domisili - belum diisi','warning']]},
  {row:16,name:'Ayu Narita',meta:'+62 813 3366 1588',severity:'warning',cats:[],issues:[['Contact Note - belum diisi','warning']]}
];
let ssFilter='all', ssPage=1;
const SS_PAGE_SIZE=5;
function validationMatches(row){
  const severity=document.getElementById('ssSeverity').value;
  const query=document.getElementById('ssSearch').value.trim().toLowerCase();
  if(ssFilter!=='all'&&!row.cats.includes(ssFilter)) return false;
  if(severity!=='all'&&row.severity!==severity) return false;
  if(query&&!`${row.row} ${row.name} ${row.meta} ${row.issues.map(x=>x[0]).join(' ')}`.toLowerCase().includes(query)) return false;
  return true;
}
function renderValidationTable(){
  const sort=document.getElementById('ssSort').value;
  let rows=VALIDATION_ISSUES.filter(validationMatches);
  rows.sort((a,b)=>{
    if(sort==='row-asc') return a.row-b.row;
    if(sort==='row-desc') return b.row-a.row;
    const first=sort==='warning-first'?'warning':'fatal';
    if(a.severity!==b.severity) return a.severity===first?-1:1;
    return a.row-b.row;
  });
  const pages=Math.max(1,Math.ceil(rows.length/SS_PAGE_SIZE));
  ssPage=Math.min(ssPage,pages);
  const from=(ssPage-1)*SS_PAGE_SIZE;
  const shown=rows.slice(from,from+SS_PAGE_SIZE);
  let lastSeverity=null, html='';
  shown.forEach(row=>{
    if(row.severity!==lastSeverity){
      const count=rows.filter(x=>x.severity===row.severity).length;
      html+=row.severity==='fatal'
        ?`<div class="ss-group fatal"><span>Fatal · ${count} ${count===1?'inquiry':'inquiries'}</span><span>Will not be added</span></div>`
        :`<div class="ss-group warning"><span>Incomplete · ${count} ${count===1?'inquiry':'inquiries'}</span><span>Will still be added</span></div>`;
      lastSeverity=row.severity;
    }
    const chips=row.issues.map(([text,type])=>`<span class="ss-chip ${type==='fatal'?'cf':type==='warning'?'cw':'cmore'}">${esc(text)}</span>`).join('');
    const issueCount=row.issues.filter(x=>x[1]!=='more').length;
    html+=`<div class="ss-datarow ${row.severity==='fatal'?'fatalrow':''}"><div class="ss-cell"><div class="ss-rownum">${row.row}</div></div><div class="ss-cell"><div class="ss-person">${row.italic?`<i>${esc(row.name)}</i>`:esc(row.name)}</div><div class="ss-meta">${esc(row.meta)}</div></div><div class="ss-cell"><div class="ss-chips">${chips}</div></div><div class="ss-cell"><span class="ss-badge ${row.severity==='fatal'?'bf':'bw'}">${issueCount} ${row.severity==='fatal'?'FATAL':'WARNING'}</span><span class="ss-result">${row.severity==='fatal'?'Terblokir':'Tetap di-import'}</span></div></div>`;
  });
  if(!shown.length) html='<div class="ss-empty">Tidak ada inquiry yang cocok dengan filter.</div>';
  document.getElementById('ssRows').innerHTML=html;
  const start=rows.length?from+1:0, finish=Math.min(from+SS_PAGE_SIZE,rows.length);
  document.getElementById('ssPagerText').textContent=`Showing ${start}–${finish} of ${rows.length} inquiries · Rows per page: ${SS_PAGE_SIZE}`;
  document.getElementById('ssPages').innerHTML=`<button type="button" class="ss-page" data-page="prev" ${ssPage===1?'disabled':''} aria-label="Previous page"><svg class="ci" viewBox="0 0 24 24"><polyline points="15 18 9 12 15 6"/></svg></button>${Array.from({length:pages},(_,i)=>`<button type="button" class="ss-page ${ssPage===i+1?'active':''}" data-page="${i+1}">${i+1}</button>`).join('')}<button type="button" class="ss-page" data-page="next" ${ssPage===pages?'disabled':''} aria-label="Next page"><svg class="ci" viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></svg></button>`;
  document.querySelectorAll('#ssPages [data-page]').forEach(btn=>btn.onclick=()=>{const p=btn.dataset.page;ssPage=p==='prev'?Math.max(1,ssPage-1):p==='next'?Math.min(pages,ssPage+1):Number(p);renderValidationTable();});
}

/* ===== Modal wiring ===== */
const overlay=document.getElementById('overlay');
let dashboardScrollY=0;
function lockDashboardScroll(){
  if(document.body.classList.contains('modal-open')) return;
  dashboardScrollY=window.scrollY;
  document.body.style.top=`-${dashboardScrollY}px`;
  document.documentElement.classList.add('modal-open');
  document.body.classList.add('modal-open');
}
function unlockDashboardScroll(){
  document.documentElement.classList.remove('modal-open');
  document.body.classList.remove('modal-open');
  document.body.style.top='';
  window.scrollTo(0,dashboardScrollY);
}
function openModal(){overlay.classList.add('show');lockDashboardScroll();}
function closeModal(){overlay.classList.remove('show');unlockDashboardScroll();resetModal();}
function resetModal(){document.getElementById('vresult').style.display='none';document.getElementById('addBtn').disabled=true;document.getElementById('addBtn').textContent='Add Inquiry';document.getElementById('fsummary').innerHTML='Tekan <b>Validate Inquiries</b> untuk mengecek dokumen.';}
document.getElementById('importBtn').onclick=openModal;
document.getElementById('closeX').onclick=closeModal;
document.getElementById('closeBtn').onclick=closeModal;
overlay.onclick=e=>{if(e.target===overlay)closeModal();};
document.getElementById('validateBtn').onclick=()=>{
  const vr=document.getElementById('vresult'); vr.style.display='block';
  ssFilter='all'; ssPage=1;
  document.querySelectorAll('#vresult .ss-ifilter').forEach(x=>x.classList.toggle('active',x.dataset.filter==='all'));
  document.getElementById('ssSearch').value=''; document.getElementById('ssSeverity').value='all'; document.getElementById('ssSort').value='fatal-first';
  syncPicker('ssSeverity'); syncPicker('ssSort');
  renderValidationTable();
  document.getElementById('fsummary').innerHTML='<span style="font-size:14px;font-weight:600;color:#3a444c">98 inquiries will be added</span><br><span style="color:var(--muted)">91 ready + 7 incomplete · 2 fatal excluded</span>';
  const ab=document.getElementById('addBtn'); ab.disabled=false; ab.textContent='Add Inquiry (98)';
  try{ vr.scrollIntoView({behavior:'smooth',block:'nearest'}); }catch(e){}
};
document.getElementById('addBtn').onclick=openConfirm;
document.querySelectorAll('#vresult .ss-ifilter').forEach(f=>f.onclick=()=>{document.querySelectorAll('#vresult .ss-ifilter').forEach(x=>x.classList.remove('active'));f.classList.add('active');ssFilter=f.dataset.filter;ssPage=1;renderValidationTable();});
document.getElementById('ssSearch').addEventListener('input',()=>{ssPage=1;renderValidationTable();});

/* ===== Add Inquiry confirmation dialog ===== */
function openConfirm(){ document.getElementById('overlay4').classList.add('show'); }
function closeConfirm(){ document.getElementById('overlay4').classList.remove('show'); }
document.getElementById('cfX').onclick=closeConfirm;
document.getElementById('cfCancel').onclick=closeConfirm;
document.getElementById('overlay4').onclick=e=>{ if(e.target===document.getElementById('overlay4')) closeConfirm(); };
document.getElementById('cfConfirm').onclick=()=>{ closeConfirm(); confirmAddInquiries(); };
function confirmAddInquiries(){
  dashRows.forEach(r=>{r._new=false;});
  const sample=[
    {branch:'HQ Training',student:'Rian',parent:'',source:'FACEBOOK',country:'Indonesia',code:'+62',phone:'51213461452',date:'—',chat:'C1',_new:true,_incomplete:true},
    {branch:'HQ Training',student:'Andaleeb',parent:'',source:'UNKNOWN',country:'Indonesia',code:'+62',phone:'55551905595',date:'—',chat:'C3',_new:true,_incomplete:true},
    {branch:'HQ Training',student:'',parent:'',source:'INSTAGRAM',country:'Indonesia',code:'+62',phone:'55655551115',date:'—',chat:'C3',_new:true,_incomplete:true}
  ];
  dashRows = sample.concat(dashRows);
  renderDash();
  closeModal();
  toast('98 inquiry ditambahkan · 2 fatal dilewati.');
}


/* =====================================================================
   PROBLEM 2, 3, 4 : New Inquiry modal + granular status + action menu
   ===================================================================== */
const SOURCES = ['Instagram','Facebook','WhatsApp (WA)','Walk in','Web','GMaps','Referral','Status Ads','Others Admin','Unknown'];
const COUNTRY_CODES = [
  {c:'Indonesia',d:'+62'},{c:'Malaysia',d:'+60'},{c:'Philippines',d:'+63'},{c:'Singapore',d:'+65'},
  {c:'Thailand',d:'+66'},{c:'Vietnam',d:'+84'},{c:'Uzbekistan',d:'+998'},{c:'India',d:'+91'},{c:'Australia',d:'+61'}
];
const el = id => document.getElementById(id);
let editIndex = null; // null = create, number = editing dashRows[idx]
let phoneFormatError = false;

/* ---- toast ---- */
let toastT;
function toast(msg){const t=el('toast');t.textContent=msg;t.classList.add('show');clearTimeout(toastT);toastT=setTimeout(()=>t.classList.remove('show'),2200);}

/* ---- duplicate key helpers (Problem 3) ---- */
function keyFromCodePhone(code,phone){
  const cd=(code||'').replace(/[^0-9]/g,'');
  let pn=(phone||'').replace(/[^0-9]/g,'');
  if(pn.startsWith('0')) pn=pn.slice(1);
  return cd+pn;
}
function findDuplicate(code,phone,skipIdx){
  const key=keyFromCodePhone(code,phone);
  if(key.replace((code||'').replace(/[^0-9]/g,''),'')==='') return null; // no local digits yet
  for(let i=0;i<dashRows.length;i++){
    if(i===skipIdx) continue;
    const r=dashRows[i];
    if(!r.phone) continue;
    if(keyFromCodePhone(r.code,r.phone)===key) return {row:r,idx:i};
  }
  return null;
}

/* ---- generic custom picker (matches status picker visual) ---- */
const PICKERS={};
function makePicker(id, options, opts){
  opts=opts||{};
  PICKERS[id]={options, onChange:opts.onChange, placeholder:opts.placeholder||'Pilih…'};
  const panel=el(id+'_panel');
  panel.innerHTML=options.map(o=>`<div class="ddi" data-v="${esc(o.value)}" onclick="pickVal('${id}',this.getAttribute('data-v'))">${esc(o.label)}<span class="ddck">✓</span></div>`).join('');
  el(id+'_btn').onclick=e=>{ e.stopPropagation(); const wrap=el(id+'_pick'); const willOpen=!wrap.classList.contains('open'); closeAllPickers(); if(willOpen)wrap.classList.add('open'); };
  syncPicker(id);
}
window.pickVal=function(id,v){ el(id).value=v; el(id+'_pick').classList.remove('open'); syncPicker(id); const p=PICKERS[id]; if(p&&p.onChange)p.onChange(v); };
function syncPicker(id){
  const p=PICKERS[id]; if(!p)return; const v=el(id).value;
  const opt=p.options.find(o=>String(o.value)===String(v));
  const lbl=el(id+'_label'); lbl.textContent=opt?opt.label:p.placeholder; lbl.classList.toggle('ph',!opt);
  el(id+'_panel').querySelectorAll('.ddi').forEach(it=>it.classList.toggle('sel', it.getAttribute('data-v')===String(v)));
}
function closeAllPickers(){ document.querySelectorAll('.pick').forEach(p=>p.classList.remove('open')); const sp=el('statusPick'); if(sp)sp.classList.remove('open'); }
document.addEventListener('click', e=>{ if(!e.target.closest('.pick')) document.querySelectorAll('.pick').forEach(p=>p.classList.remove('open')); });

/* ---- populate selects (custom pickers) ---- */
function initForm(){
  makePicker('ssSeverity',[{value:'all',label:'All'},{value:'fatal',label:'Fatal'},{value:'warning',label:'Warning'}],{onChange:()=>{ssPage=1;renderValidationTable();}});
  makePicker('ssSort',[{value:'fatal-first',label:'Fatal first'},{value:'warning-first',label:'Warning first'},{value:'row-asc',label:'Row ascending'},{value:'row-desc',label:'Row descending'}],{onChange:()=>{ssPage=1;renderValidationTable();}});
  makePicker('f_branch',[{value:'HQ Training',label:'HQ Training'},{value:'Cibubur',label:'Cibubur'},{value:'Kelapa Gading',label:'Kelapa Gading'},{value:'Transyogi',label:'Transyogi'}],{placeholder:'Pilih branch'});
  makePicker('f_source',SOURCES.map(s=>({value:s,label:s})),{placeholder:'Pilih sumber'});
  makePicker('f_country',COUNTRY_CODES.map(x=>({value:x.d,label:`${x.c} (${x.d})`})),{placeholder:'Pilih negara',onChange:(v)=>{ el('f_code').textContent=v||'+62'; checkDup(); }});
  buildStatusPicker();
  el('f_phone').addEventListener('input', e=>{
    const raw=e.target.value;
    const digits=raw.replace(/[^0-9]/g,'');
    phoneFormatError=raw!==digits;
    if(phoneFormatError) e.target.value=digits;
    checkDup();
  });
  const fd=el('f_date'); const updDate=()=>fd.classList.toggle('empty',!fd.value); fd.addEventListener('input',updDate); fd.addEventListener('change',updDate); updDate();
  renderStatusDesc();
}
function buildStatusPicker(){
  const panel=el('spPanel'); if(!panel) return;
  let last=null, html='';
  STATUS_LIST.forEach(s=>{
    if(s.cat!==last){ html+=`<div class="spgroup">${esc(s.cat)}</div>`; last=s.cat; }
    html+=`<div class="spitem" data-code="${s.code}" onclick="pickStatus('${s.code}')"><span class="d2" style="background:${CLS_COLOR[s.cls]}"></span><div class="txt"><div><span class="code2">${s.code}</span><span class="cat2">${esc(s.lbl)}</span></div><div class="desc2">${esc(s.desc)}</div></div><span class="ck">✓</span></div>`;
  });
  panel.innerHTML=html;
  el('spBtn').onclick=e=>{ e.stopPropagation(); const willOpen=!el('statusPick').classList.contains('open'); closeAllPickers(); if(willOpen)el('statusPick').classList.add('open'); };
  document.addEventListener('click', e=>{ if(!e.target.closest('#statusPick')) el('statusPick').classList.remove('open'); });
}
window.pickStatus=function(code){ el('f_status').value=code; el('statusPick').classList.remove('open'); renderStatusDesc(); };
function renderStatusDesc(){
  const code=el('f_status').value; const s=statusMeta(code);
  el('statusDesc').innerHTML = `<b>${esc(s.code)} · ${esc(s.cat)}</b> — ${esc(s.desc)}`;
  if(el('spDot')) el('spDot').style.background=CLS_COLOR[s.cls]||'#ccc';
  if(el('spLabel')) el('spLabel').innerHTML = code?`<b>${esc(s.code)}</b> · ${esc(s.cat)}`:'Pilih status…';
  document.querySelectorAll('#spPanel .spitem').forEach(it=>{ it.classList.toggle('sel', it.dataset.code===code); });
}

/* ---- live duplicate check (Problem 3: block, no force) ---- */
function checkDup(){
  const code=el('f_code').textContent, phone=el('f_phone').value;
  const localDigits=(phone||'').replace(/[^0-9]/g,'');
  const wrap=el('phonewrap'), warn=el('dupWarn'), ok=el('dupOk');
  wrap.classList.remove('err','ok'); warn.classList.remove('show'); ok.classList.remove('show');
  if(phoneFormatError){
    wrap.classList.add('err'); warn.classList.add('show');
    warn.innerHTML='⚠ Hanya angka 0–9 yang diperbolehkan. Huruf atau karakter lain sudah dihapus; periksa kembali nomor telepon.';
    el('niSave').disabled=true;
    return;
  }
  if(!localDigits){ el('niSave').disabled=false; return; }
  const dup=findDuplicate(code, phone, editIndex);
  if(dup){
    const who=dup.row.student||dup.row.parent||'(tanpa nama)';
    wrap.classList.add('err'); warn.classList.add('show');
    warn.innerHTML = `⛔ Nomor <b>${esc(code)} ${esc(phone)}</b> sudah terdaftar atas <b>${esc(who)}</b> (${esc(dup.row.branch)}, status ${esc(statusMeta(dup.row.chat).lbl)}). Nomor telepon adalah identitas unik — tidak boleh dipakai dua orang.`;
    el('niSave').disabled=true;
  } else {
    wrap.classList.add('ok'); ok.classList.add('show'); el('niSave').disabled=false;
  }
}

/* ---- open / close New Inquiry ---- */
function openNewInquiry(){
  editIndex=null; el('niTitle').textContent='New Inquiry';
  el('f_branch').value='HQ Training'; el('f_source').value=''; el('f_student').value=''; el('f_parent').value='';
  el('f_country').value='+62'; el('f_code').textContent='+62'; el('f_phone').value=''; phoneFormatError=false;
  el('f_date').value=''; el('f_sosmed').value=''; el('f_status').value='C1'; el('f_note').value='';
  ['f_branch','f_source','f_country'].forEach(syncPicker);
  el('f_date').classList.toggle('empty',!el('f_date').value);
  renderStatusDesc(); checkDup(); el('overlay2').classList.add('show');
}
function openEditInquiry(idx){
  editIndex=idx; const r=dashRows[idx];
  el('niTitle').textContent='Edit Inquiry';
  el('f_branch').value=r.branch||'HQ Training';
  el('f_source').value = SOURCES.find(s=>s.toUpperCase()===(r.source||'').toUpperCase())||'';
  el('f_student').value=r.student||''; el('f_parent').value=r.parent||'';
  const cc=COUNTRY_CODES.find(x=>x.d===r.code); el('f_country').value=cc?cc.d:'+62'; el('f_code').textContent=cc?cc.d:(r.code||'+62');
  el('f_phone').value=r.phone||''; phoneFormatError=false; el('f_date').value=''; el('f_sosmed').value=r.sosmed||'';
  el('f_status').value=r.chat||'A'; el('f_note').value=r.note||'';
  ['f_branch','f_source','f_country'].forEach(syncPicker);
  el('f_date').classList.toggle('empty',!el('f_date').value);
  renderStatusDesc(); checkDup(); el('overlay2').classList.add('show');
}
function closeNewInquiry(){ el('overlay2').classList.remove('show'); }

function saveInquiry(){
  const code=el('f_code').textContent, phone=el('f_phone').value.trim();
  if(!phone){ toast('Nomor telepon wajib diisi.'); el('phonewrap').classList.add('err'); return; }
  if(phoneFormatError||!/^[0-9]+$/.test(phone)){ toast('Nomor telepon hanya boleh berisi angka 0–9.'); checkDup(); return; }
  if(findDuplicate(code,phone,editIndex)){ toast('Tidak bisa disimpan — nomor telepon duplikat.'); return; }
  const cc=COUNTRY_CODES.find(x=>x.d===code);
  const rec={
    branch:el('f_branch').value, student:el('f_student').value.trim(), parent:el('f_parent').value.trim(),
    source:(el('f_source').value||'').toUpperCase(), country:cc?cc.c:'', code:code,
    phone:phone.replace(/[^0-9]/g,'').replace(/^0/,''), sosmed:el('f_sosmed').value.trim(),
    date: el('f_date').value?fmtDate(el('f_date').value):'—', chat:el('f_status').value, note:el('f_note').value.trim()
  };
  if(editIndex!=null){ dashRows[editIndex]={...dashRows[editIndex],...rec}; toast('Inquiry diperbarui.'); }
  else { dashRows.forEach(r=>r._new=false); rec._new=true; dashRows.unshift(rec); toast('Inquiry baru ditambahkan ke baris teratas.'); }
  renderDash(); closeNewInquiry();
}

/* =====================================================================
   Action menu — compact two-panel status workflow
   ===================================================================== */
let menuIdx=null;
let actionCategory='';
let actionQuery='';

const ACTION_ICONS={
  eye:`<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6S2.5 12 2.5 12Z"/><circle cx="12" cy="12" r="2.7"/></svg>`,
  edit:`<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20h4l11-11a2.8 2.8 0 0 0-4-4L4 16v4Z"/><path d="m13.8 6.2 4 4"/></svg>`,
  search:`<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6.5"/><path d="m16 16 4 4"/></svg>`,
  trash:`<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16"/><path d="M9 7V4h6v3"/><path d="m6.5 7 1 13h9l1-13"/><path d="M10 11v5M14 11v5"/></svg>`
};
const ACTION_CAT_META={
  'Registered':{code:'A',label:'Registered'},
  'Waiting':{code:'B',label:'Waiting'},
  'No Response':{code:'C',label:'No Response'},
  'Issues':{code:'D',label:'Issues'},
  'Low customer quality':{code:'E',label:'Low Quality'},
  'Others':{code:'F',label:'Others'}
};

function actionCatMeta(cat){return ACTION_CAT_META[cat]||{code:'?',label:cat};}
function actionButtonColor(s){return CLS_COLOR[s.cls]||CLS_COLOR.other;}
function actionStatusTitle(s){
  const short={
    A:'Sudah terdaftar', B:'Menunggu keputusan',
    C1:'Pesan belum dibaca', C2:'Sudah dibaca, belum dibalas',
    C3:'Tidak merespons setelah ditanya', C4:'Tidak merespons tawaran jadwal',
    C5:'Tidak merespons permintaan data', D1:'Terkendala harga',
    D2:'Terkendala lokasi atau jarak', D3:'Terkendala jadwal',
    D4:'Terkendala program belajar', E1:'Tidak tertarik atau belum paham',
    E2:'Inquiry iseng atau prank', F1:'Alasan lainnya', F2:'Perlu penjelasan tambahan'
  };
  return short[s.code]||s.lbl;
}

function actionMenuShell(idx){
  return `<div class="aquick">
    <button type="button" class="qbtn" onclick="actDetail(${idx})">${ACTION_ICONS.eye}<span>View Detail</span></button>
    <button type="button" class="qbtn" onclick="actEdit(${idx})">${ACTION_ICONS.edit}<span>Edit Inquiry</span></button>
  </div>
  <div class="divider"></div>
  <div class="ahead"><b>UBAH STATUS</b><span>Pilih kategori, lalu alasan yang paling sesuai</span></div>
  <label class="asearch">${ACTION_ICONS.search}<input id="actionSearch" type="search" placeholder="Cari status atau kode..." autocomplete="off" oninput="filterActionStatus(this.value)"></label>
  <div class="atwopane"><div class="acats" id="actionCats"></div><div class="adetails" id="actionDetails"></div></div>
  <button type="button" class="adelete" onclick="openDeleteConfirm(${idx})">${ACTION_ICONS.trash}<span>Delete Inquiry</span></button>`;
}

function renderActionStatusPane(){
  if(menuIdx===null) return;
  const r=dashRows[menuIdx];
  if(!r) return closeActionMenu();
  const cats=statusCats();
  const catsEl=el('actionCats');
  const detailsEl=el('actionDetails');
  if(!catsEl||!detailsEl)return;

  catsEl.innerHTML=cats.map(cat=>{
    const meta=actionCatMeta(cat);
    const first=STATUS_LIST.find(s=>s.cat===cat);
    const count=STATUS_LIST.filter(s=>s.cat===cat).length;
    return `<button type="button" class="acat${!actionQuery&&actionCategory===cat?' active':''}" onpointerdown="event.stopPropagation()" onclick="event.stopPropagation();pickActionCategory('${cat}')">
      <span class="acode" style="background:${actionButtonColor(first)}">${meta.code}</span><span>${esc(meta.label)}</span><span class="acount">${count}</span>
    </button>`;
  }).join('');

  const q=actionQuery.trim().toLowerCase();
  const choices=q
    ? STATUS_LIST.filter(s=>[s.code,s.cat,s.lbl,s.desc,actionStatusTitle(s)].join(' ').toLowerCase().includes(q))
    : STATUS_LIST.filter(s=>s.cat===actionCategory);
  const title=q?'Hasil pencarian':actionCatMeta(actionCategory).code+' · '+actionCatMeta(actionCategory).label;
  detailsEl.innerHTML=`<div class="adtitle"><b>${esc(title)}</b><span>${choices.length} ${q?'hasil':'alasan'}</span></div>`+
    (choices.length?choices.map(s=>`<button type="button" class="astatus${r.chat===s.code?' selected':''}" onclick="setStatus(${menuIdx},'${s.code}')">
      <span class="acode" style="background:${actionButtonColor(s)}">${s.code}</span>
      <span class="astxt"><b>${esc(actionStatusTitle(s))}</b><p>${esc(s.desc)}</p></span>
      ${r.chat===s.code?'<span class="acurrent">Saat ini</span>':''}
    </button>`).join(''):`<div class="aempty">Status tidak ditemukan.<br>Coba kode atau kata kunci lain.</div>`);
}

window.pickActionCategory=function(cat){
  actionCategory=cat; actionQuery='';
  const search=el('actionSearch'); if(search)search.value='';
  renderActionStatusPane();
};
window.filterActionStatus=function(value){actionQuery=value||'';renderActionStatusPane();};

function positionActionMenu(btn){
  const m=el('actionMenu');
  const rect=btn.getBoundingClientRect();
  const gap=7,pad=10,mw=m.offsetWidth;

  // Reset batas dari pembukaan sebelumnya, lalu ukur ukuran natural menu.
  m.style.maxHeight='none';
  m.style.overflowY='hidden';
  const mh=m.offsetHeight;
  const roomBelow=window.innerHeight-rect.bottom-gap-pad;
  const roomAbove=rect.top-gap-pad;
  const viewportRoom=window.innerHeight-(pad*2);

  let top,placement;
  if(roomBelow>=mh){
    // Baris atas/tengah: buka normal ke bawah.
    top=rect.bottom+gap;
    placement='bottom';
  }else if(roomAbove>=mh){
    // Baris bawah: balik ke atas tanpa mengubah ukuran menu.
    top=rect.top-gap-mh;
    placement='top';
  }else if(mh<=viewportRoom){
    // Jika menu tidak muat sepenuhnya di atas maupun di bawah tombol, jangan
    // memotongnya berdasarkan ruang di satu sisi. Geser seluruh menu ke dalam
    // viewport; popover boleh melewati posisi tombol agar semua opsi tetap utuh.
    top=Math.min(Math.max(rect.bottom+gap,pad),window.innerHeight-mh-pad);
    placement='viewport-fit';
  }else{
    // Hanya viewport yang benar-benar pendek yang memakai scroll menu.
    top=pad;
    m.style.maxHeight=viewportRoom+'px';
    m.style.overflowY='auto';
    placement='viewport-constrained';
  }

  let left=rect.right-mw;
  left=Math.max(pad,Math.min(left,window.innerWidth-mw-pad));
  m.dataset.placement=placement;
  m.style.left=Math.round(left)+'px';
  m.style.top=Math.round(top)+'px';
}

function openActionMenu(ev,idx){
  ev.stopPropagation();
  const clicked=ev.currentTarget;
  if(menuIdx===idx&&el('actionMenu').classList.contains('show')){closeActionMenu();return;}
  closeActionMenu();
  menuIdx=idx; actionQuery='';
  const current=statusMeta(dashRows[idx].chat);
  actionCategory=current.cat&&statusCats().includes(current.cat)?current.cat:statusCats()[0];
  const m=el('actionMenu');
  m.innerHTML=actionMenuShell(idx);
  m.classList.add('show');
  clicked.classList.add('menu-open');
  clicked.setAttribute('aria-expanded','true');
  renderActionStatusPane();
  positionActionMenu(clicked);
  requestAnimationFrame(()=>{const search=el('actionSearch');if(search)search.setAttribute('aria-label','Cari status atau kode');});
}
function closeActionMenu(){
  const m=el('actionMenu'); if(m)m.classList.remove('show');
  document.querySelectorAll('.dotsbtn.menu-open').forEach(b=>{b.classList.remove('menu-open');b.setAttribute('aria-expanded','false');});
  menuIdx=null; actionQuery='';
}
window.setStatus=function(idx,code){
  const previous=dashRows[idx].chat;
  if(previous===code){closeActionMenu();return;}
  dashRows[idx].chat=code; renderDash(); closeActionMenu();
  toast('Status diubah ke '+code+' · '+statusMeta(code).lbl);
};

const DICON={
 building:'<svg viewBox="0 0 24 24"><path d="M3 21h18"/><path d="M6 21V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v16"/><path d="M9 8h.01M12 8h.01M15 8h.01M9 12h.01M12 12h.01M15 12h.01"/></svg>',
 user:'<svg viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>',
 users:'<svg viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9.5" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/></svg>',
 tag:'<svg viewBox="0 0 24 24"><path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"/><line x1="7" y1="7" x2="7.01" y2="7"/></svg>',
 globe:'<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>',
 hash:'<svg viewBox="0 0 24 24"><line x1="4" y1="9" x2="20" y2="9"/><line x1="4" y1="15" x2="20" y2="15"/><line x1="10" y1="3" x2="8" y2="21"/><line x1="16" y1="3" x2="14" y2="21"/></svg>',
 phone:'<svg viewBox="0 0 24 24"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.91.34 1.85.57 2.81.7A2 2 0 0 1 22 16.92z"/></svg>',
 calendar:'<svg viewBox="0 0 24 24"><rect x="3" y="4.5" width="18" height="17" rx="2"/><line x1="16" y1="2.5" x2="16" y2="6.5"/><line x1="8" y1="2.5" x2="8" y2="6.5"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
 at:'<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M16 8v5a3 3 0 0 0 6 0v-1a10 10 0 1 0-3.92 7.94"/></svg>'
};
let detailIdx=null;
function closeDetail(){ el('overlay3').classList.remove('show'); }
window.actDetail=function(idx){
  closeActionMenu(); detailIdx=idx; const r=dashRows[idx]; const s=statusMeta(r.chat);
  const name=r.student||r.parent||'(Tanpa nama)';
  const ini=((name.replace(/[^A-Za-z ]/g,'').trim().split(/\s+/).slice(0,2).map(w=>w[0]||'').join(''))||'?').toUpperCase();
  el('dAva').textContent=ini; el('dName').textContent=name;
  el('dSub').innerHTML=`${statusPill(r.chat)}<span class="dphone">${esc(r.code||'')} ${esc(r.phone||'—')}</span>`;
  const items=[['Branch',r.branch,'building'],['Student Name',r.student,'user'],['Parent Name',r.parent,'users'],['Source',r.source,'tag'],['Country',r.country,'globe'],['Phone Code',r.code,'hash'],['Phone Number',r.phone,'phone'],['Inquiry Date',r.date,'calendar'],['Social Media',r.sosmed,'at']];
  el('dGrid').innerHTML=items.map(it=>`<div class="ditem"><div class="dic">${DICON[it[2]]||''}</div><div><div class="dlabel">${it[0]}</div><div class="dval${it[1]?'':' empty'}">${it[1]?esc(it[1]):'—'}</div></div></div>`).join('');
  el('dStatusDesc').innerHTML=`${statusPill(r.chat)}<span class="dstatustxt">${esc(s.desc)}</span>`;
  el('dNote').innerHTML=r.note?esc(r.note):'<span class="empty">Belum ada catatan.</span>';
  el('overlay3').classList.add('show');
};
window.actEdit=function(idx){ closeActionMenu(); openEditInquiry(idx); };

let deleteIdx=null;
window.openDeleteConfirm=function(idx){
  const r=dashRows[idx];
  if(!r)return;
  deleteIdx=idx;
  closeActionMenu();
  el('deleteConfirmTarget').textContent=(r.student||r.parent||'Tanpa nama')+' · '+(r.code||'')+' '+(r.phone||'');
  el('deleteConfirm').classList.add('show');
  requestAnimationFrame(()=>el('deleteNo').focus());
};
function closeDeleteConfirm(){
  el('deleteConfirm').classList.remove('show');
  deleteIdx=null;
}
function confirmDelete(){
  if(deleteIdx===null||!dashRows[deleteIdx])return closeDeleteConfirm();
  const r=dashRows[deleteIdx];
  dashRows.splice(deleteIdx,1);
  closeDeleteConfirm();
  renderDash();
  toast('Inquiry "'+(r.student||r.phone)+'" dihapus.');
}
window.openActionMenu=openActionMenu;

/* global close handlers */
document.addEventListener('click', e=>{ if(!e.target.closest('.amenu') && !e.target.closest('.dotsbtn')) closeActionMenu(); });
document.addEventListener('scroll',function(e){
  const t=e.target;
  // Scroll internal kategori/detail tidak menutup menu.
  if(t&&t.nodeType===1&&t.closest&&t.closest('.amenu,.acats,.adetails'))return;
  // Menu fixed ditutup saat halaman/tabel bergeser agar anchor tetap akurat.
  closeActionMenu();
},true);

/* wire New Inquiry modal */
el('newBtn').onclick=openNewInquiry;
el('niX').onclick=closeNewInquiry; el('niClose').onclick=closeNewInquiry;
el('niSave').onclick=saveInquiry;
el('overlay2').onclick=e=>{ if(e.target===el('overlay2')) closeNewInquiry(); };
initForm();

/* wire Detail modal */
el('dX').onclick=closeDetail; el('dClose').onclick=closeDetail;
el('overlay3').onclick=e=>{ if(e.target===el('overlay3')) closeDetail(); };
el('dEdit').onclick=()=>{ closeDetail(); openEditInquiry(detailIdx); };

/* wire Delete confirmation */
el('deleteNo').onclick=closeDeleteConfirm;
el('deleteYes').onclick=confirmDelete;
el('deleteConfirm').onclick=e=>{if(e.target===el('deleteConfirm'))closeDeleteConfirm();};
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&el('deleteConfirm').classList.contains('show'))closeDeleteConfirm();});
