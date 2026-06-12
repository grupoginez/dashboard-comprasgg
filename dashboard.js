/* ════════════════════════════════════════════════════════════
   dashboard.js — Dashboard de Compras Grupo Ginez
   ────────────────────────────────────────────────────────────
   FINANZAS: este es el ÚNICO archivo a modificar para cambios
   de visualización, nuevos indicadores o ajustes de lógica.
   No tocar connector.js ni index.html.
   ════════════════════════════════════════════════════════════ */


let ALL_DATA = {}; // populated by initDashboard() from connector.js

/* ════════════════════════════════════════════════════════════
   TRANSFORMACIÓN: RAW_BY_SHEET (connector.js) → ALL_DATA
   ════════════════════════════════════════════════════════════ */
const MES_LBL = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'};

function diffDias(d1,d2){
  if(!d1||!d2) return null;
  const a=new Date(d1+'T00:00:00'),b=new Date(d2+'T00:00:00');
  if(isNaN(a)||isNaN(b)) return null;
  return Math.round((b-a)/86400000);
}
const DENSIDAD={'KG ACIDO CLORHIDRICO':1.16,'KG HIPOCLORITO DE SODIO A 13%':1.22,'KG SOSA LIQUIDA':1.53};
const NOMBRE_DEST={'KG ACIDO CLORHIDRICO':'ACIDO CLORHIDRICO','KG HIPOCLORITO DE SODIO A 13%':'HIPOCLORITO DE SODIO A 13%','KG SOSA LIQUIDA':'SOSA LIQUIDA'};
function normUM(um){const u=String(um||'').trim().toUpperCase();if(/^KGS?$/.test(u))return'KG';if(/^(LT|LTS|LS|LIT)$/.test(u))return'LT';if(/^PZA$/.test(u))return'PZA';return'OTRO';}

function transformRecord(raw){
  const cant_oc=raw.cantidad||0, cant_ent=raw.cant_entregada||0, tieneEnt=!!raw.fecha_entrega_real;
  let periodo,anio,mes,fecha_ent;
  if(tieneEnt){periodo=raw.fecha_entrega_real.slice(0,7);anio=+raw.fecha_entrega_real.slice(0,4);mes=+raw.fecha_entrega_real.slice(5,7);fecha_ent=raw.fecha_entrega_real;}
  else{periodo='PENDIENTE';anio=raw.fecha?+raw.fecha.slice(0,4):null;mes=raw.fecha?+raw.fecha.slice(5,7):null;fecha_ent='';}
  const prod=NOMBRE_DEST[raw.producto]||raw.producto;
  const den=DENSIDAD[raw.producto];
  const cantidad=den?cant_ent/den:cant_ent;
  const cant_oc_adj=den?cant_oc/den:cant_oc;
  return{
    periodo,anio,mes,
    proveedor:raw.proveedor,producto:prod,um:normUM(raw.um),
    cantidad,cant_oc:cant_oc_adj,var_cant:cantidad-cant_oc_adj,
    cumpl_pct:cant_oc_adj>0?cantidad/cant_oc_adj*100:(cantidad>0?100:0),
    dias_ent:tieneEnt?diffDias(raw.fecha,raw.fecha_entrega_real):null,
    dias_atraso:raw.dias_atraso,fecha_oc:raw.fecha,fecha_ent,
    moneda:raw.moneda,tc:raw.tc,
    importe:raw.importe_entregado||0,total_neto:raw.total_neto,
    precio_u:raw.precio_unit||0,iva:raw.iva||0,
    precio_con_iva:raw.unitario_neto||0,
    cant_pendiente:raw.cant_pendiente,folio:raw.folio,
    entregado:tieneEnt&&cant_ent>0,estatus:raw.estatus||''
  };
}

function buildSheet(records,sheetKey){
  const UM_DEF={MP:['KG','LT'],PIGMENTOS:['KG'],FRAGANCIAS:['KG','LT'],PIPAS:['KG'],OTROS:['PZA']};
  const ps=new Set(),pvs=new Set(),prs=new Set(),ys=new Set(),ums=new Set();
  let usdS=0,usdN=0;
  records.forEach(r=>{
    if(r.periodo&&r.periodo!=='PENDIENTE')ps.add(r.periodo);
    if(r.proveedor)pvs.add(r.proveedor);if(r.producto)prs.add(r.producto);
    if(r.anio)ys.add(r.anio);if(r.um)ums.add(r.um);
    if(r.moneda==='USD'&&r.tc){usdS+=r.tc;usdN++;}
  });
  const periods=[...ps].sort();
  const periods_label={};
  periods.forEach(p=>{const[y,m]=p.split('-');periods_label[p]=(MN[+m-1]||m)+' '+y;});
  return{records,periods,periods_label,
    proveedores:[...pvs].sort(),productos:[...prs].sort(),
    years:[...ys].sort((a,b)=>a-b),
    tc_avg_global:usdN?usdS/usdN:1,
    um_present:ums.size?[...ums]:(UM_DEF[sheetKey]||['KG'])};
}

// Llamado por connector.js cuando RAW_BY_SHEET está listo
function initDashboard(){
  ALL_DATA={};
  Object.keys(RAW_BY_SHEET).forEach(sheet=>{
    const recs=(RAW_BY_SHEET[sheet]||[]).map(transformRecord);
    ALL_DATA[sheet]=buildSheet(recs,sheet);
    const tab=document.querySelector('#stab-'+sheet+' .sheet-tab-count');
    if(tab)tab.textContent=recs.length.toLocaleString('es-MX');
  });
  const chip=document.getElementById('user-chip');
  if(chip)chip.style.display='flex';
  document.getElementById('load-status').textContent='';
  rebuildFilters();
  refreshAll();
}
function reloadData(){if(typeof loadExcelData==='function')loadExcelData();}

let curSheet = "MP"; // set by initDashboard() / switchSheet()
const D = () => ALL_DATA[curSheet] || {};
const MN = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const CL = ['#5c7cfa','#20c997','#fcc419','#ff6b6b','#4dabf7','#a29bfe','#fd79a8','#74c0fc','#63e6be','#ffa94d','#e599f7','#a9e34b','#66d9e8'];
const charts = {};
let curCur = 'MXN', isDark = true, filtersFrozen = false;
const F = {yr:[], mo:[], pv:[], pr:[]};
let sortKey='qty_total', sortDir=-1, sortCycle=0;
let pvChartType='bar', evChartType='line', rankChartType='horizontal', rankImpType='stacked', rankImpAxis='horizontal';
let selPriceMats = []; // for §5b material selector

/* ── THEME ── */
function toggleTheme(){
  isDark=!isDark;
  document.documentElement.setAttribute('data-theme',isDark?'dark':'light');
  document.getElementById('sw-t').classList.toggle('on',!isDark);
  document.getElementById('sw-l').textContent=isDark?'☀️ Claro':'🌑 Oscuro';
  setTimeout(()=>refreshAll(),40);
}
const gC=()=>getComputedStyle(document.documentElement).getPropertyValue('--border').trim();
const tkC=()=>getComputedStyle(document.documentElement).getPropertyValue('--text2').trim();
const txC=()=>getComputedStyle(document.documentElement).getPropertyValue('--text').trim();
const bkC=()=>getComputedStyle(document.documentElement).getPropertyValue('--card').trim();
const dc=id=>{if(charts[id]){charts[id].destroy();delete charts[id]}};

/* ── FREEZE ── */
function toggleFreeze(){
  filtersFrozen = !filtersFrozen;
  const bar = document.getElementById('fbar');
  const btn = document.getElementById('freeze-btn');
  // fbar is always sticky; frozen just adds visual emphasis
  bar.classList.toggle('frozen', filtersFrozen);
  btn.classList.toggle('active', filtersFrozen);
  btn.textContent = filtersFrozen ? '📌 Fijado' : '📌 Fijar filtros';
  btn.title = filtersFrozen
    ? 'Filtros fijados — los filtros siempre están visibles al hacer scroll'
    : 'Los filtros ya están fijados permanentemente';
  // chips also sticky below fbar
  const chips = document.getElementById('chips');
  if(chips){
    const fbarH = bar.offsetHeight;
    chips.style.cssText = `position:sticky;top:${96+fbarH}px;z-index:188;`;
  }
}

/* ── SECTION TOGGLE ── */
function toggleSec(key){
  const body=document.getElementById('body-'+key);
  const tog=document.getElementById('tog-'+key);
  const hidden=body.classList.toggle('hidden');
  tog.classList.toggle('open',!hidden);
  tog.classList.toggle('closed',hidden);
}

/* ── CURRENCY ── */
function setCur(c){
  curCur=c;
  ['MXN','USD','ALL'].forEach(x=>document.getElementById('cbtn-'+x).classList.toggle('active',x===c));
  refreshAll();
}
/* For display: when MXN show only MXN records; USD show only USD; ALL show both in original currency */
function importeOf(r){
  // returns importe in original currency (no conversion)
  return r.importe;
}
const impLbl=()=>curCur==='ALL'?'Importe (orig.)':curCur;
const fmtN=(v,cur)=>{
  if(v==null) return '—';
  const c=cur||curCur; const sym=c==='USD'?'US$':'$';
  const abs=Math.abs(+v);
  const s=abs>=1e9?(abs/1e9).toFixed(2)+'B':abs>=1e6?(abs/1e6).toFixed(2)+'M':abs>=1e3?(abs/1e3).toFixed(1)+'K':abs.toLocaleString('es-MX',{minimumFractionDigits:0,maximumFractionDigits:0});
  return (+v<0?'-':'')+sym+s;
};
const fmt=v=>v==null?'—':Math.abs(+v)>=1e6?(+v/1e6).toFixed(2)+'M':Math.abs(+v)>=1e3?(+v/1e3).toFixed(1)+'K':(+v).toLocaleString('es-MX',{maximumFractionDigits:2});
const fmtT=v=>v==null?'—':(+v/1000).toFixed(3)+' T'; // KG → Toneladas
const fmtPct=v=>v==null?'—':((+v)>=0?'+':'')+Number(v).toFixed(1)+'%';
const pctD=(a,b)=>b===0?null:(a-b)/b*100;

/* ── MULTI-SELECT ── */
function buildMS(key,opts,ph){
  const dd=document.getElementById('msd-'+key);
  dd._opts=opts; dd._ph=ph;
  dd.innerHTML=`<div class="msd-s"><input type="text" placeholder="Buscar..." oninput="searchOpts('${key}',this.value)"></div><div class="msd-l" id="msl-${key}"></div><div class="msd-f"><button class="btn-sm" onclick="closeDD('${key}')">Cerrar</button></div>`;
  renderOpts(key,opts,'');
}
function renderOpts(key,opts,q){
  const list=document.getElementById('msl-'+key);
  if(!list) return;
  const sel=(key==='pm'?selPriceMats:F[key])||[];
  const filt=opts.filter(o=>o.lbl.toLowerCase().includes(q.toLowerCase()));
  const allSel=sel.length===0;
  list.innerHTML='';
  const aDiv=document.createElement('div'); aDiv.className='mso mso-all'+(allSel?' sel':'');
  const aCb=document.createElement('input'); aCb.type='checkbox'; aCb.id='msall-'+key; aCb.checked=allSel;
  aCb.addEventListener('change',function(){ selAll(key,aCb.checked); });
  const aLbl=document.createElement('label'); aLbl.className='mso-l'; aLbl.htmlFor='msall-'+key;
  aLbl.innerHTML='<b>Seleccionar todos</b>'; aDiv.appendChild(aCb); aDiv.appendChild(aLbl); list.appendChild(aDiv);
  filt.forEach(function(o){
    const s=sel.includes(o.val);
    const div=document.createElement('div'); div.className='mso'+(s?' sel':'');
    const cb=document.createElement('input'); cb.type='checkbox'; cb.checked=s;
    cb.addEventListener('change',function(){ togOpt(key,o.val,cb.checked); });
    const sp=document.createElement('span'); sp.className='mso-l'; sp.title=o.lbl; sp.textContent=o.lbl;
    div.appendChild(cb); div.appendChild(sp); list.appendChild(div);
  });
}
function searchOpts(key,q){ const dd=document.getElementById('msd-'+key); if(dd&&dd._opts) renderOpts(key,dd._opts,q); }
function togOpt(key,val,checked){
  const arr=key==='pm'?selPriceMats:F[key];
  if(checked){ if(!arr.includes(val)) arr.push(val); } else { const i=arr.indexOf(val); if(i>-1) arr.splice(i,1); }
  document.getElementById('msall-'+key).checked=arr.length===0;
  updateTrig(key);
  if(key==='pm') renderPrices(filtered()); else refreshAll();
}
function selAll(key,checked){
  if(key==='pm'){ selPriceMats.length=0; } else F[key]=[];
  const dd=document.getElementById('msd-'+key); if(dd&&dd._opts) renderOpts(key,dd._opts,'');
  updateTrig(key);
  if(key==='pm') renderPrices(filtered()); else refreshAll();
}
function updateTrig(key){
  const arr=key==='pm'?selPriceMats:F[key]; const n=arr.length;
  const tx=document.getElementById('mtx-'+key); const ct=document.getElementById('mct-'+key);
  const ph={yr:'Todos',mo:'Todos',pv:'Todos',pr:'Todas',pm:'Top 5 materias'};
  if(n===0){ if(tx) tx.textContent=ph[key]||'Todos'; if(ct) ct.style.display='none'; }
  else if(n===1){ if(tx) tx.textContent=arr[0]; if(ct) ct.style.display='none'; }
  else { if(tx) tx.textContent=n+' seleccionados'; if(ct){ct.textContent=n;ct.style.display='';} }
}
function toggleDD(key){
  const dd=document.getElementById('msd-'+key); const trig=document.getElementById('mst-'+key);
  const isOpen=dd.style.display!=='none';
  ['yr','mo','pv','pr','pm'].forEach(k=>{ const d=document.getElementById('msd-'+k); if(d) d.style.display='none'; const t=document.getElementById('mst-'+k); if(t) t.classList.remove('open'); });
  if(!isOpen){ dd.style.display='block'; if(trig) trig.classList.add('open'); }
}
function closeDD(key){ const d=document.getElementById('msd-'+key); if(d) d.style.display='none'; const t=document.getElementById('mst-'+key); if(t) t.classList.remove('open'); }
document.addEventListener('click',e=>{ ['yr','mo','pv','pr','pm'].forEach(k=>{ const r=document.getElementById('msr-'+k); if(r&&!r.contains(e.target)) closeDD(k); }); });
function resetAll(){
  ['yr','mo','pv','pr'].forEach(k=>{ F[k]=[]; });
  rebuildFilters();
  refreshAll();
}

/* ── CHIPS ── */
function renderChips(){
  let h='';
  F.yr.forEach(y=>h+=`<div class="chip cy">📅 ${y} <span class="chipx" onclick="rmChip('yr','${y}')">✕</span></div>`);
  F.mo.forEach(m=>h+=`<div class="chip cy">📆 ${MN[+m-1]} <span class="chipx" onclick="rmChip('mo','${m}')">✕</span></div>`);
  F.pv.forEach(p=>h+=`<div class="chip cp">🏭 ${p} <span class="chipx" onclick="rmChip('pv','${encodeURIComponent(p)}')">✕</span></div>`);
  F.pr.forEach(p=>h+=`<div class="chip cm">📦 ${p} <span class="chipx" onclick="rmChip('pr','${encodeURIComponent(p)}')">✕</span></div>`);
  document.getElementById('chips').innerHTML=h;
  const n=F.yr.length+F.mo.length+F.pv.length+F.pr.length;
  document.getElementById('fbadge').classList.toggle('show',n>0);
  if(n) document.getElementById('fbadge-n').textContent=n;
}
function rmChip(key,val){ const v=decodeURIComponent(val); F[key]=F[key].filter(x=>x!==v); updateTrig(key); const dd=document.getElementById('msd-'+key); if(dd&&dd._opts) renderOpts(key,dd._opts,''); refreshAll(); }

/* ── FILTER ── */
function filtered(){
  return D().records.filter(r=>{
    if(F.yr.length&&!F.yr.includes(String(r.anio))) return false;
    if(F.mo.length&&!F.mo.includes(String(r.mes))) return false;
    if(F.pv.length&&!F.pv.includes(r.proveedor)) return false;
    if(F.pr.length&&!F.pr.includes(r.producto)) return false;
    // currency filter: MXN→only MXN records, USD→only USD, ALL→all
    if(curCur==='MXN'&&r.moneda!=='MXN') return false;
    if(curCur==='USD'&&r.moneda!=='USD') return false;
    return true;
  });
}

/* ── REFRESH ALL ── */
function refreshAll(){
  // Update header subtitle with current sheet
  const meta = {MP:'🧪 Materias Primas',PIGMENTOS:'🎨 Pigmentos',FRAGANCIAS:'🌸 Fragancias',PIPAS:'🛢️ Pipas',OTROS:'📦 Otros'};
  const subEl = document.querySelector('.hdr-sub');
  if(subEl) subEl.textContent = (meta[curSheet]||curSheet) + ' · Solo COMPLETO · IMPORTE DE MAT. ENTREGADO';
  const recs=filtered();
  renderChips();
  renderKPIs(recs);
  renderRanking(recs);
  renderMinMax(recs);
  renderProviders(recs);
  renderEvolProv(recs);
  renderEvolProd(recs);
  renderPrices(recs);
  renderOrdenes(recs);
  renderVariacion(recs);
  renderTiempos(recs);
  renderEconomico(recs);
}

/* ── §0 KPIs ── */
function renderKPIs(recs){
  const tKg=recs.filter(r=>r.um==='KG').reduce((a,r)=>a+r.cantidad,0);
  const tTon=tKg/1000;
  const tLt=recs.filter(r=>r.um==='LT').reduce((a,r)=>a+r.cantidad,0);
  const tPza=recs.filter(r=>r.um==='PZA').reduce((a,r)=>a+r.cantidad,0);
  const tImp=recs.reduce((a,r)=>a+r.importe,0);
  const nM=12;
  const nP=new Set(recs.map(r=>r.proveedor)).size;
  const nPr=new Set(recs.map(r=>r.producto)).size;
  const vld=recs.filter(r=>r.dias_ent!=null&&r.dias_ent>=0&&r.dias_ent<1000);
  const avgD=vld.length?vld.reduce((a,r)=>a+r.dias_ent,0)/vld.length:0;
  const ums = D().um_present || [];
  const hasKG=ums.includes('KG'), hasLT=ums.includes('LT'), hasPZA=ums.includes('PZA');
  document.getElementById('kpis').innerHTML=`
    ${hasKG?`<div class="kpi b"><div class="kpi-l">Toneladas Entregadas</div><div class="kpi-v kgc">${tTon.toFixed(3)} T</div><div class="kpi-s">÷12 meses</div></div>`:''}
    ${hasLT?`<div class="kpi g"><div class="kpi-l">LTS Entregados</div><div class="kpi-v ltc">${fmt(tLt)}</div></div>`:''}
    ${hasPZA?`<div class="kpi b"><div class="kpi-l">PZA Entregadas</div><div class="kpi-v kgc">${fmt(tPza)}</div></div>`:''}
    <div class="kpi y"><div class="kpi-l">Importe (${impLbl()})</div><div class="kpi-v nc">${fmtN(tImp)}</div><div class="kpi-s">Mat. Entregado</div></div>
    <div class="kpi"><div class="kpi-l">Prom/Mes (÷12)</div><div class="kpi-v kgc">${fmt(tKg/nM)}</div></div>
    <div class="kpi r"><div class="kpi-l">Prom/Mes LTS (÷12)</div><div class="kpi-v ltc">${fmt(tLt/nM)}</div></div>
    <div class="kpi"><div class="kpi-l">Provs / Materias</div><div class="kpi-v">${nP} / ${nPr}</div></div>
    <div class="kpi"><div class="kpi-l">Días Entrega Prom.</div><div class="kpi-v">${avgD.toFixed(0)}</div><div class="kpi-s">días</div></div>`;
}

/* ── §1 RANKING ── */
let _rd=[];
let rankImpAxisMode='horizontal';
function onSortChange(){ sortKey=document.getElementById('rsort-key').value; renderRankTable(_rd); renderRankCharts(_rd); }
function toggleSortDir(){ sortDir=-sortDir; document.getElementById('rsort-dir').textContent=sortDir===-1?'↓ Mayor→Menor':'↑ Menor→Mayor'; renderRankTable(_rd); renderRankCharts(_rd); }
function cycleRankChartType(){ rankChartType=rankChartType==='horizontal'?'vertical':'horizontal'; document.getElementById('rcht').textContent=rankChartType==='horizontal'?'Horizontal':'Vertical'; renderRankCharts(_rd); }
function cycleRankImpType(){ rankImpType=rankImpType==='stacked'?'grouped':'stacked'; document.getElementById('rimp-t').textContent=rankImpType==='stacked'?'Apilado':'Agrupado'; renderRankCharts(_rd); }
function cycleRankImpAxis(){ rankImpAxisMode=rankImpAxisMode==='horizontal'?'vertical':'horizontal'; document.getElementById('rimp-a').textContent=rankImpAxisMode==='horizontal'?'Horizontal':'Vertical'; renderRankCharts(_rd); }
function setSortKey(key){ if(sortKey===key){sortDir=-sortDir;document.getElementById('rsort-dir').textContent=sortDir===-1?'↓ Mayor→Menor':'↑ Menor→Mayor';}else{sortKey=key;document.getElementById('rsort-key').value=key;} renderRankTable(_rd); renderRankCharts(_rd); }

function renderRanking(recs){
  const map={};
  recs.forEach(r=>{
    if(!map[r.producto]) map[r.producto]={kg:0,lt:0,pza:0,ot:0,imp_mxn:0,imp_usd:0,ps:new Set()};
    if(r.um==='KG') map[r.producto].kg+=r.cantidad;
    else if(r.um==='LT') map[r.producto].lt+=r.cantidad;
    else if(r.um==='PZA') map[r.producto].pza+=r.cantidad;
    else map[r.producto].ot+=r.cantidad;
    if(r.moneda==='MXN') map[r.producto].imp_mxn+=r.importe;
    else map[r.producto].imp_usd+=r.importe;
    map[r.producto].ps.add(r.periodo);
  });
  _rd=Object.entries(map).map(([k,v])=>{
    const tot=v.kg+v.lt+v.ot, ton=v.kg/1000;
    const importe=curCur==='USD'?v.imp_usd:curCur==='MXN'?v.imp_mxn:(v.imp_mxn+v.imp_usd);
    return{name:k,kg:v.kg,ton,lt:v.lt,pza:v.pza,qty_total:tot,importe,imp_mxn:v.imp_mxn,imp_usd:v.imp_usd,avg_q:tot/12,avg_n:importe/12};
  });
  renderRankTable(_rd); renderRankCharts(_rd);
}
function renderRankTable(arr){
  const sorted=[...arr].sort((a,b)=>sortDir*(b[sortKey]-a[sortKey]));
  const maxV=sorted[0]?.[sortKey]||1;
  const ums2=D().um_present||[];
  const cols=[
    {key:'qty_total',lbl:'Total Vol.'},
    ...(ums2.includes('KG')?[{key:'kg',lbl:'KG'},{key:'ton',lbl:'Toneladas'}]:[]),
    ...(ums2.includes('LT')?[{key:'lt',lbl:'LTS'}]:[]),
    ...(ums2.includes('PZA')?[{key:'pza',lbl:'PZA'}]:[]),
    {key:'importe',lbl:`Importe ${impLbl()}`},
    {key:'avg_q',lbl:'P/Mes Vol.'},{key:'avg_n',lbl:'P/Mes $'}
  ];
  let h=`<thead><tr><th class="tl">#</th><th class="tl" onclick="setSortKey('qty_total')" style="cursor:pointer">Materia</th>${cols.map(c=>`<th onclick="setSortKey('${c.key}')" style="cursor:pointer;white-space:nowrap">${c.lbl}${sortKey===c.key?(sortDir===-1?' ▼':' ▲'):' ⇅'}</th>`).join('')}</tr></thead><tbody>`;
  // Start ranking from 1 (top is index 0)
  sorted.forEach((r,i)=>{
    if(r[sortKey]===0&&i>0) return; // skip true zeros after rank 1 in vertical
    const bw=Math.max(2,Math.round(Math.abs(r[sortKey])/maxV*45));
    h+=`<tr><td class="tl"><span class="rank">${i+1}</span></td><td class="tl" title="${r.name}">${r.name}<span class="bi" style="width:${bw}px;background:${CL[i%CL.length]}"></span></td><td>${fmt(r.qty_total)}</td>${ums2.includes('KG')?`<td class="kgc">${fmt(r.kg)}</td><td class="kgc">${r.ton.toFixed(3)} T</td>`:''} ${ums2.includes('LT')?`<td class="ltc">${fmt(r.lt)}</td>`:''} ${ums2.includes('PZA')?`<td class="kgc">${fmt(r.pza)}</td>`:''}<td class="nc">${fmtN(r.importe)}</td><td>${fmt(r.avg_q)}</td><td class="nc">${fmtN(r.avg_n)}</td></tr>`;
  });
  const tK=sorted.reduce((a,r)=>a+r.kg,0),tL=sorted.reduce((a,r)=>a+r.lt,0),tI=sorted.reduce((a,r)=>a+r.importe,0);
  h+=`<tr class="tfr"><td></td><td class="tl">TOTAL</td><td>${fmt(tK+tL)}</td><td class="kgc">${fmt(tK)}</td><td class="kgc">${(tK/1000).toFixed(3)} T</td><td class="ltc">${fmt(tL)}</td><td class="nc">${fmtN(tI)}</td><td>—</td><td>—</td></tr></tbody>`;
  document.getElementById('tbl-rank').innerHTML=h;
}
function renderRankCharts(arr){
  const sorted=[...arr].sort((a,b)=>sortDir*(b[sortKey]-a[sortKey]));
  // top 12 excluding zeros
  const top12=sorted.filter(r=>r[sortKey]>0).slice(0,12);
  const isH=rankChartType==='horizontal';
  const labels=top12.map(r=>r.name.length>20?r.name.slice(0,18)+'…':r.name);
  dc('ch-rvol');
  charts['ch-rvol']=new Chart(document.getElementById('ch-rvol').getContext('2d'),{type:'bar',data:{labels,datasets:[
    {label:'KG',data:top12.map(r=>r.kg),backgroundColor:'#5c7cfa99',borderColor:'#5c7cfacc',borderWidth:1,borderRadius:3},
    {label:'LTS',data:top12.map(r=>r.lt),backgroundColor:'#20c99799',borderColor:'#20c997cc',borderWidth:1,borderRadius:3},
  ]},options:{responsive:true,maintainAspectRatio:false,indexAxis:isH?'y':'x',plugins:{legend:{labels:{color:tkC(),font:{size:11}}}},
    scales:{x:{grid:{color:gC()},ticks:{color:tkC(),callback:v=>fmt(v),font:{size:isH?10:9},maxRotation:isH?0:40}},y:{grid:{color:gC()},ticks:{color:txC(),font:{size:isH?10:9}}}}}});
  // Importe chart — MXN and USD in original currency, no conversion
  const isStk=rankImpType==='stacked', isIH=rankImpAxisMode==='horizontal';
  const dsImp=[
    {label:'🟡 MXN original',data:top12.map(r=>r.imp_mxn),backgroundColor:'#fcc41999',borderColor:'#fcc419cc',borderWidth:1,borderRadius:3,stack:isStk?'imp':undefined},
    {label:'🔵 USD original',data:top12.map(r=>r.imp_usd),backgroundColor:'#4dabf799',borderColor:'#4dabf7cc',borderWidth:1,borderRadius:3,stack:isStk?'imp':undefined},
  ];
  dc('ch-rimp');
  charts['ch-rimp']=new Chart(document.getElementById('ch-rimp').getContext('2d'),{type:'bar',data:{labels,datasets:dsImp},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:isIH?'y':'x',plugins:{legend:{labels:{color:tkC(),font:{size:11},boxWidth:12}},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${c.dataset.label.includes('USD')?fmtN(c.raw,'USD'):fmtN(c.raw,'MXN')}`}}},
      scales:{x:{stacked:isStk,grid:{color:gC()},ticks:{color:tkC(),callback:v=>v>0?'$'+fmt(v):v,font:{size:isIH?10:9},maxRotation:isIH?0:40}},y:{stacked:isStk,grid:{color:gC()},ticks:{color:txC(),font:{size:isIH?10:9}}}}}});
}

/* ── §2 MAX/MIN ── */
function renderMinMax(recs){
  const map={};
  recs.forEach(r=>{
    if(!map[r.producto]) map[r.producto]={};
    if(!map[r.producto][r.periodo]) map[r.producto][r.periodo]={qty:0,prov:r.proveedor,fecha:r.fecha_ent};
    map[r.producto][r.periodo].qty+=r.cantidad;
  });
  const rows=[];
  Object.entries(map).forEach(([p,ps])=>{
    const a=Object.entries(ps).map(([per,v])=>({per,...v}));
    if(!a.length) return;
    a.sort((x,y)=>y.qty-x.qty);
    rows.push({prod:p,mx:a[0],mn:a[a.length-1]});
  });
  rows.sort((a,b)=>b.mx.qty-a.mx.qty);
  let h=`<thead><tr><th class="tl">Materia</th><th>Máx. Vol.</th><th>Periodo Máx.</th><th>Proveedor Máx.</th><th>Mín. Vol.</th><th>Periodo Mín.</th><th>Proveedor Mín.</th></tr></thead><tbody>`;
  rows.forEach(r=>{h+=`<tr><td class="tl" title="${r.prod}">${r.prod}</td><td style="color:var(--acc2);font-weight:700">${fmt(r.mx.qty)}</td><td>${r.mx.per}</td><td>${r.mx.prov}</td><td style="color:var(--acc3);font-weight:700">${fmt(r.mn.qty)}</td><td>${r.mn.per}</td><td>${r.mn.prov}</td></tr>`;});
  document.getElementById('tbl-mm').innerHTML=h+'</tbody>';
}

/* ── §3 PROVIDERS ── */
function cyclePvType(){ pvChartType=pvChartType==='bar'?'stacked':'bar'; document.getElementById('pvt').textContent=pvChartType==='bar'?'Barras':'Apiladas'; renderProviders(filtered()); }
function renderProviders(recs){
  const map={};
  recs.forEach(r=>{ if(!map[r.proveedor]) map[r.proveedor]={kg:0,lt:0,imp:0}; if(r.um==='KG') map[r.proveedor].kg+=r.cantidad; else if(r.um==='LT') map[r.proveedor].lt+=r.cantidad; map[r.proveedor].imp+=r.importe; });
  const arr=Object.entries(map).map(([k,v])=>({name:k,...v,tot:v.kg+v.lt})).sort((a,b)=>b.tot-a.tot);
  const top15=arr.slice(0,15), gT=arr.reduce((a,r)=>a+r.tot,0);
  dc('ch-pvbar');
  charts['ch-pvbar']=new Chart(document.getElementById('ch-pvbar').getContext('2d'),{type:'bar',data:{labels:top15.map(r=>r.name),datasets:[
    {label:'KG',data:top15.map(r=>r.kg),backgroundColor:'#5c7cfa99',borderRadius:3},
    {label:'LTS',data:top15.map(r=>r.lt),backgroundColor:'#20c99799',borderRadius:3}
  ]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{labels:{color:tkC()}},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${fmt(c.raw)} (${gT?(c.raw/gT*100).toFixed(1):0}%)`}}},
    scales:{x:{stacked:pvChartType==='stacked',grid:{color:gC()},ticks:{color:tkC(),callback:v=>fmt(v)}},y:{stacked:pvChartType==='stacked',grid:{color:gC()},ticks:{color:txC(),font:{size:10}}}}}});
  const top8=arr.slice(0,8), ot=arr.slice(8).reduce((a,r)=>a+r.tot,0);
  dc('ch-pvdnt');
  charts['ch-pvdnt']=new Chart(document.getElementById('ch-pvdnt').getContext('2d'),{type:'doughnut',
    data:{labels:[...top8.map(r=>r.name),...(ot>0?['Otros']:[])],datasets:[{data:[...top8.map(r=>r.tot),...(ot>0?[ot]:[])],backgroundColor:[...CL.slice(0,8),'#6c757d'],borderColor:bkC(),borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{color:txC(),font:{size:10},boxWidth:11}},tooltip:{callbacks:{label:c=>` ${c.label}: ${fmt(c.raw)} (${gT?(c.raw/gT*100).toFixed(1):0}%)`}}}}});
  const tm=document.getElementById('treemap'); tm.innerHTML='';
  const mx=arr[0]?.tot||1;
  arr.forEach((r,i)=>{
    const w=Math.max(40,Math.round(r.tot/mx*200)),h=Math.max(34,Math.round(r.tot/mx*100));
    const el=document.createElement('div'); el.className='tm-cell';
    el.style.cssText=`width:${w}px;height:${h}px;background:${CL[i%CL.length]}`;
    el.title=`${r.name}: ${fmt(r.tot)} (${gT?(r.tot/gT*100).toFixed(1):0}%)`;
    el.textContent=r.name.length>14?r.name.slice(0,12)+'…':r.name;
    tm.appendChild(el);
  });
}

/* ── §4 EVOL PROV ── */
function cycleEvType(){ evChartType=evChartType==='line'?'bar':'line'; document.getElementById('evt').textContent=evChartType==='line'?'Línea':'Barras'; renderEvolProv(filtered()); }
function renderEvolProv(recs){
  const tot={}; recs.forEach(r=>{tot[r.proveedor]=(tot[r.proveedor]||0)+r.cantidad});
  const top8=Object.entries(tot).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([k])=>k);
  const yrs=[...new Set(recs.map(r=>r.anio))].sort();
  dc('ch-evpv');
  charts['ch-evpv']=new Chart(document.getElementById('ch-evpv').getContext('2d'),{type:evChartType,data:{labels:yrs,datasets:top8.map((p,i)=>({label:p,data:yrs.map(y=>recs.filter(r=>r.proveedor===p&&r.anio===y).reduce((a,r)=>a+r.cantidad,0)),borderColor:CL[i],backgroundColor:CL[i]+(evChartType==='line'?'22':'99'),fill:false,tension:.3,borderWidth:2,pointRadius:3,borderRadius:3}))},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:tkC(),font:{size:10},boxWidth:12}}},scales:{x:{grid:{color:gC()},ticks:{color:tkC()}},y:{grid:{color:gC()},ticks:{color:tkC(),callback:v=>fmt(v)}}}}});
  let h=`<thead><tr><th class="tl">Proveedor</th>${yrs.map(y=>`<th>${y}</th>`).join('')}${yrs.slice(1).map(y=>`<th>Δ ${y}</th>`).join('')}</tr></thead><tbody>`;
  top8.forEach((p,i)=>{
    const vs=yrs.map(y=>recs.filter(r=>r.proveedor===p&&r.anio===y).reduce((a,r)=>a+r.cantidad,0));
    const ds=vs.slice(1).map((v,j)=>pctD(v,vs[j]));
    h+=`<tr><td class="tl" style="color:${CL[i]};font-weight:700">${p}</td>${vs.map(v=>`<td>${fmt(v)}</td>`).join('')}${ds.map(d=>d==null?'<td>—</td>':`<td class="${d>=0?'gp':'gn'}">${d>=0?'▲':'▼'}${Math.abs(d).toFixed(1)}%</td>`).join('')}</tr>`;
  });
  document.getElementById('tbl-pvyr').innerHTML=h+'</tbody>';
}

/* ── §5 EVOL PROD ── */
function renderEvolProd(recs){
  const tot={}; recs.forEach(r=>{tot[r.producto]=(tot[r.producto]||0)+r.cantidad});
  const top10=Object.entries(tot).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([k])=>k);
  const yrs=[...new Set(recs.map(r=>r.anio))].sort();
  dc('ch-evpr');
  charts['ch-evpr']=new Chart(document.getElementById('ch-evpr').getContext('2d'),{type:'line',data:{labels:yrs,datasets:top10.map((p,i)=>({label:p,data:yrs.map(y=>recs.filter(r=>r.producto===p&&r.anio===y).reduce((a,r)=>a+r.cantidad,0)),borderColor:CL[i],backgroundColor:CL[i]+'22',fill:false,tension:.35,borderWidth:2,pointRadius:3}))},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:tkC(),font:{size:9},boxWidth:10}}},scales:{x:{grid:{color:gC()},ticks:{color:tkC()}},y:{grid:{color:gC()},ticks:{color:tkC(),callback:v=>fmt(v)}}}}});
  let h=`<thead><tr><th class="tl">Materia</th>${yrs.map(y=>`<th>${y}</th>`).join('')}${yrs.slice(1).map(y=>`<th>Δ ${y}</th>`).join('')}</tr></thead><tbody>`;
  top10.forEach((p,i)=>{
    const vs=yrs.map(y=>recs.filter(r=>r.producto===p&&r.anio===y).reduce((a,r)=>a+r.cantidad,0));
    const ds=vs.slice(1).map((v,j)=>pctD(v,vs[j]));
    h+=`<tr><td class="tl" style="color:${CL[i]};font-weight:700">${p}</td>${vs.map(v=>`<td>${fmt(v)}</td>`).join('')}${ds.map(d=>d==null?'<td>—</td>':`<td class="${d>=0?'gp':'gn'}">${d>=0?'▲':'▼'}${Math.abs(d).toFixed(1)}%</td>`).join('')}</tr>`;
  });
  document.getElementById('tbl-pryr').innerHTML=h+'</tbody>';
}

/* ── §5b PRICES ── */
function renderPrices(recs){
  // Build material selector options (top 10 by volume)
  const tot={}; recs.forEach(r=>{tot[r.producto]=(tot[r.producto]||0)+r.cantidad});
  const top10prods=Object.entries(tot).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([k])=>k);
  const dd=document.getElementById('msd-pm');
  const allProds=D().productos;
  dd._opts=allProds.map(p=>({val:p,lbl:p}));
  renderOpts('pm',dd._opts,'');
  // determine which materials to show
  const showMats=selPriceMats.length>0?selPriceMats:top10prods.slice(0,5);

  // Line chart: monthly price evolution per material per moneda
  const periods=D().periods;
  const datasets=[];
  showMats.forEach((prod,i)=>{
    // MXN dataset
    const mxnVals=periods.map(p=>{
      const rs=recs.filter(r=>r.producto===prod&&r.periodo===p&&r.moneda==='MXN'&&r.precio_u>0);
      if(!rs.length) return null;
      const totalQty=rs.reduce((a,r)=>a+r.cantidad,0);
      return totalQty>0?rs.reduce((a,r)=>a+r.precio_u*r.cantidad,0)/totalQty:null;
    });
    const hasAnyMXN=mxnVals.some(v=>v!=null);
    if(hasAnyMXN){
      datasets.push({label:`${prod} (MXN)`,data:periods.map((p,j)=>({x:D().periods_label[p]||p,y:mxnVals[j]})),borderColor:'#fcc419',backgroundColor:'#fcc41933',fill:false,tension:.3,borderWidth:2,pointRadius:3,borderDash:[],spanGaps:true});
    }
    // USD dataset
    const usdVals=periods.map(p=>{
      const rs=recs.filter(r=>r.producto===prod&&r.periodo===p&&r.moneda==='USD'&&r.precio_u>0);
      if(!rs.length) return null;
      const totalQty=rs.reduce((a,r)=>a+r.cantidad,0);
      return totalQty>0?rs.reduce((a,r)=>a+r.precio_u*r.cantidad,0)/totalQty:null;
    });
    const hasAnyUSD=usdVals.some(v=>v!=null);
    if(hasAnyUSD){
      datasets.push({label:`${prod} (USD)`,data:periods.map((p,j)=>({x:D().periods_label[p]||p,y:usdVals[j]})),borderColor:'#4dabf7',backgroundColor:'#4dabf733',fill:false,tension:.3,borderWidth:2,pointRadius:3,borderDash:[4,3],spanGaps:true});
    }
  });

  dc('ch-price');
  charts['ch-price']=new Chart(document.getElementById('ch-price').getContext('2d'),{type:'line',
    data:{datasets},
    options:{responsive:true,maintainAspectRatio:false,
      parsing:{xAxisKey:'x',yAxisKey:'y'},
      plugins:{legend:{labels:{color:tkC(),font:{size:9},boxWidth:10}},
        tooltip:{callbacks:{label:c=>`${c.dataset.label}: $${Number(c.parsed.y).toFixed(4)}`}}},
      scales:{x:{grid:{color:gC()},ticks:{color:tkC(),maxRotation:45,font:{size:9}}},y:{grid:{color:gC()},ticks:{color:tkC(),callback:v=>'$'+fmt(v)}}}}});

  // Stats table: min, max, avg, variation% for each material
  let h=`<thead><tr><th class="tl">Materia</th><th>Moneda</th><th>Precio Mín.</th><th>Precio Máx.</th><th>Prom. Pond.</th><th>Ú​lt. Precio</th><th>Var. vs 1er precio</th><th>N Registros</th></tr></thead><tbody>`;
  const statMats=selPriceMats.length>0?selPriceMats:top10prods;
  statMats.forEach((prod,i)=>{
    ['MXN','USD'].forEach(cur=>{
      const rs=recs.filter(r=>r.producto===prod&&r.moneda===cur&&r.precio_u>0);
      if(!rs.length) return;
      rs.sort((a,b)=>a.periodo.localeCompare(b.periodo));
      const prices=rs.map(r=>r.precio_u);
      const mn=Math.min(...prices), mx2=Math.max(...prices);
      const totalQty=rs.reduce((a,r)=>a+r.cantidad,0);
      const wavg=totalQty>0?rs.reduce((a,r)=>a+r.precio_u*r.cantidad,0)/totalQty:0;
      const first=prices[0], last=prices[prices.length-1];
      const vari=first>0?(last-first)/first*100:null;
      const sym=cur==='USD'?'US$':'$';
      h+=`<tr><td class="tl" title="${prod}">${prod}</td><td><span style="color:${cur==='USD'?'var(--acc5)':'var(--acc4)'};font-weight:700">${cur}</span></td><td>${sym}${mn.toFixed(4)}</td><td>${sym}${mx2.toFixed(4)}</td><td>${sym}${wavg.toFixed(4)}</td><td>${sym}${last.toFixed(4)}</td><td class="${vari==null?'':vari>=0?'gp':'gn'}">${vari==null?'—':(vari>=0?'▲':'▼')+Math.abs(vari).toFixed(1)+'%'}</td><td>${rs.length}</td></tr>`;
    });
  });
  document.getElementById('tbl-price').innerHTML=h+'</tbody>';

  // Annual price variation table
  const yrs=[...new Set(recs.map(r=>r.anio))].sort();
  let h2=`<thead><tr><th class="tl">Materia</th><th>Moneda</th>${yrs.map(y=>`<th>Prom ${y}</th><th>Δ ${y}</th>`).join('')}</tr></thead><tbody>`;
  statMats.forEach(prod=>{
    ['MXN','USD'].forEach(cur=>{
      const yearData=yrs.map(y=>{
        const rs=recs.filter(r=>r.producto===prod&&r.moneda===cur&&r.anio===y&&r.precio_u>0);
        if(!rs.length) return null;
        const tq=rs.reduce((a,r)=>a+r.cantidad,0);
        return tq>0?rs.reduce((a,r)=>a+r.precio_u*r.cantidad,0)/tq:null;
      });
      if(yearData.every(v=>v==null)) return;
      const sym=cur==='USD'?'US$':'$';
      let row=`<tr><td class="tl" title="${prod}">${prod}</td><td><span style="color:${cur==='USD'?'var(--acc5)':'var(--acc4)'};font-weight:700">${cur}</span></td>`;
      yearData.forEach((v,j)=>{
        const prev=yearData[j-1];
        const d=v!=null&&prev!=null?pctD(v,prev):null;
        row+=`<td>${v!=null?sym+v.toFixed(4):'—'}</td>`;
        row+=`<td class="${d==null?'':d>=0?'gp':'gn'}">${d==null?'—':(d>=0?'▲':'▼')+Math.abs(d).toFixed(1)+'%'}</td>`;
      });
      h2+=row+'</tr>';
    });
  });
  document.getElementById('tbl-price-yr').innerHTML=h2+'</tbody>';
}

/* ── §6 ÓRDENES ── */
function renderOrdenes(recs){
  const pm={};
  recs.forEach(r=>{ if(!pm[r.periodo]) pm[r.periodo]={f:new Set(),n:0}; if(r.folio&&r.folio!=='nan') pm[r.periodo].f.add(r.folio); pm[r.periodo].n+=r.importe; });
  const ps=Object.keys(pm).sort();
  const fv=ps.map(p=>pm[p].f.size);
  const tot=fv.reduce((a,b)=>a+b,0),mx=Math.max(...fv,0),mn2=Math.min(...fv,0),av=tot/12;
  document.getElementById('kpis-fo').innerHTML=`
    <div class="kpi"><div class="kpi-l">Total Órdenes</div><div class="kpi-v">${tot}</div></div>
    <div class="kpi g"><div class="kpi-l">Prom/Mes (÷12)</div><div class="kpi-v">${av.toFixed(1)}</div></div>
    <div class="kpi y"><div class="kpi-l">Máximo</div><div class="kpi-v">${mx}</div><div class="kpi-s">${ps[fv.indexOf(mx)]||'—'}</div></div>
    <div class="kpi r"><div class="kpi-l">Mínimo</div><div class="kpi-v">${mn2}</div></div>`;
  dc('ch-fo');
  charts['ch-fo']=new Chart(document.getElementById('ch-fo').getContext('2d'),{type:'bar',
    data:{labels:ps.map(p=>D().periods_label[p]||p),datasets:[
      {label:'Órdenes',data:fv,backgroundColor:'#5c7cfa88',borderRadius:3,yAxisID:'y'},
      {label:`Importe ${impLbl()}`,data:ps.map(p=>pm[p].n),type:'line',borderColor:'#fcc419',backgroundColor:'transparent',borderWidth:2,pointRadius:2,yAxisID:'y2'}
    ]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:tkC()}}},
      scales:{x:{grid:{color:gC()},ticks:{color:tkC(),maxRotation:45,font:{size:9}}},y:{grid:{color:gC()},ticks:{color:tkC()},position:'left'},y2:{grid:{drawOnChartArea:false},ticks:{color:'#fcc419',callback:v=>fmtN(v)},position:'right'}}}});
  const tN=ps.reduce((a,p)=>a+pm[p].n,0);
  let h=`<thead><tr><th class="tl">Periodo</th><th>Órdenes</th><th>Importe ${impLbl()}</th><th>Prom/Orden</th></tr></thead><tbody>`;
  ps.forEach(p=>{ const nf=pm[p].f.size; h+=`<tr><td class="tl">${D().periods_label[p]||p}</td><td>${nf}</td><td class="nc">${fmtN(pm[p].n)}</td><td>${fmtN(pm[p].n/(nf||1))}</td></tr>`; });
  h+=`<tr class="tfr"><td class="tl">TOTAL</td><td>${tot}</td><td class="nc">${fmtN(tN)}</td><td>${fmtN(tN/(tot||1))}</td></tr></tbody>`;
  document.getElementById('tbl-fo').innerHTML=h;
}

/* ── §7 VARIACIÓN ── */
function renderVariacion(recs){
  const vld=recs.filter(r=>r.cant_oc>0);
  let comp=0,sobre=0,parc=0,sin=0;
  vld.forEach(r=>{ const p=r.cumpl_pct; if(r.cantidad===0) sin++; else if(p>100) sobre++; else if(p===100) comp++; else parc++; });
  const avgC=vld.length?vld.reduce((a,r)=>a+r.cumpl_pct,0)/vld.length:0;
  document.getElementById('kpis-var').innerHTML=`
    <div class="kpi g"><div class="kpi-l">Completa (=100%)</div><div class="kpi-v">${comp}</div><div class="kpi-s">${vld.length?(comp/vld.length*100).toFixed(0):0}%</div></div>
    <div class="kpi b"><div class="kpi-l">Sobre entrega (>100%)</div><div class="kpi-v">${sobre}</div></div>
    <div class="kpi y"><div class="kpi-l">Parcial (<100%)</div><div class="kpi-v">${parc}</div></div>
    <div class="kpi r"><div class="kpi-l">Sin entrega (0%)</div><div class="kpi-v">${sin}</div></div>
    <div class="kpi"><div class="kpi-l">Cumplimiento Prom.</div><div class="kpi-v">${avgC.toFixed(1)}%</div></div>`;
  const pm={};
  vld.forEach(r=>{ if(!pm[r.proveedor]) pm[r.proveedor]={oc:0,ent:0,n:0}; pm[r.proveedor].oc+=r.cant_oc; pm[r.proveedor].ent+=r.cantidad; pm[r.proveedor].n+=r.importe; });
  const parr=Object.entries(pm).map(([k,v])=>({name:k,...v,cumpl:v.oc>0?v.ent/v.oc*100:0})).sort((a,b)=>a.cumpl-b.cumpl);
  dc('ch-vcumpl');
  charts['ch-vcumpl']=new Chart(document.getElementById('ch-vcumpl').getContext('2d'),{type:'bar',
    data:{labels:parr.map(r=>r.name),datasets:[{label:'Cumplimiento %',data:parr.map(r=>+r.cumpl.toFixed(1)),backgroundColor:parr.map(r=>r.cumpl>=100?'#20c99799':r.cumpl>=70?'#fcc41999':'#ff6b6b99'),borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.raw}% — OC:${fmt(parr[c.dataIndex].oc)} Ent:${fmt(parr[c.dataIndex].ent)}`}}},
      scales:{x:{grid:{color:gC()},ticks:{color:tkC(),callback:v=>v+'%'},min:0,max:Math.max(110,...parr.map(r=>r.cumpl))*1.02},y:{grid:{color:gC()},ticks:{color:txC(),font:{size:10}}}}}});
  dc('ch-vpie');
  charts['ch-vpie']=new Chart(document.getElementById('ch-vpie').getContext('2d'),{type:'doughnut',
    data:{labels:['Completa','Sobre entrega','Parcial','Sin entrega'],datasets:[{data:[comp,sobre,parc,sin],backgroundColor:['#20c99788','#5c7cfa88','#fcc41988','#ff6b6b88'],borderColor:['#20c997','#5c7cfa','#fcc419','#ff6b6b'],borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right',labels:{color:txC(),font:{size:11}}},tooltip:{callbacks:{label:c=>`${c.label}: ${c.raw} (${vld.length?(c.raw/vld.length*100).toFixed(1):0}%)`}}}}});
  let h=`<thead><tr><th class="tl">Proveedor</th><th>OC</th><th>Entregado</th><th>Diferencia</th><th>Cumplimiento</th><th>Importe</th></tr></thead><tbody>`;
  parr.sort((a,b)=>b.oc-a.oc).forEach(r=>{
    const cl=r.cumpl===0?'p-no':r.cumpl>=100&&r.cumpl<=100?'p-ok':r.cumpl>100?'p-ov':'p-pa';
    h+=`<tr><td class="tl">${r.name}</td><td>${fmt(r.oc)}</td><td>${fmt(r.ent)}</td><td class="${r.ent>=r.oc?'gp':'gn'}">${r.ent>=r.oc?'+':''}${fmt(r.ent-r.oc)}</td><td><span class="pill ${cl}">${r.cumpl.toFixed(1)}%</span></td><td class="nc">${fmtN(r.n)}</td></tr>`;
  });
  document.getElementById('tbl-vsum').innerHTML=h+'</tbody>';
  let dh=`<thead><tr><th class="tl">Folio</th><th class="tl">Materia</th><th class="tl">Proveedor</th><th>Cant. OC</th><th>Cant. Ent.</th><th>Dif.</th><th>Cumpl %</th><th>Fecha OC</th><th>Fecha Ent.</th><th>Días</th><th>Moneda</th><th>TC</th><th>Importe</th></tr></thead><tbody>`;
  vld.slice(0,600).forEach(r=>{
    const cl=r.cumpl_pct===0?'p-no':r.cumpl_pct>=100&&r.cumpl_pct<=100?'p-ok':r.cumpl_pct>100?'p-ov':'p-pa';
    dh+=`<tr><td class="tl">${r.folio}</td><td class="tl" title="${r.producto}">${r.producto}</td><td class="tl">${r.proveedor}</td><td>${fmt(r.cant_oc)}</td><td>${fmt(r.cantidad)}</td><td class="${r.cantidad>=r.cant_oc?'gp':'gn'}">${r.cantidad>=r.cant_oc?'+':''}${fmt(r.var_cant)}</td><td><span class="pill ${cl}">${r.cumpl_pct.toFixed(1)}%</span></td><td>${r.fecha_oc}</td><td>${r.fecha_ent}</td><td>${r.dias_ent??'—'}</td><td>${r.moneda}</td><td>${r.tc}</td><td class="nc">${fmtN(r.importe)}</td></tr>`;
  });
  document.getElementById('tbl-vdet').innerHTML=dh+'</tbody>';
}
function toggleDet(){
  const p=document.getElementById('det-pane'), btn=document.getElementById('det-btn');
  const open=p.classList.toggle('open');
  btn.textContent=open?'▼ Ocultar detalle':'▶ Ver detalle de OC';
}

/* ── §8 TIEMPOS ── */
function renderTiempos(recs){
  const vld=recs.filter(r=>r.dias_ent!=null&&r.dias_ent>=0&&r.dias_ent<1000);
  const dias=vld.map(r=>r.dias_ent);
  const av=dias.length?dias.reduce((a,b)=>a+b,0)/dias.length:0;
  const mx=dias.length?Math.max(...dias):0, mn=dias.length?Math.min(...dias):0;
  document.getElementById('kpis-tm').innerHTML=`
    <div class="kpi"><div class="kpi-l">Prom. Días</div><div class="kpi-v">${av.toFixed(0)}</div></div>
    <div class="kpi r"><div class="kpi-l">Máximo</div><div class="kpi-v">${mx}</div><div class="kpi-s">días</div></div>
    <div class="kpi g"><div class="kpi-l">Mínimo</div><div class="kpi-v">${mn}</div><div class="kpi-s">días</div></div>
    <div class="kpi b"><div class="kpi-l">Registros válidos</div><div class="kpi-v">${vld.length}</div></div>`;
  const bins=[0,7,14,30,45,60,90,120,180,366];
  const counts=new Array(bins.length-1).fill(0);
  dias.forEach(d=>{ for(let i=0;i<bins.length-1;i++) if(d>=bins[i]&&d<bins[i+1]){counts[i]++;break;} });
  dc('ch-hist');
  charts['ch-hist']=new Chart(document.getElementById('ch-hist').getContext('2d'),{type:'bar',data:{labels:bins.slice(0,-1).map((b,i)=>`${b}–${bins[i+1]-1}d`),datasets:[{label:'Entregas',data:counts,backgroundColor:'#5c7cfa99',borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:gC()},ticks:{color:tkC(),maxRotation:45,font:{size:9}}},y:{grid:{color:gC()},ticks:{color:tkC()}}}}});
  const pm={};
  vld.forEach(r=>{ if(!pm[r.proveedor]) pm[r.proveedor]=[]; pm[r.proveedor].push(r.dias_ent); });
  const pa=Object.entries(pm).map(([k,v])=>({name:k,avg:v.reduce((a,b)=>a+b,0)/v.length})).sort((a,b)=>b.avg-a.avg).slice(0,15);
  dc('ch-tpv');
  charts['ch-tpv']=new Chart(document.getElementById('ch-tpv').getContext('2d'),{type:'bar',data:{labels:pa.map(r=>r.name),datasets:[{label:'Días prom.',data:pa.map(r=>+r.avg.toFixed(1)),backgroundColor:pa.map(r=>r.avg>45?'#ff6b6b99':r.avg>14?'#fcc41999':'#20c99799'),borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{grid:{color:gC()},ticks:{color:tkC()}},y:{grid:{color:gC()},ticks:{color:txC(),font:{size:10}}}}}});
  const pm2={}; vld.forEach(r=>{ if(!pm2[r.periodo]) pm2[r.periodo]={s:0,c:0}; pm2[r.periodo].s+=r.dias_ent; pm2[r.periodo].c++; });
  const ps=Object.keys(pm2).sort();
  dc('ch-ttrend');
  charts['ch-ttrend']=new Chart(document.getElementById('ch-ttrend').getContext('2d'),{type:'line',data:{labels:ps.map(p=>D().periods_label[p]||p),datasets:[{label:'Días prom.',data:ps.map(p=>(pm2[p].s/pm2[p].c).toFixed(1)),borderColor:'#5c7cfa',backgroundColor:'#5c7cfa22',fill:true,tension:.3,borderWidth:2,pointRadius:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:gC()},ticks:{color:tkC(),maxRotation:45,font:{size:9}}},y:{grid:{color:gC()},ticks:{color:tkC()}}}}});
}

/* ── §9 ECONÓMICO ── */
function renderEconomico(recs){
  const rMXN=recs.filter(r=>r.moneda==='MXN'), rUSD=recs.filter(r=>r.moneda==='USD');
  const tMXN=rMXN.reduce((a,r)=>a+r.importe,0);
  const tUSD=rUSD.reduce((a,r)=>a+r.importe,0);
  const tcAvg=rUSD.length?rUSD.reduce((a,r)=>a+r.tc,0)/rUSD.length:D().tc_avg_global;
  document.getElementById('kpis-eco').innerHTML=`
    <div class="kpi y"><div class="kpi-l">Total MXN (original)</div><div class="kpi-v nc">${fmtN(tMXN,'MXN')}</div><div class="kpi-s">${rMXN.length} registros</div></div>
    <div class="kpi b"><div class="kpi-l">Total USD (original)</div><div class="kpi-v" style="color:var(--acc5)">${fmtN(tUSD,'USD')}</div><div class="kpi-s">${rUSD.length} registros</div></div>
    <div class="kpi"><div class="kpi-l">TC Prom. (solicitud)</div><div class="kpi-v">${tcAvg.toFixed(4)}</div><div class="kpi-s">referencia</div></div>`;
  document.getElementById('tc-info').innerHTML=`<strong style="color:var(--text)">💱 Nota de moneda:</strong> Todos los importes se muestran en su <strong>moneda original</strong> sin conversión. Fuente TC: columna TIPO DE CAMBIO (SOLICITUD). TC promedio de registros USD en filtro actual: <strong style="color:var(--acc4)">${tcAvg.toFixed(4)}</strong>. Use el filtro de moneda (MXN/USD/Ambas) para ver cada conjunto por separado.`;
  // by product
  const pmap={};
  recs.forEach(r=>{ if(!pmap[r.producto]) pmap[r.producto]={mxn:0,usd:0,n_m:0,n_u:0}; if(r.moneda==='MXN'){pmap[r.producto].mxn+=r.importe;pmap[r.producto].n_m++;}else{pmap[r.producto].usd+=r.importe;pmap[r.producto].n_u++;} });
  const parr=Object.entries(pmap).map(([k,v])=>({name:k,...v,tot_ref:v.mxn+v.usd})).sort((a,b)=>b.tot_ref-a.tot_ref).slice(0,12);
  dc('ch-ecbar');
  charts['ch-ecbar']=new Chart(document.getElementById('ch-ecbar').getContext('2d'),{type:'bar',
    data:{labels:parr.map(r=>r.name),datasets:[{label:'MXN orig.',data:parr.map(r=>r.mxn),backgroundColor:'#fcc41988',borderRadius:3},{label:'USD orig.',data:parr.map(r=>r.usd),backgroundColor:'#4dabf788',borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{labels:{color:tkC()}}},
      scales:{x:{stacked:true,grid:{color:gC()},ticks:{color:tkC(),callback:v=>'$'+fmt(v)}},y:{stacked:true,grid:{color:gC()},ticks:{color:txC(),font:{size:9}}}}}});
  const tm={};
  recs.forEach(r=>{ if(!tm[r.periodo]) tm[r.periodo]={mxn:0,usd:0}; if(r.moneda==='MXN') tm[r.periodo].mxn+=r.importe; else tm[r.periodo].usd+=r.importe; });
  const ps=Object.keys(tm).sort();
  dc('ch-ectrend');
  charts['ch-ectrend']=new Chart(document.getElementById('ch-ectrend').getContext('2d'),{type:'line',
    data:{labels:ps.map(p=>D().periods_label[p]||p),datasets:[{label:'MXN',data:ps.map(p=>tm[p].mxn),borderColor:'#fcc419',backgroundColor:'#fcc41922',fill:true,tension:.3,borderWidth:2,pointRadius:2,yAxisID:'y'},{label:'USD',data:ps.map(p=>tm[p].usd),borderColor:'#4dabf7',backgroundColor:'transparent',fill:false,tension:.3,borderWidth:2,pointRadius:2,borderDash:[4,3],yAxisID:'y'}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:tkC()}}},scales:{x:{grid:{color:gC()},ticks:{color:tkC(),maxRotation:45,font:{size:9}}},y:{grid:{color:gC()},ticks:{color:tkC(),callback:v=>'$'+fmt(v)}}}}});
  let h=`<thead><tr><th class="tl">Materia</th><th>MXN orig.</th><th>Reg. MXN</th><th>USD orig.</th><th>Reg. USD</th><th style="font-size:9px">Monedas</th></tr></thead><tbody>`;
  parr.forEach(r=>{
    const dual=r.n_m>0&&r.n_u>0;
    h+=`<tr${dual?' style="background:rgba(92,124,250,.06)"':''}><td class="tl" title="${r.name}">${r.name}${dual?' <span class="pill" style="background:rgba(92,124,250,.15);color:var(--accent);font-size:8px">2 monedas</span>':''}</td><td class="nc">${fmtN(r.mxn,'MXN')}</td><td>${r.n_m}</td><td style="color:var(--acc5)">${fmtN(r.usd,'USD')}</td><td>${r.n_u}</td><td style="font-size:9px;color:var(--text3)">${dual?'MXN + USD':r.n_m>0?'MXN':'USD'}</td></tr>`;
  });
  document.getElementById('tbl-eco').innerHTML=h+'</tbody>';
}

/* ── EXPORT ── */
function expTbl(id){ const t=document.getElementById(id); if(!t) return; const rows=[...t.querySelectorAll('tr')].map(tr=>[...tr.querySelectorAll('th,td')].map(c=>'"'+(c.textContent.trim().replace(/"/g,'""'))+'"').join(',')); const a=document.createElement('a'); a.href='data:text/csv;charset=utf-8,\uFEFF'+encodeURIComponent(rows.join('\n')); a.download=`MP_${id}_${new Date().toISOString().slice(0,10)}.csv`; a.click(); }
function exportCSV(){ const recs=filtered(); const hdrs=['periodo','anio','mes','proveedor','producto','um','cantidad','cant_oc','importe','total_neto','moneda','tc','folio','var_cant','cumpl_pct','dias_ent','fecha_oc','fecha_ent','precio_u','cant_pendiente']; const rows=[hdrs.join(','),...recs.map(r=>hdrs.map(h=>JSON.stringify(r[h]??'')).join(','))]; const a=document.createElement('a'); a.href='data:text/csv;charset=utf-8,\uFEFF'+encodeURIComponent(rows.join('\n')); a.download=`MP_datos_${new Date().toISOString().slice(0,10)}.csv`; a.click(); }


/* ── QUICK NAV ── */
function toggleQNav(){
  const nav  = document.getElementById('qnav');
  const trig = document.getElementById('qnav-trigger');
  const open = nav.classList.toggle('open');
  trig.classList.toggle('open', open);
}
function navTo(id){
  // Close menu
  document.getElementById('qnav').classList.remove('open');
  document.getElementById('qnav-trigger').classList.remove('open');
  // Open section if hidden
  const body = document.getElementById('body-' + id.replace('wrap-',''));
  if(body && body.classList.contains('hidden')) toggleSec(id.replace('wrap-',''));
  // Smooth scroll
  setTimeout(()=>{ const el=document.getElementById(id); if(el) el.scrollIntoView({behavior:'smooth',block:'start'}); }, 60);
  // Mark active
  document.querySelectorAll('.qnav-item').forEach(i=>i.classList.remove('active'));
  const btn = document.querySelector('.qnav-item[onclick*="' + id + '"]');
  if(btn) btn.classList.add('active');
}
// Close on outside click
document.addEventListener('click', e=>{
  const w=document.getElementById('qnav-wrap');
  if(w && !w.contains(e.target)){
    document.getElementById('qnav')?.classList.remove('open');
    document.getElementById('qnav-trigger')?.classList.remove('open');
  }
});
// Auto-highlight active section while scrolling
const _navObs = new IntersectionObserver(entries=>{
  entries.forEach(e=>{
    if(e.isIntersecting){
      document.querySelectorAll('.qnav-item').forEach(i=>i.classList.remove('active'));
      const btn=document.querySelector('.qnav-item[onclick*="'+e.target.id+'"]');
      if(btn) btn.classList.add('active');
    }
  });
},{threshold:0.15,rootMargin:'-60px 0px -55% 0px'});
['wrap-kpi','wrap-s1','wrap-s2','wrap-s3','wrap-s4','wrap-s5','wrap-s5b','wrap-s6','wrap-s7','wrap-s8','wrap-s9'].forEach(id=>{
  const el=document.getElementById(id); if(el) _navObs.observe(el);
});


/* ── SHEET SWITCHER ── */
function switchSheet(sheet){
  curSheet = sheet;
  // Update tab active state
  document.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
  document.getElementById('stab-' + sheet)?.classList.add('active');
  // Rebuild independent filters for this sheet
  rebuildFilters();
  refreshAll();
}

function rebuildFilters(){
  const d = D();
  // Reset filter state for new sheet
  ['yr','mo','pv','pr'].forEach(k => { F[k]=[]; });
  // Rebuild dropdowns
  buildMS('yr', d.years.map(y=>({val:String(y),lbl:String(y)})), 'Todos');
  buildMS('mo', MN.map((m,i)=>({val:String(i+1),lbl:m})), 'Todos');
  buildMS('pv', d.proveedores.map(p=>({val:p,lbl:p})), 'Todos');
  buildMS('pr', d.productos.map(p=>({val:p,lbl:p})), 'Todas');
  ['yr','mo','pv','pr'].forEach(k => updateTrig(k));
  // Rebuild price mat selector
  const pmDd = document.getElementById('msd-pm');
  if(pmDd){ pmDd._opts=d.productos.map(p=>({val:p,lbl:p})); }
  selPriceMats.length = 0;
  updateTrig('pm');
}
