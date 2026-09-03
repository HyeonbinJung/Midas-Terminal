const MIDAS_WEB_VERSION = "5.1.0";
const $ = (sel, root=document) => root.querySelector(sel);
const $$ = (sel, root=document) => [...root.querySelectorAll(sel)];
let currentRegion = 'KR';
let reportInFlight = false;
let researchInFlight = false;
let newsInFlight = false;
let marketInFlight = false;

function toast(message, error=false){
  const el=$('#toast'); el.textContent=message; el.classList.toggle('error',error); el.classList.add('show');
  clearTimeout(window.__toastTimer); window.__toastTimer=setTimeout(()=>el.classList.remove('show'),3600);
}
function log(msg, type='SYS'){
  const box=$('#command-log'); if(!box) return;
  const p=document.createElement('p'); p.innerHTML=`<time>${type}</time> ${escapeHtml(msg)}`; box.prepend(p);
}
function escapeHtml(s=''){ return String(s).replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
function money(v,currency=''){ if(v==null||Number.isNaN(Number(v))) return '—'; const n=Number(v),a=Math.abs(n); let s=a>=1e12?(a/1e12).toFixed(2)+'T':a>=1e9?(a/1e9).toFixed(2)+'B':a>=1e6?(a/1e6).toFixed(1)+'M':a.toLocaleString(undefined,{maximumFractionDigits:0}); return (n<0?'-':'')+s+(currency?' '+currency:''); }
function num(v,d=2){ return v==null||Number.isNaN(Number(v))?'—':Number(v).toLocaleString(undefined,{minimumFractionDigits:d,maximumFractionDigits:d}); }
function pct(v,d=2,ratio=false){ if(v==null||Number.isNaN(Number(v))) return '—'; return (Number(v)*(ratio?100:1)).toFixed(d)+'%'; }
function numericClass(v){ return v!=null&&!Number.isNaN(Number(v))&&Number(v)<0?'negative-number':''; }
function growthClass(v){ if(v==null||Number.isNaN(Number(v))) return ''; const n=Number(v); return n<0?'negative-number':n>=0.20?'growth-hot':''; }
function signedTextClass(text=''){ return String(text).trim().startsWith('-')?'negative-number':''; }

function switchView(name){
  $$('.view').forEach(v=>v.classList.remove('active')); $$('.fn').forEach(b=>b.classList.remove('active'));
  $(`#${name}-view`)?.classList.add('active'); $(`.fn[data-view="${name}"]`)?.classList.add('active');
  // Mobile browsers can retain horizontal scroll from wide research tables. Reset it
  // whenever the user changes F1/F2/F3/F4 so controls never appear "off-screen".
  window.scrollTo({top:0,left:0,behavior:'auto'});
  document.documentElement.scrollLeft=0; document.body.scrollLeft=0;
}
$$('.fn').forEach(btn=>btn.addEventListener('click',()=>switchView(btn.dataset.view)));
document.addEventListener('keydown', e=>{
  const map={F1:'terminal',F2:'research',F3:'news',F4:'about'};
  if(map[e.key]){e.preventDefault();switchView(map[e.key]);}
});

async function api(url, options={}){
  const res=await fetch(url, {...options, cache:'no-store'}); let data={};
  try{data=await res.json()}catch{}
  if(!res.ok){
    const retry=res.headers.get('Retry-After');
    const message=data.detail||`HTTP ${res.status}`;
    throw new Error(retry?`${message} Retry after ${retry}s.`:message);
  }
  return data;
}
function setStatus(text){ $('#footer-status').textContent=text; }
function setButtonBusy(el, busy, busyText='WORKING…'){
  if(!el) return;
  if(busy){ if(!el.dataset.originalText) el.dataset.originalText=el.innerHTML; el.disabled=true; el.classList.add('busy'); el.textContent=busyText; }
  else { el.disabled=false; el.classList.remove('busy'); if(el.dataset.originalText){ el.innerHTML=el.dataset.originalText; delete el.dataset.originalText; } }
}
function setReportButtonsBusy(busy){
  setButtonBusy($('#quick-report'),busy,'BUILDING DOCX…');
  setButtonBusy($('#download-report'),busy,'EXPORTING…');
}

async function loadMarket(ticker){
  ticker=(ticker||$('#ticker-input').value).trim().toUpperCase(); if(!ticker||marketInFlight) return;
  marketInFlight=true; setButtonBusy($('#load-market'),true,'…');
  $('#ticker-input').value=ticker; $('#research-ticker').value=ticker; $('#quote-status').textContent='LOADING'; setStatus(`LOADING ${ticker}`); log(`Requesting market snapshot for ${ticker}.`,'REQ');
  try{
    const d=await api(`/api/market/${encodeURIComponent(ticker)}`);
    $('#company-name').textContent=`${d.company_name} (${d.ticker})`;
    $('#security-meta').textContent=[d.exchange,d.sector,d.industry].filter(Boolean).join('  ·  ')||'—';
    $('#currency-label').textContent=d.currency||'—'; $('#current-price').textContent=num(d.current_price,2);
    const ch=$('#price-change'); ch.textContent=d.change_pct==null?'—':`${d.change_pct>=0?'+':''}${d.change_pct.toFixed(2)}%`; ch.className=d.change_pct==null?'':d.change_pct>=0?'positive':'negative';
    $('#market-cap').textContent=money(d.market_cap,d.currency); $('#high-52w').textContent=num(d.high_52w); $('#low-52w').textContent=num(d.low_52w); $('#beta').textContent=num(d.beta,2); $('#div-yield').textContent=pct(d.dividend_yield_pct,2); $('#avg-value').textContent=money(d.avg_daily_value,d.currency);
    $('#target-high').textContent=num(d.target_high); $('#target-mean').textContent=num(d.target_mean); $('#target-median').textContent=num(d.target_median); $('#target-low').textContent=num(d.target_low);
    $('#quote-status').textContent='LIVE'; setStatus(`${ticker} LOADED`); log(`${ticker} snapshot loaded.`,'OK');
  }catch(err){ $('#quote-status').textContent='ERROR'; setStatus('REQUEST ERROR'); toast(err.message,true); log(err.message,'ERR'); }
  finally{ marketInFlight=false; setButtonBusy($('#load-market'),false); }
}


function financialCellClass(cell, threshold){
  if(!cell) return '';
  const raw=cell.raw;
  if(raw!=null&&!Number.isNaN(Number(raw))&&Number(raw)<0) return 'negative-number';
  if(cell.kind==='growth_pct'&&raw!=null&&!Number.isNaN(Number(raw))&&Number(raw)>=Number(threshold||20)) return 'growth-hot';
  return '';
}
function renderFinancialTable(section, headSelector, bodySelector, threshold){
  const head=$(headSelector), body=$(bodySelector);
  if(!head||!body) return;
  const years=section?.years||[];
  const rows=section?.rows||[];
  head.innerHTML=`<tr><th>ITEM</th>${years.map(y=>`<th>${escapeHtml(y)}</th>`).join('')}</tr>`;
  body.innerHTML='';
  rows.forEach(row=>{
    const tr=document.createElement('tr');
    const values=row.values||[];
    tr.innerHTML=`<td>${escapeHtml(row.label||'')}</td>${years.map((_,i)=>{ const cell=values[i]||{}; const display=cell.display||'—'; return `<td class="${financialCellClass(cell,threshold)}">${escapeHtml(display)}</td>`; }).join('')}`;
    body.appendChild(tr);
  });
  if(!rows.length) body.innerHTML=`<tr><td colspan="${Math.max(2,years.length+1)}" class="placeholder">No financial data available.</td></tr>`;
}
function renderKeyFinancials(financials){
  const threshold=financials?.revenue_growth_highlight_pct??20;
  const income=financials?.income||{years:[],rows:[]};
  const balance=financials?.balance||{years:[],rows:[]};
  renderFinancialTable(income,'#income-financial-head','#income-financial-body',threshold);
  renderFinancialTable(balance,'#balance-financial-head','#balance-financial-body',threshold);
  const unit=$('#financial-units');
  const allRows=[...(income.rows||[]),...(balance.rows||[])];
  const populated=allRows.reduce((count,row)=>count+(row.values||[]).filter(cell=>cell&&cell.display).length,0);
  if(unit){
    const currency=financials?.currency||'';
    unit.textContent=`${currency}${currency?' · ':''}MM / PER-SHARE / X · ${populated} CELLS`;
    unit.classList.toggle('partial',populated===0);
  }
  const panel=$('.key-financials-panel');
  if(panel) panel.dataset.loaded=populated>0?'true':'empty';
}

async function runResearch(){
  const ticker=$('#research-ticker').value.trim().toUpperCase(); if(!ticker||researchInFlight) return;
  researchInFlight=true; setButtonBusy($('#run-research'),true,'SCANNING…');
  $('#research-status').textContent='SCANNING'; setStatus(`RESEARCH ${ticker}`); log(`Running filing/background/peer research for ${ticker}.`,'REQ');
  const warningBox=$('#research-warnings'); warningBox.hidden=true; warningBox.innerHTML='';
  try{
    const d=await api(`/api/research/${encodeURIComponent(ticker)}`);
    Object.entries(d.background||{}).forEach(([k,v])=>{ const el=$(`[data-bg="${k}"]`); if(el) el.textContent=v||'No public-source text extracted.'; });
    renderKeyFinancials(d.key_financials||{}); log(`Key financials payload loaded for ${ticker}.`,'DATA');
    const peers=$('#peer-list'); peers.innerHTML='';
    (d.peers||[]).forEach((p,i)=>{
      const node=document.createElement('div'); node.className='peer-card'; node.innerHTML=`<div class="peer-top"><strong>${escapeHtml(p.ticker)}</strong><span>PEER ${i+1}</span></div><div class="peer-name">${escapeHtml(p.company_name||'')}</div><div class="peer-stats"><div>MKT CAP<b class="${numericClass(p.market_cap)}">${money(p.market_cap)}</b></div><div>BETA<b class="${numericClass(p.beta)}">${num(p.beta,2)}</b></div><div>REV GROWTH<b class="${growthClass(p.revenue_growth)}">${pct(p.revenue_growth,1,true)}</b></div><div>EBITDA MGN<b class="${numericClass(p.ebitda_margin)}">${pct(p.ebitda_margin,1,true)}</b></div></div>`; peers.appendChild(node);
    });
    if(!(d.peers||[]).length) peers.innerHTML='<p class="placeholder">Automatic peer selection returned no eligible peers.</p>';

    const highlightList=$('#highlight-list'); highlightList.innerHTML='';
    (d.highlights||[]).forEach((h,i)=>{
      const tr=document.createElement('tr');
      tr.innerHTML=`<td>${i+1}</td><td>${escapeHtml(h.argument||'')}</td><td>${escapeHtml(h.evidence_source||'')}</td>`;
      highlightList.appendChild(tr);
    });
    if(!(d.highlights||[]).length) highlightList.innerHTML='<tr><td colspan="3" class="placeholder">No fact-based highlights generated.</td></tr>';

    const catalystList=$('#catalyst-list'); catalystList.innerHTML='';
    (d.catalysts||[]).forEach(c=>{
      const tr=document.createElement('tr');
      tr.innerHTML=`<td>${escapeHtml(c.catalyst||'')}</td><td>${escapeHtml(c.expected_date||'')}</td><td>${escapeHtml(c.probability||'')}</td><td class="${signedTextClass(c.price_impact)}">${escapeHtml(c.price_impact||'')}</td><td>${escapeHtml(c.verify||'')}</td>`;
      catalystList.appendChild(tr);
    });
    if(!(d.catalysts||[]).length) catalystList.innerHTML='<tr><td colspan="5" class="placeholder">No catalyst candidates generated.</td></tr>';

    const riskList=$('#risk-list'); riskList.innerHTML='';
    (d.risks||[]).forEach(r=>{
      const tr=document.createElement('tr');
      const likelihoodClass=String(r.likelihood||'').toLowerCase();
      const severityClass=String(r.severity||'').toLowerCase();
      tr.innerHTML=`<td>${escapeHtml(r.risk||'')}</td><td><span class="risk-level ${escapeHtml(likelihoodClass)}">${escapeHtml(r.likelihood||'')}</span></td><td><span class="risk-level ${escapeHtml(severityClass)}">${escapeHtml(r.severity||'')}</span></td><td class="${signedTextClass(r.tp_impact)}">${escapeHtml(r.tp_impact||'')}</td><td>${escapeHtml(r.mitigation||'')}</td>`;
      riskList.appendChild(tr);
    });
    if(!(d.risks||[]).length) riskList.innerHTML='<tr><td colspan="5" class="placeholder">No risk factors generated.</td></tr>';

    const sources=$('#source-list'); sources.innerHTML=''; const s=d.sources||{};
    if(s.annual_url){ const a=document.createElement('div'); a.className='source-item'; a.innerHTML=`ANNUAL · ${escapeHtml(s.annual_form||'FILING')} · ${escapeHtml(s.annual_filing_date||'')}<br><a href="${s.annual_url}" target="_blank" rel="noopener">${escapeHtml(s.annual_url)}</a>`; sources.appendChild(a); }
    if(s.proxy_url){ const a=document.createElement('div'); a.className='source-item'; a.innerHTML=`PROXY · ${escapeHtml(s.proxy_form||'DEF 14A')} · ${escapeHtml(s.proxy_filing_date||'')}<br><a href="${s.proxy_url}" target="_blank" rel="noopener">${escapeHtml(s.proxy_url)}</a>`; sources.appendChild(a); }
    if(s.event_url){ const a=document.createElement('div'); a.className='source-item'; a.innerHTML=`EVENT · ${escapeHtml(s.event_form||'8-K')} · ${escapeHtml(s.event_filing_date||'')}<br><a href="${s.event_url}" target="_blank" rel="noopener">${escapeHtml(s.event_url)}</a>`; sources.appendChild(a); }
    if(!sources.children.length) sources.innerHTML='<p class="placeholder">No SEC filing reference available for this security.</p>';

    const warnings=d.warnings||[];
    if(warnings.length){
      warningBox.hidden=false; warningBox.innerHTML=warnings.map(w=>`<div>WARN · ${escapeHtml(w)}</div>`).join('');
      $('#research-status').textContent='PARTIAL'; $('#research-status').className='partial'; setStatus(`${ticker} RESEARCH PARTIAL`);
      warnings.forEach(w=>log(w,'WARN')); toast(`${ticker} research completed with ${warnings.length} warning(s).`);
    }else{
      $('#research-status').textContent='COMPLETE'; $('#research-status').className=''; setStatus(`${ticker} RESEARCH COMPLETE`);
    }
    log(`${ticker} research complete; ${(d.peers||[]).length} peers selected.`,'OK');
  }catch(err){ $('#research-status').textContent='ERROR'; $('#research-status').className='error-text'; setStatus('RESEARCH ERROR'); toast(err.message,true); log(err.message,'ERR'); }
  finally{ researchInFlight=false; setButtonBusy($('#run-research'),false); }
}

async function loadNews(){
  if(newsInFlight) return; newsInFlight=true; setButtonBusy($('#refresh-news'),true,'LOADING…');
  $('#news-grid').innerHTML='<div class="empty-state">LOADING NEWS FEED…</div>'; $('#news-region').textContent=`${currentRegion} FEED`; setStatus(`LOADING ${currentRegion} NEWS`);
  try{
    const d=await api(`/api/news/${currentRegion}`); $('#news-count').textContent=`${d.total} ITEMS`; const grid=$('#news-grid'); grid.innerHTML='';
    (d.categories||[]).forEach(cat=>{
      const section=document.createElement('section'); section.className='news-category'; const head=document.createElement('h3'); head.textContent=cat.name; section.appendChild(head);
      (cat.items||[]).slice(0,12).forEach(item=>{
        const div=document.createElement('div'); div.className='article'; const dt=new Date(item.published_at); div.innerHTML=`<div class="article-meta"><b>${escapeHtml(item.keyword)}</b>${dt.toLocaleDateString()}<br>${dt.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})}</div><div><a href="${item.url}" target="_blank" rel="noopener">${escapeHtml(item.title)}</a><small>${escapeHtml(item.source)}</small></div>`; section.appendChild(div);
      }); grid.appendChild(section);
    });
    if(!grid.children.length) grid.innerHTML='<div class="empty-state">No recent results.</div>'; setStatus(`${currentRegion} NEWS LOADED`); log(`${currentRegion} news clipping loaded: ${d.total} items.`,'OK');
  }catch(err){ $('#news-grid').innerHTML=`<div class="empty-state">${escapeHtml(err.message)}</div>`; toast(err.message,true); setStatus('NEWS ERROR'); }
  finally{ newsInFlight=false; setButtonBusy($('#refresh-news'),false); }
}

function filenameFromDisposition(header, ticker){
  if(!header) return `Midas_${ticker}_Research.docx`;
  const utf=header.match(/filename\*=UTF-8''([^;]+)/i); if(utf) return decodeURIComponent(utf[1].replace(/["']/g,''));
  const plain=header.match(/filename="?([^";]+)"?/i); return plain?plain[1]:`Midas_${ticker}_Research.docx`;
}
async function downloadReport(ticker){
  ticker=(ticker||$('#research-ticker').value||$('#ticker-input').value).trim().toUpperCase(); if(!ticker) return;
  if(reportInFlight){ toast('A DOCX export is already running in this browser.'); return; }
  const language=($('#report-language')?.value||'en').toLowerCase();
  const langLabel=language==='ko'?'KOREAN':'ENGLISH';
  reportInFlight=true; setReportButtonsBusy(true); setStatus(`BUILDING ${ticker} ${langLabel} DOCX`); log(`Building ${langLabel} Word research template for ${ticker}.`,'JOB'); toast(`Building ${ticker} ${langLabel} report. The file will download when ready.`);
  try{
    const res=await fetch(`/api/report/${encodeURIComponent(ticker)}?language=${encodeURIComponent(language)}`,{method:'POST'});
    if(!res.ok){ let data={}; try{data=await res.json()}catch{}; const retry=res.headers.get('Retry-After'); throw new Error((data.detail||`HTTP ${res.status}`)+(retry?` Retry after ${retry}s.`:'')); }
    const blob=await res.blob();
    const fileName=filenameFromDisposition(res.headers.get('Content-Disposition'),ticker);
    const objectUrl=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=objectUrl; a.download=fileName; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(objectUrl),1000);
    const status=res.headers.get('X-Midas-Research-Status')||'complete';
    if(status==='partial'){ toast(`${ticker} ${langLabel} DOCX exported with partial upstream data.`); log(`${ticker} ${langLabel} DOCX exported with partial upstream data.`,'WARN'); }
    else { toast(`${ticker} ${langLabel} DOCX export complete.`); log(`${ticker} ${langLabel} DOCX export complete.`,'OK'); }
    setStatus(`${ticker} ${langLabel} DOCX EXPORTED`);
  }catch(err){ setStatus('EXPORT ERROR'); toast(err.message,true); log(err.message,'ERR'); }
  finally{ reportInFlight=false; setReportButtonsBusy(false); }
}

$('#load-market').addEventListener('click',()=>loadMarket()); $('#ticker-input').addEventListener('keydown',e=>{if(e.key==='Enter')loadMarket();});
$('#open-research').addEventListener('click',()=>{ $('#research-ticker').value=$('#ticker-input').value; switchView('research'); runResearch(); });
$('#quick-report').addEventListener('click',()=>downloadReport($('#ticker-input').value)); $('#run-research').addEventListener('click',runResearch); $('#research-ticker').addEventListener('keydown',e=>{if(e.key==='Enter')runResearch();}); $('#download-report').addEventListener('click',()=>downloadReport());
$$('.region-btn').forEach(b=>b.addEventListener('click',()=>{currentRegion=b.dataset.region; $$('.region-btn').forEach(x=>x.classList.remove('active')); b.classList.add('active'); loadNews();})); $('#refresh-news').addEventListener('click',loadNews);

function tick(){ const now=new Date(); $('#clock').innerHTML=`<b>LOCAL</b> ${now.toLocaleTimeString([], {hour12:false})}`; } tick(); setInterval(tick,1000);
