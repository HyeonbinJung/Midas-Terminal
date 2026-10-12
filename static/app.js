const MIDAS_WEB_VERSION = "5.11.0";
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
  // whenever the user changes terminal views so controls never appear "off-screen".
  window.scrollTo({top:0,left:0,behavior:'auto'});
  document.documentElement.scrollLeft=0; document.body.scrollLeft=0;
  if(name==='alpha') requestAnimationFrame(()=>{ if(lastAlphaPayload) renderAlphaPayload(lastAlphaPayload); else alphaResizePlots(); });
  if(name==='research') requestAnimationFrame(()=>setTimeout(()=>{
    if(typeof window.Plotly==='undefined') return;
    $$('.financial-chart').forEach(el=>{try{if(el.data) Plotly.Plots.resize(el)}catch(_){}});
  },40));
  if(name==='peer') requestAnimationFrame(()=>setTimeout(()=>{ if(lastPeerAnalysisPayload) renderPeerAnalysis(lastPeerAnalysisPayload); else peerResizePlots(); },40));
  if(name==='portfolio'){ requestAnimationFrame(()=>setTimeout(()=>{loadPortfolio(true)},25)); startPortfolioPolling(); } else { stopPortfolioPolling(); }
}
$$('.fn').forEach(btn=>btn.addEventListener('click',()=>switchView(btn.dataset.view)));
document.addEventListener('keydown', e=>{
  const map={F1:'terminal',F2:'research',F3:'news',F4:'screener',F5:'alpha',F6:'peer',F7:'portfolio',F8:'about'};
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


function financialCellClass(cell){
  if(!cell) return '';
  const raw=cell.raw;
  return raw!=null&&!Number.isNaN(Number(raw))&&Number(raw)<0?'negative-number':'';
}
function renderDetailedFinancialTable(section, headSelector, bodySelector, metaSelector){
  const head=$(headSelector), body=$(bodySelector), meta=$(metaSelector);
  if(!head||!body) return;
  const periods=section?.periods||[];
  const rows=section?.rows||[];
  head.innerHTML=`<tr><th>ITEM</th>${periods.map(p=>`<th>${escapeHtml(p)}</th>`).join('')}</tr>`;
  body.innerHTML='';
  rows.forEach(row=>{
    const vals=row.values||[];
    const tr=document.createElement('tr');
    tr.innerHTML=`<td title="${escapeHtml(row.label||'')}">${escapeHtml(row.label||'')}</td>${periods.map((_,i)=>{const cell=vals[i]||{};return `<td class="${financialCellClass(cell)}">${escapeHtml(cell.display||'—')}</td>`}).join('')}`;
    body.appendChild(tr);
  });
  if(!rows.length) body.innerHTML=`<tr><td colspan="${Math.max(2,periods.length+1)}" class="placeholder">No statement rows returned by Yahoo Finance.</td></tr>`;
  if(meta) meta.textContent=`${rows.length} ROWS · ${periods.length} PERIODS`;
}
function financialPlotLayout(){
  return {
    margin:{l:54,r:14,t:10,b:34},paper_bgcolor:'#050805',plot_bgcolor:'#050805',showlegend:true,
    legend:{orientation:'h',x:0,y:1.14,font:{family:'Cascadia Mono, monospace',size:8,color:'#9aa59b'},bgcolor:'rgba(5,8,5,.74)'},
    xaxis:{gridcolor:'#192119',zeroline:false,color:'#6f7b70',tickfont:{size:8},fixedrange:false},
    yaxis:{gridcolor:'#192119',zerolinecolor:'#303830',color:'#6f7b70',tickfont:{size:8},fixedrange:false,automargin:true},
    hovermode:'x unified',hoverlabel:{bgcolor:'#111711',bordercolor:'#526052',font:{family:'Cascadia Mono, monospace',size:9,color:'#e7eee8'}}
  };
}
const FINANCIAL_PALETTE=['#a8ff32','#e1e7e2','#73c49a','#d0bc72','#8aa3a0','#b4d68c','#9a8db5','#c98d8d'];
function renderFinancialLineChart(target, chart, {percent=false}={}){
  const el=$(target); if(!el) return;
  if(!plotlyReady()){
    el.innerHTML='<div class="financial-chart-empty">PLOTLY ENGINE UNAVAILABLE</div>'; return;
  }
  const periods=chart?.periods||[], series=chart?.series||[];
  if(!periods.length||!series.length){
    try{Plotly.purge(el)}catch(_){};
    el.innerHTML='<div class="financial-chart-empty">NO PUBLIC SERIES AVAILABLE</div>'; return;
  }
  const traces=series.map((item,i)=>({
    type:'scatter',mode:'lines+markers',name:item.label,x:periods,y:item.values,
    line:{width:1.7,color:FINANCIAL_PALETTE[i%FINANCIAL_PALETTE.length]},
    marker:{size:4,color:FINANCIAL_PALETTE[i%FINANCIAL_PALETTE.length]},
    connectgaps:false,
    hovertemplate:`${escapeHtml(item.label)}<br>%{x}<br>%{y:,.2f}${percent?'%':' MM'}<extra></extra>`
  }));
  try{Plotly.purge(el)}catch(_){}
  const layout=financialPlotLayout();
  if(percent) layout.yaxis.ticksuffix='%';
  Plotly.newPlot(el,traces,layout,{responsive:true,scrollZoom:true,displaylogo:false,displayModeBar:false,doubleClick:'reset'}).then(()=>{try{Plotly.Plots.resize(el)}catch(_){}});
}
function renderExpenseChart(target, payload){
  const chart={periods:payload?.periods||[],series:(payload?.series||[]).map(x=>({label:x.label,values:x.pct_revenue||[]}))};
  renderFinancialLineChart(target,chart,{percent:true});
}
function renderExpenseLatest(target,payload){
  const el=$(target); if(!el) return;
  const rows=(payload?.latest||[]).slice(0,8);
  if(!rows.length){el.innerHTML='<span class="placeholder">Yahoo does not expose a granular operating-expense split for this period.</span>';return;}
  el.innerHTML=rows.map((r,i)=>`<div class="expense-latest-row"><span class="expense-rank">0${i+1}</span><b>${escapeHtml(r.label||'')}</b><strong>${num(r.value,1)} MM</strong><em>${r.pct_revenue==null?'—':num(r.pct_revenue,1)+'% REV'}</em></div>`).join('');
}
function renderKeyFinancials(financials){
  const detailed=financials?.detailed||{};
  const income=detailed.income||{}, balance=detailed.balance||{}, cashflow=detailed.cashflow||{};
  renderDetailedFinancialTable(income.annual,'#income-annual-head','#income-annual-body','#income-annual-meta');
  renderDetailedFinancialTable(income.quarterly,'#income-quarterly-head','#income-quarterly-body','#income-quarterly-meta');
  renderDetailedFinancialTable(balance.annual,'#balance-annual-head','#balance-annual-body','#balance-annual-meta');
  renderDetailedFinancialTable(balance.quarterly,'#balance-quarterly-head','#balance-quarterly-body','#balance-quarterly-meta');
  renderDetailedFinancialTable(cashflow.annual,'#cashflow-annual-head','#cashflow-annual-body','#cashflow-annual-meta');
  renderDetailedFinancialTable(cashflow.quarterly,'#cashflow-quarterly-head','#cashflow-quarterly-body','#cashflow-quarterly-meta');

  const counts=financials?.counts||{};
  const incomeCount=$('#income-row-count'), balanceCount=$('#balance-row-count'), cashflowCount=$('#cashflow-row-count');
  if(incomeCount) incomeCount.textContent=`${counts.income_annual||0} FY ROWS · ${counts.income_quarterly||0} QTR ROWS`;
  if(balanceCount) balanceCount.textContent=`${counts.balance_annual||0} FY ROWS · ${counts.balance_quarterly||0} QTR ROWS`;
  if(cashflowCount) cashflowCount.textContent=`${counts.cashflow_annual||0} FY ROWS · ${counts.cashflow_quarterly||0} QTR ROWS`;

  const consensusEl=$('#income-consensus-strip');
  if(consensusEl){
    const consensus=financials?.consensus||[];
    consensusEl.innerHTML=consensus.length?consensus.map(c=>`<div><label>${escapeHtml(c.period||'')}</label><span>REV <b>${escapeHtml(c.revenue||'—')}</b></span><span>EPS <b>${escapeHtml(c.eps||'—')}</b></span></div>`).join(''):'<span class="placeholder">Public consensus unavailable.</span>';
  }

  const populated=Object.values(counts).reduce((a,v)=>a+(Number(v)||0),0);
  const unit=$('#financial-units');
  if(unit){const currency=financials?.currency||'';unit.textContent=`${currency}${currency?' · ':''}${financials?.units||'MM'} · ${populated} TOTAL ROWS`;unit.classList.toggle('partial',populated===0);}
  const panel=$('.key-financials-panel'); if(panel) panel.dataset.loaded=populated>0?'true':'empty';

  // Plot after the newly populated tables have laid out; this prevents zero-width Plotly canvases.
  requestAnimationFrame(()=>setTimeout(()=>{
    renderFinancialLineChart('#income-annual-chart',income.annual_chart);
    renderFinancialLineChart('#income-quarterly-chart',income.quarterly_chart);
    renderExpenseChart('#expense-annual-chart',income.expense_annual);
    renderExpenseChart('#expense-quarterly-chart',income.expense_quarterly);
    renderFinancialLineChart('#balance-annual-chart',balance.annual_chart);
    renderFinancialLineChart('#balance-quarterly-chart',balance.quarterly_chart);
    renderFinancialLineChart('#cashflow-annual-chart',cashflow.annual_chart);
    renderFinancialLineChart('#cashflow-quarterly-chart',cashflow.quarterly_chart);
    renderExpenseLatest('#expense-annual-latest',income.expense_annual);
    renderExpenseLatest('#expense-quarterly-latest',income.expense_quarterly);
  },40));
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





const FACTOR_AXES=[
  {key:'market_cap',label:'MKT CAP',long:'Market Cap'},
  {key:'sector',label:'SECTOR',long:'Sector Regime'},
  {key:'volatility',label:'VOL',long:'Volatility'},
  {key:'beta',label:'BETA',long:'Market Sensitivity'},
  {key:'style',label:'STYLE',long:'Style (Value ↔ Growth)'},
  {key:'growth',label:'GROWTH',long:'Growth'},
  {key:'profitability',label:'PROFIT',long:'Profitability'},
  {key:'leverage',label:'LEV',long:'Leverage'},
  {key:'cash_flow',label:'CASH',long:'Cash Flow'},
  {key:'dividend',label:'DIV',long:'Dividend'},
  {key:'momentum',label:'MOM',long:'Momentum'},
  {key:'liquidity',label:'LIQ',long:'Liquidity'},
  {key:'cyclicality',label:'CYC',long:'Cyclicality'},
];
const PLOT_BG='#050805';
let lastScreenerRows=[];

function clampScore(v, fallback=50){
  const n=Number(v);
  if(!Number.isFinite(n)) return fallback;
  return Math.max(0,Math.min(100,Math.round(n)));
}
function logScaleScore(v, low, high, fallback=50){
  const n=Number(v);
  if(!Number.isFinite(n)||n<=0) return fallback;
  const lo=Math.log10(low), hi=Math.log10(high);
  return clampScore(((Math.log10(n)-lo)/(hi-lo))*100,fallback);
}
function labelScore(value, map, fallback=50){
  return Object.prototype.hasOwnProperty.call(map,value)?map[value]:fallback;
}
/*
 * Convert every screener classification into a 0-100 factor-exposure score.
 * These are visualization scores, not investment recommendations: a higher score
 * means "more exposure to the named factor" (e.g. higher leverage or volatility),
 * not necessarily "better".
 */
function scoreFor(r,key){
  const labels=r?.labels||{};
  switch(key){
    case 'market_cap':
      return logScaleScore(r?.market_cap,3e8,2e11,50);
    case 'sector': {
      // Sector itself is categorical, so encode the sector's broad market-regime
      // exposure rather than pretending one industry is inherently "better".
      const sector=String(r?.sector||'');
      if(['Technology','Communication Services'].includes(sector)) return 85;
      if(['Consumer Cyclical','Industrials','Energy','Basic Materials','Financial Services'].includes(sector)) return 65;
      if(['Real Estate'].includes(sector)) return 50;
      if(['Healthcare'].includes(sector)) return 40;
      if(['Consumer Defensive','Utilities'].includes(sector)) return 25;
      return 50;
    }
    case 'volatility': {
      const vol=Number(r?.annualized_volatility);
      if(Number.isFinite(vol)) return clampScore(((vol-0.10)/(0.70-0.10))*100,50);
      return labelScore(labels.volatility,{Low:25,Medium:55,High:85});
    }
    case 'beta': {
      const beta=Number(r?.beta);
      if(Number.isFinite(beta)) return clampScore(((beta-0.40)/(1.80-0.40))*100,50);
      return labelScore(labels.beta,{Low:25,Medium:55,High:85});
    }
    case 'style':
      return labelScore(labels.style,{'Deep Value':15,Value:35,GARP:60,Growth:90});
    case 'growth': {
      const g=Number(r?.revenue_growth);
      if(Number.isFinite(g)) return clampScore(((g+0.10)/0.50)*100,50);
      return labelScore(labels.growth,{Low:25,Medium:55,High:85});
    }
    case 'profitability': {
      const vals=[];
      const roe=Number(r?.roe), margin=Number(r?.net_margin);
      if(Number.isFinite(roe)) vals.push(clampScore(((roe+0.05)/0.45)*100));
      if(Number.isFinite(margin)) vals.push(clampScore(((margin+0.05)/0.35)*100));
      if(vals.length) return Math.round(vals.reduce((a,b)=>a+b,0)/vals.length);
      return labelScore(labels.profitability,{'Loss-making':15,Average:55,'High Quality':90});
    }
    case 'leverage': {
      const nde=Number(r?.net_debt_ebitda);
      if(Number.isFinite(nde)) return clampScore(((nde+0.5)/5.0)*100,50);
      return labelScore(labels.leverage,{'Net Cash':10,Moderate:50,High:90});
    }
    case 'cash_flow':
      return labelScore(labels.cash_flow,{'Cash Burn':10,'CapEx Heavy':35,'Positive FCF':65,'FCF Compounder':90});
    case 'dividend':
      return labelScore(labels.dividend,{'Non-dividend':10,Dividend:60,'High Yield':90});
    case 'momentum': {
      const m=Number(r?.ret_6m);
      if(Number.isFinite(m)) return clampScore(((m+0.30)/0.80)*100,50);
      return labelScore(labels.momentum,{Negative:15,Neutral:50,Positive:85});
    }
    case 'liquidity': {
      const adtv=Number(r?.avg_daily_value);
      if(Number.isFinite(adtv)&&adtv>0) return logScaleScore(adtv,5e6,1e9,50);
      return labelScore(labels.liquidity,{Low:20,Medium:55,High:90});
    }
    case 'cyclicality':
      return labelScore(labels.cyclicality,{Defensive:20,Cyclical:60,'Secular Growth':90});
    default:
      return 50;
  }
}
function factorScores(r){
  return FACTOR_AXES.map(ax=>scoreFor(r,ax.key));
}
function resolvePlotlyEngine(){
  // plotly.py 6.x can ship a browser bundle whose UMD wrapper exports as
  // window.moduleName instead of window.Plotly. Recover that export so every
  // screener/alpha chart uses the same engine.
  if((!window.Plotly || typeof window.Plotly.newPlot!=='function') && window.moduleName && typeof window.moduleName.newPlot==='function'){
    window.Plotly=window.moduleName;
  }
  return window.Plotly;
}
function plotlyReady(){
  const p=resolvePlotlyEngine();
  return !!(p && typeof p.newPlot==='function' && p.Plots && typeof p.Plots.resize==='function');
}
function purgePlot(el){
  if(!el||!plotlyReady()) return;
  try{ Plotly.purge(el); }catch(_){ }
  el.innerHTML='';
  // Plotly can leave inline white backgrounds after a WebGL context reset.
  el.style.background='#050805';
}
function drawFactor3D(rows=[]){
  const plot=$('#factor-3d-plot'); if(!plot) return;
  if(!plotlyReady()){
    plot.innerHTML='<div class="deck-empty">3D ENGINE UNAVAILABLE — CHECK NETWORK ACCESS TO PLOTLY CDN.</div>';
    return;
  }
  // Rebuild the WebGL scene on every scan. This avoids stale/lost WebGL contexts
  // that could leave the second screener run as a white rectangle.
  purgePlot(plot);
  const palette=['#a8ff32','#d8dfd8','#7fa98a','#d0bc72','#74c9ae'];
  const traces=(rows||[]).slice(0,5).map((r,idx)=>{
    const z=factorScores(r);
    const y=FACTOR_AXES.map(()=>idx+1);
    return {
      type:'scatter3d', mode:'lines+markers', name:r.ticker,
      x:FACTOR_AXES.map((_,i)=>i), y, z,
      text:FACTOR_AXES.map((f,i)=>`${r.ticker}<br>${f.long}: ${z[i]}`), hoverinfo:'text',
      line:{width:5,color:palette[idx]}, marker:{size:4,color:palette[idx],symbol:'circle'}
    };
  });
  if(!traces.length){
    traces.push({type:'scatter3d',mode:'markers',x:[0],y:[1],z:[0],marker:{size:1,color:'#050805'},hoverinfo:'skip',showlegend:false});
  }
  const tickers=(rows||[]).slice(0,5).map(r=>r.ticker);
  const layout={
    autosize:true,
    margin:{l:0,r:0,t:8,b:0},paper_bgcolor:'#050805',plot_bgcolor:'#050805',showlegend:rows.length>0,
    legend:{orientation:'h',x:0.01,y:0.99,font:{family:'Cascadia Mono, monospace',size:9,color:'#9ba79b'},bgcolor:'rgba(5,8,5,.78)'},
    scene:{
      bgcolor:'#050805',dragmode:'orbit',aspectmode:'manual',aspectratio:{x:1.9,y:1.05,z:1.2},
      camera:{eye:{x:1.48,y:1.58,z:1.06}},
      xaxis:{title:'FACTOR',tickmode:'array',tickvals:FACTOR_AXES.map((_,i)=>i),ticktext:FACTOR_AXES.map(f=>f.label),range:[-.35,FACTOR_AXES.length-.65],showbackground:true,backgroundcolor:'#071007',gridcolor:'#283128',zerolinecolor:'#354035',color:'#829082',titlefont:{size:9},tickfont:{size:8}},
      yaxis:{title:'SECURITY',tickmode:'array',tickvals:[1,2,3,4,5],ticktext:tickers.length?tickers:['—','—','—','—','—'],range:[.5,5.5],showbackground:true,backgroundcolor:'#070a07',gridcolor:'#222922',zerolinecolor:'#354035',color:'#829082',titlefont:{size:9},tickfont:{size:8}},
      zaxis:{title:'SCORE',range:[0,100],dtick:20,showbackground:true,backgroundcolor:'#070a07',gridcolor:'#2b362b',zerolinecolor:'#354035',color:'#829082',titlefont:{size:9},tickfont:{size:8}}
    }
  };
  const config={responsive:true,scrollZoom:true,displaylogo:false,doubleClick:'reset',modeBarButtonsToRemove:['toImage','sendDataToCloud','lasso3d','select2d']};
  Plotly.newPlot(plot,traces,layout,config).then(()=>{
    requestAnimationFrame(()=>{
      try{ Plotly.Plots.resize(plot); Plotly.relayout(plot,{'paper_bgcolor':'#050805','plot_bgcolor':'#050805','scene.bgcolor':'#050805'}); }catch(_){ }
    });
  }).catch(()=>{
    plot.innerHTML='<div class="deck-empty">3D RENDER RESET FAILED — RUN SEARCH AGAIN.</div>';
  });
}
function purgePricePlots(){
  if(!plotlyReady()) return;
  $$('.price-plot, #price-overlay-plot').forEach(p=>purgePlot(p));
}
function renderPriceOverlay(rows=[]){
  const el=$('#price-overlay-plot'); if(!el) return;
  purgePlot(el);
  if(!rows.length){ el.innerHTML='<div class="deck-empty">Run a search to compare normalized 3-year performance.</div>'; return; }
  if(!plotlyReady()) return;
  const palette=['#a8ff32','#d8dfd8','#7fa98a','#d0bc72','#74c9ae'];
  const traces=[];
  rows.slice(0,5).forEach((r,i)=>{
    const h=Array.isArray(r.price_history_3y)?r.price_history_3y:[];
    if(!h.length) return;
    const first=Number(h[0].close);
    if(!Number.isFinite(first)||first===0) return;
    traces.push({
      type:'scatter',mode:'lines',name:r.ticker,
      x:h.map(p=>p.date),y:h.map(p=>Number(p.close)/first*100),
      hovertemplate:`${escapeHtml(r.ticker)}<br>%{x}<br>INDEX %{y:.1f}<extra></extra>`,
      line:{width:1.7,color:palette[i]}
    });
  });
  if(!traces.length){el.innerHTML='<div class="deck-empty">PRICE HISTORY UNAVAILABLE</div>';return;}
  const layout={
    autosize:true,margin:{l:46,r:12,t:10,b:28},paper_bgcolor:'#050805',plot_bgcolor:'#050805',
    showlegend:true,legend:{orientation:'h',x:0,y:1.08,font:{family:'Cascadia Mono, monospace',size:8,color:'#829082'},bgcolor:'rgba(5,8,5,.7)'},
    xaxis:{gridcolor:'#172017',zeroline:false,color:'#657165',tickfont:{size:8},rangeslider:{visible:false}},
    yaxis:{title:'START = 100',gridcolor:'#172017',zeroline:false,color:'#657165',tickfont:{size:8},titlefont:{size:8},fixedrange:false},
    hovermode:'x unified'
  };
  const config={responsive:true,scrollZoom:true,displaylogo:false,displayModeBar:false,doubleClick:'reset'};
  Plotly.newPlot(el,traces,layout,config);
}
function renderPriceDeck(rows=[]){
  const deck=$('#price-deck'); if(!deck) return;
  // Critical leak fix: destroy old Plotly instances before removing their DOM nodes.
  // The old scattergl charts consumed WebGL contexts and could make the next 3D plot white.
  purgePricePlots();
  if(!rows.length){deck.innerHTML='<div class="deck-empty">Run a search to stream 3-year prices for the five matches.</div>'; renderPriceOverlay([]); return;}
  deck.innerHTML=rows.slice(0,5).map((r,i)=>`<article class="price-card"><div class="price-card-head"><span>0${i+1} // ${escapeHtml(r.ticker)} · 3Y PRICE</span><span>${escapeHtml(r.name||'')}</span></div><div class="price-plot" id="price-plot-${i}"></div></article>`).join('');
  if(!plotlyReady()) return;
  rows.slice(0,5).forEach((r,i)=>{
    const h=Array.isArray(r.price_history_3y)?r.price_history_3y:[];
    const el=$(`#price-plot-${i}`); if(!el)return;
    if(!h.length){el.innerHTML='<div class="deck-empty">PRICE HISTORY UNAVAILABLE</div>';return;}
    // Use SVG scatter rather than scattergl. Only the 3D factor scene needs WebGL;
    // keeping the five price charts in SVG prevents browser WebGL context exhaustion.
    const trace={type:'scatter',mode:'lines',x:h.map(p=>p.date),y:h.map(p=>p.close),hovertemplate:`${escapeHtml(r.ticker)}<br>%{x}<br>$%{y:.2f}<extra></extra>`,line:{width:1.4,color:'#a8ff32'},fill:'tozeroy',fillcolor:'rgba(168,255,50,.035)'};
    const layout={autosize:true,margin:{l:42,r:10,t:8,b:25},paper_bgcolor:'#050805',plot_bgcolor:'#050805',showlegend:false,xaxis:{gridcolor:'#172017',zeroline:false,color:'#657165',tickfont:{size:8},rangeslider:{visible:false}},yaxis:{gridcolor:'#172017',zeroline:false,color:'#657165',tickfont:{size:8},fixedrange:false},hovermode:'x unified'};
    const config={responsive:true,scrollZoom:true,displaylogo:false,displayModeBar:false,doubleClick:'reset'};
    Plotly.newPlot(el,[trace],layout,config);
  });
  renderPriceOverlay(rows);
}
function renderScreenerCards(rows=[]){
  const wrap=$('#screener-cards'); if(!wrap)return;
  if(!rows.length){wrap.innerHTML='<div class="screener-card-empty">NO MATCHED SECURITIES.</div>';return;}
  wrap.innerHTML=rows.slice(0,5).map((r,i)=>`<article class="screener-result-card" data-ticker="${escapeHtml(r.ticker)}"><div class="src-top"><div><span class="src-rank">0${i+1}</span><div class="src-ticker">${escapeHtml(r.ticker)}</div><div class="src-name">${escapeHtml(r.name||'')}</div></div><span class="src-open">OPEN ↗</span></div><div class="src-grid"><div><label>MKT CAP</label><strong>${money(r.market_cap)}</strong></div><div><label>SECTOR</label><strong>${escapeHtml(r.sector||'—')}</strong></div><div><label>BETA</label><strong>${num(r.beta,2)}</strong></div><div><label>VOL</label><strong>${fmtRatio(r.annualized_volatility)}</strong></div><div><label>6M MOM</label><strong>${fmtRatio(r.ret_6m)}</strong></div><div><label>FWD P/E</label><strong>${num(r.forward_pe,1)}</strong></div></div><div class="score-mini-strip">${FACTOR_AXES.map(ax=>`<span><b>${ax.label}</b>${scoreFor(r,ax.key)}</span>`).join('')}</div><div class="src-tags">${tagHtml(r.labels||{})}</div></article>`).join('');
}
function initCockpitResizers(){
  const cockpit=$('#quant-cockpit'); if(!cockpit) return;
  $$('.cockpit-splitter').forEach(split=>{
    split.addEventListener('pointerdown',e=>{
      if(window.innerWidth<=820)return;
      e.preventDefault(); split.setPointerCapture(e.pointerId); split.classList.add('dragging');
      const rect=cockpit.getBoundingClientRect(); const children=[...cockpit.children];
      const left=$('#cockpit-left'),center=$('#cockpit-center'),right=$('#cockpit-right');
      const startX=e.clientX, lw=left.getBoundingClientRect().width, cw=center.getBoundingClientRect().width, rw=right.getBoundingClientRect().width;
      const move=ev=>{
        const dx=ev.clientX-startX;
        let nl=lw,nc=cw,nr=rw;
        if(split.dataset.split==='left'){nl=Math.max(210,Math.min(lw+dx,lw+cw-390));nc=lw+cw-nl;}
        else {nc=Math.max(390,Math.min(cw+dx,cw+rw-240));nr=cw+rw-nc;}
        cockpit.style.gridTemplateColumns=`${nl}px 6px ${nc}px 6px ${nr}px`;
        if(plotlyReady()){const p=$('#factor-3d-plot'); if(p)Plotly.Plots.resize(p); rowsPriceResize();}
      };
      const up=()=>{split.classList.remove('dragging');split.removeEventListener('pointermove',move);split.removeEventListener('pointerup',up);};
      split.addEventListener('pointermove',move); split.addEventListener('pointerup',up);
    });
  });
}
function rowsPriceResize(){ $$('.price-plot, #price-overlay-plot').forEach(p=>{try{Plotly.Plots.resize(p)}catch(_){}}); }
function initDragScroll(){
  $$('.drag-scroll').forEach(el=>{
    let down=false,sx=0,sl=0;
    el.addEventListener('pointerdown',e=>{if(e.target.closest('button'))return;down=true;sx=e.clientX;sl=el.scrollLeft;el.classList.add('dragging-scroll');el.setPointerCapture(e.pointerId)});
    el.addEventListener('pointermove',e=>{if(down)el.scrollLeft=sl-(e.clientX-sx)});
    const stop=()=>{down=false;el.classList.remove('dragging-scroll')}; el.addEventListener('pointerup',stop); el.addEventListener('pointercancel',stop);
  });
}
window.addEventListener('resize',()=>{if(plotlyReady()){const p=$('#factor-3d-plot');if(p)try{Plotly.Plots.resize(p)}catch(_){};rowsPriceResize();}});
let screenerInFlight=false;
function screenerFilters(){
  const out={};
  $$('#screener-view [data-filter]').forEach(el=>out[el.dataset.filter]=el.value);
  return out;
}
function updateScreenerQuery(){
  const f=screenerFilters();
  const active=Object.entries(f).filter(([,v])=>v&&v!=='Any').map(([k,v])=>`${k.toUpperCase()}=${v}`);
  const el=$('#active-query'); if(el) el.textContent=active.length?active.join('  //  '):'NO FILTERS — BROAD UNIVERSE';
}
function fmtRatio(v,d=1){ return v==null||Number.isNaN(Number(v))?'—':`${(Number(v)*100).toFixed(d)}%`; }
function tagHtml(labels={}){
  const keys=['market_cap','volatility','style','profitability','leverage','momentum','cyclicality'];
  return keys.filter(k=>labels[k]&&labels[k]!=='Unknown').map(k=>`<span class="factor-tag ${['High Quality','Positive','Mega','Growth','FCF Compounder'].includes(labels[k])?'hot':''}">${escapeHtml(labels[k])}</span>`).join('');
}
async function runScreener(){
  if(screenerInFlight) return;
  screenerInFlight=true; const btn=$('#run-screener'); setButtonBusy(btn,true,'SCANNING YAHOO…');
  $('#screener-status').textContent='SCANNING'; setStatus('STOCK SCREENER RUNNING');
  const body=$('#screener-results'); body.innerHTML='<tr><td colspan="13" class="screener-empty">FETCHING LIVE YFINANCE METRICS · CLASSIFYING UNIVERSE…</td></tr>';
  try{
    const params=new URLSearchParams(); Object.entries(screenerFilters()).forEach(([k,v])=>{if(v!=='Any')params.set(k,v)});
    const d=await api(`/api/screener?${params.toString()}`);
    body.innerHTML='';
    (d.results||[]).forEach((r,i)=>{
      const tr=document.createElement('tr');
      const mom=r.ret_6m==null?'':r.ret_6m>=0?'pos':'neg';
      tr.innerHTML=`<td class="screener-rank">0${i+1}</td><td class="screener-ticker">${escapeHtml(r.ticker)}</td><td class="screener-company">${escapeHtml(r.name||'')}</td><td class="screener-sector"><b>${escapeHtml(r.sector||'—')}</b><small>${escapeHtml(r.industry||'—')}</small></td><td>${money(r.market_cap)}</td><td>${num(r.beta,2)}</td><td>${fmtRatio(r.annualized_volatility)}</td><td class="screener-metric ${mom}">${fmtRatio(r.ret_6m)}</td><td>${num(r.forward_pe,1)}</td><td class="screener-metric ${r.revenue_growth!=null&&r.revenue_growth>=0?'pos':'neg'}">${fmtRatio(r.revenue_growth)}</td><td>${fmtRatio(r.roe)}</td><td><div class="tag-stack">${tagHtml(r.labels||{})}</div></td><td><button class="screener-open" data-ticker="${escapeHtml(r.ticker)}">OPEN</button></td>`;
      body.appendChild(tr);
    });
    if(!(d.results||[]).length) body.innerHTML='<tr><td colspan="13" class="screener-empty">NO EXACT MATCHES IN THE CURRENT LIQUID-US UNIVERSE. RELAX ONE OR MORE FILTERS.</td></tr>';
    lastScreenerRows=(d.results||[]).slice(0,5); drawFactor3D(lastScreenerRows); renderPriceDeck(lastScreenerRows); renderScreenerCards(lastScreenerRows);
    $('#screener-count').textContent=`${(d.results||[]).length} SHOWN · ${d.count_matched||0} MATCHED · ${d.count_scanned||0} SCANNED`;
    $('#screener-status').textContent='COMPLETE'; $('#screener-note').textContent=d.note||''; setStatus('SCREENER COMPLETE');
    log(`Screener: ${(d.results||[]).length} shown / ${d.count_matched||0} matched from ${d.count_scanned||0} live snapshots.`,'SCAN');
  }catch(err){ body.innerHTML=`<tr><td colspan="13" class="screener-empty">${escapeHtml(err.message)}</td></tr>`; $('#screener-status').textContent='ERROR'; setStatus('SCREENER ERROR'); toast(err.message,true); }
  finally{screenerInFlight=false;setButtonBusy(btn,false)}
}
$$('#screener-view [data-filter]').forEach(el=>el.addEventListener('change',updateScreenerQuery));
$('#run-screener')?.addEventListener('click',runScreener);
$('#reset-screener')?.addEventListener('click',()=>{$$('#screener-view [data-filter]').forEach(el=>el.value='Any');updateScreenerQuery();$('#screener-count').textContent='0 RESULTS';$('#screener-status').textContent='READY';$('#screener-results').innerHTML='<tr><td colspan="13" class="screener-empty">Configure the matrix above and press SEARCH 5 STOCKS.</td></tr>';lastScreenerRows=[];drawFactor3D([]);renderPriceDeck([]);renderScreenerCards([]);});
$('#screener-results')?.addEventListener('click',e=>{const b=e.target.closest('[data-ticker]');if(!b)return;$('#ticker-input').value=b.dataset.ticker;loadMarket(b.dataset.ticker);switchView('terminal');});
$('#screener-cards')?.addEventListener('click',e=>{const c=e.target.closest('[data-ticker]');if(!c)return;$('#ticker-input').value=c.dataset.ticker;loadMarket(c.dataset.ticker);switchView('terminal');});
updateScreenerQuery();
initCockpitResizers(); initDragScroll();
requestAnimationFrame(()=>{drawFactor3D([]);renderPriceDeck([]);renderScreenerCards([]);});


// === v5.7 Alpha Factor Modeling ===
let alphaInFlight=false;
let lastAlphaPayload=null;
let alphaCatalogCache=null;
const ALPHA_DEFAULTS=['momentum_6m','pe_ttm','ev_ebitda_ttm','revenue_growth_yoy','roic_ttm','fcf_margin_ttm','debt_equity','beta_1y'];
const ALPHA_PALETTE=['#a8ff32','#d9dfd9','#79b894','#d7bd70','#71c7ad','#c98b79','#8ca8d8','#b49bd8','#d7d77d','#7db8c9','#a0c77d','#c7a17d','#a8a8ff','#ffb27d','#83d4ff','#d4ff83'];

function selectedAlphaFactors(){
  return $$('[data-alpha-factor]:checked').map(el=>el.dataset.alphaFactor);
}
function setAlphaSelections(keys){
  const wanted=new Set(keys||[]);
  $$('[data-alpha-factor]').forEach(el=>{el.checked=wanted.has(el.dataset.alphaFactor)});
  updateAlphaSelectedCount();
}
function updateAlphaSelectedCount(){
  const n=selectedAlphaFactors().length;
  const el=$('#alpha-selected-count'); if(el) el.textContent=`${n} SELECTED`;
  const count=$('#alpha-factor-count'); if(count) count.textContent=lastAlphaPayload?`${n}/${lastAlphaPayload.availability?.available_factors??'—'}`:String(n);
}
function alphaCatalogIndex(payload){
  const source=payload?.catalog||alphaCatalogCache||{};
  const out={};
  Object.entries(source).forEach(([group,items])=>{(items||[]).forEach(item=>out[item.key]={...item,group})});
  return out;
}
function alphaGroupHint(group){
  const map={
    'Momentum':'PRICE TREND','Value / Multiple':'VALUATION','Growth':'FUNDAMENTALS','Profitability / Quality':'ECONOMIC QUALITY',
    'Cash Flow Quality':'EARNINGS QUALITY','Leverage / Balance Sheet':'CAPITAL STRUCTURE','Efficiency':'OPERATING EFFICIENCY',
    'Volatility / Risk':'MARKET RISK','Size / Liquidity':'TRADABILITY','Estimate Revision':'EXPECTATIONS','Earnings Surprise':'EVENT SIGNALS',
    'Sentiment':'POSITIONING / OPINION','Shareholder Return':'CAPITAL RETURN','Capital Allocation':'REINVESTMENT'
  };
  return map[group]||'FACTOR GROUP';
}
function renderAlphaSelector(catalog){
  const root=$('#alpha-factor-selector-groups'); if(!root||!catalog)return;
  const previous=new Set(selectedAlphaFactors());
  const firstLoad=!root.dataset.ready;
  root.innerHTML=Object.entries(catalog).map(([group,items])=>`<section class="alpha-factor-group"><h3>${escapeHtml(group.toUpperCase())}<span>${escapeHtml(alphaGroupHint(group))}</span></h3>${(items||[]).map(item=>`<label class="alpha-factor-option" title="${escapeHtml(item.description||'')}"><input type="checkbox" data-alpha-factor="${escapeHtml(item.key)}"><span><b>${escapeHtml(item.label.toUpperCase())}</b><small>${escapeHtml(item.description||'')}</small></span></label>`).join('')}</section>`).join('');
  root.dataset.ready='1';
  $$('[data-alpha-factor]',root).forEach(el=>el.addEventListener('change',()=>{updateAlphaSelectedCount();if(lastAlphaPayload)renderAlphaPayload(lastAlphaPayload)}));
  if(firstLoad) setAlphaSelections(ALPHA_DEFAULTS); else setAlphaSelections([...previous]);
}
async function loadAlphaCatalog(){
  try{
    const d=await api('/api/alpha-factor-catalog');
    alphaCatalogCache=d.catalog||{};
    renderAlphaSelector(alphaCatalogCache);
  }catch(err){
    const root=$('#alpha-factor-selector-groups');
    if(root) root.innerHTML=`<div class="deck-empty">FACTOR CATALOG ERROR · ${escapeHtml(err.message)}</div>`;
  }
}
function alphaPlotValues(values,unit){
  return (values||[]).map(v=>v==null?null:(unit==='percent'?Number(v)*100:Number(v)));
}
function alphaValueText(value,meta,digits=2){
  if(value==null||Number.isNaN(Number(value))) return '—';
  const n=Number(value);
  if(meta?.unit==='percent') return `${(n*100).toFixed(digits)}%`;
  if(meta?.unit==='multiple') return `${n.toFixed(digits)}x`;
  if(meta?.unit==='currency') return money(n);
  if(meta?.unit==='days') return `${n.toFixed(1)}d`;
  if(meta?.unit==='score') return n.toFixed(digits);
  return n.toFixed(digits);
}
function alphaAxisSuffix(meta){ return meta?.unit==='percent'?'%':meta?.unit==='multiple'?'x':meta?.unit==='days'?'d':''; }
function alphaAxisTitle(meta){
  if(meta?.unit==='percent')return 'PERCENT'; if(meta?.unit==='multiple')return 'MULTIPLE'; if(meta?.unit==='currency')return 'USD'; if(meta?.unit==='days')return 'DAYS'; if(meta?.unit==='score')return 'SCORE'; return 'VALUE';
}
function purgeAlphaPlots(){
  if(!plotlyReady()) return;
  ['#alpha-overlay-plot','#alpha-price-plot'].forEach(sel=>{const el=$(sel);if(el)purgePlot(el)});
  $$('.alpha-factor-chart').forEach(el=>purgePlot(el));
}
function alphaBaseLayout(extra={}){
  return {
    autosize:true,margin:{l:54,r:14,t:12,b:31},paper_bgcolor:PLOT_BG,plot_bgcolor:PLOT_BG,
    font:{family:'Cascadia Mono, monospace',size:8,color:'#788578'},
    xaxis:{gridcolor:'#182018',zeroline:false,color:'#687468',tickfont:{size:8},fixedrange:false,automargin:true},
    yaxis:{gridcolor:'#182018',zerolinecolor:'#2e3a2e',color:'#687468',tickfont:{size:8},fixedrange:false,automargin:true},
    hovermode:'x unified',...extra
  };
}
function alphaPlotError(el,err){
  if(!el)return;
  console.error('Alpha Plotly render error',err);
  try{ if(plotlyReady()) Plotly.purge(el); }catch(_){ }
  el.innerHTML=`<div class="deck-empty">GRAPH RENDER ERROR · ${escapeHtml(err?.message||String(err||'UNKNOWN'))}</div>`;
  el.style.background=PLOT_BG;
}
function alphaArray(value){ return Array.isArray(value)?value:[]; }
function alphaQuarters(payload){ return alphaArray(payload?.quarters).map(v=>String(v)); }
function alphaSeries(payload,key,quarters){
  const raw=alphaArray(payload?.series?.[key]);
  return quarters.map((_,i)=>{
    if(i>=raw.length || raw[i]==null) return null;
    const n=Number(raw[i]);
    return Number.isFinite(n)?n:null;
  });
}
/*
 * Use the exact same render lifecycle as the working F4 Screener chart:
 * purge -> Plotly.newPlot -> resize -> force graphite backgrounds.
 * Do not manually set width/height or relayout them to null; that lifecycle was
 * the source of the blank Alpha canvas on some Chromium/Plotly combinations.
 */
function alphaPlot(el,traces,layout,config={}){
  if(!el) return Promise.resolve(null);
  if(!plotlyReady()){
    alphaPlotError(el,new Error('Plotly engine failed to initialize'));
    return Promise.resolve(null);
  }
  purgePlot(el);
  el.style.background=PLOT_BG;
  const finalLayout={
    autosize:true,
    paper_bgcolor:PLOT_BG,
    plot_bgcolor:PLOT_BG,
    font:{family:'Cascadia Mono, monospace',size:8,color:'#788578'},
    ...layout
  };
  const finalConfig={responsive:true,scrollZoom:true,displaylogo:false,doubleClick:'reset',...config};
  try{
    return Plotly.newPlot(el,traces,finalLayout,finalConfig).then(()=>{
      requestAnimationFrame(()=>{
        try{
          Plotly.Plots.resize(el);
          const patch={'paper_bgcolor':PLOT_BG,'plot_bgcolor':PLOT_BG};
          if(finalLayout.scene) patch['scene.bgcolor']=PLOT_BG;
          Plotly.relayout(el,patch);
        }catch(_){ }
      });
      return el;
    }).catch(err=>{alphaPlotError(el,err);return null;});
  }catch(err){
    alphaPlotError(el,err);
    return Promise.resolve(null);
  }
}
function zNormalize(raw){
  const src=alphaArray(raw);
  const clean=src.filter(v=>Number.isFinite(v));
  if(!clean.length)return src.map(()=>null);
  const mean=clean.reduce((a,b)=>a+b,0)/clean.length;
  const variance=clean.reduce((a,b)=>a+(b-mean)*(b-mean),0)/clean.length;
  const sd=Math.sqrt(variance);
  return src.map(v=>Number.isFinite(v)?(sd>0?(v-mean)/sd:0):null);
}
let alphaOverlayMode='3d';
function alphaQuarterNumber(label){
  const m=String(label||'').match(/^(\d{4})Q([1-4])$/i);
  if(!m) return null;
  return Number(m[1])*4+(Number(m[2])-1);
}
function alphaQuarterDateSeries(payload,quarters){
  const supplied=alphaArray(payload?.quarter_dates).map(v=>String(v));
  return quarters.map((q,i)=>{
    if(supplied[i] && !Number.isNaN(Date.parse(supplied[i]))) return supplied[i];
    const qn=alphaQuarterNumber(q); if(qn==null) return null;
    const year=Math.floor(qn/4), qi=qn%4;
    const month=[3,6,9,12][qi];
    return `${year}-${String(month).padStart(2,'0')}-${month===3?'31':month===6?'30':month===9?'30':'31'}`;
  });
}
function alphaPriceDenseSeries(payload){
  const rows=alphaArray(payload?.price_history_3y)
    .filter(p=>p&&p.date!=null&&Number.isFinite(Number(p.close))&&!Number.isNaN(Date.parse(String(p.date))))
    .map(p=>({date:String(p.date),close:Number(p.close)}))
    .sort((a,b)=>Date.parse(a.date)-Date.parse(b.date));
  const raw=rows.map(p=>p.close);
  return {dates:rows.map(p=>p.date),raw,z:zNormalize(raw)};
}
function alphaDayOffset(date,originMs){
  const ms=Date.parse(String(date));
  return Number.isFinite(ms)&&Number.isFinite(originMs)?(ms-originMs)/86400000:null;
}
function alphaSetOverlayMode(mode){
  alphaOverlayMode=mode==='2d'?'2d':'3d';
  $('#alpha-view-3d')?.classList.toggle('active',alphaOverlayMode==='3d');
  $('#alpha-view-2d')?.classList.toggle('active',alphaOverlayMode==='2d');
  const note=$('#alpha-overlay-axis-note'), help=$('#alpha-overlay-help'), plot=$('#alpha-overlay-plot');
  if(note) note.textContent=alphaOverlayMode==='3d'?'3D · X=DATE · Y=FACTOR/PRICE · Z=Z-SCORE':'2D · X=DATE · Y=NORMALIZED Z-SCORE';
  if(help) help.textContent=alphaOverlayMode==='3d'?'DRAG: ROTATE · SHIFT+DRAG: PAN · WHEEL: ZOOM · DOUBLE CLICK: RESET · PRICE=DAILY PATH':'DRAG: PAN · WHEEL: ZOOM · DOUBLE CLICK: RESET · PRICE=DAILY PATH · FACTORS=QUARTERLY';
  if(plot){plot.classList.toggle('alpha-overlay-3d',alphaOverlayMode==='3d');plot.classList.toggle('alpha-overlay-2d',alphaOverlayMode==='2d');}
  if(lastAlphaPayload) renderAlphaOverlay(lastAlphaPayload,selectedAlphaFactors());
}
function renderAlphaOverlay(payload,selected){
  const el=$('#alpha-overlay-plot'); if(!el) return;
  if(!plotlyReady()){el.innerHTML='<div class="deck-empty">PLOTLY ENGINE UNAVAILABLE</div>';return;}
  if(!selected.length){purgePlot(el);el.innerHTML='<div class="deck-empty">SELECT AT LEAST ONE FACTOR</div>';return;}
  const catalog=alphaCatalogIndex(payload), quarters=alphaQuarters(payload);
  if(!quarters.length){purgePlot(el);el.innerHTML='<div class="deck-empty">NO QUARTER AXIS RETURNED BY MODEL</div>';return;}

  const quarterDates=alphaQuarterDateSeries(payload,quarters);
  const factorSeries=[];
  selected.forEach((key,i)=>{
    const raw=alphaSeries(payload,key,quarters), z=zNormalize(raw), meta=catalog[key]||{label:key,unit:'raw'};
    if(!z.some(Number.isFinite)) return;
    factorSeries.push({key,label:meta.label||key,meta,raw,z,color:ALPHA_PALETTE[i%ALPHA_PALETTE.length]});
  });
  // Keep price at the native 3-year market-data frequency (daily in v5.7.5).
  // Factors remain quarterly; the two are intentionally not downsampled to the same cadence.
  const price=alphaPriceDenseSeries(payload);
  const hasPrice=price.z.some(Number.isFinite);
  if(!factorSeries.length&&!hasPrice){purgePlot(el);el.innerHTML='<div class="deck-empty">NO OBSERVATIONS FOR THE SELECTED FACTORS</div>';return;}

  purgePlot(el); el.innerHTML=''; el.style.background=PLOT_BG;
  const config={responsive:true,scrollZoom:true,displaylogo:false,doubleClick:'reset',displayModeBar:true,
    modeBarButtonsToRemove:['toImage','sendDataToCloud','lasso3d','select2d']};

  if(alphaOverlayMode==='2d'){
    const traces=factorSeries.map((f,i)=>({
      type:'scatter',mode:'lines+markers',name:f.label,x:quarterDates,y:f.z,connectgaps:false,
      line:{width:1.6,color:f.color},marker:{size:4,color:f.color},
      customdata:f.raw.map(v=>alphaValueText(v,f.meta)),
      hovertemplate:`${escapeHtml(f.label)}<br>%{x}<br>Z %{y:.2f}<br>RAW %{customdata}<extra></extra>`
    }));
    if(hasPrice) traces.push({
      type:'scatter',mode:'lines',name:`${payload.ticker||'PRICE'} PRICE · DAILY`,x:price.dates,y:price.z,connectgaps:false,
      line:{width:2.8,color:'#f1f5f1'},
      customdata:price.raw.map(v=>Number.isFinite(v)?`$${Number(v).toFixed(2)}`:'—'),
      hovertemplate:`${escapeHtml(payload.ticker||'')} PRICE<br>%{x}<br>Z %{y:.2f}<br>PRICE %{customdata}<extra></extra>`
    });
    const layout={autosize:true,margin:{l:52,r:18,t:12,b:38},paper_bgcolor:PLOT_BG,plot_bgcolor:PLOT_BG,
      font:{family:'Cascadia Mono, monospace',size:8,color:'#788578'},hovermode:'closest',showlegend:true,
      legend:{orientation:'h',x:0,y:1.08,font:{size:7,color:'#9ba79b'},bgcolor:'rgba(5,8,5,.55)'},
      xaxis:{title:'DATE · PRICE DAILY / FACTORS QUARTERLY',type:'date',gridcolor:'#1b241b',zeroline:false,color:'#7a877a',tickfont:{size:8},fixedrange:false,automargin:true},
      yaxis:{title:'NORMALIZED Z-SCORE',range:[-3.5,3.5],dtick:1,gridcolor:'#202a20',zeroline:true,zerolinecolor:'#657065',color:'#7a877a',tickfont:{size:8},fixedrange:false,automargin:true}};
    alphaPlot(el,traces,layout,{displayModeBar:true});
    return;
  }

  // scatter3d is most stable with numeric X values. Use elapsed days from the
  // first price observation, so daily price points form a continuous path while
  // quarterly factors sit on their actual quarter-end dates.
  const validDates=[...price.dates,...quarterDates].filter(d=>d&&!Number.isNaN(Date.parse(d)));
  const originMs=validDates.length?Math.min(...validDates.map(d=>Date.parse(d))):Date.now();
  const priceX=price.dates.map(d=>alphaDayOffset(d,originMs));
  const quarterX=quarterDates.map(d=>alphaDayOffset(d,originMs));
  const traces=[], names=[];
  factorSeries.forEach((f,i)=>{
    const factorIndex=names.length, xs=[],ys=[],zs=[],hover=[];
    for(let q=0;q<quarters.length;q++){
      if(!Number.isFinite(f.z[q])||!Number.isFinite(quarterX[q])) continue;
      xs.push(Number(quarterX[q]));ys.push(factorIndex);zs.push(Number(f.z[q]));
      hover.push(`${f.label}<br>${quarters[q]} · ${quarterDates[q]||''}<br>Z ${Number(f.z[q]).toFixed(2)}<br>RAW ${alphaValueText(f.raw[q],f.meta)}`);
    }
    names.push(f.label);
    traces.push({type:'scatter3d',mode:zs.length>1?'lines+markers':'markers',name:f.label,x:xs,y:ys,z:zs,text:hover,hoverinfo:'text',
      line:{width:5,color:f.color},marker:{size:4,color:f.color,symbol:'circle',opacity:1}});
  });
  if(hasPrice){
    const yi=names.length, xs=[],ys=[],zs=[],hover=[];
    for(let i=0;i<price.z.length;i++){
      if(!Number.isFinite(price.z[i])||!Number.isFinite(priceX[i])) continue;
      xs.push(Number(priceX[i]));ys.push(yi);zs.push(Number(price.z[i]));
      hover.push(`${payload.ticker||''} PRICE<br>${price.dates[i]}<br>Z ${Number(price.z[i]).toFixed(2)}<br>PRICE $${Number(price.raw[i]).toFixed(2)}`);
    }
    names.push(`${payload.ticker||'PRICE'} PRICE · DAILY`);
    traces.push({type:'scatter3d',mode:'lines',name:`${payload.ticker||''} PRICE · DAILY`,x:xs,y:ys,z:zs,text:hover,hoverinfo:'text',
      line:{width:8,color:'#f1f5f1'}});
  }
  const allX=[...priceX.filter(Number.isFinite),...quarterX.filter(Number.isFinite)];
  const minX=allX.length?Math.min(...allX):0, maxX=allX.length?Math.max(...allX):1, pad=Math.max(10,(maxX-minX)*0.025);
  const tickVals=[],tickText=[];
  quarterX.forEach((v,i)=>{if(Number.isFinite(v)){tickVals.push(v);tickText.push(quarters[i]);}});
  const layout={autosize:true,margin:{l:0,r:0,t:8,b:0},paper_bgcolor:PLOT_BG,plot_bgcolor:PLOT_BG,
    font:{family:'Cascadia Mono, monospace',size:8,color:'#788578'},showlegend:false,
    scene:{bgcolor:PLOT_BG,dragmode:'orbit',aspectmode:'manual',aspectratio:{x:2.15,y:Math.max(.8,Math.min(1.8,names.length/7)),z:1.15},camera:{eye:{x:1.48,y:1.58,z:1.10}},
      xaxis:{title:'DATE · PRICE DAILY / FACTORS QUARTERLY',tickmode:'array',tickvals:tickVals,ticktext:tickText,range:[minX-pad,maxX+pad],showbackground:true,backgroundcolor:'#071007',gridcolor:'#283128',zerolinecolor:'#354035',color:'#829082',titlefont:{size:8},tickfont:{size:7}},
      yaxis:{title:'FACTOR / PRICE',tickmode:'array',tickvals:names.map((_,i)=>i),ticktext:names,range:[-.5,Math.max(.5,names.length-.5)],showbackground:true,backgroundcolor:'#070a07',gridcolor:'#222922',zerolinecolor:'#354035',color:'#829082',titlefont:{size:9},tickfont:{size:6}},
      zaxis:{title:'NORMALIZED Z-SCORE',range:[-3.5,3.5],dtick:1,showbackground:true,backgroundcolor:'#070a07',gridcolor:'#2b362b',zerolinecolor:'#637063',color:'#829082',titlefont:{size:9},tickfont:{size:7}}}};
  try{
    Plotly.newPlot(el,traces,layout,config).then(()=>requestAnimationFrame(()=>{try{Plotly.Plots.resize(el);Plotly.relayout(el,{'paper_bgcolor':PLOT_BG,'plot_bgcolor':PLOT_BG,'scene.bgcolor':PLOT_BG});}catch(_){}})).catch(err=>alphaPlotError(el,err));
  }catch(err){alphaPlotError(el,err);}
}

function renderAlphaFactorDeck(payload,selected){
  const deck=$('#alpha-factor-deck'); if(!deck) return;
  if(plotlyReady()) $$('.alpha-factor-chart',deck).forEach(el=>purgePlot(el));
  if(!selected.length){deck.innerHTML='<div class="deck-empty">SELECT AT LEAST ONE FACTOR</div>';return;}
  const catalog=alphaCatalogIndex(payload), quarters=alphaQuarters(payload);
  deck.innerHTML=selected.map((key,i)=>{
    const meta=catalog[key]||{label:key,unit:'raw',group:'FACTOR'};
    const coverage=(payload.diagnostics?.[key]?.coverage??0);
    return `<article class="alpha-factor-card"><div class="alpha-factor-card-head"><span>${escapeHtml(meta.label||key)}</span><em>${escapeHtml(meta.group||'FACTOR')} · N=${coverage}</em></div><div class="alpha-factor-chart" id="alpha-factor-${i}"></div></article>`;
  }).join('');
  if(!plotlyReady()) return;
  selected.forEach((key,i)=>{
    const meta=catalog[key]||{label:key,unit:'raw'};
    const raw=alphaSeries(payload,key,quarters), vals=alphaPlotValues(raw,meta.unit);
    const el=$(`#alpha-factor-${i}`); if(!el)return;
    const finite=vals.filter(v=>v!=null&&Number.isFinite(Number(v)));
    if(!finite.length){el.innerHTML='<div class="deck-empty">NO HISTORICAL DATA EXPOSED BY YAHOO</div>';return;}
    const suffix=alphaAxisSuffix(meta);
    const trace={type:'scatter',mode:'lines+markers',x:quarters,y:vals,name:meta.label||key,connectgaps:false,
      line:{width:1.8,color:ALPHA_PALETTE[i%ALPHA_PALETTE.length]},marker:{size:4,color:ALPHA_PALETTE[i%ALPHA_PALETTE.length]},
      hovertemplate:`${escapeHtml(meta.label||key)}<br>%{x}<br>%{y:.2f}${suffix}<extra></extra>`};
    const layout={margin:{l:55,r:10,t:8,b:30},showlegend:false,hovermode:'x unified',
      xaxis:{gridcolor:'#182018',zeroline:false,color:'#687468',tickfont:{size:8},fixedrange:false,automargin:true},
      yaxis:{title:alphaAxisTitle(meta),gridcolor:'#182018',zeroline:true,zerolinecolor:'#303a30',color:'#687468',tickfont:{size:7},titlefont:{size:7},ticksuffix:suffix,fixedrange:false,automargin:true}};
    alphaPlot(el,[trace],layout,{displayModeBar:false});
  });
}
function renderAlphaPrice(payload){
  const el=$('#alpha-price-plot'); if(!el)return;
  const h=alphaArray(payload?.price_history_3y).filter(p=>p&&p.date!=null&&Number.isFinite(Number(p.close)));
  if(!plotlyReady()){el.innerHTML='<div class="deck-empty">PLOTLY ENGINE UNAVAILABLE</div>';return;}
  if(!h.length){purgePlot(el);el.innerHTML='<div class="deck-empty">3-YEAR PRICE HISTORY UNAVAILABLE</div>';return;}
  const trace={type:'scatter',mode:'lines',x:h.map(p=>String(p.date)),y:h.map(p=>Number(p.close)),name:payload.ticker||'',
    line:{width:1.7,color:'#a8ff32'},fill:'tozeroy',fillcolor:'rgba(168,255,50,.035)',
    hovertemplate:`${escapeHtml(payload.ticker||'')}<br>%{x}<br>$%{y:.2f}<extra></extra>`};
  const layout={margin:{l:50,r:12,t:12,b:30},showlegend:false,hovermode:'x unified',
    xaxis:{gridcolor:'#182018',zeroline:false,color:'#687468',tickfont:{size:8},rangeslider:{visible:false},fixedrange:false,automargin:true},
    yaxis:{title:'PRICE',gridcolor:'#182018',zeroline:false,color:'#687468',tickfont:{size:8},titlefont:{size:8},fixedrange:false,automargin:true}};
  alphaPlot(el,[trace],layout,{displayModeBar:false});
}
function renderAlphaDiagnostics(payload,selected){
  const body=$('#alpha-diagnostics'); if(!body)return;
  const catalog=alphaCatalogIndex(payload);
  if(!selected.length){body.innerHTML='<tr><td colspan="5" class="screener-empty">NO FACTORS SELECTED</td></tr>';return;}
  body.innerHTML=selected.map(key=>{
    const meta=catalog[key]||{label:key,unit:'raw'}, d=payload.diagnostics?.[key]||{};
    const corr=Number(d.forward_3m_spearman), corrClass=Number.isFinite(corr)?(corr>0?'alpha-corr-pos':corr<0?'alpha-corr-neg':''):'';
    return `<tr><td title="${escapeHtml(meta.description||'')}">${escapeHtml(meta.label)}</td><td>${alphaValueText(d.latest,meta)}</td><td>${d.zscore==null?'—':num(d.zscore,2)}</td><td class="${corrClass}">${d.forward_3m_spearman==null?'—':num(d.forward_3m_spearman,2)}</td><td>${d.coverage??d.observations??0}</td></tr>`;
  }).join('');
}
function renderAlphaPayload(payload){
  if(payload?.catalog){alphaCatalogCache=payload.catalog;if(!$('#alpha-factor-selector-groups')?.dataset.ready)renderAlphaSelector(payload.catalog)}
  const selected=selectedAlphaFactors();
  $('#alpha-symbol').textContent=payload.ticker||'—';
  $('#alpha-company').textContent=payload.company_name||payload.ticker||'—';
  $('#alpha-meta').textContent=[payload.sector,payload.industry].filter(Boolean).join(' // ')||'YAHOO FINANCE';
  $('#alpha-price-qtrs').textContent=payload.availability?.price_quarters??'—';
  $('#alpha-fund-qtrs').textContent=payload.availability?.fundamental_quarters??'—';
  $('#alpha-factor-count').textContent=`${selected.length}/${payload.availability?.available_factors??'—'}`;
  $('#alpha-note').textContent=payload.note||'';
  // Render the WebGL 3D scene first, exactly like F4 Screener. Stagger the
  // SVG factor/price charts so they cannot contend with the initial WebGL scene.
  renderAlphaOverlay(payload,selected);
  renderAlphaDiagnostics(payload,selected);
  setTimeout(()=>renderAlphaFactorDeck(payload,selected),60);
  setTimeout(()=>renderAlphaPrice(payload),120);
  setTimeout(alphaResizePlots,220);
  setTimeout(alphaResizePlots,520);
}
function alphaResizePlots(){
  if(!plotlyReady()) return;
  ['#alpha-overlay-plot','#alpha-price-plot'].forEach(sel=>{const el=$(sel);if(el&&el.data)try{Plotly.Plots.resize(el)}catch(_){}});
  $$('.alpha-factor-chart').forEach(el=>{if(el.data)try{Plotly.Plots.resize(el)}catch(_){}});
}
async function runAlphaModel(){
  if(alphaInFlight) return;
  const ticker=$('#alpha-ticker')?.value.trim().toUpperCase();
  if(!ticker){toast('Enter a ticker.',true);return;}
  const selected=selectedAlphaFactors();
  if(!selected.length){toast('Select at least one alpha factor.',true);return;}
  alphaInFlight=true; const btn=$('#run-alpha'); setButtonBusy(btn,true,'BUILDING FACTORS…');
  $('#alpha-status').textContent='LOADING'; setStatus(`ALPHA MODEL ${ticker}`);
  purgeAlphaPlots();
  const top=$('#alpha-overlay-plot'); if(top) top.innerHTML='<div class="deck-empty">STREAMING PRICE + QUARTERLY FUNDAMENTALS…</div>';
  try{
    const payload=await api(`/api/alpha-factors/${encodeURIComponent(ticker)}`);
    lastAlphaPayload=payload; renderAlphaPayload(payload);
    $('#alpha-status').textContent='COMPLETE'; setStatus(`${ticker} FACTOR MODEL COMPLETE`); log(`${ticker} expanded quarterly alpha-factor model loaded.`,'ALPHA');
  }catch(err){
    $('#alpha-status').textContent='ERROR'; setStatus('ALPHA MODEL ERROR'); toast(err.message,true); log(err.message,'ERR');
    const el=$('#alpha-overlay-plot'); if(el)el.innerHTML=`<div class="deck-empty">MODEL ERROR · ${escapeHtml(err.message)}</div>`;
  }finally{alphaInFlight=false;setButtonBusy(btn,false)}
}
$('#alpha-defaults')?.addEventListener('click',()=>{setAlphaSelections(ALPHA_DEFAULTS);if(lastAlphaPayload)renderAlphaPayload(lastAlphaPayload)});
$('#alpha-select-all')?.addEventListener('click',()=>{setAlphaSelections($$('[data-alpha-factor]').map(el=>el.dataset.alphaFactor));if(lastAlphaPayload)renderAlphaPayload(lastAlphaPayload)});
$('#alpha-clear')?.addEventListener('click',()=>{setAlphaSelections([]);if(lastAlphaPayload)renderAlphaPayload(lastAlphaPayload)});
$('#alpha-view-3d')?.addEventListener('click',()=>alphaSetOverlayMode('3d'));
$('#alpha-view-2d')?.addEventListener('click',()=>alphaSetOverlayMode('2d'));
$('#run-alpha')?.addEventListener('click',runAlphaModel);
$('#alpha-ticker')?.addEventListener('keydown',e=>{if(e.key==='Enter')runAlphaModel()});
window.addEventListener('resize',alphaResizePlots);
loadAlphaCatalog();

// ==============================
// F6 // PEER ANALYSIS
// ==============================
let peerAnalysisInFlight=false;
let lastPeerAnalysisPayload=null;

function peerPctRatio(v,d=1){
  return v==null||Number.isNaN(Number(v))?'—':(Number(v)*100).toFixed(d)+'%';
}
function peerPctPoint(v,d=1){
  return v==null||Number.isNaN(Number(v))?'—':Number(v).toFixed(d)+'%';
}
function peerMultiple(v,d=1){
  return v==null||Number.isNaN(Number(v))?'—':Number(v).toFixed(d)+'x';
}
function peerPlotLayout(extra={}){
  return Object.assign({
    margin:{l:56,r:18,t:14,b:46},paper_bgcolor:'#050805',plot_bgcolor:'#050805',
    font:{family:'Cascadia Mono, monospace',size:9,color:'#97a298'},
    legend:{orientation:'h',x:0,y:1.14,font:{size:8,color:'#a4aea5'}},
    xaxis:{gridcolor:'#182018',zeroline:false,color:'#687468',tickfont:{size:8},automargin:true},
    yaxis:{gridcolor:'#182018',zerolinecolor:'#333a33',color:'#687468',tickfont:{size:8},automargin:true},
    hoverlabel:{bgcolor:'#111711',bordercolor:'#526052',font:{family:'Cascadia Mono, monospace',size:9,color:'#e7eee8'}},
    hovermode:'closest'
  },extra);
}
function clearPeerPlot(target,message='NO DATA AVAILABLE'){
  const el=$(target); if(!el) return;
  if(plotlyReady()){try{Plotly.purge(el)}catch(_){}}
  el.innerHTML=`<div class="deck-empty">${escapeHtml(message)}</div>`;
}
function renderPeerGroupedChart(target,companies,metrics,{percent=false}={}){
  const el=$(target); if(!el) return;
  if(!plotlyReady()){el.innerHTML='<div class="deck-empty">PLOTLY ENGINE UNAVAILABLE</div>';return;}
  const tickers=companies.map(c=>c.ticker);
  const traces=metrics.map(m=>({
    type:'bar',name:m.label,x:tickers,
    y:companies.map(c=>{
      const v=c[m.key];
      if(v==null||Number.isNaN(Number(v))) return null;
      return percent ? Number(v)*(m.points?1:100) : Number(v);
    }),
    hovertemplate:`%{x}<br>${escapeHtml(m.label)}: %{y:,.2f}${percent?'%':'x'}<extra></extra>`
  }));
  try{Plotly.purge(el)}catch(_){}
  const layout=peerPlotLayout({barmode:'group',hovermode:'x unified'});
  if(percent) layout.yaxis.ticksuffix='%';
  Plotly.newPlot(el,traces,layout,{responsive:true,displaylogo:false,displayModeBar:false}).then(()=>{try{Plotly.Plots.resize(el)}catch(_){}});
}
function renderPeerRevenueChart(target,companies,key){
  const el=$(target); if(!el) return;
  if(!plotlyReady()){el.innerHTML='<div class="deck-empty">PLOTLY ENGINE UNAVAILABLE</div>';return;}
  const traces=companies.map(c=>({
    type:'scatter',mode:'lines+markers',name:c.ticker,
    x:(c[key]||[]).map(p=>p.period),y:(c[key]||[]).map(p=>p.value),
    line:{width:1.8},marker:{size:5},connectgaps:false,
    hovertemplate:`${escapeHtml(c.ticker)}<br>%{x}<br>%{y:,.1f} MM<extra></extra>`
  })).filter(t=>t.x.length);
  if(!traces.length){clearPeerPlot(target,'NO REVENUE HISTORY RETURNED');return;}
  try{Plotly.purge(el)}catch(_){}
  const layout=peerPlotLayout({hovermode:'x unified'});
  layout.yaxis.title={text:'Revenue (MM)',font:{size:8,color:'#687468'}};
  Plotly.newPlot(el,traces,layout,{responsive:true,displaylogo:false,displayModeBar:false}).then(()=>{try{Plotly.Plots.resize(el)}catch(_){}});
}
function renderPeerPriceChart(companies){
  const el=$('#peer-price-chart'); if(!el) return;
  if(!plotlyReady()){el.innerHTML='<div class="deck-empty">PLOTLY ENGINE UNAVAILABLE</div>';return;}
  const traces=companies.map((c,i)=>({
    type:'scatter',mode:'lines',name:c.ticker,
    x:(c.normalized_price||[]).map(p=>p.date),y:(c.normalized_price||[]).map(p=>p.value),
    line:{width:i===0?2.8:1.7},connectgaps:false,
    hovertemplate:`${escapeHtml(c.ticker)}<br>%{x}<br>Indexed: %{y:.2f}<extra></extra>`
  })).filter(t=>t.x.length);
  if(!traces.length){clearPeerPlot('#peer-price-chart','NO PRICE HISTORY RETURNED');return;}
  try{Plotly.purge(el)}catch(_){}
  const layout=peerPlotLayout({hovermode:'x unified',margin:{l:60,r:20,t:14,b:42}});
  layout.yaxis.title={text:'Indexed price (start = 100)',font:{size:8,color:'#687468'}};
  layout.yaxis.zeroline=false;
  layout.shapes=[{type:'line',xref:'paper',x0:0,x1:1,y0:100,y1:100,line:{color:'#596459',width:1,dash:'dot'}}];
  Plotly.newPlot(el,traces,layout,{responsive:true,scrollZoom:true,displaylogo:false,displayModeBar:true,doubleClick:'reset'}).then(()=>{try{Plotly.Plots.resize(el)}catch(_){}});
}
function peerPeriodSort(period,quarterly=false){
  if(!quarterly){const m=String(period).match(/FY(\d{4})/);return m?Number(m[1])*10:0;}
  const m=String(period).match(/(\d)Q(\d{4})/);return m?Number(m[2])*10+Number(m[1]):0;
}
function renderPeerRevenueTable(companies,key,headSelector,bodySelector,quarterly=false){
  const head=$(headSelector),body=$(bodySelector); if(!head||!body) return;
  const periods=[...new Set(companies.flatMap(c=>(c[key]||[]).map(x=>x.period)))].sort((a,b)=>peerPeriodSort(a,quarterly)-peerPeriodSort(b,quarterly));
  head.innerHTML=`<tr><th>COMPANY</th><th>CCY</th>${periods.map(p=>`<th>${escapeHtml(p)}</th>`).join('')}</tr>`;
  body.innerHTML='';
  companies.forEach(c=>{
    const values=Object.fromEntries((c[key]||[]).map(x=>[x.period,x.value]));
    const tr=document.createElement('tr');
    tr.innerHTML=`<td><b>${escapeHtml(c.ticker)}</b><small>${escapeHtml(c.company_name||'')}</small></td><td>${escapeHtml(c.currency||'—')}</td>${periods.map(p=>`<td class="${numericClass(values[p])}">${values[p]==null?'—':num(values[p],1)}</td>`).join('')}`;
    body.appendChild(tr);
  });
  if(!companies.length) body.innerHTML='<tr><td colspan="3" class="screener-empty">NO DATA</td></tr>';
}
function renderPeerMatrix(companies){
  const body=$('#peer-matrix-body'); if(!body) return;
  body.innerHTML='';
  companies.forEach((c,i)=>{
    const tr=document.createElement('tr'); if(i===0) tr.className='peer-subject-row';
    tr.innerHTML=`
      <td><div class="peer-company-cell"><b>${escapeHtml(c.ticker)}</b><span>${escapeHtml(c.company_name||'')}</span><small>${escapeHtml(c.industry||'')}</small></div></td>
      <td>${c.current_price==null?'—':num(c.current_price,2)}</td>
      <td>${money(c.market_cap,c.currency||'')}</td><td>${money(c.enterprise_value,c.currency||'')}</td><td class="${numericClass(c.net_debt)}">${money(c.net_debt,c.currency||'')}</td>
      <td>${peerMultiple(c.trailing_pe)}</td><td>${peerMultiple(c.forward_pe)}</td><td>${peerMultiple(c.ev_revenue)}</td><td>${peerMultiple(c.ev_ebitda)}</td><td>${peerMultiple(c.price_sales)}</td><td>${peerMultiple(c.price_book)}</td>
      <td class="${growthClass(c.revenue_growth)}">${peerPctRatio(c.revenue_growth)}</td><td>${peerPctRatio(c.ebitda_margin)}</td><td>${peerPctRatio(c.operating_margin)}</td><td>${peerPctRatio(c.net_margin)}</td><td>${peerPctRatio(c.roe)}</td><td>${peerPctRatio(c.roa)}</td>
      <td class="${numericClass(c.net_debt_ebitda)}">${peerMultiple(c.net_debt_ebitda)}</td><td class="${growthClass(c.fcf_yield==null?null:c.fcf_yield/100)}">${peerPctPoint(c.fcf_yield)}</td><td>${num(c.beta,2)}</td>`;
    body.appendChild(tr);
  });
  if(!companies.length) body.innerHTML='<tr><td colspan="20" class="screener-empty">NO PEER DATA RETURNED</td></tr>';
}
function renderPeerAnalysis(payload){
  lastPeerAnalysisPayload=payload;
  const companies=payload?.companies||[];
  $('#peer-company-count').textContent=`${companies.length} COMPANIES · ${Math.max(0,companies.length-1)} PEERS`;
  $('#peer-status').textContent=companies.length?'LOADED':'EMPTY';
  renderPeerMatrix(companies);
  renderPeerGroupedChart('#peer-multiples-chart',companies,[
    {key:'forward_pe',label:'Forward P/E'},{key:'ev_revenue',label:'EV/Revenue'},{key:'ev_ebitda',label:'EV/EBITDA'}
  ]);
  renderPeerGroupedChart('#peer-quality-chart',companies,[
    {key:'ebitda_margin',label:'EBITDA Margin'},{key:'operating_margin',label:'Operating Margin'},{key:'roe',label:'ROE'},{key:'fcf_yield',label:'FCF Yield',points:true}
  ],{percent:true});
  renderPeerRevenueChart('#peer-revenue-fy-chart',companies,'annual_revenue');
  renderPeerRevenueChart('#peer-revenue-q-chart',companies,'quarterly_revenue');
  renderPeerRevenueTable(companies,'annual_revenue','#peer-fy-head','#peer-fy-body',false);
  renderPeerRevenueTable(companies,'quarterly_revenue','#peer-q-head','#peer-q-body',true);
  renderPeerPriceChart(companies);
  const note=$('#peer-analysis-notes');
  if(note){
    const warnings=payload?.warnings||[],notes=payload?.notes||[];
    note.innerHTML=[...notes,...warnings.map(w=>'WARNING · '+w)].map(x=>`<span>${escapeHtml(x)}</span>`).join('');
  }
  requestAnimationFrame(()=>setTimeout(peerResizePlots,60));
}
function peerResizePlots(){
  if(!plotlyReady()) return;
  ['#peer-multiples-chart','#peer-quality-chart','#peer-revenue-fy-chart','#peer-revenue-q-chart','#peer-price-chart'].forEach(sel=>{
    const el=$(sel); if(!el) return; try{if(el.data) Plotly.Plots.resize(el)}catch(_){}
  });
}
function syncPeerSubject(value){
  const v=String(value||'').trim().toUpperCase();
  if($('#peer-subject')) $('#peer-subject').value=v;
  if($('#peer-input-0')) $('#peer-input-0').value=v;
}
function collectPeerInputs(){
  const subject=($('#peer-input-0')?.value||$('#peer-subject')?.value||'').trim().toUpperCase();
  const peers=[];
  for(let i=1;i<=7;i++){
    const v=($(`#peer-input-${i}`)?.value||'').trim().toUpperCase();
    if(v && v!==subject && !peers.includes(v)) peers.push(v);
  }
  return {subject,peers};
}
async function runPeerAnalysis(){
  if(peerAnalysisInFlight) return;
  const {subject,peers}=collectPeerInputs();
  if(!subject){toast('Enter a subject ticker.',true);return;}
  syncPeerSubject(subject);
  peerAnalysisInFlight=true; setButtonBusy($('#run-peer-analysis'),true,'LOADING PEERS…');
  $('#peer-status').textContent='LOADING'; setStatus(`PEER ANALYSIS ${subject}`); log(`Loading ${subject} + ${peers.length} analyst-selected peers.`,'REQ');
  try{
    const qs=peers.length?`?peers=${encodeURIComponent(peers.join(','))}`:'';
    const payload=await api(`/api/peer-analysis/${encodeURIComponent(subject)}${qs}`);
    renderPeerAnalysis(payload);
    setStatus(`${subject} PEER ANALYSIS LOADED`);
    log(`Peer analysis loaded: ${(payload.companies||[]).map(x=>x.ticker).join(', ')}.`,'OK');
    if((payload.warnings||[]).length) toast(`Loaded with ${(payload.warnings||[]).length} peer warning(s).`);
  }catch(err){
    $('#peer-status').textContent='ERROR'; setStatus('PEER ANALYSIS ERROR'); toast(err.message,true); log(err.message,'ERR');
  }finally{peerAnalysisInFlight=false; setButtonBusy($('#run-peer-analysis'),false);}
}

$('#run-peer-analysis')?.addEventListener('click',runPeerAnalysis);
$('#peer-subject')?.addEventListener('input',e=>syncPeerSubject(e.target.value));
$('#peer-input-0')?.addEventListener('input',e=>{if($('#peer-subject')) $('#peer-subject').value=String(e.target.value||'').toUpperCase();});
$('#peer-subject')?.addEventListener('keydown',e=>{if(e.key==='Enter')runPeerAnalysis();});
$$('.peer-input input').forEach(el=>el.addEventListener('keydown',e=>{if(e.key==='Enter')runPeerAnalysis();}));
$('#peer-clear')?.addEventListener('click',()=>{for(let i=1;i<=7;i++){const el=$(`#peer-input-${i}`);if(el)el.value='';}});

// === v5.11 Shared Virtual Trading Desk ===
let portfolioInFlight=false;
let portfolioTradeInFlight=false;
let portfolioQuoteInFlight=false;
let portfolioChartInFlight=false;
let portfolioLastPayload=null;
let portfolioLastQuote=null;
let portfolioChartPayload=null;
let portfolioInterval='1m';
let portfolioPollTimer=null;
let portfolioQuoteTimer=null;
let portfolioChartTimer=null;

function usd(v,d=2){
  if(v==null||Number.isNaN(Number(v))) return '—';
  const n=Number(v), sign=n<0?'-$':'$';
  return `${sign}${Math.abs(n).toLocaleString(undefined,{minimumFractionDigits:d,maximumFractionDigits:d})}`;
}
function portfolioSignedClass(v){ return Number(v)<0?'negative-number':Number(v)>0?'positive':''; }
function portfolioFormatEt(value){
  if(!value) return '—';
  try{return new Intl.DateTimeFormat('en-US',{timeZone:'America/New_York',month:'short',day:'2-digit',year:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}).format(new Date(value))+' ET'}catch(_){return String(value)}
}
function portfolioCompact(v){
  if(v==null||!Number.isFinite(Number(v))) return '—';
  return Number(v).toLocaleString(undefined,{notation:'compact',maximumFractionDigits:2});
}
function portfolioPct(v,d=2){
  if(v==null||!Number.isFinite(Number(v))) return '—';
  const n=Number(v);return `${n>=0?'+':''}${n.toFixed(d)}%`;
}
function renderPortfolioMarketStatus(status){
  const chip=$('#portfolio-market-chip'), ticket=$('#portfolio-ticket-market');
  const open=!!status?.is_open;
  if(chip){chip.textContent=open?'NYSE OPEN':'NYSE CLOSED';chip.classList.toggle('open',open);chip.classList.toggle('closed',!open);}
  if(ticket) ticket.textContent=open?'OPEN / REGULAR':'CLOSED';
  [$('#portfolio-buy'),$('#portfolio-sell')].forEach(btn=>{if(btn) btn.disabled=!open||portfolioTradeInFlight;});
  const rule=$('#portfolio-trade-rule');
  if(rule){
    const close=status?.session_close_et?` · Session close ${portfolioFormatEt(status.session_close_et)}`:'';
    rule.textContent=`New BUY/SELL orders are accepted only during the U.S. regular session${close}. Market orders are IOC. Limit orders may remain OPEN/PARTIAL and are checked by the server monitor. Yahoo/yfinance is not an exchange-certified real-time or Level-II feed.`;
  }
}
function renderPortfolioQuote(q){
  portfolioLastQuote=q;
  if(!q) return;
  $('#portfolio-ticket-price').textContent=usd(q.price);
  $('#portfolio-ticket-bid').textContent=usd(q.bid);
  $('#portfolio-ticket-ask').textContent=usd(q.ask);
  $('#portfolio-ticket-spread').textContent=q.spread_bps==null?'—':`${num(q.spread_bps,2)} bp`;
  $('#portfolio-spread-source').textContent=q.spread_source==='YAHOO_BID_ASK'?'YAHOO BID/ASK':'INDICATIVE MODEL';
  $('#portfolio-bid-size').textContent=q.bid_size?`SIZE ${portfolioCompact(q.bid_size)}`:'SIZE N/A';
  $('#portfolio-ask-size').textContent=q.ask_size?`SIZE ${portfolioCompact(q.ask_size)}`:'SIZE N/A';
  $('#portfolio-ticket-volume').textContent=portfolioCompact(q.volume);
  $('#portfolio-avg-volume').textContent=q.avg_volume?`ADV ${portfolioCompact(q.avg_volume)}`:'ADV —';
  $('#portfolio-minute-volume').textContent=portfolioCompact(q.minute_volume);
  const dc=$('#portfolio-day-change'); if(dc){dc.textContent=q.day_change_pct==null?'DAY —':`DAY ${portfolioPct(q.day_change_pct)}`;dc.className=portfolioSignedClass(q.day_change_pct);}
  $('#portfolio-quote-time').textContent=portfolioFormatEt(q.quoted_at);
  $('#portfolio-quote-disclaimer').textContent=q.spread_source==='YAHOO_BID_ASK'?'Yahoo bid/ask available · not Level II':'Bid/ask unavailable · indicative spread model active';
  if($('#portfolio-order-type')?.value==='LIMIT'&&!$('#portfolio-limit-price')?.value&&q.price) $('#portfolio-limit-price').value=Number(q.price).toFixed(2);
  renderPortfolioMarketStatus(q.market_status||portfolioLastPayload?.market_status||{});
  updatePortfolioTicketNotional();
  updatePortfolioLivePriceLine();
}
function renderPortfolioPositions(rows){
  const body=$('#portfolio-positions-body'); if(!body) return;
  $('#portfolio-position-count').textContent=`${rows.length} POSITION${rows.length===1?'':'S'}`;
  if(!rows.length){body.innerHTML='<tr><td colspan="12" class="screener-empty">NO OPEN POSITIONS</td></tr>';return;}
  body.innerHTML=rows.map(r=>`<tr>
    <td><b>${escapeHtml(r.ticker)}</b></td><td>${num(r.quantity,4)}</td><td>${num(r.available_to_sell,4)}</td>
    <td>${usd(r.avg_cost)}</td><td>${usd(r.last_price)}</td><td>${usd(r.market_value)}</td><td>${usd(r.cost_basis)}</td>
    <td class="${portfolioSignedClass(r.unrealized_pnl)}">${usd(r.unrealized_pnl)}</td>
    <td class="${portfolioSignedClass(r.return_pct)}">${portfolioPct(r.return_pct)}</td><td>${num(r.weight_pct,2)}%</td>
    <td class="${portfolioSignedClass(r.realized_pnl)}">${usd(r.realized_pnl)}</td><td>${escapeHtml(portfolioFormatEt(r.last_price_at))}</td>
  </tr>`).join('');
}
function portfolioOrderStatusClass(status){
  const s=String(status||'').toLowerCase();
  return s.includes('filled')?'filled':s.includes('cancel')||s.includes('expired')?'cancelled':s.includes('partial')?'partial':'open';
}
function renderPortfolioOpenOrders(rows){
  const body=$('#portfolio-orders-body'); if(!body) return;
  $('#portfolio-open-order-label').textContent=`${rows.length} ACTIVE · LIMIT ORDERS AUTO-CHECK WHILE SERVER IS RUNNING`;
  if(!rows.length){body.innerHTML='<tr><td colspan="13" class="screener-empty">NO OPEN ORDERS</td></tr>';return;}
  body.innerHTML=rows.map(r=>{
    const rem=Math.max(0,Number(r.quantity)-Number(r.filled_quantity||0));
    return `<tr><td>${r.id}</td><td>${escapeHtml(portfolioFormatEt(r.created_at))}</td>
      <td><span class="portfolio-side ${String(r.side).toLowerCase()}">${escapeHtml(r.side)}</span></td><td><b>${escapeHtml(r.ticker)}</b></td>
      <td>${escapeHtml(r.order_type)}</td><td>${num(r.quantity,4)}</td><td>${num(r.filled_quantity,4)}</td><td>${num(rem,4)}</td>
      <td>${r.limit_price==null?'—':usd(r.limit_price)}</td><td>${r.avg_fill_price==null?'—':usd(r.avg_fill_price)}</td><td>${escapeHtml(r.tif)}</td>
      <td><span class="portfolio-order-status ${portfolioOrderStatusClass(r.status)}">${escapeHtml(r.status)}</span></td>
      <td><button class="portfolio-cancel-order" data-order-id="${r.id}">CANCEL</button></td></tr>`;
  }).join('');
  $$('.portfolio-cancel-order').forEach(btn=>btn.addEventListener('click',()=>cancelPortfolioOrder(Number(btn.dataset.orderId))));
}
function renderPortfolioOrderHistory(rows){
  const body=$('#portfolio-order-history-body'); if(!body)return;
  if(!rows?.length){body.innerHTML='<tr><td colspan="13" class="screener-empty">NO ORDERS YET</td></tr>';return;}
  body.innerHTML=rows.map(r=>`<tr>
    <td>${r.id}</td><td>${escapeHtml(portfolioFormatEt(r.created_at))}</td><td><span class="portfolio-side ${String(r.side).toLowerCase()}">${escapeHtml(r.side)}</span></td>
    <td><b>${escapeHtml(r.ticker)}</b></td><td>${escapeHtml(r.order_type)}</td><td>${num(r.quantity,4)}</td><td>${num(r.filled_quantity,4)}</td>
    <td>${r.limit_price==null?'—':usd(r.limit_price)}</td><td>${r.avg_fill_price==null?'—':usd(r.avg_fill_price)}</td><td>${escapeHtml(r.tif)}</td>
    <td><span class="portfolio-order-status ${portfolioOrderStatusClass(r.status)}">${escapeHtml(r.status)}</span></td>
    <td class="portfolio-note-cell">${escapeHtml(r.note||'—')}</td><td class="portfolio-note-cell">${escapeHtml(r.cancel_reason||'—')}</td>
  </tr>`).join('');
}
function renderPortfolioTransactions(rows){
  const body=$('#portfolio-transactions-body'); if(!body) return;
  if(!rows?.length){body.innerHTML='<tr><td colspan="16" class="screener-empty">NO FILLS YET</td></tr>';return;}
  body.innerHTML=rows.map(r=>`<tr>
    <td>${r.id}</td><td>${r.order_id??'—'}</td><td>${escapeHtml(portfolioFormatEt(r.executed_at))}</td>
    <td><span class="portfolio-side ${String(r.side).toLowerCase()}">${escapeHtml(r.side)}</span></td><td><b>${escapeHtml(r.ticker)}</b></td>
    <td>${escapeHtml(r.order_type||'MARKET')}</td><td>${num(r.quantity,4)}</td><td>${usd(r.price)}</td><td>${usd(r.bid)}</td><td>${usd(r.ask)}</td>
    <td>${r.slippage_bps==null?'—':`${num(r.slippage_bps,2)} bp`}</td><td>${r.participation_pct==null?'—':`${num(r.participation_pct,2)}%`}</td>
    <td>${usd(r.gross_amount)}</td><td class="${portfolioSignedClass(r.realized_pnl)}">${usd(r.realized_pnl)}</td><td>${usd(r.cash_after)}</td><td class="portfolio-note-cell">${escapeHtml(r.note||'—')}</td>
  </tr>`).join('');
}
function renderPortfolioEquityChart(history){
  const el=$('#portfolio-equity-chart'); if(!el) return;
  if(!plotlyReady()){el.innerHTML='<div class="deck-empty">PLOTLY ENGINE UNAVAILABLE</div>';return;}
  if(!history?.length){el.innerHTML='<div class="deck-empty">NO SNAPSHOTS YET</div>';return;}
  const x=history.map(h=>h.captured_at);
  const layout={margin:{l:62,r:14,t:18,b:42},paper_bgcolor:'#050805',plot_bgcolor:'#050805',hovermode:'x unified',showlegend:true,
    legend:{orientation:'h',x:0,y:1.12,font:{family:'Cascadia Mono, monospace',size:8,color:'#8e998f'}},
    xaxis:{gridcolor:'#192119',color:'#697469',tickfont:{size:8}},yaxis:{gridcolor:'#192119',color:'#697469',tickfont:{size:8},tickprefix:'$',tickformat:',.0f'},
    hoverlabel:{bgcolor:'#111711',bordercolor:'#526052',font:{family:'Cascadia Mono, monospace',size:9,color:'#e7eee8'}}};
  const traces=[{x,y:history.map(h=>h.total_equity),type:'scatter',mode:'lines',name:'Total Equity',line:{width:2}},{x,y:history.map(h=>h.cash),type:'scatter',mode:'lines',name:'Cash',line:{width:1,dash:'dot'}}];
  Plotly.react(el,traces,layout,{responsive:true,displaylogo:false,scrollZoom:false});
}
function portfolioTradeMarkerTraces(ticker){
  const tx=(portfolioLastPayload?.transactions||[]).filter(r=>String(r.ticker).toUpperCase()===ticker);
  const buys=tx.filter(r=>r.side==='BUY'), sells=tx.filter(r=>r.side==='SELL');
  const traces=[];
  if(buys.length) traces.push({x:buys.map(r=>r.executed_at),y:buys.map(r=>r.price),type:'scatter',mode:'markers',name:'Buy fills',marker:{symbol:'triangle-up',size:9},text:buys.map(r=>`BUY ${num(r.quantity,4)} @ ${usd(r.price)}`),hovertemplate:'%{text}<extra></extra>',yaxis:'y'});
  if(sells.length) traces.push({x:sells.map(r=>r.executed_at),y:sells.map(r=>r.price),type:'scatter',mode:'markers',name:'Sell fills',marker:{symbol:'triangle-down',size:9},text:sells.map(r=>`SELL ${num(r.quantity,4)} @ ${usd(r.price)}`),hovertemplate:'%{text}<extra></extra>',yaxis:'y'});
  return traces;
}
function renderPortfolioMarketChart(payload){
  portfolioChartPayload=payload;
  const el=$('#portfolio-market-chart');if(!el)return;
  if(!plotlyReady()){el.innerHTML='<div class="deck-empty">PLOTLY ENGINE UNAVAILABLE</div>';return;}
  const rows=payload?.rows||[]; if(!rows.length){el.innerHTML='<div class="deck-empty">NO MARKET DATA</div>';return;}
  const x=rows.map(r=>r.time);
  const traces=[
    {x,open:rows.map(r=>r.open),high:rows.map(r=>r.high),low:rows.map(r=>r.low),close:rows.map(r=>r.close),type:'candlestick',name:payload.ticker,yaxis:'y',increasing:{line:{color:'#a8ff32'}},decreasing:{line:{color:'#d58b76'}}},
    {x,y:rows.map(r=>r.volume),type:'bar',name:'Volume',yaxis:'y2',opacity:.35,hovertemplate:'VOL %{y:,.0f}<extra></extra>'}
  ];
  if($('#portfolio-show-vwap')?.checked) traces.push({x,y:rows.map(r=>r.vwap),type:'scatter',mode:'lines',name:'VWAP',line:{width:1,dash:'dot'},yaxis:'y'});
  if($('#portfolio-show-sma')?.checked) traces.push({x,y:rows.map(r=>r.sma20),type:'scatter',mode:'lines',name:'SMA20',line:{width:1},yaxis:'y'});
  if($('#portfolio-show-ema')?.checked) traces.push({x,y:rows.map(r=>r.ema20),type:'scatter',mode:'lines',name:'EMA20',line:{width:1},yaxis:'y'});
  traces.push(...portfolioTradeMarkerTraces(payload.ticker));
  const last=portfolioLastQuote?.ticker===payload.ticker?Number(portfolioLastQuote.price):Number(rows[rows.length-1].close);
  const layout={margin:{l:60,r:54,t:22,b:38},paper_bgcolor:'#050805',plot_bgcolor:'#050805',hovermode:'x unified',showlegend:true,
    legend:{orientation:'h',x:0,y:1.08,font:{family:'Cascadia Mono, monospace',size:7,color:'#8e998f'}},
    xaxis:{gridcolor:'#172017',color:'#697469',rangeslider:{visible:false},tickfont:{size:8}},
    yaxis:{domain:[.27,1],side:'right',gridcolor:'#172017',color:'#697469',tickprefix:'$',tickformat:',.2f',tickfont:{size:8}},
    yaxis2:{domain:[0,.19],side:'right',gridcolor:'#111711',color:'#5f695f',tickformat:'~s',tickfont:{size:7}},
    hoverlabel:{bgcolor:'#111711',bordercolor:'#526052',font:{family:'Cascadia Mono, monospace',size:9,color:'#e7eee8'}},
    shapes:Number.isFinite(last)?[{type:'line',xref:'paper',x0:0,x1:1,yref:'y',y0:last,y1:last,line:{width:1,dash:'dot'}}]:[]};
  Plotly.react(el,traces,layout,{responsive:true,displaylogo:false,scrollZoom:true});
  $('#portfolio-chart-title').textContent=`${payload.ticker} · ${payload.interval_label}`;
  $('#portfolio-chart-updated').textContent=`UPDATED ${portfolioFormatEt(payload.updated_at)}`;
}
function updatePortfolioLivePriceLine(){
  const el=$('#portfolio-market-chart');
  if(!plotlyReady()||!el?.data||!portfolioLastQuote||portfolioChartPayload?.ticker!==portfolioLastQuote.ticker)return;
  const p=Number(portfolioLastQuote.price);if(!Number.isFinite(p))return;
  try{Plotly.relayout(el,{'shapes[0].y0':p,'shapes[0].y1':p});}catch(_){ }
}
function updatePortfolioTicketNotional(){
  const qty=Number($('#portfolio-quantity')?.value||0), px=Number(portfolioLastQuote?.price||0);
  const el=$('#portfolio-ticket-notional'); if(el) el.textContent=(qty>0&&px>0)?usd(qty*px):'—';
  const est=$('#portfolio-est-fill');
  if(est){
    const maxP=Number(portfolioLastPayload?.execution_model?.max_participation_pct||5)/100;
    const mv=Number(portfolioLastQuote?.minute_volume||0);
    const cap=mv>0?Math.max(1,mv*maxP):null;
    est.textContent=cap?`SIM MAX ≈ ${num(Math.min(qty,cap),2)} SH / CYCLE`:'LIQUIDITY EST. N/A';
  }
}
function renderPortfolio(payload){
  portfolioLastPayload=payload;
  const a=payload?.account||{};
  $('#portfolio-cash').textContent=usd(a.cash);$('#portfolio-buying-power').textContent=usd(a.buying_power);
  $('#portfolio-reserved-cash').textContent=`${usd(a.reserved_cash)} RESERVED`;$('#portfolio-cash-pct').textContent=`${num(a.cash_pct,1)}% OF NAV`;
  $('#portfolio-market-value').textContent=usd(a.positions_value);$('#portfolio-gross-exposure').textContent=`${num(a.gross_exposure_pct,1)}% GROSS`;
  $('#portfolio-equity').textContent=usd(a.total_equity);$('#portfolio-unrealized').textContent=usd(a.unrealized_pnl);$('#portfolio-realized').textContent=usd(a.realized_pnl);
  $('#portfolio-largest-position').textContent=`${num(a.largest_position_pct,2)}%`;$('#portfolio-open-order-count').textContent=String(a.open_order_count||0);
  $('#portfolio-unrealized').className=portfolioSignedClass(a.unrealized_pnl);$('#portfolio-realized').className=portfolioSignedClass(a.realized_pnl);$('#portfolio-equity').className=portfolioSignedClass(a.total_pnl);
  const ret=$('#portfolio-total-return');if(ret){ret.textContent=`${portfolioPct(a.total_return_pct)} SINCE INCEPTION`;ret.className=portfolioSignedClass(a.total_return_pct);}
  renderPortfolioMarketStatus(payload?.market_status||{});renderPortfolioOpenOrders(payload?.open_orders||[]);renderPortfolioPositions(payload?.positions||[]);
  renderPortfolioTransactions(payload?.transactions||[]);renderPortfolioOrderHistory(payload?.orders||[]);renderPortfolioEquityChart(payload?.history||[]);
  const model=payload?.execution_model||{};if($('#portfolio-max-participation'))$('#portfolio-max-participation').textContent=`${num(model.max_participation_pct||5,1)}%`;
  if($('#portfolio-execution-note'))$('#portfolio-execution-note').textContent=`${model.slippage||''}. ${model.partial_fill||''}. ${model.disclaimer||''}`;
  const warnings=$('#portfolio-warnings');if(warnings){warnings.innerHTML=(payload?.warnings||[]).map(w=>`<span>// EXECUTION / QUOTE WARNING · ${escapeHtml(w)}</span>`).join('');}
}
function portfolioResizePlot(){
  if(!plotlyReady()) return;['#portfolio-equity-chart','#portfolio-market-chart'].forEach(sel=>{const el=$(sel);try{if(el?.data)Plotly.Plots.resize(el)}catch(_){}});
}
async function loadPortfolio(refresh=true,silent=false){
  if(portfolioInFlight)return;
  portfolioInFlight=true;if(!silent)setButtonBusy($('#portfolio-refresh'),true,'REFRESHING…');
  try{const d=await api(`/api/portfolio?refresh=${refresh?'true':'false'}`);renderPortfolio(d);if(!silent)setStatus(`PORTFOLIO ${d?.market_status?.is_open?'MARKET OPEN':'MARKET CLOSED'}`);}
  catch(err){if(!silent){toast(err.message,true);setStatus('PORTFOLIO ERROR');}}
  finally{portfolioInFlight=false;if(!silent)setButtonBusy($('#portfolio-refresh'),false);}
}
async function getPortfolioQuote(silent=false){
  if(portfolioQuoteInFlight)return;
  const ticker=String($('#portfolio-ticker')?.value||'').trim().toUpperCase();if(!ticker)return;
  portfolioQuoteInFlight=true;if(!silent)setButtonBusy($('#portfolio-get-quote'),true,'QUOTE…');
  try{const d=await api(`/api/portfolio/quote/${encodeURIComponent(ticker)}`);$('#portfolio-ticker').value=d.ticker;renderPortfolioQuote(d);if(!silent)toast(`${d.ticker} ${usd(d.price)} · bid ${usd(d.bid)} / ask ${usd(d.ask)}`);}
  catch(err){if(!silent){toast(err.message,true);} }
  finally{portfolioQuoteInFlight=false;if(!silent)setButtonBusy($('#portfolio-get-quote'),false);}
}
async function loadPortfolioChart(silent=false){
  if(portfolioChartInFlight)return;
  const ticker=String($('#portfolio-ticker')?.value||'').trim().toUpperCase();if(!ticker)return;
  portfolioChartInFlight=true;
  try{const d=await api(`/api/portfolio/chart/${encodeURIComponent(ticker)}?interval=${encodeURIComponent(portfolioInterval)}`);renderPortfolioMarketChart(d);}
  catch(err){if(!silent)toast(err.message,true);}
  finally{portfolioChartInFlight=false;}
}
function syncOrderTypeUi(){
  const type=$('#portfolio-order-type')?.value||'MARKET';const lp=$('#portfolio-limit-price'),tif=$('#portfolio-tif');
  if(lp){lp.disabled=type!=='LIMIT';if(type==='LIMIT'&&!lp.value&&portfolioLastQuote?.price)lp.value=Number(portfolioLastQuote.price).toFixed(2);if(type!=='LIMIT')lp.value='';}
  if(tif){tif.disabled=type==='MARKET';if(type==='MARKET')tif.value='DAY';}
}
async function submitPortfolioTrade(side){
  if(portfolioTradeInFlight)return;
  const ticker=String($('#portfolio-ticker')?.value||'').trim().toUpperCase(),quantity=Number($('#portfolio-quantity')?.value||0);
  const password=$('#portfolio-password')?.value||'',note=$('#portfolio-note')?.value||'',order_type=$('#portfolio-order-type')?.value||'MARKET';
  const limit_price=order_type==='LIMIT'?Number($('#portfolio-limit-price')?.value||0):null,tif=$('#portfolio-tif')?.value||'DAY';
  if(!ticker||!Number.isFinite(quantity)||quantity<=0){toast('Enter a valid ticker and quantity.',true);return;}
  if(order_type==='LIMIT'&&(!Number.isFinite(limit_price)||limit_price<=0)){toast('Enter a valid limit price.',true);return;}
  if(!password){toast('Trade password is required.',true);return;}
  portfolioTradeInFlight=true;const btn=side==='BUY'?$('#portfolio-buy'):$('#portfolio-sell');setButtonBusy(btn,true,`${side}…`);[$('#portfolio-buy'),$('#portfolio-sell')].forEach(b=>{if(b)b.disabled=true});
  try{
    const d=await api('/api/portfolio/trade',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({side,ticker,quantity,password,order_type,limit_price,tif,note})});
    if(d.quote)renderPortfolioQuote({...d.quote,market_status:d.portfolio?.market_status});renderPortfolio(d.portfolio);$('#portfolio-password').value='';
    const o=d.order||{},f=d.fill;
    if(f)toast(`${side} ${num(f.quantity,4)} ${o.ticker} @ ${usd(f.price)} · ${num(f.slippage_bps,2)}bp impact · ${o.status}`);
    else toast(`${side} ${num(o.quantity,4)} ${o.ticker} ${o.order_type} order #${o.id} placed · ${o.status}`);
    log(`ORDER #${o.id} ${side} ${quantity} ${ticker} ${order_type} -> ${o.status}`,'PORT');
    await loadPortfolioChart(true);
  }catch(err){toast(err.message,true);log(err.message,'ERR');}
  finally{portfolioTradeInFlight=false;setButtonBusy(btn,false);renderPortfolioMarketStatus(portfolioLastPayload?.market_status||{});}
}
async function cancelPortfolioOrder(orderId){
  const password=$('#portfolio-password')?.value||'';if(!password){toast('Enter trade password in the ticket to cancel an order.',true);return;}
  try{const d=await api(`/api/portfolio/order/${orderId}/cancel`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password})});renderPortfolio(d.portfolio);$('#portfolio-password').value='';toast(`Order #${orderId} cancelled.`);}
  catch(err){toast(err.message,true);}
}
function startPortfolioPolling(){
  stopPortfolioPolling();
  getPortfolioQuote(true);loadPortfolioChart(true);
  portfolioQuoteTimer=setInterval(()=>{if($('#portfolio-view')?.classList.contains('active'))getPortfolioQuote(true)},5000);
  portfolioChartTimer=setInterval(()=>{if($('#portfolio-view')?.classList.contains('active'))loadPortfolioChart(true)},10000);
  portfolioPollTimer=setInterval(()=>{if($('#portfolio-view')?.classList.contains('active'))loadPortfolio(true,true)},15000);
}
function stopPortfolioPolling(){
  if(portfolioPollTimer){clearInterval(portfolioPollTimer);portfolioPollTimer=null;}
  if(portfolioQuoteTimer){clearInterval(portfolioQuoteTimer);portfolioQuoteTimer=null;}
  if(portfolioChartTimer){clearInterval(portfolioChartTimer);portfolioChartTimer=null;}
}

$('#portfolio-refresh')?.addEventListener('click',async()=>{await loadPortfolio(true);await getPortfolioQuote(true);await loadPortfolioChart(true);});
$('#portfolio-get-quote')?.addEventListener('click',async()=>{await getPortfolioQuote(false);await loadPortfolioChart(true);});
$('#portfolio-buy')?.addEventListener('click',()=>submitPortfolioTrade('BUY'));
$('#portfolio-sell')?.addEventListener('click',()=>submitPortfolioTrade('SELL'));
$('#portfolio-order-type')?.addEventListener('change',syncOrderTypeUi);
$('#portfolio-ticker')?.addEventListener('input',e=>{e.target.value=String(e.target.value||'').toUpperCase();portfolioLastQuote=null;portfolioChartPayload=null;['#portfolio-ticket-price','#portfolio-ticket-bid','#portfolio-ticket-ask','#portfolio-ticket-spread'].forEach(sel=>{if($(sel))$(sel).textContent='—'});updatePortfolioTicketNotional();});
$('#portfolio-ticker')?.addEventListener('keydown',async e=>{if(e.key==='Enter'){await getPortfolioQuote(false);await loadPortfolioChart(true);}});
$('#portfolio-quantity')?.addEventListener('input',updatePortfolioTicketNotional);
$$('#portfolio-chart-intervals button').forEach(btn=>btn.addEventListener('click',()=>{portfolioInterval=btn.dataset.interval||'1m';$$('#portfolio-chart-intervals button').forEach(x=>x.classList.toggle('active',x===btn));loadPortfolioChart(false);}));
['#portfolio-show-vwap','#portfolio-show-sma','#portfolio-show-ema'].forEach(sel=>$(sel)?.addEventListener('change',()=>{if(portfolioChartPayload)renderPortfolioMarketChart(portfolioChartPayload)}));
window.addEventListener('resize',()=>{if($('#portfolio-view')?.classList.contains('active'))portfolioResizePlot();});
syncOrderTypeUi();

