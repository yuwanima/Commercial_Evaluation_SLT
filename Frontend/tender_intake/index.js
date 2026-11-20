
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const strip=n=>n.replace(/\.[^.]+$/,'');
function guessRaw(n){return strip(n).replace(/_/g," ").replace(/\b(GCC|SCC|compliance|for|general conditions of contract|special conditions of contract|copy of)\b/ig,' ').replace(/[_\-]+/g,' ').replace(/\s+/g,' ').trim()}
function guess(n){return guessRaw(n)||strip(n)}
const EXT=/\.(xlsx|xlsm|xls|docx|doc|pdf|zip)$/i, EVALNAME=/evaluation|\bTET\b|SSET/i;
// Bidder name from a folder path; category folders are not bidder names.
function folderName(parts){
  if(!parts||parts.length<2)return '';
  const parent=parts[parts.length-2];
  if(/^(price documents?.*|gcc\s*&\s*scc.*|gcc|scc|technical|evaluation|documents?)$/i.test(parent.trim()))return '';
  return parts.length>=3?parent:'';
}

function panel(root,mode){
  const items=[],zone=root.querySelector('.zone'),ul=root.querySelector('ul.files'),cnt=root.querySelector('.count'),out=root.querySelector('.out'),go=root.querySelector('.go'),goText=go.textContent;
  const progress=root.querySelector('#evaluation-progress');
  const progressTitle=root.querySelector('#evaluation-progress-title');
  const progressMessage=root.querySelector('#evaluation-progress-message');
  const progressFiles=root.querySelector('#evaluation-progress-files');
  const progressTrack=root.querySelector('#evaluation-progress-track');
  const progressSteps=progress?[...progress.querySelectorAll('[data-progress-step]')]:[];
  function setProgress(state, message){
    if(!progress)return;
    progress.classList.remove('hidden');
    progress.dataset.state=state;
    progress.setAttribute('aria-busy',state==='processing'?'true':'false');
    progressTitle.textContent=state==='processing'?'Submitting and analysing your documents…':state==='complete'?'Evaluation response received':'We could not complete the analysis';
    const submittedTypes=[...new Set(items.map(it=>(it.file.name.match(/\.([^.]+)$/)||[])[1]?.toUpperCase()).filter(Boolean))];
    progressMessage.textContent=state==='processing'&&submittedTypes.length
      ?`${message} File types in this request: ${submittedTypes.join(', ')}.`
      :message;
    progressTrack.setAttribute('aria-valuetext',state==='processing'?'Waiting for the evaluation service response':state==='complete'?'Processing response received':'Processing request failed');
    progressFiles.replaceChildren();
    items.forEach(it=>{
      const chip=document.createElement('span');
      chip.className='progress-file';
      chip.textContent=it.path||it.file.name;
      chip.title=chip.textContent;
      progressFiles.append(chip);
    });
    progressSteps.forEach((step,index)=>{
      const status=step.querySelector('[data-step-status]');
      step.classList.toggle('is-active',state==='processing'&&index===0);
      step.classList.toggle('is-complete',state==='complete'&&index<2);
      step.classList.toggle('is-error',state==='error'&&index===1);
      if(status)status.textContent=state==='processing'
        ?(index===0?'Submitting selected files':index===1?'Waiting for processing response':'Not available yet')
        :state==='complete'
          ?(index<2?'Completed': 'Results available below')
          :(index===0?'Submission attempted':index===1?'Result could not be confirmed':'Not available');
    });
    if(state==='complete')progressSteps[2]?.classList.add('is-complete');
    if(state==='error')progressSteps[1]?.classList.add('is-error');
  }
  const kindOf=n=>/\.zip$/i.test(n)?'ZIP':/gct/i.test(n)?'GCT':mode==='all'?'DOC':mode==='eval'?'EVAL':mode==='price'?'PRICE':/\.(xlsx|xlsm|xls)$/i.test(n)?'GCC':/\.pdf$/i.test(n)?(/scc/i.test(n)?'SCC':/gcc/i.test(n)?'GCC':'PDF'):'SCC';
  const named=it=>it.kind!=='ZIP'&&it.kind!=='EVAL';
  function render(extra){
    if(progress){progress.classList.add('hidden');progress.removeAttribute('data-state');progress.setAttribute('aria-busy','false')}
    ul.innerHTML='';
    if(mode==='all'&&items.length){
      const groups={Technical:[],Price:[], 'GCC/SCC':[],GCT:[],Other:[]};
      items.forEach(it=>{
        const text=`${it.path||''} ${it.file.name}`.toLowerCase();
        const group=/technical|evaluation|\btet\b|\bsset\b/.test(text)?'Technical':
          /price|schedule|financial/.test(text)?'Price':
          /gct|general conditions of tender/.test(text)?'GCT':
          /gcc|scc|general conditions|special conditions/.test(text)?'GCC/SCC':'Other';
        groups[group].push(it);
      });
      const groupSummary=Object.entries(groups).filter(([,files])=>files.length).map(([group,files])=>{
        const types={};
        files.forEach(it=>{const ext=(it.file.name.match(/\.([^.]+)$/)||[])[1];if(ext)types[ext.toLowerCase()]=(types[ext.toLowerCase()]||0)+1});
        return `<div><b>${group}</b><span>${files.length} document(s) · ${Object.entries(types).map(([type,count])=>`${count} ${type.toUpperCase()}`).join(' · ')}</span></div>`;
      }).join('');
      ul.innerHTML=`<li class="upload-selection-summary"><b>${items.length} document(s) selected</b><div class="upload-group-summary">${groupSummary}</div></li>`;
    }else items.forEach((it,i)=>{
      const li=document.createElement('li');
      li.innerHTML=`<span class="fn"><span class="tag ${it.kind}">${it.kind}</span> ${esc(it.path||it.file.name)}</span>`+(named(it)?`<input type="text" value="${esc(it.name)}" title="Bidder name" placeholder="Bidder name">`:'')+`<button type="button" title="Remove">×</button>`;
      const inp=li.querySelector('input');if(inp)inp.oninput=e=>it.name=e.target.value;
      li.querySelector('button').onclick=()=>{items.splice(i,1);render()};
      ul.append(li);
    });
    const c=k=>items.filter(x=>x.kind===k).length;
    cnt.textContent=!items.length?'':mode==='all'?`${items.length} document(s) ready`:mode==='eval'?`${items.length} report file(s) ready`:mode==='price'?`${items.length} price schedule file(s) ready`:`${c('GCC')} GCC · ${c('SCC')} SCC`+(c('PDF')?` · ${c('PDF')} PDF (type detected automatically)`:'')+(c('ZIP')?` · ${c('ZIP')} zip`:'')+' ready';
    out.innerHTML=(extra&&extra.length)?`<div class="msg">${extra.map(esc).join('<br>')}</div>`:'';
  }
  function add(list){           // list: [{file, parts}] parts = path segments or null
    const msgs=[];let skipped=0,duplicates=0;
    list.forEach(({file,parts})=>{
      if(!EXT.test(file.name)){skipped++;return}
      if(file.name.startsWith('~$')){return}
      if(mode==='bids'&&/\.(docx|doc|pdf)$/i.test(file.name)&&EVALNAME.test(file.name)){msgs.push(`${file.name}: looks like an evaluation report — upload it in the Technical Evaluation tab`);return}
      const fromFolder=parts&&parts.length>1;
      const path=fromFolder?parts.join('/'):null;
      const duplicate=items.some(it=>it.file.name===file.name&&it.file.size===file.size&&
        it.file.lastModified===file.lastModified&&(it.path||'')===(path||''));
      if(duplicate){duplicates++;return}
      items.push({file,kind:kindOf(file.name),path:path,
        name:mode==='eval'?strip(file.name):(fromFolder&&(folderName(parts)||guessRaw(file.name)||parts[0]))||guess(file.name)});
    });
    if(skipped)msgs.push(`${skipped} file(s) skipped (unsupported type)`);
    if(duplicates)msgs.push(`${duplicates} duplicate file(s) already selected`);
    render(msgs);
  }
  root.querySelector('.f-files').onchange=e=>{add([...e.target.files].map(f=>({file:f,parts:null})));e.target.value=''};
  root.querySelector('.f-dir').onchange=e=>{add([...e.target.files].map(f=>({file:f,parts:(f.webkitRelativePath||f.name).split('/')})));e.target.value=''};
  zone.onclick=e=>{if(!e.target.closest('.browse-button'))root.querySelector('.f-files').click()};
  ['dragenter','dragover'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.add('over')}));
  ['dragleave','drop'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.remove('over')}));
  const readEntry=(en,path)=>new Promise(res=>{
    if(en.isFile)en.file(f=>res([{file:f,parts:(path+en.name).split('/')}]),()=>res([]));
    else{const rd=en.createReader(),all=[];(function more(){rd.readEntries(async b=>{
      if(!b.length){res((await Promise.all(all.map(x=>readEntry(x,path+en.name+'/')))).flat());return}
      all.push(...b);more()},()=>res([]))})()}
  });
  zone.addEventListener('drop',async e=>{
    const its=[...(e.dataTransfer.items||[])].map(i=>i.webkitGetAsEntry&&i.webkitGetAsEntry()).filter(Boolean);
    if(its.length){add((await Promise.all(its.map(en=>readEntry(en,'')))).flat().map(x=>x.parts.length===1?{file:x.file,parts:null}:x))}
    else add([...e.dataTransfer.files].map(f=>({file:f,parts:null})));
  });
  go.onclick=async()=>{
    if(!items.length){out.innerHTML='<div class="msg">Add files or a folder first.</div>';return}
    const fd=new FormData();fd.append('mode',mode);
    items.forEach(it=>{fd.append('files',it.file);fd.append('names',named(it)||mode==='eval'?it.name.trim():'')});
    go.disabled=true;go.innerHTML='<span class="spin"></span>Processing…';out.innerHTML='';
    setProgress('processing','Submitting the selected files to the evaluation service. The service response will confirm when processing is complete.');
    try{
      const r=await fetch('/api/process',{method:'POST',body:fd});
      const contentType=r.headers.get('content-type')||'';
      if(!contentType.includes('application/json')){
        const detail=await r.text();
        throw new Error(`Server returned HTTP ${r.status} instead of JSON${detail?`: ${detail.slice(0,160)}`:''}`);
      }
      const j=await r.json();
      const visibleProblems=(mode==='all'?(j.problems||[]).filter(p=>{
        const text=String(p).toLowerCase();
        return !((text.includes('no scc table 1 / table 2 found')&&j.price)||
          (text.includes('no filled-in price schedule found')&&(j.gcc||j.gct||j.scc||!j.price)));
      }):(j.problems||[]));
      const probs=visibleProblems.length?'<b style="color:var(--bad);font-size:14px">Some files could not be read:</b><ul class="warn">'+visibleProblems.map(p=>'<li>'+esc(p)+'</li>').join('')+'</ul>':'';
      if(!r.ok){
        setProgress('error',j.error||'The evaluation service returned an error. Review the message below and try again.');
        out.innerHTML=`<div class="msg">${esc(j.error||'Something went wrong')}${j.problems?'<ul class="warn">'+j.problems.map(p=>'<li>'+esc(p)+'</li>').join('')+'</ul>':''}</div>`;return
      }
      const completedTypes=[];
      if(j.eval)completedTypes.push('technical evaluation reports');
      if(j.gcc||j.gct||j.scc)completedTypes.push('GCC/GCT/SCC compliance documents');
      if(j.price)completedTypes.push('price documents');
      const completionMessage=completedTypes.length
        ?`The service returned results for ${completedTypes.join(', ')}. Review the results and any extraction notes below.`
        :'The service returned a response. Review the results and any extraction notes below.';
      setProgress('complete',completionMessage);
      if(mode==='all'){out.innerHTML=allBox(j,probs);return}
      if(mode==='eval'){out.innerHTML=(j.evals||[]).map(evBox(j)).join('')+(probs?`<div class="res">${probs}</div>`:'');return}
      if(mode==='price'){out.innerHTML=priceBox(j,probs);return}
      const dl=(k,t,b)=>`<a class="${j[k]?'':'off'}" href="/api/download/${j.job}/${k}"><b>⬇ ${t}</b><span>${j[k]?b.length+' bidder(s): '+esc(b.join(', ')):'No files uploaded'}</span></a>`;
      const summaries=(j.document_summaries||[]).map(s=>`<div class="doc-summary"><h4>${esc(s.bidder||'Unknown bidder')} — ${esc(s.type||'Document')} summary</h4><ul>${s.points.filter(p=>!/^signature section:/i.test(p)&&!/^bidder name(?: source)?:/i.test(p)).map(p=>`<li>${esc(p)}</li>`).join('')}</ul></div>`).join('');
      out.innerHTML=`<div class="res"><h3>✓ Comparison ready</h3><div style="color:var(--mut);font-size:14px">Download your Excel files:</div>
        <div class="dl">${dl('gcc','GCC_Compliance_Comparison.xlsx',j.gcc_bidders)}${dl('gct','GCT_Compliance_Comparison.xlsx',j.gct_bidders)}${dl('scc','SCC_Compliance_Comparison.xlsx',j.scc_bidders)}</div>${summaries?`<div class="doc-summaries"><h4>Document checks</h4>${summaries}</div>`:''}${probs}</div>`;
    }catch(e){setProgress('error','The request did not return a usable processing response. Review the error below before retrying.');out.innerHTML='<div class="msg">Could not reach the server: '+esc(String(e))+'</div>'}
    finally{go.disabled=false;go.textContent=goText}
  };
}

const savedEvaluationRoot=document.getElementById('saved-evaluations');
const savedEvaluationRefresh=document.getElementById('refresh-evaluations');
function savedEvaluationMessage(message,error=false){
  if(savedEvaluationRoot)savedEvaluationRoot.innerHTML=`<p class="saved-status${error?' saved-error':''}">${esc(message)}</p>`;
}
async function loadSavedEvaluations(){
  if(!savedEvaluationRoot)return;
  savedEvaluationMessage('Loading saved evaluations...');
  try{
    const response=await fetch('/api/tenders');
    const data=await response.json();
    if(!response.ok)throw new Error(data.error||'The saved evaluations could not be loaded.');
    const tenders=Array.isArray(data.tenders)?data.tenders:[];
    if(!tenders.length){savedEvaluationMessage('No generated Excel files have been saved yet. Upload documents to create an evaluation.');return;}
    const records=await Promise.all(tenders.map(async tender=>{
      const result=await fetch(`/api/tender-files?tender_no=${encodeURIComponent(tender.tender_no)}`);
      const detail=await result.json();
      if(!result.ok)throw new Error(detail.error||'Unable to load tender files.');
      return detail;
    }));
    savedEvaluationRoot.innerHTML=records.map(record=>{
      const files=(record.files||[]).filter(file=>file.filename);
      const rows=files.length?files.map(file=>`<div class="saved-file-row"><span class="saved-file-label">${esc(file.label)}</span><a class="saved-download" href="${esc(file.url)}" download>Download</a></div>`).join(''):'<p class="saved-status">No generated summaries are available for this tender.</p>';
      const first=files[0]||{};
      return `<article class="saved-tender"><div class="saved-tender-meta"><span class="saved-date">${esc(first.date||'Date unavailable')} ${esc(first.time||'')}</span><strong>Tender No : ${esc(record.tender_no)}</strong></div><div class="saved-file-list">${rows}</div></article>`;
    }).join('');
  }catch(error){savedEvaluationMessage(error.message||'The saved evaluations could not be loaded.',true);}
}
if(savedEvaluationRefresh)savedEvaluationRefresh.onclick=loadSavedEvaluations;

function allBox(j,probs){
  const dl=(k,t,b)=>`<a class="${j[k]?'':'off'}" href="/api/download/${j.job}/${k}"><b>⬇ ${t}</b><span>${j[k]?(b||[]).length+' bidder(s): '+esc((b||[]).join(', ')):'No matching files uploaded'}</span></a>`;
  const evaluations=j.evals||[], priceRows=j.rows||[];
  const comparison=j.gcc||j.gct||j.scc?`<div class="all-section"><h2>GCC / GCT / SCC Comparison</h2><div class="res"><h3>✓ Comparison ready</h3><div class="dl">${dl('gcc','GCC_Compliance_Comparison.xlsx',j.gcc_bidders)}${dl('gct','GCT_Compliance_Comparison.xlsx',j.gct_bidders)}${dl('scc','SCC_Compliance_Comparison.xlsx',j.scc_bidders)}</div></div></div>`:`<div class="all-section empty-section"><h2>GCC / GCT / SCC Comparison</h2><p>No GCC, GCT, or SCC documents detected.</p></div>`;
  const evaluationDetails=evaluations.map((ev,i)=>`<details class="result-details"><summary>${esc(ev.label||`Technical report ${i+1}`)}</summary>${evBox(j)(ev,i)}</details>`).join('');
  const evaluationSection=evaluations.length?`<div class="all-section"><h2>Technical Evaluation</h2>${evaluationDetails}</div>`:`<div class="all-section empty-section"><h2>Technical Evaluation</h2><p>No technical evaluation reports detected.</p></div>`;
  const priceSection=j.price?`<div class="all-section"><h2>Price Schedule</h2><details class="result-details" open><summary>View all price results (${priceRows.length})</summary>${priceBox(j,'')}</details></div>`:`<div class="all-section empty-section"><h2>Price Schedule</h2><p>No price schedules detected.</p></div>`;
  return `<div class="all-results">${priceSection}${evaluationSection}${comparison}${probs?`<div class="res">${probs}</div>`:''}</div>`;
}

function evBox(j){return (ev,idx)=>{
  const tender=ev.meta&&ev.meta['Tender No.']?String(ev.meta['Tender No.']):'Not stated';
  const reportDetailOrder=['Tender No.','Tender Published Date','Bid Bond Expiry Date(s)'];
  const metaRows=Object.entries(ev.meta||{})
    .filter(([,value])=>String(value||'').trim())
    .sort(([keyA],[keyB])=>{
      const indexA=reportDetailOrder.indexOf(keyA);
      const indexB=reportDetailOrder.indexOf(keyB);
      const rankA=indexA===-1?reportDetailOrder.length:indexA;
      const rankB=indexB===-1?reportDetailOrder.length:indexB;
      return rankA-rankB;
    })
    .map(([key,value])=>`<div class="detail"><b>${esc(key)}</b><span>${esc(value)}</span></div>`).join('');
  const bidDetails=ev.bids&&Array.isArray(ev.bids.headers)&&Array.isArray(ev.bids.rows)?`<h4>Details of Bidders</h4><div class="bidder-details-grid">${ev.bids.rows.map(row=>{
    const bidder=row[0]||'Unnamed bidder';
    const fields=ev.bids.headers.slice(1).map((header,i)=>({header,value:row[i+1]||''})).filter(x=>String(x.value).trim());
    return `<article class="bidder-detail"><h5>${esc(bidder)}</h5>${fields.length?fields.map(x=>`<div class="detail"><b>${esc(x.header)}</b><span>${esc(x.value)}</span></div>`).join(''):'<div class="meta">No additional bidder details detected.</div>'}</article>`;
  }).join('')}</div>`:'';
  const nonEligibility=Array.isArray(ev.recommendations)?ev.recommendations.filter(r=>r.eligibility_status==='excluded').map(r=>`<article class="non-eligibility"><h5>${esc(r.bidder)} — excluded</h5>${r.eligibility_reason?`<p>${esc(r.eligibility_reason)}</p>`:''}${Array.isArray(r.issues)&&r.issues.length?`<ul>${r.issues.map(x=>`<li><b>${esc(x.area)}:</b> ${esc(x.status)}</li>`).join('')}</ul>`:''}</article>`).join(''):'';
  const needsReview=Array.isArray(ev.recommendations)?ev.recommendations.filter(r=>r.eligibility_status==='review_required').map(r=>`<article class="non-eligibility review"><h5>${esc(r.bidder)} — manual TET review required</h5><p>${esc(r.eligibility_reason||'Technical eligibility could not be confirmed automatically.')}</p></article>`).join(''):'';
  const reasons=(nonEligibility||needsReview)?`<div class="non-eligibility-section"><h4>Technical eligibility decisions requiring attention</h4>${nonEligibility}${needsReview}</div>`:'';
  const scoreRows=Array.isArray(ev.recommendations)?ev.recommendations.map(r=>`<tr><td>${esc(r.bidder)}</td><td>${r.technical_score==null?'Not available':esc(Number(r.technical_score).toFixed(2))}</td><td>${esc(r.technical_marks??0)} / ${esc(r.maximum_technical_marks??0)}</td><td>${esc(r.technical_score_source==='reported_in_technical_report'?'Reported in technical report':r.technical_score_source==='calculated_from_compliance'?'Calculated from compliance labels':'Not available')}</td><td>${esc(r.criteria_scored??0)}</td><td>${esc(r.complied??0)}</td><td>${esc(r.partially_complied??0)}</td><td>${esc(r.not_complied??0)}</td></tr>`).join(''):'';
  const scoreTable=scoreRows?`<div class="technical-score"><h4>Technical score by company</h4><p class="meta">${esc(ev.technical_score_basis||'Scores are based on the technical report evidence.')}</p><div class="tw"><table class="mx"><thead><tr><th>Company</th><th>Score /100</th><th>Marks</th><th>Source</th><th>Criteria scored</th><th>Complied</th><th>Partial</th><th>Non complied</th></tr></thead><tbody>${scoreRows}</tbody></table></div></div>`:'';
  const mats=ev.matrices.map(m=>`<div class="tw"><table class="mx"><thead><tr>${m.headers.map(x=>`<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>`+
    m.rows.map(r=>`<tr><td>${esc(r.area)}</td>${r.cells.map(c=>`<td class="${c.c}">${esc(c.t)}</td>`).join('')}</tr>`).join('')+
    (m.overall?`<tr class="ov"><td>Overall result (as stated in report)</td>${m.overall.map(c=>`<td class="${c.c}">${esc(c.t)}</td>`).join('')}</tr>`:'')+
    `</tbody></table></div>`).join('<br>');
  return `<div class="evalbox"><h3>${ev.type==='SSET'?'Sole Source Evaluation':'Technical Evaluation'} — Tender No: ${esc(tender)}</h3><div class="meta">File: ${esc(ev.label)}</div><div class="report-details"><h4>Report details</h4>${metaRows||'<div class="meta">No additional report details were detected.</div>'}</div><div class="recommendations">${bidDetails||'<h4>Details of Bidders</h4><div class="meta">No bidder details were detected.</div>'}</div>${scoreTable}${mats}${reasons}
    ${idx===0?`<a class="xl" href="/api/download/${j.job}/eval">⬇ Download Excel (Tender No: ${esc(tender)})</a>`:''}</div>`}}
function fmtMoney(value){
  return value===null||value===undefined||!Number.isFinite(Number(value))?'Not available':Number(value).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
}
function priceBox(j,probs){
  const rows=(Array.isArray(j.rows)?j.rows:[]).slice().sort((a,b)=>{
    const rankA=a.rank||999, rankB=b.rank||999;
    return rankA-rankB||String(a.bidder||'').localeCompare(String(b.bidder||''));
  });
  const currencyGroups=new Map();
  rows.forEach(r=>{
    const currency=String(r.currency||'Not stated');
    if(!currencyGroups.has(currency))currencyGroups.set(currency,[]);
    currencyGroups.get(currency).push(r);
  });
  const tables=[...currencyGroups].map(([currency,group])=>`<div class="price-currency-group"><h4>Currency: ${esc(currency)}</h4><table class="price-table"><thead><tr><th>Bidder</th><th>Currency</th><th>Evaluated price</th><th>Price rank</th><th>Price score / 100</th></tr></thead><tbody>${group.map(r=>`<tr><td>${esc(r.bidder)}</td><td>${esc(r.currency||'Not stated')}</td><td>${esc(fmtMoney(r.evaluated_total??r.total))}</td><td>${esc(r.rank?`${r.rank}${r.rank===1?'st':r.rank===2?'nd':r.rank===3?'rd':'th'}`:'Not ranked')}</td><td>${r.price_score==null?'Not calculated':esc(Number(r.price_score).toFixed(2))}</td></tr>`).join('')}</tbody></table></div>`).join('')||'<table class="price-table"><tbody><tr><td colspan="5">No bidder prices detected.</td></tr></tbody></table>';
  const notes=(j.notes||[]).concat(rows.flatMap(r=>r.warnings||[]));
  return `<div class="res"><h3>✓ Price schedule extraction ready</h3><div class="price-summary">Extracted bidder price schedules (${rows.length})</div>${tables}${notes.length?`<ul class="warn">${notes.map(esc).map(x=>`<li>${x}</li>`).join('')}</ul>`:''}<a class="xl" href="/api/download/${j.job}/price">⬇ Download Price Schedules Excel</a>${probs?`<div>${probs}</div>`:''}</div>`;
}
function awardBox(j){
  const a=j.award_assessment||{};
  const ready=a.decision==='recommendation_ready';
  const provisional=a.decision==='provisional_price_leader';
  const badgeClass=ready?'award-ready':provisional?'award-provisional':'award-review';
  const money=p=>p&&p.amount!==null&&p.amount!==undefined?`${esc(p.currency||'')} ${esc(fmtMoney(p.amount))}`:'Not available';
  const rows=(a.bidders||[]).map(b=>{
    const p=b.price||{};
    const rank=b.price_rank?`${b.price_rank}${b.price_rank===1?'st':b.price_rank===2?'nd':b.price_rank===3?'rd':'th'}${b.price_rank_scope?` (${esc(b.price_rank_scope)})`:''}`:'—';
    return `<tr><td>${esc(b.bidder)}</td><td>${esc((b.eligibility_status||'review_required').replaceAll('_',' '))}</td><td>${money(p)}</td><td>${rank}</td><td>${b.technical_score==null?'Not provided':esc(Number(b.technical_score).toFixed(2))}</td><td>${b.weighted_technical_score==null?'Not calculated':esc(Number(b.weighted_technical_score).toFixed(2))}</td><td>${b.commercial_score==null?'Not provided':esc(Number(b.commercial_score).toFixed(2))}</td><td>${b.weighted_commercial_score==null?'Not calculated':esc(Number(b.weighted_commercial_score).toFixed(2))}</td><td>${b.overall_score==null?'Not calculated':esc(Number(b.overall_score).toFixed(2))}</td></tr>`;
  }).join('');
  const weights=a.weights&&a.weights.technical!==undefined&&a.weights.commercial!==undefined?`Technical ${esc(a.weights.technical)}% / Commercial ${esc(a.weights.commercial)}% (${a.weights.approved?'approval evidence found':'approval not verified'})`:'Not found in uploaded approved strategy documents';
  const blockers=(a.blockers||[]).map(x=>`<li>${esc(x)}</li>`).join('');
  const notes=(a.manual_notes||[]).map(x=>`<li>${esc(x)}</li>`).join('');
  const awardConfirmation=a.award_confirmation_status==='confirmed_from_document'
    ?`<div class="award-confirmed"><span>Confirmed awarded bidder (explicit approval/award document)</span><strong>${esc(a.awarded_bidder||'')}</strong><small>Source: ${esc((a.award_confirmation_sources||[]).join(', '))}</small></div>`
    :`<div class="award-pending"><span>Actual awarded bidder</span><strong>${a.award_confirmation_status==='review_required'?'Award document needs manual verification':'Not established by uploaded award/approval documents'}</strong></div>`;
  return `<div class="res award-res">${awardConfirmation}<div class="award-status ${badgeClass}">${esc(a.status_text||'Award assessment unavailable')}</div>${a.recommended_bidder?`<div class="award-winner"><span>${ready?'Preferred bidder for Tender Board consideration':'Provisional price front-runner'}</span><strong>${esc(a.recommended_bidder)}</strong>${a.recommended_score!=null?`<small>Overall weighted score: ${esc(Number(a.recommended_score).toFixed(2))}/100</small>`:a.recommended_price?`<small>Evaluated price: ${money(a.recommended_price)}</small>`:''}</div>`:`<p class="award-no-winner">The system cannot identify a defensible preferred bidder from the uploaded evidence yet.</p>`}<div class="award-meta"><span><b>Weighting record:</b> ${weights}</span><span><b>Eligible bidders:</b> ${esc(a.eligible_count??0)}</span><span><b>Excluded:</b> ${esc(a.excluded_count??0)}</span><span><b>Manual review:</b> ${esc(a.review_count??0)}</span></div><p class="award-basis"><b>Assessment basis:</b> ${esc(a.basis||'')}</p><div class="tw"><table class="price-table award-table"><thead><tr><th>Bidder</th><th>TET eligibility</th><th>Evaluated price</th><th>Price rank</th><th>Technical score /100</th><th>Weighted technical score</th><th>Commercial score /100</th><th>Weighted commercial score</th><th>Overall score</th></tr></thead><tbody>${rows||'<tr><td colspan="9">No bidder assessment data extracted.</td></tr>'}</tbody></table></div>${blockers?`<div class="award-blockers"><h4>Information still required before a final recommendation</h4><ul>${blockers}</ul></div>`:''}${notes?`<details class="award-notes"><summary>Governance and manual notes</summary><ul>${notes}</ul></details>`:''}<a class="xl" href="/api/download/${j.job}/award">⬇ Download Award Assessment Excel</a></div>`;
}

panel(document.getElementById('p-all'),'all');
document.querySelectorAll('.tabs button').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('.tabs button').forEach(x=>x.classList.toggle('on',x===b));
  document.querySelectorAll('.panel').forEach(p=>p.hidden=p.id!=='p-'+b.dataset.tab);
  try{location.hash=b.dataset.tab}catch(e){}
  if(b.dataset.tab==='view')loadSavedEvaluations();
});
if(location.hash&&document.querySelector(`[data-tab="${location.hash.slice(1)}"]`))document.querySelector(`[data-tab="${location.hash.slice(1)}"]`).click();

// Reuse the existing same-origin chatbot page in a resizable drawer.
const comopsLauncher = document.getElementById('comops-launcher');
const comopsPanel = document.getElementById('comops-panel');
const comopsFrame = document.getElementById('comops-frame');
const comopsClose = document.getElementById('comops-close');
const comopsExpand = document.getElementById('comops-expand');
const comopsResizeHandle = document.getElementById('comops-resize-handle');
let comopsPanelWidth = 520;
let comopsRestoreWidth = comopsPanelWidth;
let comopsMaximized = false;
let comopsResizeStart = null;

function comopsWidthLimits(){
  const viewport = window.innerWidth;
  const minimum = Math.min(viewport <= 600 ? 280 : 360, viewport);
  const maximum = viewport <= 600
    ? viewport
    : Math.min(900, Math.max(minimum, viewport - 280));
  return {minimum, maximum};
}

function setComopsPanelWidth(width, allowFullViewport=false){
  const {minimum, maximum:availableMaximum} = comopsWidthLimits();
  const maximum = allowFullViewport ? window.innerWidth : availableMaximum;
  comopsPanelWidth = Math.round(Math.max(minimum, Math.min(maximum, width)));
  comopsPanel.style.setProperty('--comops-panel-width', `${comopsPanelWidth}px`);
  comopsResizeHandle.setAttribute('aria-valuemin', String(minimum));
  comopsResizeHandle.setAttribute('aria-valuemax', String(maximum));
  comopsResizeHandle.setAttribute('aria-valuenow', String(comopsPanelWidth));
}

function openComopsPanel(){
  setComopsPanelWidth(comopsMaximized ? window.innerWidth : comopsPanelWidth, comopsMaximized);
  comopsPanel.classList.add('is-open');
  comopsPanel.setAttribute('aria-hidden', 'false');
  comopsLauncher.setAttribute('aria-expanded', 'true');
  if(comopsFrame.dataset.loaded !== 'true'){
    comopsFrame.src = comopsFrame.dataset.src;
    comopsFrame.dataset.loaded = 'true';
  }
  comopsClose.focus({preventScroll:true});
}

function closeComopsPanel(){
  comopsPanel.classList.remove('is-open');
  comopsPanel.setAttribute('aria-hidden', 'true');
  comopsLauncher.setAttribute('aria-expanded', 'false');
  comopsLauncher.focus({preventScroll:true});
}

comopsLauncher.addEventListener('click', event=>{
  event.preventDefault();
  openComopsPanel();
});
comopsClose.addEventListener('click', closeComopsPanel);
comopsExpand.addEventListener('click', ()=>{
  if(comopsMaximized){
    comopsMaximized = false;
    comopsExpand.setAttribute('aria-pressed', 'false');
    comopsExpand.setAttribute('aria-label', 'Maximize ComOps panel');
    comopsExpand.title = 'Maximize panel';
    setComopsPanelWidth(comopsRestoreWidth);
  }else{
    comopsRestoreWidth = comopsPanelWidth;
    comopsMaximized = true;
    comopsExpand.setAttribute('aria-pressed', 'true');
    comopsExpand.setAttribute('aria-label', 'Restore ComOps panel size');
    comopsExpand.title = 'Restore panel size';
    setComopsPanelWidth(window.innerWidth, true);
  }
});

comopsResizeHandle.addEventListener('pointerdown', event=>{
  if(!comopsPanel.classList.contains('is-open') || comopsMaximized) return;
  if(event.button !== undefined && event.button !== 0) return;
  comopsResizeStart = {x:event.clientX, width:comopsPanelWidth, pointerId:event.pointerId};
  comopsResizeHandle.setPointerCapture(event.pointerId);
  document.body.classList.add('comops-resizing');
  event.preventDefault();
});
comopsResizeHandle.addEventListener('pointermove', event=>{
  if(!comopsResizeStart || event.pointerId !== comopsResizeStart.pointerId) return;
  setComopsPanelWidth(comopsResizeStart.width + comopsResizeStart.x - event.clientX);
});
function finishComopsResize(event){
  if(!comopsResizeStart || (event.pointerId !== undefined && event.pointerId !== comopsResizeStart.pointerId)) return;
  comopsResizeStart = null;
  document.body.classList.remove('comops-resizing');
}
comopsResizeHandle.addEventListener('pointerup', finishComopsResize);
comopsResizeHandle.addEventListener('pointercancel', finishComopsResize);
comopsResizeHandle.addEventListener('keydown', event=>{
  if(event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
  event.preventDefault();
  setComopsPanelWidth(comopsPanelWidth + (event.key === 'ArrowLeft' ? 24 : -24));
});

document.addEventListener('keydown', event=>{
  if(event.key === 'Escape' && comopsPanel.classList.contains('is-open')) closeComopsPanel();
});
window.addEventListener('resize', ()=>{
  if(comopsMaximized) setComopsPanelWidth(window.innerWidth, true);
  else setComopsPanelWidth(comopsPanelWidth);
});
setComopsPanelWidth(comopsPanelWidth);
