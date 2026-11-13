
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
  const kindOf=n=>/\.zip$/i.test(n)?'ZIP':/gct/i.test(n)?'GCT':mode==='all'?'DOC':mode==='eval'?'EVAL':mode==='price'?'PRICE':/\.(xlsx|xlsm|xls)$/i.test(n)?'GCC':/\.pdf$/i.test(n)?(/scc/i.test(n)?'SCC':/gcc/i.test(n)?'GCC':'PDF'):'SCC';
  const named=it=>it.kind!=='ZIP'&&it.kind!=='EVAL';
  function render(extra){
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
  zone.onclick=()=>root.querySelector('.f-files').click();
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
      if(!r.ok){out.innerHTML=`<div class="msg">${esc(j.error||'Something went wrong')}${j.problems?'<ul class="warn">'+j.problems.map(p=>'<li>'+esc(p)+'</li>').join('')+'</ul>':''}</div>`;return}
      if(mode==='all'){out.innerHTML=allBox(j,probs);return}
      if(mode==='eval'){out.innerHTML=(j.evals||[]).map(evBox(j)).join('')+(probs?`<div class="res">${probs}</div>`:'');return}
      if(mode==='price'){out.innerHTML=priceBox(j,probs);return}
      const dl=(k,t,b)=>`<a class="${j[k]?'':'off'}" href="/api/download/${j.job}/${k}"><b>⬇ ${t}</b><span>${j[k]?b.length+' bidder(s): '+esc(b.join(', ')):'No files uploaded'}</span></a>`;
      const extracted=s=>(s.points||[]).find(p=>/^bidder name source:\s*official frank \/ stamp \/ seal$/i.test(p));
      const summaries=(j.document_summaries||[]).map(s=>`<div class="doc-summary"><h4>${esc(s.bidder||'Unknown bidder')} — ${esc(s.type||'Document')} summary</h4><ul>${s.points.filter(p=>!/^signature section:/i.test(p)&&!/^bidder name(?: source)?:/i.test(p)).map(p=>`<li>${esc(p)}</li>`).join('')}</ul></div>`).join('');
      const names=(j.document_summaries||[]).filter(extracted).map(s=>`<div><b>${esc(s.bidder||'Unknown')}</b><span>Official frank / stamp / seal</span></div>`).join('');
      out.innerHTML=`<div class="res"><h3>✓ Comparison ready</h3><div style="color:var(--mut);font-size:14px">Download your Excel files:</div>
        <div class="dl">${dl('gcc','GCC_Compliance_Comparison.xlsx',j.gcc_bidders)}${dl('gct','GCT_Compliance_Comparison.xlsx',j.gct_bidders)}${dl('scc','SCC_Compliance_Comparison.xlsx',j.scc_bidders)}</div>${names?`<div class="detected-names"><h4>Detected bidder name(s)</h4>${names}</div>`:''}${summaries?`<div class="doc-summaries"><h4>Document checks</h4>${summaries}</div>`:''}${probs}</div>`;
    }catch(e){out.innerHTML='<div class="msg">Could not reach the server: '+esc(String(e))+'</div>'}
    finally{go.disabled=false;go.textContent=goText}
  };
}

function allBox(j,probs){
  const dl=(k,t,b)=>`<a class="${j[k]?'':'off'}" href="/api/download/${j.job}/${k}"><b>⬇ ${t}</b><span>${j[k]?(b||[]).length+' bidder(s): '+esc((b||[]).join(', ')):'No matching files uploaded'}</span></a>`;
  const extracted=s=>(s.points||[]).find(p=>/^bidder name source:\s*official frank \/ stamp \/ seal$/i.test(p));
  const summaries=(j.document_summaries||[]).map(s=>`<div class="doc-summary"><h4>${esc(s.bidder||'Unknown bidder')} — ${esc(s.type||'Document')} summary</h4><ul>${(s.points||[]).filter(p=>!/^signature section:/i.test(p)&&!/^bidder name(?: source)?:/i.test(p)).map(p=>`<li>${esc(p)}</li>`).join('')}</ul></div>`).join('');
  const names=(j.document_summaries||[]).filter(extracted).map(s=>`<div><b>${esc(s.bidder||'Unknown')}</b><span>Official frank / stamp / seal</span></div>`).join('');
  const evaluations=j.evals||[], priceRows=j.rows||[];
  const comparison=j.gcc||j.gct||j.scc?`<div class="all-section"><h2>GCC / GCT / SCC Comparison</h2><div class="res"><h3>✓ Comparison ready</h3><div class="dl">${dl('gcc','GCC_Compliance_Comparison.xlsx',j.gcc_bidders)}${dl('gct','GCT_Compliance_Comparison.xlsx',j.gct_bidders)}${dl('scc','SCC_Compliance_Comparison.xlsx',j.scc_bidders)}</div>${names?`<div class="detected-names"><h4>Detected bidder name(s)</h4>${names}</div>`:''}${summaries?`<details class="result-details"${j.gct?' open':''}><summary>View document checks (${(j.document_summaries||[]).length})</summary><div class="doc-summaries">${summaries}</div></details>`:''}</div></div>`:`<div class="all-section empty-section"><h2>GCC / GCT / SCC Comparison</h2><p>No GCC, GCT, or SCC documents detected.</p></div>`;
  const evaluationDetails=evaluations.map((ev,i)=>`<details class="result-details"><summary>${esc(ev.label||`Technical report ${i+1}`)}</summary>${evBox(j)(ev,i)}</details>`).join('');
  const evaluationSection=evaluations.length?`<div class="all-section"><h2>Technical Evaluation</h2>${evaluationDetails}</div>`:`<div class="all-section empty-section"><h2>Technical Evaluation</h2><p>No technical evaluation reports detected.</p></div>`;
  const priceSection=j.price?`<div class="all-section"><h2>Price Schedule</h2><details class="result-details" open><summary>View all price results (${priceRows.length})</summary>${priceBox(j,'')}</details></div>`:`<div class="all-section empty-section"><h2>Price Schedule</h2><p>No price schedules detected.</p></div>`;
  return `<div class="all-results">${priceSection}${evaluationSection}${comparison}${probs?`<div class="res">${probs}</div>`:''}</div>`;
}

function evBox(j){return (ev,idx)=>{
  const tender=ev.meta&&ev.meta['Tender No.']?String(ev.meta['Tender No.']):'Not stated';
  const metaRows=Object.entries(ev.meta||{}).filter(([,value])=>String(value||'').trim()).map(([key,value])=>`<div class="detail"><b>${esc(key)}</b><span>${esc(value)}</span></div>`).join('');
  const bidDetails=ev.bids&&Array.isArray(ev.bids.headers)&&Array.isArray(ev.bids.rows)?`<h4>Details of Bidders</h4><div class="bidder-details-grid">${ev.bids.rows.map(row=>{
    const bidder=row[0]||'Unnamed bidder';
    const fields=ev.bids.headers.slice(1).map((header,i)=>({header,value:row[i+1]||''})).filter(x=>String(x.value).trim());
    return `<article class="bidder-detail"><h5>${esc(bidder)}</h5>${fields.length?fields.map(x=>`<div class="detail"><b>${esc(x.header)}</b><span>${esc(x.value)}</span></div>`).join(''):'<div class="meta">No additional bidder details detected.</div>'}</article>`;
  }).join('')}</div>`:'';
  const nonEligibility=Array.isArray(ev.recommendations)?ev.recommendations.filter(r=>!r.eligible&&Array.isArray(r.issues)&&r.issues.length).map(r=>`<article class="non-eligibility"><h5>${esc(r.bidder)}</h5><ul>${r.issues.map(x=>`<li><b>${esc(x.area)}:</b> ${esc(x.status)}</li>`).join('')}</ul></article>`).join(''):'';
  const reasons=nonEligibility?`<div class="non-eligibility-section"><h4>Why bidders are not eligible for commercial evaluation</h4>${nonEligibility}</div>`:'';
  const recommendationTable=Array.isArray(ev.recommendation_table)&&ev.recommendation_table.length?`<section class="recommendation-table-section"><h4>5. Recommendations for Commercial Evaluation</h4><p class="recommendation-table-note">Review each bidder below. Partially complied bidders may proceed; non-complied bidders are excluded.</p><div class="tw"><table class="recommendation-table"><thead><tr><th>FIELD</th>${ev.recommendation_table.map(r=>`<th>${esc(r.bidder)}</th>`).join('')}</tr></thead><tbody>
    <tr><th>Recommended for Financial Bid Opening (Yes/No)</th>${ev.recommendation_table.map(r=>`<td class="${r.financial_bid_opening==='Yes'?'ok':'bad'}">${esc(r.financial_bid_opening)}</td>`).join('')}</tr>
    <tr><th>Evaluated BOM included? (Yes/No)</th>${ev.recommendation_table.map(r=>`<td>${esc(r.evaluated_bom_included)}</td>`).join('')}</tr>
    <tr><th>Missing price items needing separate sealed envelope? (Yes/No — list items)</th>${ev.recommendation_table.map(r=>`<td>${esc(r.missing_price_items)}</td>`).join('')}</tr>
    <tr><th>Previous purchase price (with period & quantity)</th>${ev.recommendation_table.map(r=>`<td>${esc(r.previous_purchase_price)}</td>`).join('')}</tr>
    <tr><th>Reference price of similar items (market/previous)</th>${ev.recommendation_table.map(r=>`<td>${esc(r.reference_price)}</td>`).join('')}</tr>
    <tr><th>More Suitable</th>${ev.recommendation_table.map(r=>`<td>${esc(r.more_suitable)}</td>`).join('')}</tr>
  </tbody></table></div></section>`:'';
  const mats=ev.matrices.map(m=>`<div class="tw"><table class="mx"><thead><tr>${m.headers.map(x=>`<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>`+
    m.rows.map(r=>`<tr><td>${esc(r.area)}</td>${r.cells.map(c=>`<td class="${c.c}">${esc(c.t)}</td>`).join('')}</tr>`).join('')+
    (m.overall?`<tr class="ov"><td>Overall result (as stated in report)</td>${m.overall.map(c=>`<td class="${c.c}">${esc(c.t)}</td>`).join('')}</tr>`:'')+
    `</tbody></table></div>`).join('<br>');
  return `<div class="evalbox"><h3>${ev.type==='SSET'?'Sole Source Evaluation':'Technical Evaluation'} — compliance by company</h3><div class="meta">File: ${esc(ev.label)}</div><div class="report-details"><h4>Report details</h4>${metaRows||'<div class="meta">No additional report details were detected.</div>'}</div><div class="recommendations">${bidDetails||'<h4>Details of Bidders</h4><div class="meta">No bidder details were detected.</div>'}</div>${mats}${recommendationTable}${reasons}
    ${idx===0?`<a class="xl" href="/api/download/${j.job}/eval">⬇ Download Excel (Tender No: ${esc(tender)})</a>`:''}</div>`}}
function priceBox(j,probs){
  const rows=(Array.isArray(j.rows)?j.rows:[]).slice().sort((a,b)=>{
    return String(a.bidder||'').localeCompare(String(b.bidder||''));
  });
  const amount=value=>value===null||value===undefined?'Not available':Number(value).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
  const body=rows.map(r=>`<tr><td>${esc(r.bidder)}</td><td>${esc(r.currency||'Not stated')}</td><td>${esc(amount(r.unit_price_excl_taxes))}</td><td>${esc(amount(r.unit_price_incl_taxes))}</td></tr>`).join('');
  const notes=(j.notes||[]).concat(rows.flatMap(r=>r.warnings||[]));
  return `<div class="res"><h3>✓ Price evaluation ready</h3><div class="price-summary">All bidder prices (${rows.length})</div><table class="price-table"><thead><tr><th>Bidder</th><th>Currency</th><th>Total Unit Price (excl. taxes)</th><th>Total Unit Price (incl. taxes)</th></tr></thead><tbody>${body||'<tr><td colspan="4">No bidder prices detected.</td></tr>'}</tbody></table>${notes.length?`<ul class="warn">${notes.map(esc).map(x=>`<li>${x}</li>`).join('')}</ul>`:''}<a class="xl" href="/api/download/${j.job}/price">⬇ Download Price Schedules Excel</a>${probs?`<div>${probs}</div>`:''}</div>`;
}
panel(document.getElementById('p-all'),'all');
document.querySelectorAll('.tabs button').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('.tabs button').forEach(x=>x.classList.toggle('on',x===b));
  document.querySelectorAll('.panel').forEach(p=>p.hidden=p.id!=='p-'+b.dataset.tab);
  try{location.hash=b.dataset.tab}catch(e){}});
if(location.hash&&document.querySelector(`[data-tab="${location.hash.slice(1)}"]`))document.querySelector(`[data-tab="${location.hash.slice(1)}"]`).click();
