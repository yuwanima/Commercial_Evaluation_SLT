
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const strip=n=>n.replace(/\.[^.]+$/,'');
function guessRaw(n){return strip(n).replace(/_/g," ").replace(/\b(GCC|SCC|compliance|for|general conditions of contract|special conditions of contract|copy of)\b/ig,' ').replace(/[_\-]+/g,' ').replace(/\s+/g,' ').trim()}
function guess(n){return guessRaw(n)||strip(n)}
const EXT=/\.(xlsx|xlsm|xls|docx|doc|pdf|zip)$/i, EVALNAME=/evaluation|\bTET\b|SSET/i;
// bidder name from a folder path: Bidders/Alpha/GCC.xlsx -> Alpha ; Alpha/GCC.xlsx -> Alpha
function folderName(parts){return parts.length>=3?parts[1]:''}

function panel(root,mode){
  const items=[],zone=root.querySelector('.zone'),ul=root.querySelector('ul.files'),cnt=root.querySelector('.count'),out=root.querySelector('.out'),go=root.querySelector('.go'),goText=go.textContent;
  const kindOf=n=>/\.zip$/i.test(n)?'ZIP':mode==='eval'?'EVAL':mode==='price'?'PRICE':/\.(xlsx|xlsm|xls)$/i.test(n)?'GCC':/\.pdf$/i.test(n)?(/scc/i.test(n)?'SCC':/gcc/i.test(n)?'GCC':'PDF'):'SCC';
  const named=it=>it.kind!=='ZIP'&&it.kind!=='EVAL';
  function render(extra){
    ul.innerHTML='';
    items.forEach((it,i)=>{
      const li=document.createElement('li');
      li.innerHTML=`<span class="fn"><span class="tag ${it.kind}">${it.kind}</span> ${esc(it.path||it.file.name)}</span>`+(named(it)?`<input type="text" value="${esc(it.name)}" title="Bidder name" placeholder="Bidder name">`:'')+`<button type="button" title="Remove">×</button>`;
      const inp=li.querySelector('input');if(inp)inp.oninput=e=>it.name=e.target.value;
      li.querySelector('button').onclick=()=>{items.splice(i,1);render()};
      ul.append(li);
    });
    const c=k=>items.filter(x=>x.kind===k).length;
    cnt.textContent=!items.length?'':mode==='eval'?`${items.length} report file(s) ready`:mode==='price'?`${items.length} price schedule file(s) ready`:`${c('GCC')} GCC · ${c('SCC')} SCC`+(c('PDF')?` · ${c('PDF')} PDF (type detected automatically)`:'')+(c('ZIP')?` · ${c('ZIP')} zip`:'')+' ready';
    out.innerHTML=(extra&&extra.length)?`<div class="msg">${extra.map(esc).join('<br>')}</div>`:'';
  }
  function add(list){           // list: [{file, parts}] parts = path segments or null
    const msgs=[];let skipped=0;
    list.forEach(({file,parts})=>{
      if(!EXT.test(file.name)){skipped++;return}
      if(file.name.startsWith('~$')){return}
      if(mode==='bids'&&/\.(docx|doc|pdf)$/i.test(file.name)&&EVALNAME.test(file.name)){msgs.push(`${file.name}: looks like an evaluation report — upload it in the Technical Evaluation tab`);return}
      const fromFolder=parts&&parts.length>1;
      items.push({file,kind:kindOf(file.name),path:fromFolder?parts.join('/'):null,
        name:mode==='eval'?strip(file.name):(fromFolder&&(folderName(parts)||guessRaw(file.name)||parts[0]))||guess(file.name)});
    });
    if(skipped)msgs.push(`${skipped} file(s) skipped (unsupported type)`);
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
      const probs=(j.problems||[]).length?'<b style="color:var(--bad);font-size:14px">Some files could not be read:</b><ul class="warn">'+j.problems.map(p=>'<li>'+esc(p)+'</li>').join('')+'</ul>':'';
      if(!r.ok){out.innerHTML=`<div class="msg">${esc(j.error||'Something went wrong')}${j.problems?'<ul class="warn">'+j.problems.map(p=>'<li>'+esc(p)+'</li>').join('')+'</ul>':''}</div>`;return}
      if(mode==='eval'){out.innerHTML=(j.evals||[]).map(evBox(j)).join('')+(probs?`<div class="res">${probs}</div>`:'');return}
      if(mode==='price'){out.innerHTML=priceBox(j,probs);return}
      const dl=(k,t,b)=>`<a class="${j[k]?'':'off'}" href="/api/download/${j.job}/${k}"><b>⬇ ${t}</b><span>${j[k]?b.length+' bidder(s): '+esc(b.join(', ')):'No files uploaded'}</span></a>`;
      out.innerHTML=`<div class="res"><h3>✓ Comparison ready</h3><div style="color:var(--mut);font-size:14px">Download your Excel files:</div>
        <div class="dl">${dl('gcc','GCC_Compliance_Comparison.xlsx',j.gcc_bidders)}${dl('scc','SCC_Compliance_Comparison.xlsx',j.scc_bidders)}</div>${probs}</div>`;
    }catch(e){out.innerHTML='<div class="msg">Could not reach the server: '+esc(String(e))+'</div>'}
    finally{go.disabled=false;go.textContent=goText}
  };
}

function evBox(j){return (ev,idx)=>{
  const meta=[['Tender No.',ev.meta['Tender No.']],['Subject',ev.meta['Subject']],['User Department',ev.meta['User Department']]].filter(x=>x[1]).map(x=>`<b>${x[0]}:</b> ${esc(x[1])}`).join(' &nbsp;•&nbsp; ');
  const mats=ev.matrices.map(m=>`<div class="tw"><table class="mx"><thead><tr>${m.headers.map(x=>`<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>`+
    m.rows.map(r=>`<tr><td>${esc(r.area)}</td>${r.cells.map(c=>`<td class="${c.c}">${esc(c.t)}</td>`).join('')}</tr>`).join('')+
    (m.overall?`<tr class="ov"><td>Overall result (as stated in report)</td>${m.overall.map(c=>`<td class="${c.c}">${esc(c.t)}</td>`).join('')}</tr>`:'')+
    (m.eligible?`<tr class="ov"><td>Eligible for Commercial Evaluation</td>${m.eligible.map(e=>`<td class="${e?'ok':'bad'}">${e?'Yes — proceed':'No — excluded'}</td>`).join('')}</tr>`:'')+
    `</tbody></table></div>`).join('<br>');
  return `<div class="evalbox"><h3>${ev.type==='SSET'?'Sole Source Evaluation':'Technical Evaluation'} — compliance by company</h3><div class="meta">${meta||''}</div><div class="meta" style="margin-top:-8px">File: ${esc(ev.label)}</div>${mats}
    ${idx===0?`<a class="xl" href="/api/download/${j.job}/eval">⬇ Download Excel (editable)</a>`:''}</div>`}}
function priceBox(j,probs){
  const rows=Array.isArray(j.rows)?j.rows:[];
  const body=rows.map(r=>`<tr><td>${esc(r.bidder)}</td><td>${esc(r.currency||'Not stated')}</td><td>${Number(r.total||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>${esc(r.status||'')}</td><td>${esc(r.source||'')}</td></tr>`).join('');
  const notes=(j.notes||[]).concat(rows.flatMap(r=>r.warnings||[]));
  return `<div class="res"><h3>✓ Price evaluation ready</h3><table class="price-table"><thead><tr><th>Bidder</th><th>Currency</th><th>Total</th><th>Ranking</th><th>Source</th></tr></thead><tbody>${body}</tbody></table>${notes.length?`<ul class="warn">${notes.map(esc).map(x=>`<li>${x}</li>`).join('')}</ul>`:''}<a class="xl" href="/api/download/${j.job}/price">⬇ Download Price Schedules Excel</a>${probs?`<div>${probs}</div>`:''}</div>`;
}
panel(document.getElementById('p-bids'),'bids');panel(document.getElementById('p-eval'),'eval');
panel(document.getElementById('p-price'),'price');
document.querySelectorAll('.tabs button').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('.tabs button').forEach(x=>x.classList.toggle('on',x===b));
  document.querySelectorAll('.panel').forEach(p=>p.hidden=p.id!=='p-'+b.dataset.tab);
  try{location.hash=b.dataset.tab}catch(e){}});
if(location.hash==='#eval')document.querySelector('[data-tab=eval]').click();
