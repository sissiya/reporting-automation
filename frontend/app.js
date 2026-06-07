async function postJSON(url, data){
  const r = await fetch(url, {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  return r.json();
}

let currentRows = [];

function renderTable(rows){
  const tbl = document.getElementById('tbl');
  const thead = tbl.querySelector('thead');
  const tbody = tbl.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';
  if(!rows || rows.length===0) return;
  const keys = Object.keys(rows[0]);
  const tr = document.createElement('tr');
  keys.forEach(k=>{const th=document.createElement('th'); th.textContent=k; tr.appendChild(th)});
  thead.appendChild(tr);
  rows.forEach(r=>{
    const tr2 = document.createElement('tr');
    keys.forEach(k=>{const td=document.createElement('td'); td.textContent = r[k]===null? '': r[k]; tr2.appendChild(td)});
    tbody.appendChild(tr2);
  });
}

function populateEntities(rows){
  const sel = document.getElementById('entity');
  const set = new Set();
  rows.forEach(r=>{ if(r['Entité']) set.add(r['Entité']) });
  sel.innerHTML = '<option value="__all__">All</option>' + Array.from(set).map(e=>`<option value="${e}">${e}</option>`).join('');
}

document.getElementById('upload').addEventListener('click',async ()=>{
  const f = document.getElementById('file').files[0];
  if(!f){ alert('Select a JSON file'); return }
  const form = new FormData();
  form.append('file', f);
  const r = await fetch('/upload_json', {method:'POST', body: form});
  const j = await r.json();
  if(j.status!=='ok'){ document.getElementById('notice').textContent = 'Upload failed: '+ (j.detail||''); return }
  currentRows = j.rows;
  renderTable(currentRows);
  populateEntities(currentRows);
  document.getElementById('notice').textContent = 'Loaded '+ currentRows.length + ' rows';
});

document.getElementById('entity').addEventListener('change', ()=>{
  const v = document.getElementById('entity').value;
  if(v==='__all__') renderTable(currentRows);
  else renderTable(currentRows.filter(r=>r['Entité']===v));
});

document.getElementById('export').addEventListener('click', async ()=>{
  const v = document.getElementById('entity').value;
  let rows = currentRows;
  let fname = 'incident_analysis_api.pptx';
  if(v!=='__all__'){
    rows = rows.filter(r=>r['Entité']===v);
    fname = `incident_analysis_${v.replace(/[^a-z0-9]/gi,'_')}.pptx`;
  }
  if(!rows || rows.length===0){ alert('No rows to export'); return }
  document.getElementById('notice').textContent = 'Exporting...';
  const resp = await postJSON('/generate_ppt_file', {data: rows, filename: fname});
  if(resp.status==='ok'){
    const url = '/outputs/' + fname;
    document.getElementById('notice').innerHTML = `Done: <a href="${url}" target="_blank">Download PPT</a>`;
  } else {
    document.getElementById('notice').textContent = 'Export failed: ' + (resp.detail||'');
  }
});
