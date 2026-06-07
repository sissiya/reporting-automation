document.getElementById('upload').addEventListener('click', async () => {
  const f = document.getElementById('file').files[0]
  if (!f) { alert('Select a JSON file'); return }
  const fd = new FormData()
  fd.append('file', f)
  document.getElementById('message').innerText = 'Uploading...'
  const res = await fetch('/upload_json', { method: 'POST', body: fd })
  const data = await res.json()
  console.log('upload_json response', data)
  document.getElementById('message').innerText = `Rows: ${data.count}`
  const sel = document.getElementById('entitySelect')
  sel.innerHTML = ''
  let entities = data.entities || []
  // fallback: if entities empty but tables keys exist, use them
  const tablesFallback = data.tables || {}
  if ((!entities || entities.length===0) && Object.keys(tablesFallback).length>0){
    entities = Object.keys(tablesFallback)
  }
  entities.forEach(e => {
    const o = document.createElement('option'); o.value = e; o.text = e; sel.appendChild(o)
  })
  // store tables for preview
  window.__preview_tables = data.tables || {}
  window.__preview_rows = data.rows || []
  if (!data.tables || Object.keys(data.tables).length===0){
    console.warn('No per-entity tables returned from /upload_json')
    // show raw rows and a diagnostic blob to help debugging
    document.getElementById('message').innerText += ' — no summary tables returned'
    const dbg = document.createElement('pre')
    dbg.style.maxHeight = '200px'
    dbg.style.overflow = 'auto'
    dbg.innerText = JSON.stringify(Object.keys(data).reduce((acc,k)=>{acc[k]=Array.isArray(data[k])? (data[k].length>5? data[k].slice(0,5): data[k]) : data[k]; return acc},{}), null, 2)
    const container = document.getElementById('table')
    container.innerHTML = ''
    container.appendChild(dbg)
  }
  updateSummariesFromSelection()
})

// On page load, warm-up and render the overall summary so users see it immediately
(async function preloadAndRenderSummary(){
  try{
    document.getElementById('message').innerText = 'Loading default summary...'
    const res = await fetch('/dev_tables')
    if (!res.ok) { document.getElementById('message').innerText = ''; return }
    const data = await res.json()
    if (data && data.tables){
      window.__preview_tables = data.tables || {}
      document.getElementById('message').innerText = `Loaded ${Object.keys(window.__preview_tables).length} entities`
      renderSummary()
      // clear message after short delay
      setTimeout(()=>{ document.getElementById('message').innerText = '' }, 1000)
    }
  }catch(e){
    console.error('preload summary error', e)
    document.getElementById('message').innerText = ''
  }
})()

// Select All / Clear buttons
document.getElementById('selectAll').addEventListener('click', (ev)=>{
  ev.preventDefault()
  const sel = document.getElementById('entitySelect')
  Array.from(sel.options).forEach(o=>o.selected = true)
  updateSummariesFromSelection()
})
document.getElementById('clearSelection').addEventListener('click', (ev)=>{
  ev.preventDefault()
  const sel = document.getElementById('entitySelect')
  Array.from(sel.options).forEach(o=>o.selected = false)
  updateSummariesFromSelection()
})
document.getElementById('previewSelected').addEventListener('click', (ev)=>{ ev.preventDefault(); updateSummariesFromSelection() })

function renderTable(rows){
  const container = document.getElementById('table')
  container.innerHTML = ''
  if (!rows || rows.length===0) { container.innerText = 'No preview'; return }
  const table = document.createElement('table')
  const hdr = document.createElement('tr')
  Object.keys(rows[0]).slice(0,10).forEach(k => { const th = document.createElement('th'); th.innerText = k; hdr.appendChild(th) })
  table.appendChild(hdr)
  rows.forEach(r => {
    const tr = document.createElement('tr')
    Object.keys(r).slice(0,10).forEach(k => { const td = document.createElement('td'); td.innerText = r[k]; tr.appendChild(td) })
    table.appendChild(tr)
  })
  container.appendChild(table)
}

function renderPreview(){
  const sel = document.getElementById('entitySelect')
  const selected = Array.from(sel.selectedOptions).map(o=>o.value).filter(v=>v && v.trim())
  const container = document.getElementById('table')
  container.innerHTML = ''
  const tables = window.__preview_tables || {}
  if (!selected || selected.length===0){
    renderSummary()
    return
  }
  // render one section per selected entity
  selected.forEach(val=>{
    if (!tables[val]) return
    const rows = tables[val]
      // show raw rows preview (first 20 rows for the selected entity)
      const rawRows = (window.__preview_rows || []).filter(r=> (r['Entité']||r['Entity']||r['entite']||r['Entite'])==val).slice(0,20)
      if (rawRows && rawRows.length>0){
        const rawLabel = document.createElement('div')
        rawLabel.className = 'ppt-entity-label'
        rawLabel.innerText = `Preview rows for: ${val}`
        container.appendChild(rawLabel)
        const rawTable = document.createElement('table')
        rawTable.className = 'preview-rows'
        const hdr = document.createElement('tr')
        Object.keys(rawRows[0]).slice(0,6).forEach(k => { const th=document.createElement('th'); th.innerText = k; hdr.appendChild(th) })
        rawTable.appendChild(hdr)
        rawRows.forEach(r=>{
          const tr=document.createElement('tr')
          Object.keys(r).slice(0,6).forEach(k=>{ const td=document.createElement('td'); td.innerText = r[k]; tr.appendChild(td) })
          rawTable.appendChild(tr)
        })
        container.appendChild(rawTable)
      }

    const label = document.createElement('div')
    label.className = 'ppt-entity-label'
    label.innerText = `Entité: ${val}`
    container.appendChild(label)

    const table = document.createElement('table')
    table.className = 'ppt-table'
    const thead = document.createElement('thead')
    const hdr = document.createElement('tr')
    ['Catégorie','Volume','Part du total'].forEach(h => { const th = document.createElement('th'); th.innerText = h; hdr.appendChild(th) })
    thead.appendChild(hdr)
    table.appendChild(thead)
    const tbody = document.createElement('tbody')
    const getCell = (obj, names)=>{
      for(const n of names){ if (n in obj) return obj[n] }
      // try normalized names
      const nrm = names.map(x=>String(x).normalize('NFD').replace(/\p{Diacritic}/gu, '') )
      for(const k of Object.keys(obj||{})){
        const kn = String(k).normalize('NFD').replace(/\p{Diacritic}/gu, '')
        for(const nn of nrm) if (kn.toLowerCase()===nn.toLowerCase()) return obj[k]
      }
      return ''
    }

    rows.forEach((r, idx) => {
      const tr = document.createElement('tr')
      const tdCat = document.createElement('td')
      tdCat.className = 'category-cell'
      tdCat.innerText = getCell(r, ['Catégorie','Categorie','Cat�gorie','Cat'])
      tr.appendChild(tdCat)

      const tdVol = document.createElement('td')
      tdVol.innerText = getCell(r, ['Volume','Vol','volume'])
      tr.appendChild(tdVol)

      const tdPct = document.createElement('td')
      let pct = String(getCell(r, ['Part du total','Part_du_total','Part_du_total','Part_du_total']))
      pct = pct.replace(/\s*%/, ' %')
      tdPct.innerText = pct
      tr.appendChild(tdPct)

      tbody.appendChild(tr)
    })
    table.appendChild(tbody)
    container.appendChild(table)
  })
}

// New: when selection changes, fetch summaries from backend
async function fetchEntitySummaries(entities){
  if (!entities || entities.length===0) return {}
  try{
    const res = await fetch('/entity-summary', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({entities})
    })
    if (!res.ok) throw new Error('fetch failed')
    const data = await res.json()
    return data
  }catch(e){
    console.error('fetchEntitySummaries error', e)
    return {}
  }
}

async function updateSummariesFromSelection(){
  const sel = document.getElementById('entitySelect')
  const selected = Array.from(sel.selectedOptions).map(o=>o.value).filter(v=>v && v.trim())
  if (!selected || selected.length===0){
    // clear and show overall summary
    renderSummary()
    return
  }
  document.getElementById('message').innerText = 'Loading summaries...'
  const summaries = await fetchEntitySummaries(selected)
  // store returned tables for rendering and export
  window.__preview_tables = summaries || {}
  document.getElementById('message').innerText = `Loaded ${Object.keys(window.__preview_tables).length} summaries`
  renderPreview()
}

// wire select change to update summaries
document.getElementById('entitySelect').addEventListener('change', updateSummariesFromSelection)

// Show overall summary button: use cached tables or fetch /dev_tables
document.getElementById('showSummary').addEventListener('click', async (ev)=>{
  ev && ev.preventDefault && ev.preventDefault()
  const btn = document.getElementById('showSummary')
  btn.disabled = true
  const container = document.getElementById('table')
  container.innerHTML = ''
  document.getElementById('message').innerText = 'Loading overall summary...'
  if (window.__preview_tables && Object.keys(window.__preview_tables).length>0){
    renderSummary()
    document.getElementById('message').innerText = 'Summary (from cached tables)'
    btn.disabled = false
    return
  }
  try{
    // If user has selected a file, POST it to /upload_json to compute tables using same logic as PPT
    const f = document.getElementById('file').files[0]
    if (f){
      const fd = new FormData()
      fd.append('file', f)
      const res = await fetch('/upload_json', { method: 'POST', body: fd })
      if (!res.ok) throw new Error('upload_json failed')
      const data = await res.json()
      if (data && data.tables){
        window.__preview_tables = data.tables || {}
        document.getElementById('message').innerText = `Loaded ${Object.keys(window.__preview_tables).length} entities (from uploaded file)`
        renderSummary()
        btn.disabled = false
        return
      }
    }
    // fallback to server sample cache
    const res2 = await fetch('/dev_tables')
    if (!res2.ok) throw new Error('dev_tables fetch failed')
    const data2 = await res2.json()
    if (data2 && data2.tables){
      window.__preview_tables = data2.tables || {}
      document.getElementById('message').innerText = `Loaded ${Object.keys(window.__preview_tables).length} entities from server sample`
      renderSummary()
      btn.disabled = false
      return
    }
    document.getElementById('message').innerText = 'No summary tables returned from server'
    btn.disabled = false
  }catch(e){
    console.error('showSummary error', e)
    document.getElementById('message').innerText = 'Error loading summary'
    btn.disabled = false
  }
})

// Export handler: include selected entities (comma-separated) in the form
document.getElementById('export').addEventListener('click', async () => {
  const f = document.getElementById('file').files[0]
  if (!f) { alert('Select a JSON file'); return }
  const sel = document.getElementById('entitySelect')
  const selected = Array.from(sel.selectedOptions).map(o => o.value).filter(v=>v && v.trim())
  // Use async background export endpoint and poll job status
  const fd = new FormData()
  fd.append('file', f)
  if (selected.length>0) fd.append('entities', JSON.stringify(selected))
  document.getElementById('message').innerText = 'Starting background export...'
  try{
    const res = await fetch('/generate_ppt_async', { method: 'POST', body: fd })
    const data = await res.json()
    if (!data.job_id){ document.getElementById('message').innerText = 'Failed to start job'; return }
    const jobId = data.job_id
    document.getElementById('message').innerText = `Job started: ${jobId}. Polling...`
    // poll status
    const poll = setInterval(async ()=>{
      try{
        const st = await fetch(`/job-status/${jobId}`)
        if (!st.ok) throw new Error('status fetch failed')
        const js = await st.json()
        if (js.status === 'pending' || js.status === 'running'){
          document.getElementById('message').innerText = `Job ${jobId}: ${js.status}`
          return
        }
        if (js.status === 'done'){
          clearInterval(poll)
          const a = document.createElement('a')
          a.href = js.ppt_path
          a.innerText = 'Download PPT'
          a.target = '_blank'
          document.getElementById('message').innerHTML = ''
          document.getElementById('message').appendChild(a)
          return
        }
        if (js.status === 'failed'){
          clearInterval(poll)
          document.getElementById('message').innerText = `Job failed: ${js.error || 'unknown'}`
          return
        }
      }catch(e){
        clearInterval(poll)
        document.getElementById('message').innerText = 'Error polling job';
      }
    }, 1200)
  }catch(e){
    console.error(e)
    document.getElementById('message').innerText = 'Error starting export'
  }
})

function renderSummary(){
  const container = document.getElementById('table')
  container.innerHTML = ''
  const tables = window.__preview_tables || {}
  const entities = Object.keys(tables)
  if (!entities || entities.length===0){
    // fall back to raw rows if no tables
    renderTable(window.__preview_rows || [])
    return
  }
  const table = document.createElement('table')
  const hdr = document.createElement('tr')
  ['Entity','Total','Top categories'].forEach(h => { const th = document.createElement('th'); th.innerText = h; hdr.appendChild(th) })
  table.appendChild(hdr)
  entities.forEach(ent => {
    const rows = tables[ent]
    let total = 0
    const cats = []
    rows.forEach(r => { const vol = Number(r['Volume']||0); total += vol; if (vol>0) cats.push({k: r['Catégorie'], v: vol}) })
    cats.sort((a,b)=>b.v-a.v)
    const top = cats.slice(0,3).map(c=>`${c.k}: ${c.v}`).join(', ')
    const tr = document.createElement('tr')
    tr.style.cursor = 'pointer'
    tr.addEventListener('click', ()=>{
      const sel = document.getElementById('entitySelect')
      Array.from(sel.options).forEach(o=> o.selected = (o.value === ent))
      renderPreview()
    })
    const tdEnt = document.createElement('td'); tdEnt.innerText = ent; tr.appendChild(tdEnt)
    const tdTotal = document.createElement('td'); tdTotal.innerText = total; tr.appendChild(tdTotal)
    const tdTop = document.createElement('td'); tdTop.innerText = top; tr.appendChild(tdTop)
    table.appendChild(tr)
  })
  container.appendChild(table)
}

// (export handler is defined earlier and supports selected entities)
