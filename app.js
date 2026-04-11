// ─── Stato applicazione ───────────────────────────────────────
  const state = {
    sidebarOpen: false,
    theme: 'dark',
    activeSection: 'inserimento',
  };

  // ─── Apps Script endpoint (JSONP — aggira CORS) ───────────────
  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbysjN2wGXA8Pn_MloHyw2pUcwzwcz4e4escZD4ggGUwPWeQeFEfnm_S3VoPB1P_DGfUQQ/exec';

  function asCall(params) {
    return new Promise((resolve, reject) => {
      const cbName = '_bm_cb_' + Date.now();
      const script = document.createElement('script');
      const timeout = setTimeout(() => {
        cleanup();
        reject(new Error('Timeout'));
      }, 15000);

      window[cbName] = (data) => {
        cleanup();
        resolve(data);
      };

      function cleanup() {
        clearTimeout(timeout);
        delete window[cbName];
        if (script.parentNode) script.parentNode.removeChild(script);
      }

      const allParams = { ...params, callback: cbName };
      script.src = APPS_SCRIPT_URL + '?' + new URLSearchParams(allParams).toString();
      script.onerror = () => { cleanup(); reject(new Error('Script error')); };
      document.head.appendChild(script);
    });
  }

  // ─── Elementi DOM ─────────────────────────────────────────────
  const layout       = document.getElementById('layout');
  const hamburgerBtn = document.getElementById('hamburgerBtn');
  const themeToggle  = document.getElementById('themeToggle');
  const navItems     = document.querySelectorAll('.nav-item');
  const sections     = document.querySelectorAll('.section');

  // ─── Sidebar toggle ───────────────────────────────────────────
  function toggleSidebar(force) {
    state.sidebarOpen = (force !== undefined) ? force : !state.sidebarOpen;
    layout.classList.toggle('sidebar-open', state.sidebarOpen);
    hamburgerBtn.setAttribute('aria-expanded', state.sidebarOpen);
  }

  hamburgerBtn.addEventListener('click', () => toggleSidebar());

  // ─── Navigazione sezioni ──────────────────────────────────────
  navItems.forEach(item => {
    item.addEventListener('click', () => {
      const target = item.dataset.section;
      if (target === state.activeSection) return;

      navItems.forEach(n => n.classList.remove('active'));
      sections.forEach(s => s.classList.remove('active'));

      item.classList.add('active');
      document.getElementById('sec-' + target).classList.add('active');
      state.activeSection = target;
    });
  });

  // ─── Tema ─────────────────────────────────────────────────────
  themeToggle.addEventListener('click', () => {
    state.theme = (state.theme === 'dark') ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', state.theme === 'light' ? 'light' : '');
    updateThemeIcon();
    localStorage.setItem('instrumenta-theme', state.theme);
  });

  updateThemeIcon();

  function updateThemeIcon() {
    const isDark = state.theme === 'dark';
    themeToggle.innerHTML = isDark
      ? `<svg width="16" height="16" viewBox="0 0 16 16" fill="none">
           <circle cx="8" cy="8" r="3.5" stroke="currentColor" stroke-width="1.4"/>
           <line x1="8" y1="1" x2="8" y2="2.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
           <line x1="8" y1="13.5" x2="8" y2="15" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
           <line x1="1" y1="8" x2="2.5" y2="8" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
           <line x1="13.5" y1="8" x2="15" y2="8" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
           <line x1="3.05" y1="3.05" x2="4.11" y2="4.11" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
           <line x1="11.89" y1="11.89" x2="12.95" y2="12.95" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
           <line x1="12.95" y1="3.05" x2="11.89" y2="4.11" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
           <line x1="4.11" y1="11.89" x2="3.05" y2="12.95" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
         </svg>`
      : `<svg width="16" height="16" viewBox="0 0 16 16" fill="none">
           <path d="M13.5 10.5A6 6 0 0 1 5.5 2.5a6 6 0 1 0 8 8z" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
         </svg>`;
  }

  // Ripristino tema salvato (default: light)
  const savedTheme = localStorage.getItem('instrumenta-theme') ?? 'light';
  state.theme = savedTheme;
  if (savedTheme === 'light') {
    document.documentElement.setAttribute('data-theme', 'light');
    updateThemeIcon();
  }

  // ─── Form Inserimento ─────────────────────────────────────────
  const CATEGORIE = {
    uscita: {
      principali: ['Abbonamenti','Auto','Cibo','Mutuo','Previdenza c.','Svago','Spesa'],
      extra: ['Acquisti online','Altro','Beauty','Bolletta gas','Bolletta luce','Bolletta rifiuti','Bolletta wifi','Casa','Regali','Riscaldamento','Scarpe','Spese condominiali','Spese mediche','Telefonia','Trasporti','Vestiti','Viaggi']
    },
    entrata: {
      principali: ['Stipendio','Affitto','Regali','Ripetizioni','Trasferimenti','Altro'],
      extra: []
    }
  };

  let currentType = 'uscita';
  let selectedCategory = '';
  let selectedPayer = 'Silvio';

  const categoryGrid = document.getElementById('categoryGrid');
  const typeToggle   = document.getElementById('typeToggle');
  const payerToggle  = document.getElementById('payerToggle');
  const btnSubmit    = document.getElementById('btnSubmit');
  const btnReset     = document.getElementById('btnReset');
  const formFeedback = document.getElementById('formFeedback');

  // Imposta data odierna di default
  const fieldData = document.getElementById('fieldData');
  fieldData.valueAsDate = new Date();

  function makePill(cat) {
    const pill = document.createElement('div');
    pill.className = 'cat-pill';
    pill.textContent = cat;
    pill.addEventListener('click', () => {
      document.querySelectorAll('.cat-pill').forEach(p => p.classList.remove('selected','uscita','entrata'));
      pill.classList.add('selected', currentType);
      selectedCategory = cat;
    });
    return pill;
  }

  function renderCategories() {
    selectedCategory = '';
    categoryGrid.innerHTML = '';
    const cats = CATEGORIE[currentType];

    cats.principali.forEach(cat => categoryGrid.appendChild(makePill(cat)));

    if (cats.extra.length > 0) {
      const toggleBtn = document.createElement('div');
      toggleBtn.className = 'cat-pill cat-more';
      toggleBtn.innerHTML = `<span>altro</span> <svg width="10" height="10" viewBox="0 0 10 10" fill="none"><path d="M2 4l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
      toggleBtn.dataset.open = 'false';

      const extraWrap = document.createElement('div');
      extraWrap.className = 'cat-extra-wrap';
      cats.extra.forEach(cat => extraWrap.appendChild(makePill(cat)));

      toggleBtn.addEventListener('click', () => {
        const isOpen = toggleBtn.dataset.open === 'true';
        toggleBtn.dataset.open = isOpen ? 'false' : 'true';
        extraWrap.classList.toggle('visible', !isOpen);
        toggleBtn.innerHTML = !isOpen
          ? `<span>meno</span> <svg width="10" height="10" viewBox="0 0 10 10" fill="none"><path d="M2 6l3-3 3 3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>`
          : `<span>altro</span> <svg width="10" height="10" viewBox="0 0 10 10" fill="none"><path d="M2 4l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
      });

      categoryGrid.appendChild(toggleBtn);
      categoryGrid.appendChild(extraWrap);
    }
  }

  function setType(type) {
    currentType = type;
    document.querySelectorAll('.type-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.type === type);
    });
    btnSubmit.className = 'btn-submit' + (type === 'entrata' ? ' entrata' : '');
    renderCategories();
  }

  typeToggle.querySelectorAll('.type-btn').forEach(btn => {
    btn.addEventListener('click', () => setType(btn.dataset.type));
  });

  payerToggle.querySelectorAll('.payer-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      payerToggle.querySelectorAll('.payer-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      selectedPayer = btn.dataset.payer;
    });
  });

  function showFeedback(msg, type) {
    formFeedback.textContent = msg;
    formFeedback.className = `form-feedback visible ${type}`;
    setTimeout(() => { formFeedback.className = 'form-feedback'; }, 4000);
  }

  function resetForm() {
    document.getElementById('fieldData').valueAsDate = new Date();
    document.getElementById('fieldCosto').value = '';
    document.getElementById('fieldDescrizione').value = '';
    document.getElementById('fieldCondivisa').checked = false;
    selectedCategory = '';
    renderCategories();
  }

  btnReset.addEventListener('click', resetForm);

  btnSubmit.addEventListener('click', async () => {
    const data        = document.getElementById('fieldData').value;
    const costo       = document.getElementById('fieldCosto').value;
    const descrizione = document.getElementById('fieldDescrizione').value.trim();
    const condivisa   = document.getElementById('fieldCondivisa').checked;

    if (!data || !costo || !descrizione || !selectedCategory) {
      showFeedback('⚠ Compila tutti i campi obbligatori.', 'error');
      return;
    }

    const dateObj       = new Date(data);
    const mese          = dateObj.getMonth() + 1;
    const anno          = dateObj.getFullYear();
    const tipo          = currentType === 'entrata' ? 'Entrate' : 'Uscite';
    const costoNum      = parseFloat(costo.replace(',', '.'));
    const costoFmt      = costoNum.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const row = [data, costoFmt, descrizione, selectedCategory, tipo, mese, anno, selectedPayer, condivisa ? 'TRUE' : 'FALSE'];

    btnSubmit.textContent = 'Salvataggio…';
    btnSubmit.disabled = true;

    try {
      const json = await asCall({ action: 'append', row: JSON.stringify(row) });
      if (json.status !== 'ok') throw new Error(json.message);
      showFeedback('✓ Voce salvata correttamente su Google Sheets.', 'success');
      resetForm();
    } catch(e) {
      showFeedback('✗ Errore durante il salvataggio. Riprova.', 'error');
    } finally {
      btnSubmit.innerHTML = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M2 7h10M8 3l4 4-4 4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg> Salva su Sheets`;
      btnSubmit.disabled = false;
    }
  });

  // ─── Anno corrente nel badge ───────────────────────────────────
  document.getElementById('yearBadge').textContent = new Date().getFullYear();
  renderCategories();

  // ─── Elenco spese ─────────────────────────────────────────────
  const TUTTE_CAT = [...CATEGORIE.uscita.principali, ...CATEGORIE.uscita.extra, ...CATEGORIE.entrata.principali].filter((v,i,a)=>a.indexOf(v)===i).sort();

  let allRows = [];
  let sortDir = 'desc'; // default: più recente prima
  let deleteTarget = null;

  const speseBody    = document.getElementById('speseBody');
  const speseTable   = document.getElementById('speseTable');
  const tableLoading = document.getElementById('tableLoading');
  const tableEmpty   = document.getElementById('tableEmpty');
  const elencoCount  = document.getElementById('elencoCount');
  const deleteModal  = document.getElementById('deleteModal');
  const modalDesc    = document.getElementById('modalDesc');

  // Popola select categoria nei filtri
  const fCatEl = document.getElementById('fCategoria');
  TUTTE_CAT.forEach(c => { const o = document.createElement('option'); o.value=c; o.textContent=c; fCatEl.appendChild(o); });

  // Carica dati da Apps Script
  // ─── Caricamento globale ──────────────────────────────────────
  // Banner di caricamento iniziale
  const loadingBanner = (() => {
    const el = document.createElement('div');
    el.id = 'globalLoading';
    el.style.cssText = `position:fixed;bottom:20px;right:20px;z-index:999;
      background:var(--bg2);border:1px solid var(--border-idle);border-radius:12px;
      padding:12px 18px;font-size:13px;color:var(--text-secondary);
      display:flex;align-items:center;gap:10px;box-shadow:0 4px 20px rgba(0,0,0,0.15);`;
    el.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" class="spin" style="flex-shrink:0;">
      <circle cx="8" cy="8" r="6" stroke="var(--border-hover)" stroke-width="2"/>
      <path d="M8 2a6 6 0 0 1 6 6" stroke="var(--accent-blue)" stroke-width="2" stroke-linecap="round"/>
    </svg> Caricamento dati…`;
    document.body.appendChild(el);
    return el;
  })();

  async function initApp() {
    try {
      const json = await asCall({ action: 'read_all' });
      if (json.status !== 'ok') throw new Error(json.message);

      // Popola allRows (spese)
      allRows = json.rows || [];

      // Popola mAllRate (mutuo)
      mAllRate = json.mutuo || [];

      loadingBanner.remove();

      // Se la sezione attiva ha bisogno di dati, renderizzala subito
      const s = state.activeSection;
      if (s === 'elenco')   { populateAnnoFilter(); applyFilters(); }
      if (s === 'annuale')  renderDashAnnuale();
      if (s === 'generale') renderDashGenerale();
      if (s === 'tabelle')  renderTabelle();
      if (s === 'utenze')   renderUtenze();
      if (s === 'mutuo')    buildMutuo(mAllRate);

    } catch(e) {
      loadingBanner.innerHTML = `<span style="color:var(--accent-red);">⚠ Errore caricamento. Ricarica la pagina.</span>`;
      setTimeout(() => loadingBanner.remove(), 5000);
    }
  }

  async function loadSheetData() {
    // Ora è solo un alias che riapplica i filtri se i dati ci sono già
    if (allRows.length) {
      tableLoading.style.display = 'none';
      populateAnnoFilter();
      applyFilters();
      return;
    }
    // Altrimenti mostra loading e aspetta initApp
    tableLoading.style.display = 'flex';
    speseTable.style.display = 'none';
    tableEmpty.style.display = 'none';
  }

  function populateAnnoFilter() {
    const fAnnoEl = document.getElementById('fAnno');
    const anni = [...new Set(allRows.map(r => r.anno).filter(Boolean))].sort((a,b) => b-a);
    fAnnoEl.innerHTML = '<option value="">Tutti</option>';
    anni.forEach(a => { const o = document.createElement('option'); o.value=a; o.textContent=a; fAnnoEl.appendChild(o); });
  }

  function applyFilters() {
    const fAnno     = document.getElementById('fAnno').value;
    const fMese     = document.getElementById('fMese').value;
    const fTipo     = document.getElementById('fTipo').value;
    const fCat      = document.getElementById('fCategoria').value;
    const fPag      = document.getElementById('fPagante').value;
    const fCond     = document.getElementById('fCondiviso').value;

    const filtered = allRows.filter(r => {
      if (fAnno && r.anno !== fAnno) return false;
      if (fMese && String(r.mese) !== fMese) return false;
      if (fTipo && r.tipo !== fTipo) return false;
      if (fCat  && r.categoria !== fCat) return false;
      if (fPag  && r.pagante !== fPag) return false;
      if (fCond && r.condiviso.toUpperCase() !== fCond) return false;
      return true;
    });

    // Ordina per data
    filtered.sort((a, b) => {
      const da = new Date(a.data), db = new Date(b.data);
      return sortDir === 'desc' ? db - da : da - db;
    });

    elencoCount.textContent = `${filtered.length} ${filtered.length === 1 ? 'voce' : 'voci'}`;
    renderTable(filtered);
  }

  // Sort click sul th Data
  document.getElementById('thData').addEventListener('click', () => {
    sortDir = sortDir === 'desc' ? 'asc' : 'desc';
    const th = document.getElementById('thData');
    th.classList.toggle('sort-asc', sortDir === 'asc');
    th.classList.toggle('sort-desc', sortDir === 'desc');
    applyFilters();
  });

  function renderTable(rows) {
    tableLoading.style.display = 'none';
    speseBody.innerHTML = '';

    if (rows.length === 0) {
      speseTable.style.display = 'none';
      tableEmpty.textContent = allRows.length === 0 ? 'Nessun dato trovato. Connettiti e premi Aggiorna.' : 'Nessuna voce trovata con i filtri selezionati.';
      tableEmpty.style.display = 'block';
      return;
    }

    speseTable.style.display = 'table';
    tableEmpty.style.display = 'none';

    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.dataset.rowIndex = r.rowIndex;
      const isUscita = r.tipo === 'Uscite';
      const condLabel = r.condiviso.toUpperCase() === 'TRUE' ? 'Sì' : 'No';

      tr.innerHTML = `
        <td>${formatDataDisplay(r.data)}</td>
        <td>${escHtml(r.descrizione)}</td>
        <td><span class="tag tag-neutral">${escHtml(r.categoria)}</span></td>
        <td><span class="tag ${isUscita ? 'tag-red' : 'tag-green'}">${escHtml(r.tipo)}</span></td>
        <td class="td-costo ${isUscita ? 'uscita' : 'entrata'}">${escHtml(r.costo)} €</td>
        <td>${escHtml(r.pagante)}</td>
        <td>${condLabel}</td>
        <td>
          <div class="td-actions">
            <button class="btn-row edit" title="Modifica" onclick="startEdit(this, ${r.rowIndex})">
              <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M9 2l2 2-7 7H2V9L9 2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/></svg>
            </button>
            <button class="btn-row delete" title="Elimina" onclick="askDelete(${r.rowIndex}, '${escAttr(r.descrizione)}', '${escAttr(r.costo)}')">
              <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M2 3.5h9M5 3.5V2.5h3v1M4 3.5l.5 7h4l.5-7" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </button>
          </div>
        </td>`;
      speseBody.appendChild(tr);
    });
  }

  function formatDataDisplay(d) {
    if (!d) return '';
    const p = d.split('-');
    if (p.length === 3) return `${p[2]}/${p[1]}/${p[0]}`;
    return d;
  }
  function escHtml(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function escAttr(s) { return String(s).replace(/'/g,"\\'"); }

  // ─── Modifica inline ──────────────────────────────────────────
  window.startEdit = function(btn, rowIndex) {
    const tr = document.querySelector(`tr[data-row-index="${rowIndex}"]`);
    if (!tr || tr.classList.contains('editing')) return;
    const row = allRows.find(r => r.rowIndex === rowIndex);
    if (!row) return;
    tr.classList.add('editing');

    const allCat = [...CATEGORIE.uscita.principali, ...CATEGORIE.uscita.extra, ...CATEGORIE.entrata.principali].filter((v,i,a)=>a.indexOf(v)===i).sort();
    const catOpts = allCat.map(c => `<option value="${c}" ${c===row.categoria?'selected':''}>${c}</option>`).join('');
    const tipoOpts = ['Uscite','Entrate'].map(t => `<option value="${t}" ${t===row.tipo?'selected':''}>${t}</option>`).join('');
    const pagOpts  = ['Silvio','Giulia'].map(p => `<option value="${p}" ${p===row.pagante?'selected':''}>${p}</option>`).join('');
    const condOpts = [['TRUE','Sì'],['FALSE','No']].map(([v,l]) => `<option value="${v}" ${row.condiviso.toUpperCase()===v?'selected':''}>${l}</option>`).join('');

    tr.cells[0].innerHTML = `<input class="inline-input" id="ei-data-${rowIndex}" type="date" value="${row.data}" style="width:130px">`;
    tr.cells[1].innerHTML = `<input class="inline-input" id="ei-desc-${rowIndex}" type="text" value="${escHtml(row.descrizione)}" style="min-width:140px">`;
    tr.cells[2].innerHTML = `<select class="inline-select" id="ei-cat-${rowIndex}">${catOpts}</select>`;
    tr.cells[3].innerHTML = `<select class="inline-select" id="ei-tipo-${rowIndex}">${tipoOpts}</select>`;
    tr.cells[4].innerHTML = `<input class="inline-input" id="ei-costo-${rowIndex}" type="text" value="${row.costo}" style="width:80px;text-align:right">`;
    tr.cells[5].innerHTML = `<select class="inline-select" id="ei-pag-${rowIndex}">${pagOpts}</select>`;
    tr.cells[6].innerHTML = `<select class="inline-select" id="ei-cond-${rowIndex}">${condOpts}</select>`;
    tr.cells[7].innerHTML = `
      <div class="td-actions">
        <button class="btn-row save" title="Salva" onclick="saveEdit(${rowIndex})">
          <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M2 7l3.5 3.5L11 3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </button>
        <button class="btn-row cancel" title="Annulla" onclick="cancelEdit(${rowIndex})">
          <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M3 3l7 7M10 3l-7 7" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
        </button>
      </div>`;
  };

  window.cancelEdit = function(rowIndex) {
    applyFilters();
  };

  window.saveEdit = async function(rowIndex) {
    const data       = document.getElementById(`ei-data-${rowIndex}`).value;
    const desc       = document.getElementById(`ei-desc-${rowIndex}`).value;
    const cat        = document.getElementById(`ei-cat-${rowIndex}`).value;
    const tipo       = document.getElementById(`ei-tipo-${rowIndex}`).value;
    const costoRaw   = document.getElementById(`ei-costo-${rowIndex}`).value.replace(',','.');
    const pag        = document.getElementById(`ei-pag-${rowIndex}`).value;
    const cond       = document.getElementById(`ei-cond-${rowIndex}`).value;

    const dateObj  = new Date(data);
    const mese     = dateObj.getMonth() + 1;
    const anno     = dateObj.getFullYear();
    const costoFmt = parseFloat(costoRaw).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2});

    try {
      const json = await asCall({ action: 'update', rowIndex, row: JSON.stringify([data, costoFmt, desc, cat, tipo, mese, anno, pag, cond]) });
      if (json.status !== 'ok') throw new Error(json.message);
      const idx = allRows.findIndex(r => r.rowIndex === rowIndex);
      if (idx > -1) allRows[idx] = { rowIndex, data, costo: costoFmt, descrizione: desc, categoria: cat, tipo, mese: String(mese), anno: String(anno), pagante: pag, condiviso: cond };
      applyFilters();
    } catch(e) {
      alert('Errore nel salvataggio. Riprova.');
    }
  };

  // ─── Eliminazione ─────────────────────────────────────────────
  window.askDelete = function(rowIndex, desc, costo) {
    deleteTarget = rowIndex;
    modalDesc.textContent = `"${desc}" — ${costo} €`;
    deleteModal.style.display = 'flex';
  };

  document.getElementById('btnCancelDelete').addEventListener('click', () => {
    deleteModal.style.display = 'none';
    deleteTarget = null;
  });

  document.getElementById('btnConfirmDelete').addEventListener('click', async () => {
    if (!deleteTarget) return;
    const rowIndex = deleteTarget;
    deleteModal.style.display = 'none';
    deleteTarget = null;

    try {
      const json = await asCall({ action: 'delete', rowIndex });
      if (json.status !== 'ok') throw new Error(json.message);
      allRows = allRows.filter(r => r.rowIndex !== rowIndex);
      allRows.forEach(r => { if (r.rowIndex > rowIndex) r.rowIndex--; });
      applyFilters();
    } catch(e) {
      alert('Errore nell\'eliminazione. Riprova.');
    }
  });

  // Filtri: listener
  ['fAnno','fMese','fTipo','fCategoria','fPagante','fCondiviso'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyFilters);
  });

  document.getElementById('btnClearFilters').addEventListener('click', () => {
    ['fAnno','fMese','fTipo','fCategoria','fPagante','fCondiviso'].forEach(id => {
      document.getElementById(id).value = '';
    });
    applyFilters();
  });

  document.getElementById('btnReload').addEventListener('click', () => {
    allRows = [];
    mAllRate = [];
    initApp().then(() => { populateAnnoFilter(); applyFilters(); });
  });

  // ─── Dashboard Annuale ────────────────────────────────────────
  const MESI_LABEL    = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
  const MESI_COMPLETI = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];

  let chartMensile = null, chartConfronto = null, chartDonut = null, chartWaterfall = null, chartCondivise = null;

  function destroyCharts() {
    [chartMensile, chartConfronto, chartDonut, chartWaterfall, chartCondivise].forEach(c => { if (c) { c.destroy(); } });
    chartMensile = chartConfronto = chartDonut = chartWaterfall = chartCondivise = null;
    ['chartMensile','chartConfronto','chartDonut','chartWaterfall','chartCondivise'].forEach(id => {
      const old = document.getElementById(id);
      if (old) { const nc = document.createElement('canvas'); nc.id = id; old.parentNode.replaceChild(nc, old); }
    });
  }

  function fmtEur(val) {
    return val.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
  }

  function parseCosto(s) {
    if (!s) return 0;
    return parseFloat(String(s).trim()) || 0;
  }

  function getDashFilters() {
    return {
      anno:    document.getElementById('annoDashSelect').value,
      mese:    document.getElementById('meseDashSelect').value,
      pagante: document.getElementById('paganteDashSelect').value,
    };
  }

  function populateAnnoSelect() {
    const sel  = document.getElementById('annoDashSelect');
    const anni = [...new Set(allRows.map(r => r.anno).filter(Boolean))].sort((a,b) => b - a);
    sel.innerHTML = anni.map(a => `<option value="${a}">${a}</option>`).join('');
    const cur = String(new Date().getFullYear());
    if (anni.includes(cur)) sel.value = cur;
  }

  function filterRows(rows, { anno, mese, pagante }) {
    return rows.filter(r => {
      if (anno    && r.anno    !== anno)    return false;
      if (mese    && String(r.mese) !== mese) return false;
      if (pagante && r.pagante !== pagante) return false;
      return true;
    });
  }

  function renderDashAnnuale() {
    if (!allRows.length) return;
    populateAnnoSelect();
    buildDash();
  }

  ['annoDashSelect','meseDashSelect','paganteDashSelect'].forEach(id => {
    document.getElementById(id).addEventListener('change', buildDash);
  });

  function buildDash() {
    destroyCharts();
    const f       = getDashFilters();
    const rows    = filterRows(allRows, f);
    const entrate = rows.filter(r => r.tipo === 'Entrate');
    const uscite  = rows.filter(r => r.tipo === 'Uscite');

    const totEnt = entrate.reduce((s,r) => s + parseCosto(r.costo), 0);
    const totUsc = uscite.reduce((s,r)  => s + parseCosto(r.costo), 0);
    const saldo  = totEnt - totUsc;

    // ── KPI ──
    document.getElementById('kpiEntrate').textContent = fmtEur(totEnt);
    document.getElementById('kpiUscite').textContent  = fmtEur(totUsc);
    const kpiSaldoEl = document.getElementById('kpiSaldo');
    kpiSaldoEl.textContent = fmtEur(saldo);
    kpiSaldoEl.className = 'kpi-value ' + (saldo >= 0 ? 'kpi-green' : 'kpi-red');
    const pct = totEnt > 0 ? ((saldo / totEnt) * 100).toFixed(1) + '%' : '—';
    document.getElementById('kpiSaldoPct').textContent = pct;

    // Barre percentuali entrate/uscite
    const totale = totEnt + totUsc;
    const pctEnt = totale > 0 ? (totEnt / totale * 100).toFixed(1) : 0;
    const pctUsc = totale > 0 ? (totUsc / totale * 100).toFixed(1) : 0;
    document.getElementById('kpiEntrateBar').style.width = pctEnt + '%';
    document.getElementById('kpiUsciteBar').style.width  = pctUsc + '%';

    // Tasso di risparmio
    const risparmio = totEnt > 0 ? ((saldo / totEnt) * 100) : 0;
    const kpiRisp = document.getElementById('kpiRisparmio');
    kpiRisp.textContent = risparmio.toFixed(1) + '%';
    kpiRisp.className = 'kpi-value ' + (risparmio >= 0 ? 'kpi-blue' : 'kpi-red');

    // Spesa media giornaliera
    const giorniPeriodo = (() => {
      if (f.mese) {
        return new Date(parseInt(f.anno), parseInt(f.mese), 0).getDate();
      }
      // Anno: conta i giorni con almeno una spesa
      const oggi = new Date();
      const fineAnno = new Date(parseInt(f.anno), 11, 31);
      const riferimento = fineAnno < oggi ? fineAnno : oggi;
      const inizioAnno = new Date(parseInt(f.anno), 0, 1);
      return Math.max(1, Math.round((riferimento - inizioAnno) / 86400000) + 1);
    })();
    const mediaGiorno = totUsc / giorniPeriodo;
    document.getElementById('kpiMediaGiorno').textContent = fmtEur(mediaGiorno);
    document.getElementById('kpiMediaGiornoSub').textContent = `su ${giorniPeriodo} giorni`;

    // ── Top 10 uscite (per descrizione) ──
    const top10map = {};
    uscite.forEach(r => { top10map[r.descrizione] = (top10map[r.descrizione] || 0) + parseCosto(r.costo); });
    const top10 = Object.entries(top10map).sort((a,b) => b[1]-a[1]).slice(0,10);
    const top10max = top10[0]?.[1] || 1;
    const top10El = document.getElementById('top10List');
    top10El.innerHTML = top10.length ? top10.map(([nome, val]) => `
      <div class="topn-row">
        <span class="topn-name">${nome}</span>
        <div class="topn-bar-wrap"><div class="topn-bar" style="width:${(val/top10max*100).toFixed(1)}%"></div></div>
        <span class="topn-amount">${fmtEur(val)}</span>
      </div>`).join('') : '<p style="color:var(--text-hint);font-size:13px;padding:8px 0;">Nessun dato</p>';

    // ── Top 5 categorie — grafico donut ──
    const catMap = {};
    uscite.forEach(r => { catMap[r.categoria] = (catMap[r.categoria] || 0) + parseCosto(r.costo); });
    const top5cat = Object.entries(catMap).sort((a,b) => b[1]-a[1]).slice(0,5);
    const donutColors = ['#ff3b3b','#ff6b6b','#ff9a9a','#ffbebe','#ffdada'];

    chartDonut = new Chart(document.getElementById('chartDonut'), {
      type: 'doughnut',
      data: {
        labels: top5cat.map(c => c[0]),
        datasets: [{ data: top5cat.map(c => c[1]), backgroundColor: donutColors, borderWidth: 0, hoverOffset: 8 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        cutout: '58%',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed) } }
        },
        layout: { padding: 32 }
      },
      plugins: [{
        id: 'outerLabels',
        afterDraw(chart) {
          const { ctx, chartArea: { width, height, left, top } } = chart;
          const cx = left + width / 2;
          const cy = top  + height / 2;
          const meta = chart.getDatasetMeta(0);
          ctx.save();
          meta.data.forEach((arc, i) => {
            const angle  = (arc.startAngle + arc.endAngle) / 2;
            const r      = arc.outerRadius + 18;
            const x      = cx + Math.cos(angle) * r;
            const y      = cy + Math.sin(angle) * r;
            const label  = chart.data.labels[i];
            const val    = chart.data.datasets[0].data[i];
            const total  = chart.data.datasets[0].data.reduce((a,b) => a+b, 0);
            const pctLbl = total > 0 ? (val / total * 100).toFixed(0) + '%' : '';
            const align  = x < cx ? 'right' : 'left';
            ctx.fillStyle = donutColors[i];
            ctx.font      = 'bold 11px DM Sans, sans-serif';
            ctx.textAlign = align;
            ctx.fillText(label, x, y - 6);
            ctx.fillStyle = isDark ? '#888' : '#666';
            ctx.font      = '10px DM Sans, sans-serif';
            ctx.fillText(pctLbl, x, y + 7);
          });
          ctx.restore();
        }
      }]
    });

    const isDark    = document.documentElement.getAttribute('data-theme') !== 'light';
    const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    const textColor = isDark ? '#888' : '#666';
    Chart.defaults.color      = textColor;
    Chart.defaults.font.family = 'DM Sans';
    Chart.defaults.font.size   = 11;

    const ticksY = { color: textColor, callback: v => v.toLocaleString('it-IT') + ' €' };
    const ticksX = { color: textColor };

    // ── Grafico andamento: se filtro mese → barre orizzontali entrate/uscite, altrimenti barre mensili ──
    if (f.mese) {
      // Vista mese: barra orizzontale entrate vs uscite
      document.getElementById('chartMensileTitle').textContent = `Entrate / Uscite — ${MESI_COMPLETI[parseInt(f.mese)-1]}`;
      chartMensile = new Chart(document.getElementById('chartMensile'), {
        type: 'bar',
        data: {
          labels: ['Entrate', 'Uscite'],
          datasets: [{
            data: [totEnt, totUsc],
            backgroundColor: ['rgba(46,204,113,0.75)', 'rgba(255,59,59,0.75)'],
            borderRadius: 6, borderSkipped: false
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed.x) } } },
          scales: {
            x: { grid: { color: gridColor }, ticks: ticksY },
            y: { grid: { display: false }, ticks: ticksX }
          }
        }
      });
    } else {
      // Vista anno: barre mensili entrate + uscite
      const ent12 = Array(12).fill(0);
      const usc12 = Array(12).fill(0);
      entrate.forEach(r => { const m = parseInt(r.mese); if (m >= 1 && m <= 12) ent12[m-1] += parseCosto(r.costo); });
      uscite.forEach(r  => { const m = parseInt(r.mese); if (m >= 1 && m <= 12) usc12[m-1] += parseCosto(r.costo); });
      document.getElementById('chartMensileTitle').textContent = 'Entrate / Uscite mensili';
      chartMensile = new Chart(document.getElementById('chartMensile'), {
        type: 'bar',
        data: { labels: MESI_LABEL, datasets: [
          { label: 'Entrate', data: ent12, backgroundColor: 'rgba(46,204,113,0.7)', borderRadius: 6, borderSkipped: false },
          { label: 'Uscite',  data: usc12, backgroundColor: 'rgba(255,59,59,0.7)',  borderRadius: 6, borderSkipped: false }
        ]},
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' } },
          scales: { x: { grid: { color: gridColor }, ticks: ticksX }, y: { grid: { color: gridColor }, ticks: ticksY } } }
      });
    }

    // ── Waterfall guadagno mensile ──
    {
      const ent12 = Array(12).fill(0);
      const usc12 = Array(12).fill(0);
      entrate.forEach(r => { const m = parseInt(r.mese); if (m>=1&&m<=12) ent12[m-1] += parseCosto(r.costo); });
      uscite.forEach(r  => { const m = parseInt(r.mese); if (m>=1&&m<=12) usc12[m-1] += parseCosto(r.costo); });

      const delta = ent12.map((e,i) => e - usc12[i]); // guadagno mensile
      let cumulo = 0;
      // Barre flottanti: [base, top]
      const barData = delta.map(d => {
        const base = cumulo;
        cumulo += d;
        return d >= 0 ? [base, base + d] : [base + d, base];
      });
      const barColors = delta.map(d => d >= 0 ? 'rgba(46,204,113,0.85)' : 'rgba(255,59,59,0.85)');

      // Barra totale finale
      const totale = delta.reduce((a,b) => a+b, 0);
      barData.push([0, totale]);
      barColors.push('rgba(180,200,0,0.85)');
      const wfLabels = [...MESI_LABEL, 'Totale'];

      chartWaterfall = new Chart(document.getElementById('chartWaterfall'), {
        type: 'bar',
        data: {
          labels: wfLabels,
          datasets: [{
            data: barData,
            backgroundColor: barColors,
            borderRadius: 4,
            borderSkipped: false,
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: ctx => {
                  const [lo, hi] = ctx.raw;
                  const val = hi - lo;
                  return ' ' + (val >= 0 ? '+' : '') + fmtEur(val);
                }
              }
            }
          },
          scales: {
            x: { grid: { color: gridColor }, ticks: ticksX },
            y: {
              grid: { color: gridColor },
              ticks: { color: textColor, callback: v => v.toLocaleString('it-IT') }
            }
          }
        }
      });
    }

    // ── Donut spese condivise vs personali ──
    {
      const totCond = uscite.filter(r => r.condiviso.toUpperCase()==='TRUE').reduce((s,r) => s+parseCosto(r.costo), 0);
      const totPers = uscite.filter(r => r.condiviso.toUpperCase()!=='TRUE').reduce((s,r) => s+parseCosto(r.costo), 0);
      chartCondivise = new Chart(document.getElementById('chartCondivise'), {
        type: 'doughnut',
        data: {
          labels: ['Condivise', 'Personali'],
          datasets: [{ data: [totCond, totPers], backgroundColor: ['#5b9bff','#ff3b3b'], borderWidth: 0, hoverOffset: 6 }]
        },
        options: {
          responsive: true, maintainAspectRatio: false, cutout: '60%',
          plugins: {
            legend: { position: 'bottom', labels: { boxWidth: 12, padding: 16, font: { size: 12 } } },
            tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed) } }
          }
        },
        plugins: [{
          id: 'centerLabel',
          afterDraw(chart) {
            const { ctx, chartArea: { width, height, left, top } } = chart;
            const cx = left + width/2, cy = top + height/2 - 14;
            const tot = totCond + totPers;
            const pct = tot > 0 ? (totCond/tot*100).toFixed(0)+'%' : '—';
            ctx.save();
            ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            ctx.fillStyle = '#5b9bff';
            ctx.font = 'bold 20px DM Sans, sans-serif';
            ctx.fillText(pct, cx, cy);
            ctx.fillStyle = isDark ? '#888' : '#666';
            ctx.font = '11px DM Sans, sans-serif';
            ctx.fillText('condivise', cx, cy + 20);
            ctx.restore();
          }
        }]
      });
    }

    // ── Heatmap mese × categoria ──
    {
      // Categorie con almeno una spesa nel periodo
      const catSet = [...new Set(uscite.map(r => r.categoria).filter(Boolean))].sort();
      const matrix = {}; // matrix[cat][mese] = totale
      catSet.forEach(cat => { matrix[cat] = Array(12).fill(0); });
      uscite.forEach(r => {
        const m = parseInt(r.mese); if (m>=1&&m<=12&&matrix[r.categoria]) matrix[r.categoria][m-1] += parseCosto(r.costo);
      });
      // Max globale per normalizzare colori
      const allVals = catSet.flatMap(cat => matrix[cat]);
      const maxVal  = Math.max(...allVals, 1);

      const isDarkHm = document.documentElement.getAttribute('data-theme') !== 'light';
      const hmBg = isDarkHm ? '#1a1a1a' : '#ffffff';
      const hmBorder = isDarkHm ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';

      let html = `<table class="heatmap-table">
        <thead><tr><th></th>${MESI_LABEL.map(m=>`<th>${m}</th>`).join('')}</tr></thead><tbody>`;
      catSet.forEach(cat => {
        html += `<tr><td class="hm-label">${cat}</td>`;
        matrix[cat].forEach(val => {
          if (val === 0) {
            html += `<td style="background:${hmBorder};color:var(--text-hint);">—</td>`;
          } else {
            const intensity = val / maxVal;
            const alpha = (0.12 + intensity * 0.75).toFixed(2);
            html += `<td style="background:rgba(255,59,59,${alpha});color:${intensity>0.55?'#fff':'var(--text-primary)'};">${val.toLocaleString('it-IT',{minimumFractionDigits:0,maximumFractionDigits:0})}</td>`;
          }
        });
        html += '</tr>';
      });
      html += '</tbody></table>';
      document.getElementById('heatmapWrap').innerHTML = html;
    }

    // ── Grafico confronto periodo precedente (linee cumulate) ──
    let labelsCfr, dataCurr, dataPrev, titleCfr;

    if (f.mese) {
      // Confronto con mese precedente — accumulo giornaliero
      const meseNum  = parseInt(f.mese);
      const prevMese = meseNum === 1 ? 12 : meseNum - 1;
      const prevAnno = meseNum === 1 ? String(parseInt(f.anno) - 1) : f.anno;
      const giorniMese = new Date(parseInt(f.anno), meseNum, 0).getDate();

      const byDay     = Array(giorniMese).fill(0);
      const byDayPrev = Array(giorniMese).fill(0);

      uscite.forEach(r => {
        const d = new Date(r.data); if (isNaN(d)) return;
        const g = d.getDate() - 1; if (g >= 0 && g < giorniMese) byDay[g] += parseCosto(r.costo);
      });

      filterRows(allRows.filter(r => r.tipo === 'Uscite'), { anno: prevAnno, mese: String(prevMese), pagante: f.pagante }).forEach(r => {
        const d = new Date(r.data); if (isNaN(d)) return;
        const g = d.getDate() - 1; const dMax = new Date(parseInt(prevAnno), prevMese, 0).getDate();
        if (g >= 0 && g < Math.min(giorniMese, dMax)) byDayPrev[g] += parseCosto(r.costo);
      });

      // Accumulo cumulato
      let cumCurr = 0, cumPrev = 0;
      dataCurr = byDay.map(v     => { cumCurr += v; return cumCurr; });
      dataPrev = byDayPrev.map(v => { cumPrev += v; return cumPrev; });
      labelsCfr = Array.from({length: giorniMese}, (_,i) => i+1);
      const pm = MESI_COMPLETI[prevMese-1].slice(0,3);
      const py = prevAnno !== f.anno ? ` ${prevAnno}` : '';
      titleCfr  = `Confronto uscite — ${MESI_COMPLETI[meseNum-1]} vs ${pm}${py}`;

    } else {
      // Confronto con anno precedente — cumulo mensile
      const prevAnno = String(parseInt(f.anno) - 1);
      const usc12Curr = Array(12).fill(0);
      const usc12Prev = Array(12).fill(0);
      uscite.forEach(r => { const m = parseInt(r.mese); if (m>=1&&m<=12) usc12Curr[m-1] += parseCosto(r.costo); });
      filterRows(allRows.filter(r=>r.tipo==='Uscite'), { anno: prevAnno, mese: '', pagante: f.pagante })
        .forEach(r => { const m = parseInt(r.mese); if (m>=1&&m<=12) usc12Prev[m-1] += parseCosto(r.costo); });
      let cumC = 0, cumP = 0;
      dataCurr  = usc12Curr.map(v => { cumC += v; return cumC; });
      dataPrev  = usc12Prev.map(v => { cumP += v; return cumP; });
      labelsCfr = MESI_LABEL;
      titleCfr  = `Confronto uscite — ${f.anno} vs ${prevAnno}`;
    }

    document.getElementById('chartConfrontoTitle').textContent = titleCfr;
    chartConfronto = new Chart(document.getElementById('chartConfronto'), {
      type: 'line',
      data: { labels: labelsCfr, datasets: [
        { label: 'Corrente', data: dataCurr, borderColor: '#ff3b3b', backgroundColor: 'rgba(255,59,59,0.08)', borderWidth: 2, pointRadius: 2, fill: true, tension: 0.3 },
        { label: 'Precedente', data: dataPrev, borderColor: 'rgba(255,59,59,0.35)', backgroundColor: 'transparent', borderWidth: 1.5, pointRadius: 1, borderDash: [4,3], tension: 0.3 }
      ]},
      options: { responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top', labels: { boxWidth: 12, padding: 16 } },
          tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed.y) } } },
        scales: { x: { grid: { color: gridColor }, ticks: ticksX }, y: { grid: { color: gridColor }, ticks: ticksY } } }
    });

    // ── Tabella riepilogo mensile ──
    {
      const ent12 = Array(12).fill(0);
      const usc12 = Array(12).fill(0);
      entrate.forEach(r => { const m=parseInt(r.mese); if(m>=1&&m<=12) ent12[m-1]+=parseCosto(r.costo); });
      uscite.forEach(r  => { const m=parseInt(r.mese); if(m>=1&&m<=12) usc12[m-1]+=parseCosto(r.costo); });

      let html = `<thead><tr>
        <th>Mese</th><th>Entrate</th><th>Uscite</th><th>Saldo</th><th>Δ% mese prec.</th>
      </tr></thead><tbody>`;

      let prevUsc = null;
      MESI_COMPLETI.forEach((nome, i) => {
        const e = ent12[i], u = usc12[i], s = e - u;
        const delta = (prevUsc !== null && prevUsc > 0) ? ((u - prevUsc) / prevUsc * 100) : null;
        const deltaStr = delta === null ? '<span class="neu">—</span>'
          : `<span class="${delta > 0 ? 'neg' : delta < 0 ? 'pos' : 'neu'}">${delta > 0 ? '+' : ''}${delta.toFixed(1)}%</span>`;
        const rowClass = (e===0 && u===0) ? ' style="opacity:0.35"' : '';
        html += `<tr${rowClass}>
          <td>${nome}</td>
          <td class="pos">${fmtEur(e)}</td>
          <td class="neg">${fmtEur(u)}</td>
          <td class="${s>=0?'pos':'neg'}">${fmtEur(s)}</td>
          <td>${deltaStr}</td>
        </tr>`;
        if (u > 0) prevUsc = u;
      });

      // Riga totale
      const totE = ent12.reduce((a,b)=>a+b,0), totU = usc12.reduce((a,b)=>a+b,0);
      html += `</tbody><tfoot><tr class="tot-row">
        <td>Totale</td>
        <td class="pos">${fmtEur(totE)}</td>
        <td class="neg">${fmtEur(totU)}</td>
        <td class="${totE-totU>=0?'pos':'neg'}">${fmtEur(totE-totU)}</td>
        <td></td>
      </tr></tfoot>`;
      document.getElementById('tabellaRiepilogoMensile').innerHTML = html;
    }

    // ── Alert anomalie (mesi con uscite > media + 1.5σ) ──
    {
      const usc12 = Array(12).fill(0);
      uscite.forEach(r => { const m=parseInt(r.mese); if(m>=1&&m<=12) usc12[m-1]+=parseCosto(r.costo); });
      const attivi = usc12.filter(v => v > 0);
      if (attivi.length >= 3) {
        const media = attivi.reduce((a,b)=>a+b,0) / attivi.length;
        const varianza = attivi.reduce((s,v)=>s+Math.pow(v-media,2),0) / attivi.length;
        const sigma = Math.sqrt(varianza);
        const soglia = media + 1.5 * sigma;
        const anomalie = usc12.map((v,i)=>({mese:MESI_COMPLETI[i],val:v,i})).filter(x=>x.val>soglia);
        const wrap = document.getElementById('anomalieWrap');
        if (anomalie.length > 0) {
          wrap.style.display = 'block';
          document.getElementById('anomalieList').innerHTML = anomalie.map(a =>
            `<div class="anomalia-card">
              <div class="anomalia-dot"></div>
              <div class="anomalia-text">
                <strong>${a.mese}</strong> — uscite di <strong>${fmtEur(a.val)}</strong>,
                il <strong>${((a.val/media-1)*100).toFixed(0)}% sopra</strong> la media mensile (${fmtEur(media)}).
              </div>
            </div>`).join('');
        } else {
          wrap.style.display = 'none';
        }
      } else {
        document.getElementById('anomalieWrap').style.display = 'none';
      }
    }
  }

  // ─── Dashboard Generale ───────────────────────────────────────
  let gChartArea = null, gChartWaterfall = null, gChartDonutUsc = null, gChartDonutEnt = null;

  function destroyGeneraleCharts() {
    [gChartArea, gChartWaterfall, gChartDonutUsc, gChartDonutEnt].forEach(c => { if (c) c.destroy(); });
    gChartArea = gChartWaterfall = gChartDonutUsc = gChartDonutEnt = null;
    ['gChartArea','gChartWaterfall','gChartDonutUsc','gChartDonutEnt'].forEach(id => {
      const old = document.getElementById(id);
      if (old) { const nc = document.createElement('canvas'); nc.id = id; old.parentNode.replaceChild(nc, old); }
    });
  }

  function populateGAnnoSelect() {
    const sel  = document.getElementById('gAnnoSelect');
    const anni = [...new Set(allRows.map(r => r.anno).filter(Boolean))].sort((a,b) => b-a);
    const cur  = sel.value;
    sel.innerHTML = '<option value="">Tutti</option>' + anni.map(a => `<option value="${a}">${a}</option>`).join('');
    if (cur) sel.value = cur;
  }

  ['gAnnoSelect','gPaganteSelect'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderDashGenerale);
  });

  function buildDonut(canvasId, dataMap, colors, isDark) {
    const top5 = Object.entries(dataMap).sort((a,b)=>b[1]-a[1]).slice(0,5);
    if (!top5.length) return null;
    return new Chart(document.getElementById(canvasId), {
      type: 'doughnut',
      data: {
        labels: top5.map(c=>c[0]),
        datasets: [{ data: top5.map(c=>c[1]), backgroundColor: colors, borderWidth: 0, hoverOffset: 8 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '58%',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed) } }
        },
        layout: { padding: 30 }
      },
      plugins: [{
        id: 'outerLabels_' + canvasId,
        afterDraw(chart) {
          const { ctx, chartArea: { width, height, left, top } } = chart;
          const cx = left + width/2, cy = top + height/2;
          const meta = chart.getDatasetMeta(0);
          ctx.save();
          meta.data.forEach((arc, i) => {
            const angle = (arc.startAngle + arc.endAngle) / 2;
            const r = arc.outerRadius + 18;
            const x = cx + Math.cos(angle)*r, y = cy + Math.sin(angle)*r;
            const total = chart.data.datasets[0].data.reduce((a,b)=>a+b,0);
            const pct   = total > 0 ? (chart.data.datasets[0].data[i]/total*100).toFixed(0)+'%' : '';
            ctx.textAlign = x < cx ? 'right' : 'left';
            ctx.fillStyle = colors[i];
            ctx.font = 'bold 11px DM Sans, sans-serif';
            ctx.fillText(chart.data.labels[i], x, y-6);
            ctx.fillStyle = isDark ? '#888' : '#666';
            ctx.font = '10px DM Sans, sans-serif';
            ctx.fillText(pct, x, y+7);
          });
          ctx.restore();
        }
      }]
    });
  }

  function renderDashGenerale() {
    if (!allRows.length) return;
    populateGAnnoSelect();
    destroyGeneraleCharts();

    const fAnno    = document.getElementById('gAnnoSelect').value;
    const fPagante = document.getElementById('gPaganteSelect').value;

    const filtered = allRows.filter(r => {
      if (fAnno    && r.anno    !== fAnno)    return false;
      if (fPagante && r.pagante !== fPagante) return false;
      return true;
    });

    const uscite  = filtered.filter(r => r.tipo === 'Uscite');
    const entrate = filtered.filter(r => r.tipo === 'Entrate');
    const totEnt  = entrate.reduce((s,r) => s+parseCosto(r.costo), 0);
    const totUsc  = uscite.reduce((s,r)  => s+parseCosto(r.costo), 0);
    const saldo   = totEnt - totUsc;

    // KPI
    document.getElementById('gKpiEntrate').textContent = fmtEur(totEnt);
    document.getElementById('gKpiUscite').textContent  = fmtEur(totUsc);
    const gSaldoEl = document.getElementById('gKpiSaldo');
    gSaldoEl.textContent = fmtEur(saldo);
    gSaldoEl.className   = 'kpi-value ' + (saldo >= 0 ? 'kpi-green' : 'kpi-red');
    const totale = totEnt + totUsc;
    document.getElementById('gKpiEntrateBar').style.width = totale > 0 ? (totEnt/totale*100).toFixed(1)+'%' : '0%';
    document.getElementById('gKpiUsciteBar').style.width  = totale > 0 ? (totUsc/totale*100).toFixed(1)+'%' : '0%';
    const risparmio = totEnt > 0 ? (saldo/totEnt*100).toFixed(1) : 0;
    const gRisp = document.getElementById('gKpiRisparmio');
    gRisp.textContent = risparmio + '%';
    gRisp.className   = 'kpi-value ' + (risparmio >= 0 ? 'kpi-blue' : 'kpi-red');
    document.getElementById('gKpiSaldoPct').textContent = totEnt > 0 ? (saldo/totEnt*100).toFixed(1)+'%' : '—';
    const mesiUnici = new Set(uscite.map(r => r.anno+'-'+r.mese).filter(Boolean));
    const mediaM = mesiUnici.size ? totUsc/mesiUnici.size : 0;
    document.getElementById('gKpiMediaMensile').textContent = fmtEur(mediaM);
    document.getElementById('gKpiMediaMensileSub').textContent = `su ${mesiUnici.size} mesi`;

    // Top 10 uscite
    const top10map = {};
    uscite.forEach(r => { top10map[r.descrizione] = (top10map[r.descrizione]||0)+parseCosto(r.costo); });
    const top10 = Object.entries(top10map).sort((a,b)=>b[1]-a[1]).slice(0,10);
    const top10max = top10[0]?.[1] || 1;
    document.getElementById('gTop10List').innerHTML = top10.length
      ? top10.map(([nome,val]) => `<div class="topn-row"><span class="topn-name">${nome}</span><div class="topn-bar-wrap"><div class="topn-bar" style="width:${(val/top10max*100).toFixed(1)}%"></div></div><span class="topn-amount">${fmtEur(val)}</span></div>`).join('')
      : '<p style="color:var(--text-hint);font-size:13px;padding:8px 0;">Nessun dato</p>';

    const isDark    = document.documentElement.getAttribute('data-theme') !== 'light';
    const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    const textColor = isDark ? '#888' : '#666';
    Chart.defaults.color = textColor; Chart.defaults.font.family = 'DM Sans'; Chart.defaults.font.size = 11;
    const ticksY = { color: textColor, callback: v => v.toLocaleString('it-IT')+' €' };
    const ticksX = { color: textColor };

    // Donut uscite
    const catUscMap = {};
    uscite.forEach(r => { catUscMap[r.categoria] = (catUscMap[r.categoria]||0)+parseCosto(r.costo); });
    gChartDonutUsc = buildDonut('gChartDonutUsc', catUscMap, ['#ff3b3b','#ff6b6b','#ff9a9a','#ffbebe','#ffdada'], isDark);

    // Donut entrate
    const catEntMap = {};
    entrate.forEach(r => { catEntMap[r.categoria] = (catEntMap[r.categoria]||0)+parseCosto(r.costo); });
    gChartDonutEnt = buildDonut('gChartDonutEnt', catEntMap, ['#2ecc71','#5dd68a','#8ee0a3','#bfebbc','#d8f5d5'], isDark);

    // Anni filtrati
    const anni = [...new Set(filtered.map(r=>r.anno).filter(Boolean))].sort();

    // Grafico area
    const entPerAnno   = anni.map(a => entrate.filter(r=>r.anno===a).reduce((s,r)=>s+parseCosto(r.costo),0));
    const uscPerAnno   = anni.map(a => uscite.filter(r=>r.anno===a).reduce((s,r)=>s+parseCosto(r.costo),0));
    const saldoPerAnno = anni.map((_,i) => entPerAnno[i]-uscPerAnno[i]);
    gChartArea = new Chart(document.getElementById('gChartArea'), {
      type: 'line',
      data: { labels: anni, datasets: [
        { label: 'Entrate', data: entPerAnno,   borderColor: '#2ecc71', backgroundColor: 'rgba(46,204,113,0.12)', borderWidth: 2, fill: true,  tension: 0.35, pointRadius: 4 },
        { label: 'Uscite',  data: uscPerAnno,   borderColor: '#ff3b3b', backgroundColor: 'rgba(255,59,59,0.10)',  borderWidth: 2, fill: true,  tension: 0.35, pointRadius: 4 },
        { label: 'Saldo',   data: saldoPerAnno, borderColor: '#5b9bff', backgroundColor: 'transparent', borderWidth: 1.5, borderDash: [5,4], fill: false, tension: 0.35, pointRadius: 3 }
      ]},
      options: { responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top', labels: { boxWidth: 10, padding: 14 } }, tooltip: { callbacks: { label: ctx => ' '+ctx.dataset.label+': '+fmtEur(ctx.parsed.y) } } },
        scales: { x: { grid: { color: gridColor }, ticks: ticksX }, y: { grid: { color: gridColor }, ticks: ticksY } } }
    });

    // Waterfall trimestrale
    const trimestri = [], wfData = [], wfColors = [];
    anni.forEach(anno => {
      [1,2,3,4].forEach(q => {
        const mesiQ = q===1?[1,2,3]:q===2?[4,5,6]:q===3?[7,8,9]:[10,11,12];
        const eQ = entrate.filter(r=>r.anno===anno&&mesiQ.includes(parseInt(r.mese))).reduce((s,r)=>s+parseCosto(r.costo),0);
        const uQ = uscite.filter(r=>r.anno===anno&&mesiQ.includes(parseInt(r.mese))).reduce((s,r)=>s+parseCosto(r.costo),0);
        const delta = eQ - uQ;
        if (eQ > 0 || uQ > 0) { trimestri.push(`Q${q} ${anno}`); wfData.push(delta); wfColors.push(delta >= 0 ? 'rgba(46,204,113,0.8)' : 'rgba(255,59,59,0.8)'); }
      });
    });
    let cum = 0;
    const wfFloat = wfData.map(d => { const base=cum; cum+=d; return d>=0?[base,base+d]:[base+d,base]; });
    // Barra totale finale
    const wfTotale = wfData.reduce((a,b)=>a+b,0);
    trimestri.push('Totale');
    wfFloat.push([0, wfTotale]);
    wfColors.push('rgba(180,200,0,0.85)');

    gChartWaterfall = new Chart(document.getElementById('gChartWaterfall'), {
      type: 'bar',
      data: { labels: trimestri, datasets: [{ data: wfFloat, backgroundColor: wfColors, borderRadius: 4, borderSkipped: false }] },
      options: { responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => { const [lo,hi]=ctx.raw; const v=hi-lo; return ' '+(v>=0?'+':'')+fmtEur(v); } } }
        },
        scales: { x: { grid: { color: gridColor }, ticks: { color: textColor, maxRotation: 45 } }, y: { grid: { color: gridColor }, ticks: ticksY } } }
    });

    // Heatmap anno × categoria
    const catSet  = [...new Set(uscite.map(r=>r.categoria).filter(Boolean))].sort();
    const matrix  = {};
    catSet.forEach(cat => { matrix[cat] = {}; anni.forEach(a => { matrix[cat][a] = 0; }); });
    uscite.forEach(r => { if (matrix[r.categoria]) matrix[r.categoria][r.anno] = (matrix[r.categoria][r.anno]||0)+parseCosto(r.costo); });
    const allVals  = catSet.flatMap(cat => anni.map(a => matrix[cat][a]));
    const maxVal   = Math.max(...allVals, 1);
    const hmBorder = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    let html = `<table class="heatmap-table"><thead><tr><th></th>${anni.map(a=>`<th>${a}</th>`).join('')}</tr></thead><tbody>`;
    catSet.forEach(cat => {
      html += `<tr><td class="hm-label">${cat}</td>`;
      anni.forEach(a => {
        const val = matrix[cat][a] || 0;
        if (val===0) { html += `<td style="background:${hmBorder};color:var(--text-hint);">—</td>`; }
        else { const i=val/maxVal; const al=(0.12+i*0.75).toFixed(2); html += `<td style="background:rgba(255,59,59,${al});color:${i>0.55?'#fff':'var(--text-primary)'};">${val.toLocaleString('it-IT',{maximumFractionDigits:0})}</td>`; }
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    document.getElementById('gHeatmapWrap').innerHTML = html;

    // ── Tabella riepilogo per anno ──
    {
      let thtml = `<thead><tr>
        <th>Anno</th><th>Entrate</th><th>Uscite</th><th>Saldo</th>
        <th>Risp. %</th><th>Media mensile</th><th>Δ% anno prec.</th>
      </tr></thead><tbody>`;

      let prevUscAnno = null;
      anni.forEach(anno => {
        const eA = entrate.filter(r=>r.anno===anno).reduce((s,r)=>s+parseCosto(r.costo),0);
        const uA = uscite.filter(r=>r.anno===anno).reduce((s,r)=>s+parseCosto(r.costo),0);
        const sA = eA - uA;
        const rA = eA > 0 ? (sA/eA*100).toFixed(1) : '—';
        const mesiA = new Set(uscite.filter(r=>r.anno===anno).map(r=>r.mese).filter(Boolean));
        const medA  = mesiA.size ? uA/mesiA.size : 0;
        const delta = (prevUscAnno !== null && prevUscAnno > 0) ? ((uA-prevUscAnno)/prevUscAnno*100) : null;
        const deltaStr = delta===null ? '<span class="neu">—</span>'
          : `<span class="${delta>0?'neg':delta<0?'pos':'neu'}">${delta>0?'+':''}${delta.toFixed(1)}%</span>`;
        thtml += `<tr>
          <td>${anno}</td>
          <td class="pos">${fmtEur(eA)}</td>
          <td class="neg">${fmtEur(uA)}</td>
          <td class="${sA>=0?'pos':'neg'}">${fmtEur(sA)}</td>
          <td class="${parseFloat(rA)>=0?'pos':'neg'}">${rA !== '—' ? rA+'%' : '—'}</td>
          <td>${fmtEur(medA)}</td>
          <td>${deltaStr}</td>
        </tr>`;
        if (uA > 0) prevUscAnno = uA;
      });

      const totE2 = entrate.reduce((s,r)=>s+parseCosto(r.costo),0);
      const totU2 = uscite.reduce((s,r)=>s+parseCosto(r.costo),0);
      const totS2 = totE2 - totU2;
      thtml += `</tbody><tfoot><tr class="tot-row">
        <td>Totale</td>
        <td class="pos">${fmtEur(totE2)}</td>
        <td class="neg">${fmtEur(totU2)}</td>
        <td class="${totS2>=0?'pos':'neg'}">${fmtEur(totS2)}</td>
        <td></td><td></td><td></td>
      </tr></tfoot>`;
      document.getElementById('tabellaRiepilogoAnno').innerHTML = thtml;
    }
  }

  // ─── Sezione Tabelle ──────────────────────────────────────────

  // Tab switching (Uscite / Entrate)
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    });
  });

  // Filtri tabelle
  function populateTabAnnoSelect() {
    const sel  = document.getElementById('tabAnnoSelect');
    const anni = [...new Set(allRows.map(r => r.anno).filter(Boolean))].sort((a,b) => b-a);
    const cur  = sel.value;
    sel.innerHTML = anni.map(a => `<option value="${a}">${a}</option>`).join('');
    if (cur && anni.includes(cur)) sel.value = cur;
    else { const y = String(new Date().getFullYear()); if (anni.includes(y)) sel.value = y; }
  }

  function buildPivotAnno(rows, tipo, annoSel) {
    // Righe = categorie, colonne = mesi
    const cats = [...new Set(rows.filter(r=>r.tipo===tipo).map(r=>r.categoria).filter(Boolean))].sort();
    const table = {};
    cats.forEach(c => { table[c] = Array(12).fill(0); });
    rows.filter(r=>r.tipo===tipo).forEach(r => {
      const m = parseInt(r.mese); if (m>=1&&m<=12 && table[r.categoria]) table[r.categoria][m-1] += parseCosto(r.costo);
    });
    const totRow = Array(12).fill(0);
    cats.forEach(c => { table[c].forEach((v,i) => { totRow[i] += v; }); });

    const isEnt = tipo === 'Entrate';
    const isSaldo = tipo === '__saldo__';

    let h = `<thead><tr><th>${annoSel}</th>`;
    MESI_LABEL.forEach(m => { h += `<th>${m}</th>`; });
    h += `<th class="col-tot">Totale</th></tr></thead><tbody>`;

    cats.forEach(c => {
      const tot = table[c].reduce((a,b)=>a+b,0);
      if (tot === 0) return;
      h += `<tr><td>${c}</td>`;
      table[c].forEach(v => {
        h += v===0 ? `<td class="zero">—</td>` : `<td>${fmtEur(v)}</td>`;
      });
      h += `<td class="col-tot ${isEnt?'pos':isSaldo?(tot>=0?'pos':'neg'):'neg'}">${fmtEur(tot)}</td></tr>`;
    });

    const totTot = totRow.reduce((a,b)=>a+b,0);
    h += `</tbody><tfoot><tr class="tot-row"><td>Totale</td>`;
    totRow.forEach(v => { h += `<td class="${isEnt?'pos':isSaldo?(v>=0?'pos':'neg'):'neg'}">${fmtEur(v)}</td>`; });
    h += `<td class="col-tot ${isEnt?'pos':isSaldo?(totTot>=0?'pos':'neg'):'neg'}">${fmtEur(totTot)}</td></tr></tfoot>`;
    return h;
  }

  function buildPivotTutti(rows, tipo) {
    // Righe = categorie, colonne = anni
    const anni = [...new Set(rows.map(r=>r.anno).filter(Boolean))].sort();
    const cats = [...new Set(rows.filter(r=>r.tipo===tipo).map(r=>r.categoria).filter(Boolean))].sort();
    const table = {};
    cats.forEach(c => { table[c] = {}; anni.forEach(a => { table[c][a] = 0; }); });
    rows.filter(r=>r.tipo===tipo).forEach(r => {
      if (table[r.categoria]) table[r.categoria][r.anno] = (table[r.categoria][r.anno]||0) + parseCosto(r.costo);
    });

    const isEnt = tipo === 'Entrate';
    const isSaldo = tipo === '__saldo__';

    let h = `<thead><tr><th>Categoria</th>`;
    anni.forEach(a => { h += `<th>${a}</th>`; });
    h += `<th class="col-tot">Totale</th></tr></thead><tbody>`;

    cats.forEach(c => {
      const tot = anni.reduce((s,a) => s+(table[c][a]||0), 0);
      if (tot === 0) return;
      h += `<tr><td>${c}</td>`;
      anni.forEach(a => {
        const v = table[c][a]||0;
        h += v===0 ? `<td class="zero">—</td>` : `<td>${fmtEur(v)}</td>`;
      });
      h += `<td class="col-tot ${isEnt?'pos':isSaldo?(tot>=0?'pos':'neg'):'neg'}">${fmtEur(tot)}</td></tr>`;
    });

    // Riga totali
    const totAnno = {};
    anni.forEach(a => { totAnno[a] = cats.reduce((s,c)=>s+(table[c][a]||0),0); });
    const totTot = anni.reduce((s,a)=>s+totAnno[a],0);
    h += `</tbody><tfoot><tr class="tot-row"><td>Totale</td>`;
    anni.forEach(a => { h += `<td class="${isEnt?'pos':isSaldo?(totAnno[a]>=0?'pos':'neg'):'neg'}">${fmtEur(totAnno[a])}</td>`; });
    h += `<td class="col-tot ${isEnt?'pos':isSaldo?(totTot>=0?'pos':'neg'):'neg'}">${fmtEur(totTot)}</td></tr></tfoot>`;
    return h;
  }

  function buildPivotSaldoAnno(rows, annoSel) {
    // Saldo = entrate - uscite per categoria x mese
    const cats = [...new Set(rows.map(r=>r.categoria).filter(Boolean))].sort();
    const tableEnt = {}, tableUsc = {};
    cats.forEach(c => { tableEnt[c] = Array(12).fill(0); tableUsc[c] = Array(12).fill(0); });
    rows.filter(r=>r.tipo==='Entrate').forEach(r => { const m=parseInt(r.mese); if(m>=1&&m<=12&&tableEnt[r.categoria]) tableEnt[r.categoria][m-1]+=parseCosto(r.costo); });
    rows.filter(r=>r.tipo==='Uscite').forEach(r  => { const m=parseInt(r.mese); if(m>=1&&m<=12&&tableUsc[r.categoria]) tableUsc[r.categoria][m-1]+=parseCosto(r.costo); });

    let h = `<thead><tr><th>${annoSel} — Saldo</th>`;
    MESI_LABEL.forEach(m => { h += `<th>${m}</th>`; });
    h += `<th class="col-tot">Totale</th></tr></thead><tbody>`;

    cats.forEach(c => {
      const saldi = Array(12).fill(0).map((_,i) => (tableEnt[c][i]||0)-(tableUsc[c][i]||0));
      const tot = saldi.reduce((a,b)=>a+b,0);
      if (saldi.every(v=>v===0)) return;
      h += `<tr><td>${c}</td>`;
      saldi.forEach(v => { h += v===0?`<td class="zero">—</td>`:`<td class="${v>=0?'pos':'neg'}">${fmtEur(v)}</td>`; });
      h += `<td class="col-tot ${tot>=0?'pos':'neg'}">${fmtEur(tot)}</td></tr>`;
    });

    const totMese = Array(12).fill(0).map((_,i) => cats.reduce((s,c)=>(s+(tableEnt[c][i]||0)-(tableUsc[c][i]||0)),0));
    const totTot  = totMese.reduce((a,b)=>a+b,0);
    h += `</tbody><tfoot><tr class="tot-row"><td>Totale</td>`;
    totMese.forEach(v => { h += `<td class="${v>=0?'pos':'neg'}">${fmtEur(v)}</td>`; });
    h += `<td class="col-tot ${totTot>=0?'pos':'neg'}">${fmtEur(totTot)}</td></tr></tfoot>`;
    return h;
  }

  function buildPivotSaldoTutti(rows) {
    const anni = [...new Set(rows.map(r=>r.anno).filter(Boolean))].sort();
    const cats = [...new Set(rows.map(r=>r.categoria).filter(Boolean))].sort();
    const tableEnt = {}, tableUsc = {};
    cats.forEach(c => { tableEnt[c] = {}; tableUsc[c] = {}; anni.forEach(a => { tableEnt[c][a]=0; tableUsc[c][a]=0; }); });
    rows.filter(r=>r.tipo==='Entrate').forEach(r => { if(tableEnt[r.categoria]) tableEnt[r.categoria][r.anno]=(tableEnt[r.categoria][r.anno]||0)+parseCosto(r.costo); });
    rows.filter(r=>r.tipo==='Uscite').forEach(r  => { if(tableUsc[r.categoria]) tableUsc[r.categoria][r.anno]=(tableUsc[r.categoria][r.anno]||0)+parseCosto(r.costo); });

    let h = `<thead><tr><th>Categoria — Saldo</th>`;
    anni.forEach(a => { h += `<th>${a}</th>`; });
    h += `<th class="col-tot">Totale</th></tr></thead><tbody>`;

    cats.forEach(c => {
      const saldi = {};
      anni.forEach(a => { saldi[a] = (tableEnt[c][a]||0)-(tableUsc[c][a]||0); });
      const tot = anni.reduce((s,a)=>s+saldi[a],0);
      if (anni.every(a=>saldi[a]===0)) return;
      h += `<tr><td>${c}</td>`;
      anni.forEach(a => { h += saldi[a]===0?`<td class="zero">—</td>`:`<td class="${saldi[a]>=0?'pos':'neg'}">${fmtEur(saldi[a])}</td>`; });
      h += `<td class="col-tot ${tot>=0?'pos':'neg'}">${fmtEur(tot)}</td></tr>`;
    });

    const totAnno = {};
    anni.forEach(a => { totAnno[a] = cats.reduce((s,c)=>(s+(tableEnt[c][a]||0)-(tableUsc[c][a]||0)),0); });
    const totTot = anni.reduce((s,a)=>s+totAnno[a],0);
    h += `</tbody><tfoot><tr class="tot-row"><td>Totale</td>`;
    anni.forEach(a => { h += `<td class="${totAnno[a]>=0?'pos':'neg'}">${fmtEur(totAnno[a])}</td>`; });
    h += `<td class="col-tot ${totTot>=0?'pos':'neg'}">${fmtEur(totTot)}</td></tr></tfoot>`;
    return h;
  }

  function renderTabelle() {
    if (!allRows.length) return;
    populateTabAnnoSelect();

    const modo    = document.getElementById('tabModoSelect').value;
    const annoSel = document.getElementById('tabAnnoSelect').value;
    const pag     = document.getElementById('tabPaganteSelect').value;

    document.getElementById('tabAnnoWrap').style.display = modo==='anno' ? 'flex' : 'none';

    const filtered = allRows.filter(r => {
      if (modo==='anno' && r.anno !== annoSel) return false;
      if (pag && r.pagante !== pag) return false;
      return true;
    });

    if (modo === 'anno') {
      document.getElementById('pivotUscite').innerHTML  = buildPivotAnno(filtered, 'Uscite', annoSel);
      document.getElementById('pivotEntrate').innerHTML = buildPivotAnno(filtered, 'Entrate', annoSel);
    } else {
      document.getElementById('pivotUscite').innerHTML  = buildPivotTutti(filtered, 'Uscite');
      document.getElementById('pivotEntrate').innerHTML = buildPivotTutti(filtered, 'Entrate');
    }
  }

  // ── Export CSV ──
  document.getElementById('btnDownloadCsv').addEventListener('click', () => {
    const modo    = document.getElementById('tabModoSelect').value;
    const annoSel = document.getElementById('tabAnnoSelect').value;
    const pag     = document.getElementById('tabPaganteSelect').value;
    const activeTab = document.querySelector('.tab-btn.active')?.dataset.tab || 'uscite';

    const filtered = allRows.filter(r => {
      if (modo==='anno' && r.anno !== annoSel) return false;
      if (pag && r.pagante !== pag) return false;
      return true;
    });

    const tipo   = activeTab === 'uscite' ? 'Uscite' : 'Entrate';
    const rows   = filtered.filter(r => r.tipo === tipo);
    const cols   = modo==='anno' ? MESI_LABEL : [...new Set(rows.map(r=>r.anno).filter(Boolean))].sort();
    const cats   = [...new Set(rows.map(r=>r.categoria).filter(Boolean))].sort();
    const keyFn  = modo==='anno' ? (r) => parseInt(r.mese)-1 : (r) => cols.indexOf(r.anno);

    const table = {};
    cats.forEach(c => { table[c] = Array(cols.length).fill(0); });
    rows.forEach(r => { const idx=keyFn(r); if(idx>=0&&table[r.categoria]) table[r.categoria][idx]+=parseCosto(r.costo); });

    let csv = ['Categoria,' + cols.join(',') + ',Totale'];
    cats.forEach(c => {
      const vals = table[c];
      const tot  = vals.reduce((a,b)=>a+b,0);
      if (tot === 0) return;
      csv.push([c, ...vals.map(v=>v.toFixed(2)), tot.toFixed(2)].join(','));
    });
    const totRow = Array(cols.length).fill(0);
    cats.forEach(c => { table[c].forEach((v,i) => { totRow[i]+=v; }); });
    csv.push(['Totale', ...totRow.map(v=>v.toFixed(2)), totRow.reduce((a,b)=>a+b,0).toFixed(2)].join(','));

    const content = '\uFEFF' + csv.join('\n');
    const encoded = 'data:text/csv;charset=utf-8,' + encodeURIComponent(content);
    window.open(encoded, '_blank');
  });

  ['tabModoSelect','tabAnnoSelect','tabPaganteSelect'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderTabelle);
  });

  // Aggiorna navigazione — i dati sono già in memoria, solo render
  document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', () => {
      const s = item.dataset.section;
      if (s === 'elenco')   loadSheetData();
      if (s === 'annuale')  renderDashAnnuale();
      if (s === 'generale') renderDashGenerale();
      if (s === 'tabelle')  renderTabelle();
      if (s === 'utenze')   renderUtenze();
      if (s === 'mutuo')    { if (mAllRate.length) buildMutuo(mAllRate); }
    });
  });

  // ─── Sezione Utenze domestiche ────────────────────────────────
  const UTENZE = [
    { cat: 'Spese condominiali', color: '#5b9bff', label: 'Condominio' },
    { cat: 'Riscaldamento',      color: '#ff6b35', label: 'Riscaldamento' },
    { cat: 'Bolletta luce',      color: '#ffb400', label: 'Luce' },
    { cat: 'Bolletta gas',       color: '#2ecc71', label: 'Gas' },
    { cat: 'Telefonia',          color: '#a78bfa', label: 'Telefonia' },
    { cat: 'Bolletta rifiuti',   color: '#888',    label: 'Rifiuti' },
    { cat: 'Bolletta wifi',      color: '#06b6d4', label: 'Wi-Fi' },
  ];

  let uChartPeso = null, uChartArea = null, uChartStacked = null;


  function destroyUtenzeCharts() {
    [uChartPeso, uChartArea, uChartStacked].forEach(c => { if (c) c.destroy(); });
    uChartPeso = uChartArea = uChartStacked = null;
    ['uChartPeso','uChartArea','uChartStacked'].forEach(id => {
      const old = document.getElementById(id);
      if (old) { const nc = document.createElement('canvas'); nc.id = id; old.parentNode.replaceChild(nc, old); }
    });
  }

  function populateUAnnoSelect() {
    const sel  = document.getElementById('uAnnoSelect');
    const anni = [...new Set(allRows.map(r => r.anno).filter(Boolean))].sort((a,b) => b-a);
    const cur  = sel.value;
    sel.innerHTML = anni.map(a => `<option value="${a}">${a}</option>`).join('');
    if (cur && anni.includes(cur)) sel.value = cur;
    else { const y = String(new Date().getFullYear()); if (anni.includes(y)) sel.value = y; }
  }

  ['uAnnoSelect','uPaganteSelect'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderUtenze);
  });

  function renderUtenze() {
    if (!allRows.length) return;
    populateUAnnoSelect();
    destroyUtenzeCharts();

    const anno    = document.getElementById('uAnnoSelect').value;
    const pagante = document.getElementById('uPaganteSelect').value;
    const annoPrec = String(parseInt(anno) - 1);

    const filterFn = (a) => allRows.filter(r => {
      if (r.anno !== a) return false;
      if (pagante && r.pagante !== pagante) return false;
      if (r.tipo !== 'Uscite') return false;
      return true;
    });

    // Tutte le uscite filtrate per pagante (tutti gli anni) — per grafici storici
    const filteredRows = allRows.filter(r => {
      if (pagante && r.pagante !== pagante) return false;
      return true;
    });

    const rows     = filterFn(anno);
    const rowsPrec = filterFn(annoPrec);

    const isDark    = document.documentElement.getAttribute('data-theme') !== 'light';
    const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    const textColor = isDark ? '#888' : '#666';
    Chart.defaults.color = textColor; Chart.defaults.font.family = 'DM Sans'; Chart.defaults.font.size = 11;
    const ticksY = { color: textColor, callback: v => v.toLocaleString('it-IT') + ' €' };
    const ticksX = { color: textColor };

    // Calcola totali per utenza
    const totali = {}, totaliPrec = {}, mensili = {}, mensiliPrec = {};
    UTENZE.forEach(u => {
      totali[u.cat]      = rows.filter(r=>r.categoria===u.cat).reduce((s,r)=>s+parseCosto(r.costo),0);
      totaliPrec[u.cat]  = rowsPrec.filter(r=>r.categoria===u.cat).reduce((s,r)=>s+parseCosto(r.costo),0);
      mensili[u.cat]     = Array(12).fill(0);
      mensiliPrec[u.cat] = Array(12).fill(0);
      rows.filter(r=>r.categoria===u.cat).forEach(r => { const m=parseInt(r.mese); if(m>=1&&m<=12) mensili[u.cat][m-1]+=parseCosto(r.costo); });
      rowsPrec.filter(r=>r.categoria===u.cat).forEach(r => { const m=parseInt(r.mese); if(m>=1&&m<=12) mensiliPrec[u.cat][m-1]+=parseCosto(r.costo); });
    });

    // ── KPI cards ──
    const kpiGrid = document.getElementById('uKpiGrid');
    kpiGrid.innerHTML = '';
    UTENZE.forEach(u => {
      const tot  = totali[u.cat];
      const prec = totaliPrec[u.cat];
      const mesiAttivi = mensili[u.cat].filter(v=>v>0).length || 1;
      const media = tot / mesiAttivi;
      const delta = prec > 0 ? ((tot-prec)/prec*100) : null;
      const deltaHtml = delta === null
        ? `<span class="u-kpi-delta neu">— vs ${annoPrec}</span>`
        : `<span class="u-kpi-delta ${delta<=0?'pos':'neg'}">${delta>0?'+':''}${delta.toFixed(1)}% vs ${annoPrec}</span>`;
      kpiGrid.innerHTML += `
        <div class="u-kpi-card" style="--accent-color:${u.color}">
          <span class="u-kpi-label">${u.label}</span>
          <span class="u-kpi-value">${fmtEur(tot)}</span>
          <span class="u-kpi-sub">media ${fmtEur(media)}/mese</span>
          ${deltaHtml}
        </div>`;
    });

    // ── Grafico peso — donut ──
    const utenzeAttive = UTENZE.filter(u=>totali[u.cat]>0);
    uChartPeso = new Chart(document.getElementById('uChartPeso'), {
      type: 'doughnut',
      data: {
        labels: utenzeAttive.map(u=>u.label),
        datasets: [{ data: utenzeAttive.map(u=>totali[u.cat]), backgroundColor: utenzeAttive.map(u=>u.color), borderWidth: 0, hoverOffset: 8 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '58%',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed) } }
        },
        layout: { padding: 30 }
      },
      plugins: [{
        id: 'uDonutLabels',
        afterDraw(chart) {
          const { ctx, chartArea: { width, height, left, top } } = chart;
          const cx = left + width/2, cy = top + height/2;
          const meta = chart.getDatasetMeta(0);
          ctx.save();
          meta.data.forEach((arc, i) => {
            const angle = (arc.startAngle + arc.endAngle) / 2;
            const r = arc.outerRadius + 16;
            const x = cx + Math.cos(angle)*r, y = cy + Math.sin(angle)*r;
            const total = chart.data.datasets[0].data.reduce((a,b)=>a+b,0);
            const pct   = total > 0 ? (chart.data.datasets[0].data[i]/total*100).toFixed(0)+'%' : '';
            ctx.textAlign = x < cx ? 'right' : 'left';
            ctx.fillStyle = utenzeAttive[i].color;
            ctx.font = 'bold 11px DM Sans, sans-serif';
            ctx.fillText(chart.data.labels[i], x, y-5);
            ctx.fillStyle = isDark ? '#888' : '#666';
            ctx.font = '10px DM Sans, sans-serif';
            ctx.fillText(pct, x, y+7);
          });
          ctx.restore();
        }
      }]
    });

    // ── Pill + Grafico area aggregato per anno (una utenza alla volta) ──
    {
      const anniTutti = [...new Set(allRows.filter(r=>r.tipo==='Uscite').map(r=>r.anno).filter(Boolean))].sort();
      const utenzeDisp = UTENZE.filter(u => allRows.some(r=>r.tipo==='Uscite'&&r.categoria===u.cat));
      if (!window._uAreaActive || !utenzeDisp.find(u=>u.cat===window._uAreaActive)) {
        window._uAreaActive = utenzeDisp[0]?.cat;
      }

      // Pill
      const pillWrap = document.getElementById('uAreaPills');
      pillWrap.innerHTML = '';
      utenzeDisp.forEach(u => {
        const pill = document.createElement('div');
        pill.className = 'u-pill' + (u.cat===window._uAreaActive?' active':'');
        pill.textContent = u.label;
        if (u.cat===window._uAreaActive) pill.style.background = u.color;
        pill.addEventListener('click', () => {
          window._uAreaActive = u.cat;
          document.querySelectorAll('#uAreaPills .u-pill').forEach(p=>{ p.classList.remove('active'); p.style.background=''; });
          pill.classList.add('active'); pill.style.background = u.color;
          buildUAreaChart(anniTutti, u, filteredRows, gridColor, textColor, ticksY);
        });
        pillWrap.appendChild(pill);
      });

      buildUAreaChart(anniTutti, utenzeDisp.find(u=>u.cat===window._uAreaActive)||utenzeDisp[0], filteredRows, gridColor, textColor, ticksY);
    }

    // ── Heatmap con totali (celle vuote bianche) ──
    {
      const allVals = UTENZE.flatMap(u => mensili[u.cat]);
      const maxVal  = Math.max(...allVals, 1);
      const totRow  = Array(12).fill(0);

      let th = `<thead><tr>
        <th style="text-align:left">Utenza</th>
        ${MESI_LABEL.map(m=>`<th>${m}</th>`).join('')}
        <th class="col-tot">Totale</th>
        <th class="col-tot">Media</th>
      </tr></thead><tbody>`;

      UTENZE.forEach(u => {
        if (totali[u.cat]===0) return;
        const tot  = totali[u.cat];
        const mAtt = mensili[u.cat].filter(v=>v>0).length||1;
        mensili[u.cat].forEach((v,i) => { totRow[i]+=v; });

        const hex = u.color.replace('#','');
        const r   = parseInt(hex.slice(0,2),16);
        const g   = parseInt(hex.slice(2,4),16);
        const b   = parseInt(hex.slice(4,6),16);

        th += `<tr>`;
        th += `<td style="font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;color:var(--text-secondary);">${u.label}</td>`;
        mensili[u.cat].forEach(v => {
          if (v===0) {
            th += `<td style="text-align:right;font-family:'Space Mono',monospace;font-size:11px;background:transparent;"></td>`;
          } else {
            const intensity = v/maxVal;
            const alpha     = (0.1 + intensity*0.7).toFixed(2);
            const textCol   = intensity > 0.5 ? '#fff' : 'var(--text-primary)';
            th += `<td style="text-align:right;font-family:'Space Mono',monospace;font-size:11px;background:rgba(${r},${g},${b},${alpha});color:${textCol};">${v.toLocaleString('it-IT',{maximumFractionDigits:0})}</td>`;
          }
        });
        th += `<td class="col-tot" style="font-weight:700;">${fmtEur(tot)}</td>`;
        th += `<td class="col-tot" style="color:var(--text-secondary);">${fmtEur(tot/mAtt)}</td>`;
        th += `</tr>`;
      });

      const totTot = totRow.reduce((a,b)=>a+b,0);
      th += `</tbody><tfoot><tr class="tot-row"><td>Totale</td>`;
      totRow.forEach(v=>{ th+=`<td>${fmtEur(v)}</td>`; });
      th += `<td class="col-tot">${fmtEur(totTot)}</td><td class="col-tot">${fmtEur(totTot/(totRow.filter(v=>v>0).length||1))}</td>`;
      th += `</tr></tfoot>`;

      document.getElementById('uTabellaRiepilogo').innerHTML = th;
    }

    // ── Pill + barre mensili per categoria (tutti gli anni) ──
    {
      const anniTutti = [...new Set(allRows.filter(r=>r.tipo==='Uscite').map(r=>r.anno).filter(Boolean))].sort();
      const utenzeDisp = UTENZE.filter(u => filteredRows.some(r=>r.tipo==='Uscite'&&r.categoria===u.cat));
      if (!window._uStackedActive || !utenzeDisp.find(u=>u.cat===window._uStackedActive)) {
        window._uStackedActive = utenzeDisp[0]?.cat;
      }

      const pillWrap2 = document.getElementById('uStackedPills');
      pillWrap2.innerHTML = '';
      utenzeDisp.forEach(u => {
        const pill = document.createElement('div');
        pill.className = 'u-pill' + (u.cat===window._uStackedActive?' active':'');
        pill.textContent = u.label;
        if (u.cat===window._uStackedActive) pill.style.background = u.color;
        pill.addEventListener('click', () => {
          window._uStackedActive = u.cat;
          document.querySelectorAll('#uStackedPills .u-pill').forEach(p=>{ p.classList.remove('active'); p.style.background=''; });
          pill.classList.add('active'); pill.style.background = u.color;
          buildUMonthlyChart(anniTutti, utenzeDisp.find(u=>u.cat===window._uStackedActive), filteredRows, gridColor, textColor, ticksY);
        });
        pillWrap2.appendChild(pill);
      });

      buildUMonthlyChart(anniTutti, utenzeDisp.find(u=>u.cat===window._uStackedActive)||utenzeDisp[0], filteredRows, gridColor, textColor, ticksY);
    }
  }

  function buildUAreaChart(anniTutti, u, filteredRows, gridColor, textColor, ticksY) {
    if (uChartArea) uChartArea.destroy();
    const oldA = document.getElementById('uChartArea');
    if (oldA) { const nc = document.createElement('canvas'); nc.id='uChartArea'; oldA.parentNode.replaceChild(nc,oldA); }

    // Aggrega per anno usando solo le righe già filtrate per pagante
    const data = anniTutti.map(a =>
      filteredRows.filter(r=>r.tipo==='Uscite'&&r.categoria===u.cat&&r.anno===a)
                  .reduce((s,r)=>s+parseCosto(r.costo),0)
    );

    const hex = u.color.replace('#','');
    const ri=parseInt(hex.slice(0,2),16), gi=parseInt(hex.slice(2,4),16), bi=parseInt(hex.slice(4,6),16);

    uChartArea = new Chart(document.getElementById('uChartArea'), {
      type: 'line',
      data: {
        labels: anniTutti,
        datasets: [{
          label: u.label, data,
          borderColor: u.color,
          backgroundColor: `rgba(${ri},${gi},${bi},0.15)`,
          borderWidth: 2.5, fill: true, tension: 0.35, pointRadius: 5,
          pointBackgroundColor: u.color
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false },
          tooltip: { callbacks: { label: ctx => ' '+fmtEur(ctx.parsed.y) } } },
        scales: {
          x: { grid: { color: gridColor }, ticks: { color: textColor } },
          y: { grid: { color: gridColor }, ticks: ticksY }
        }
      }
    });
  }


  function buildUMonthlyChart(anniTutti, u, filteredRows, gridColor, textColor, ticksY) {
    if (uChartStacked) uChartStacked.destroy();
    const oldS = document.getElementById('uChartStacked');
    if (oldS) { const nc = document.createElement('canvas'); nc.id='uChartStacked'; oldS.parentNode.replaceChild(nc,oldS); }

    // Etichette: "Gen 2024", "Feb 2024", ... per tutti gli anni
    const labels = [];
    const data   = [];
    anniTutti.forEach(anno => {
      MESI_LABEL.forEach((ml, mi) => {
        labels.push(`${ml} ${anno}`);
        const v = filteredRows.filter(r=>r.tipo==='Uscite'&&r.categoria===u.cat&&r.anno===anno&&parseInt(r.mese)===mi+1)
                              .reduce((s,r)=>s+parseCosto(r.costo),0);
        data.push(v || null); // null per celle vuote (non disegna barra)
      });
    });

    const hex = u.color.replace('#','');
    const ri=parseInt(hex.slice(0,2),16), gi=parseInt(hex.slice(2,4),16), bi=parseInt(hex.slice(4,6),16);

    uChartStacked = new Chart(document.getElementById('uChartStacked'), {
      type: 'bar',
      data: {
        labels,
        datasets: [{
          label: u.label,
          data,
          backgroundColor: `rgba(${ri},${gi},${bi},0.75)`,
          borderRadius: 3,
          borderSkipped: false,
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ctx.parsed.y ? ' ' + fmtEur(ctx.parsed.y) : '' } }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              maxRotation: 45,
              callback(val, idx) {
                // Mostra solo le etichette di gennaio
                return labels[idx].startsWith('Gen') ? labels[idx].replace('Gen ','') : '';
              }
            }
          },
          y: { grid: { color: gridColor }, ticks: ticksY }
        }
      }
    });
  }

  // ─── Sezione Mutuo ────────────────────────────────────────────
  const MUTUO_SHEET = 'Mutuo';
  let mChartDonut = null, mChartCapInt = null;
  let mAllRate = [];
  let mTabAttivo = 'tutte';

  window.mSetTab = function(tab) {
    mTabAttivo = tab;
    document.querySelectorAll('#mTabTutte,#mTabPagate,#mTabResiduo').forEach(b => b.classList.remove('active'));
    document.getElementById('mTab' + tab.charAt(0).toUpperCase() + tab.slice(1)).classList.add('active');
    renderMutuoTabella(mAllRate);
  };

  function renderMutuo() {
    if (!mAllRate.length) return;
    buildMutuo(mAllRate);
  }

  function parseMutuoNum(s) {
    if (!s && s !== 0) return 0;
    if (typeof s === 'number') return s;
    const str = String(s).trim();
    // Formato italiano con virgola decimale: "34.712,62" → 34712.62
    if (str.includes(',')) {
      return parseFloat(str.replace(/\./g, '').replace(',', '.')) || 0;
    }
    // Numero con punto decimale anglosassone (come arriva dal JSON): "34712.62" → 34712.62
    return parseFloat(str) || 0;
  }

  function formatDataMutuo(d) {
    if (!d) return '';
    const s = String(d).trim();
    // ISO datetime: 2019-05-31T22:00:00.000Z
    const dt = new Date(s);
    if (!isNaN(dt)) {
      const dd = String(dt.getUTCDate()).padStart(2,'0');
      const mm = String(dt.getUTCMonth()+1).padStart(2,'0');
      const yy = dt.getUTCFullYear();
      return `${dd}/${mm}/${yy}`;
    }
    return s;
  }

  function buildMutuo(rows) {
    if (!rows.length) return;

    // Colonne: Nr, Data, Debito, Debito pagato, Percentuale, Rata, Interesse, Capitale, Pagata
    // Indici: 0    1     2       3               4            5     6           7         8

    const oggi = new Date();
    oggi.setHours(0,0,0,0);

    // Funzione per parsare la data dalla riga (può essere Date JS, stringa ISO, stringa italiana)
    function parseDataRata(val) {
      if (!val) return null;
      const dt = new Date(String(val).trim());
      if (isNaN(dt)) return null;
      // Usa UTC per evitare sfasamenti di fuso orario
      return new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()));
    }

    // Rate pagate = data <= oggi (indipendente dal campo Pagata)
    const ratePagate  = rows.filter(r => { const d = parseDataRata(r[1]); return d && d <= oggi; });
    const rateResidue = rows.filter(r => { const d = parseDataRata(r[1]); return !d || d > oggi; });

    // Debito iniziale = debito della prima rata
    const debitoIniziale = parseMutuoNum(rows[0]?.[2]);

    // Debito residuo = debito dell'ultima rata pagata (o iniziale se nessuna pagata)
    const ultimaPagata  = ratePagate[ratePagate.length - 1];
    const debitoResiduo = ultimaPagata ? parseMutuoNum(ultimaPagata[2]) : debitoIniziale;
    const debitoPagato  = debitoIniziale - debitoResiduo;
    const percPagata    = debitoIniziale > 0 ? (debitoPagato / debitoIniziale * 100) : 0;

    // Tempo rimanente
    const mesiRimanenti = rateResidue.length;
    const anni  = Math.floor(mesiRimanenti / 12);
    const mesi  = mesiRimanenti % 12;
    const tempoStr = anni > 0 ? `${anni}a ${mesi}m` : `${mesi} mesi`;

    // Prossima rata
    const prossima     = rateResidue[0];
    const prossimaData = prossima ? formatDataMutuo(prossima[1]) : '—';

    // Totale interessi/capitale pagati
    const totCapPagato = ratePagate.reduce((s,r)=>s+parseMutuoNum(r[7]),0);
    const totIntPagati = ratePagate.reduce((s,r)=>s+parseMutuoNum(r[6]),0);

    // ── KPI ──
    document.getElementById('mKpiTotale').textContent  = fmtEur(debitoIniziale);
    document.getElementById('mKpiPagato').textContent  = fmtEur(debitoPagato);
    document.getElementById('mKpiResiduo').textContent = fmtEur(debitoResiduo);
    document.getElementById('mKpiPerc').textContent    = percPagata.toFixed(1) + '%';
    document.getElementById('mKpiPercBar').style.width = percPagata.toFixed(1) + '%';
    document.getElementById('mKpiTempo').textContent   = tempoStr;
    document.getElementById('mKpiTempoSub').textContent = prossima ? `Prossima: ${prossimaData}` : 'Completato';

    const isDark    = document.documentElement.getAttribute('data-theme') !== 'light';
    const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    const textColor = isDark ? '#888' : '#666';
    Chart.defaults.color = textColor; Chart.defaults.font.family = 'DM Sans'; Chart.defaults.font.size = 11;

    // ── Donut stato rimborso ──
    if (mChartDonut) mChartDonut.destroy();
    const oldD = document.getElementById('mChartDonut');
    if (oldD) { const nc = document.createElement('canvas'); nc.id='mChartDonut'; oldD.parentNode.replaceChild(nc,oldD); }

    mChartDonut = new Chart(document.getElementById('mChartDonut'), {
      type: 'doughnut',
      data: {
        labels: ['Pagato', 'Residuo'],
        datasets: [{ data: [debitoPagato, debitoResiduo], backgroundColor: ['#2ecc71','#ff3b3b'], borderWidth: 0, hoverOffset: 6 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '62%',
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 12, padding: 16 } },
          tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed) } }
        }
      },
      plugins: [{
        id: 'mutuoCenter',
        afterDraw(chart) {
          const { ctx, chartArea: { width, height, left, top } } = chart;
          const cx = left+width/2, cy = top+height/2 - 16;
          ctx.save();
          ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
          ctx.fillStyle = '#2ecc71';
          ctx.font = 'bold 22px DM Sans, sans-serif';
          ctx.fillText(percPagata.toFixed(1)+'%', cx, cy);
          ctx.fillStyle = isDark ? '#888' : '#666';
          ctx.font = '12px DM Sans, sans-serif';
          ctx.fillText('rimborsato', cx, cy+22);
          ctx.restore();
        }
      }]
    });

    // ── Donut capitale vs interessi pagati ──
    if (mChartCapInt) mChartCapInt.destroy();
    const oldC = document.getElementById('mChartCapInt');
    if (oldC) { const nc = document.createElement('canvas'); nc.id='mChartCapInt'; oldC.parentNode.replaceChild(nc,oldC); }

    mChartCapInt = new Chart(document.getElementById('mChartCapInt'), {
      type: 'doughnut',
      data: {
        labels: ['Capitale', 'Interessi'],
        datasets: [{ data: [totCapPagato, totIntPagati], backgroundColor: ['#5b9bff','#ffb400'], borderWidth: 0, hoverOffset: 6 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '60%',
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 12, padding: 16 } },
          tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed) } }
        }
      },
      plugins: [{
        id: 'capIntCenter',
        afterDraw(chart) {
          const { ctx, chartArea: { width, height, left, top } } = chart;
          const cx = left+width/2, cy = top+height/2 - 10;
          const tot = totCapPagato + totIntPagati;
          ctx.save();
          ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
          ctx.fillStyle = isDark ? '#f0f0f0' : '#111';
          ctx.font = 'bold 15px DM Sans, sans-serif';
          ctx.fillText(fmtEur(tot), cx, cy);
          ctx.fillStyle = isDark ? '#888' : '#666';
          ctx.font = '11px DM Sans, sans-serif';
          ctx.fillText('totale versato', cx, cy+18);
          ctx.restore();
        }
      }]
    });

    // ── Tabella ──
    renderMutuoTabella(rows);
  }

  function renderMutuoTabella(rows) {
    const oggi = new Date(); oggi.setHours(0,0,0,0);
    function parseDataRata(val) {
      if (!val) return null;
      if (val instanceof Date) return val;
      const s = String(val).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return new Date(s);
      if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) { const [dd,mm,yy]=s.split('/'); return new Date(`${yy}-${mm}-${dd}`); }
      return new Date(s);
    }
    let filtered;
    if (mTabAttivo === 'pagate')  filtered = rows.filter(r => { const d=parseDataRata(r[1]); return d && d<=oggi; });
    else if (mTabAttivo === 'residue') filtered = rows.filter(r => { const d=parseDataRata(r[1]); return !d || d>oggi; });
    else filtered = rows;

    let h = `<thead><tr>
      <th style="text-align:left">Nr</th>
      <th style="text-align:left">Data</th>
      <th>Debito</th>
      <th>Debito pagato</th>
      <th>%</th>
      <th>Rata</th>
      <th>Interesse</th>
      <th>Capitale</th>
    </tr></thead><tbody>`;

    filtered.forEach(r => {
      const d = parseDataRata(r[1]);
      const pagata = d && d <= oggi;
      const rowStyle = pagata ? '' : 'opacity:0.55';
      h += `<tr style="${rowStyle}">
        <td style="font-family:'DM Sans',sans-serif;font-size:12px;">${r[0]||''}</td>
        <td style="font-family:'DM Sans',sans-serif;font-size:12px;">${formatDataMutuo(r[1])}</td>
        <td>${r[2]?fmtEur(parseMutuoNum(r[2])):''}</td>
        <td class="pos">${r[3]?fmtEur(parseMutuoNum(r[3])):''}</td>
        <td>${r[4]?parseFloat(String(r[4]).replace(',','.')).toFixed(1)+'%':''}</td>
        <td>${r[5]?fmtEur(parseMutuoNum(r[5])):''}</td>
        <td style="color:var(--accent-amber);">${r[6]?fmtEur(parseMutuoNum(r[6])):''}</td>
        <td class="pos">${r[7]?fmtEur(parseMutuoNum(r[7])):''}</td>
      </tr>`;
    });

    h += '</tbody>';
    document.getElementById('mTabella').innerHTML = h;
  }



  // ── Avvio app: carica tutto in una volta ──
  initApp();
