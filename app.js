// ─── Stato applicazione ───────────────────────────────────────
  const state = {
    sidebarOpen: false,
    theme: 'dark',
    activeSection: 'inserimento',
  };

  // ─── Backend Firestore (stessa interfaccia del vecchio Apps Script) ──
  // Ogni row porta il suo docId (sopravvive alla cache in sessionStorage);
  // update/delete risolvono rowIndex → docId tramite allRows.
  function _docIdFor(rowIndex) {
    const row = allRows.find(r => r.rowIndex === rowIndex);
    return row && row.docId;
  }

  function _rowArrayToDoc([data, costo, descrizione, categoria, tipo, mese, anno]) {
    return {
      data: String(data),
      costo: String(costo),
      descrizione: String(descrizione),
      categoria: String(categoria),
      tipo: String(tipo),
      mese: String(mese),
      anno: String(anno),
    };
  }

  // Dati per-utente: ogni account ha il suo spazio users/{uid}
  function _userBase() {
    const u = firebase.auth().currentUser;
    if (!u) throw new Error('Utente non autenticato');
    return 'users/' + u.uid;
  }

  async function asCall(params) {
    const db   = firebase.firestore();
    const base = _userBase();
    const col  = db.collection(base + '/spese');
    try {
      switch (params.action) {

        case 'read_all': {
          const [snap, muSnap, paSnap, riSnap] = await Promise.all([
            col.orderBy('ord').get(),
            db.doc(base + '/meta/mutuo').get(),
            db.doc(base + '/meta/patrimonio').get(),
            db.doc(base + '/meta/ricorrenti').get(),
          ]);
          const rows = snap.docs.map((d, i) => {
            const v = d.data();
            return {
              rowIndex: i + 1, docId: d.id,
              data: v.data, costo: v.costo, descrizione: v.descrizione,
              categoria: v.categoria, tipo: v.tipo, mese: v.mese, anno: v.anno,
            };
          });
          return {
            status: 'ok',
            rows,
            mutuo:      muSnap.exists ? JSON.parse(muSnap.data().json) : [],
            patrimonio: paSnap.exists ? JSON.parse(paSnap.data().json) : [],
            ricorrenti: riSnap.exists ? JSON.parse(riSnap.data().json) : [],
          };
        }

        case 'append': {
          const doc = _rowArrayToDoc(JSON.parse(params.row));
          doc.ord = Date.now();
          await col.add(doc);
          sessionStorage.removeItem('bm_data_cache');
          return { status: 'ok' };
        }

        case 'update': {
          const id = _docIdFor(params.rowIndex);
          if (!id) throw new Error('rowIndex non trovato');
          await col.doc(id).update(_rowArrayToDoc(JSON.parse(params.row)));
          sessionStorage.removeItem('bm_data_cache');
          return { status: 'ok' };
        }

        case 'delete': {
          const id = _docIdFor(params.rowIndex);
          if (!id) throw new Error('rowIndex non trovato');
          await col.doc(id).delete();
          sessionStorage.removeItem('bm_data_cache');
          return { status: 'ok' };
        }

        // Import CSV: molte righe in batch (max 500 op/batch)
        case 'append_many': {
          const rows = JSON.parse(params.rows);
          let ord = Date.now();
          for (let i = 0; i < rows.length; i += 450) {
            const batch = db.batch();
            rows.slice(i, i + 450).forEach(r => {
              const doc = _rowArrayToDoc(r);
              doc.ord = ord++;
              batch.set(col.doc(), doc);
            });
            await batch.commit();
          }
          sessionStorage.removeItem('bm_data_cache');
          return { status: 'ok' };
        }

        // Salva mutuo o patrimonio (doc meta/{nome}, campo json)
        case 'save_meta': {
          if (!['mutuo','patrimonio','ricorrenti'].includes(params.doc)) throw new Error('doc non valido');
          await db.doc(base + '/meta/' + params.doc).set({ json: params.json });
          sessionStorage.removeItem('bm_data_cache');
          return { status: 'ok' };
        }

        default:
          throw new Error('Azione sconosciuta: ' + params.action);
      }
    } catch (e) {
      return { status: 'error', message: e.message };
    }
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

  // Mobile: pulsante "Altro" della bottom-nav e backdrop
  document.getElementById('bnavMenu').addEventListener('click', () => toggleSidebar());
  document.getElementById('sidebarBackdrop').addEventListener('click', () => toggleSidebar(false));

  // ─── Navigazione sezioni (sidebar + bottom-nav) ───────────────
  const bnavItems = document.querySelectorAll('.bnav-item[data-section]');

  function showSection(target) {
    if (target === state.activeSection) return;

    navItems.forEach(n => n.classList.toggle('active', n.dataset.section === target));
    bnavItems.forEach(n => n.classList.toggle('active', n.dataset.section === target));
    sections.forEach(s => s.classList.remove('active'));

    document.getElementById('sec-' + target).classList.add('active');
    state.activeSection = target;
    renderActiveSection(target);

    // Mobile/tablet: chiudi la sidebar dopo la scelta
    if (window.matchMedia('(max-width: 1024px)').matches) toggleSidebar(false);
    window.scrollTo(0, 0);
  }

  navItems.forEach(item  => item.addEventListener('click', () => showSection(item.dataset.section)));
  bnavItems.forEach(item => item.addEventListener('click', () => showSection(item.dataset.section)));

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
      extra: ['Acquisti online','Altro','Bolletta gas','Bolletta luce','Bolletta rifiuti','Bolletta wifi','Casa','Regali','Riscaldamento','Spese condominiali','Spese mediche','Telefonia','Trasporti','Vestiti','Viaggi']
    },
    entrata: {
      principali: ['Stipendio','Affitto','Regali','Ripetizioni','Trasferimenti','Altro'],
      extra: []
    }
  };

  let currentType = 'uscita';
  let selectedCategory = '';

  const categoryGrid = document.getElementById('categoryGrid');
  const typeToggle   = document.getElementById('typeToggle');
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

  function showFeedback(msg, type) {
    formFeedback.textContent = msg;
    formFeedback.className = `form-feedback visible ${type}`;
    setTimeout(() => { formFeedback.className = 'form-feedback'; }, 4000);
  }

  function resetForm() {
    document.getElementById('fieldData').valueAsDate = new Date();
    document.getElementById('fieldCosto').value = '';
    document.getElementById('fieldDescrizione').value = '';
    selectedCategory = '';
    renderCategories();
  }

  btnReset.addEventListener('click', resetForm);

  btnSubmit.addEventListener('click', async () => {
    const data        = document.getElementById('fieldData').value;
    const costo       = document.getElementById('fieldCosto').value;
    const descrizione = document.getElementById('fieldDescrizione').value.trim();

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

    const row = [data, costoFmt, descrizione, selectedCategory, tipo, mese, anno];

    btnSubmit.textContent = 'Salvataggio…';
    btnSubmit.disabled = true;

    try {
      const json = await asCall({ action: 'append', row: JSON.stringify(row) });
      if (json.status !== 'ok') throw new Error(json.message);
      showFeedback('✓ Voce salvata correttamente.', 'success');
      resetForm();
    } catch(e) {
      showFeedback('✗ Errore durante il salvataggio. Riprova.', 'error');
    } finally {
      btnSubmit.innerHTML = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M2 7h10M8 3l4 4-4 4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg> Salva`;
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

  let patrimonioRows = [];
  let ricorrentiRows = [];
  const CACHE_KEY = 'bm_data_cache';
  const CACHE_TTL = 5 * 60 * 1000; // 5 minuti

  function saveCache(data) {
    try {
      const uid = firebase.auth().currentUser?.uid;
      sessionStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), uid, data }));
    } catch(e) {}
  }

  function loadCache() {
    try {
      const raw = sessionStorage.getItem(CACHE_KEY);
      if (!raw) return null;
      const { ts, uid, data } = JSON.parse(raw);
      if (Date.now() - ts > CACHE_TTL) { sessionStorage.removeItem(CACHE_KEY); return null; }
      if (uid !== firebase.auth().currentUser?.uid) { sessionStorage.removeItem(CACHE_KEY); return null; }
      return data;
    } catch(e) { return null; }
  }

  function applyData(json) {
    allRows        = json.rows       || [];
    mAllRate       = json.mutuo      || [];
    patrimonioRows = json.patrimonio || [];
    ricorrentiRows = json.ricorrenti || [];
    loadingBanner.remove();
    updateNotifiche();
    const s = state.activeSection;
    if (s === 'elenco')     { populateAnnoFilter(); applyFilters(); }
    if (s === 'annuale')    renderDashAnnuale();
    if (s === 'generale')   renderDashGenerale();
    if (s === 'tabelle')    renderTabelle();
    if (s === 'utenze')     renderUtenze();
    if (s === 'mutuo')      buildMutuo(mAllRate);
    if (s === 'cashflow')   renderCashflow();
    if (s === 'patrimonio') renderPatrimonio();
    if (s === 'calendario')  { renderCalendario(); }
  }

  async function initApp(forceRefresh = false) {
    try {
      // Usa cache se disponibile e non forzato refresh
      if (!forceRefresh) {
        const cached = loadCache();
        if (cached) { applyData(cached); return; }
      }

      const json = await asCall({ action: 'read_all' });
      if (json.status !== 'ok') throw new Error(json.message);

      saveCache(json);
      applyData(json);

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

    const filtered = allRows.filter(r => {
      if (fAnno && r.anno !== fAnno) return false;
      if (fMese && String(r.mese) !== fMese) return false;
      if (fTipo && r.tipo !== fTipo) return false;
      if (fCat  && r.categoria !== fCat) return false;
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

      const actionsHtml = `
          <div class="td-actions">
            <button class="btn-row edit" title="Modifica" onclick="startEdit(this, ${r.rowIndex})">
              <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M9 2l2 2-7 7H2V9L9 2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/></svg>
            </button>
            <button class="btn-row delete" title="Elimina" onclick="askDelete(${r.rowIndex}, '${escAttr(r.descrizione)}', '${escAttr(r.costo)}')">
              <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M2 3.5h9M5 3.5V2.5h3v1M4 3.5l.5 7h4l.5-7" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </button>
          </div>`;

      tr.innerHTML = `
        <td>${formatDataDisplay(r.data)}</td>
        <td>${escHtml(r.descrizione)}</td>
        <td><span class="tag tag-neutral">${escHtml(r.categoria)}</span></td>
        <td><span class="tag ${isUscita ? 'tag-red' : 'tag-green'}">${escHtml(r.tipo)}</span></td>
        <td class="td-costo ${isUscita ? 'uscita' : 'entrata'}">${escHtml(r.costo)} €</td>
        <td>${actionsHtml}</td>`;
      speseBody.appendChild(tr);

      // Riga dettaglio (visibile solo su mobile, espansa al tap sulla riga)
      const trD = document.createElement('tr');
      trD.className = 'm-detail';
      trD.innerHTML = `
        <td colspan="3">
          <div class="m-detail-row">
            <span class="m-detail-label">Descrizione</span>
            <span class="m-detail-val">${escHtml(r.descrizione)}</span>
          </div>
          <div class="m-detail-row">
            <span class="m-detail-label">Tipo</span>
            <span class="m-detail-val"><span class="tag ${isUscita ? 'tag-red' : 'tag-green'}">${escHtml(r.tipo)}</span></span>
            <span class="m-detail-acts">${actionsHtml}</span>
          </div>
        </td>`;
      speseBody.appendChild(trD);

      tr.addEventListener('click', (e) => {
        if (!window.matchMedia('(max-width: 768px)').matches) return;
        if (e.target.closest('button') || tr.classList.contains('editing')) return;
        tr.classList.toggle('m-open');
        trD.classList.toggle('m-visible');
      });
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

    // Mobile: chiudi la riga dettaglio mentre si modifica
    tr.classList.remove('m-open');
    const nx = tr.nextElementSibling;
    if (nx && nx.classList.contains('m-detail')) nx.classList.remove('m-visible');

    const allCat = [...CATEGORIE.uscita.principali, ...CATEGORIE.uscita.extra, ...CATEGORIE.entrata.principali].filter((v,i,a)=>a.indexOf(v)===i).sort();
    const catOpts = allCat.map(c => `<option value="${c}" ${c===row.categoria?'selected':''}>${c}</option>`).join('');
    const tipoOpts = ['Uscite','Entrate'].map(t => `<option value="${t}" ${t===row.tipo?'selected':''}>${t}</option>`).join('');

    tr.cells[0].innerHTML = `<input class="inline-input" id="ei-data-${rowIndex}" type="date" value="${row.data}" style="width:130px">`;
    tr.cells[1].innerHTML = `<input class="inline-input" id="ei-desc-${rowIndex}" type="text" value="${escHtml(row.descrizione)}" style="min-width:140px">`;
    tr.cells[2].innerHTML = `<select class="inline-select" id="ei-cat-${rowIndex}">${catOpts}</select>`;
    tr.cells[3].innerHTML = `<select class="inline-select" id="ei-tipo-${rowIndex}">${tipoOpts}</select>`;
    tr.cells[4].innerHTML = `<input class="inline-input" id="ei-costo-${rowIndex}" type="text" value="${row.costo}" style="width:80px;text-align:right">`;
    tr.cells[5].innerHTML = `
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

    const dateObj  = new Date(data);
    const mese     = dateObj.getMonth() + 1;
    const anno     = dateObj.getFullYear();
    const costoFmt = parseFloat(costoRaw).toLocaleString('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2});

    try {
      const json = await asCall({ action: 'update', rowIndex, row: JSON.stringify([data, costoFmt, desc, cat, tipo, mese, anno]) });
      if (json.status !== 'ok') throw new Error(json.message);
      const idx = allRows.findIndex(r => r.rowIndex === rowIndex);
      if (idx > -1) allRows[idx] = { ...allRows[idx], data, costo: costoFmt, descrizione: desc, categoria: cat, tipo, mese: String(mese), anno: String(anno) };
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
  ['fAnno','fMese','fTipo','fCategoria'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyFilters);
  });

  document.getElementById('btnClearFilters').addEventListener('click', () => {
    ['fAnno','fMese','fTipo','fCategoria'].forEach(id => {
      document.getElementById(id).value = '';
    });
    applyFilters();
  });

  document.getElementById('btnReload').addEventListener('click', () => {
    allRows = [];
    mAllRate = [];
    sessionStorage.removeItem(CACHE_KEY);
    initApp(true).then(() => { populateAnnoFilter(); applyFilters(); });
  });

  // ─── Dashboard Annuale ────────────────────────────────────────
  const MESI_LABEL    = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
  const MESI_COMPLETI = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];

  let chartMensile = null, chartConfronto = null, chartDonut = null, chartWaterfall = null;

  function destroyCharts() {
    [chartMensile, chartConfronto, chartDonut, chartWaterfall].forEach(c => { if (c) { c.destroy(); } });
    chartMensile = chartConfronto = chartDonut = chartWaterfall = null;
    ['chartMensile','chartConfronto','chartDonut','chartWaterfall'].forEach(id => {
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
    };
  }

  function populateAnnoSelect() {
    const sel  = document.getElementById('annoDashSelect');
    const anni = [...new Set(allRows.map(r => r.anno).filter(Boolean))].sort((a,b) => b - a);
    sel.innerHTML = anni.map(a => `<option value="${a}">${a}</option>`).join('');
    const cur = String(new Date().getFullYear());
    if (anni.includes(cur)) sel.value = cur;
  }

  function filterRows(rows, { anno, mese }) {
    return rows.filter(r => {
      if (anno && r.anno !== anno) return false;
      if (mese && String(r.mese) !== mese) return false;
      return true;
    });
  }

  function renderDashAnnuale() {
    if (!allRows.length) return;
    populateAnnoSelect();
    buildDash();
  }

  ['annoDashSelect','meseDashSelect'].forEach(id => {
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

      filterRows(allRows.filter(r => r.tipo === 'Uscite'), { anno: prevAnno, mese: String(prevMese) }).forEach(r => {
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
      filterRows(allRows.filter(r=>r.tipo==='Uscite'), { anno: prevAnno, mese: '' })
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

  ['gAnnoSelect'].forEach(id => {
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

    const fAnno = document.getElementById('gAnnoSelect').value;

    const filtered = allRows.filter(r => {
      if (fAnno && r.anno !== fAnno) return false;
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

    document.getElementById('tabAnnoWrap').style.display = modo==='anno' ? 'flex' : 'none';

    const filtered = allRows.filter(r => {
      if (modo==='anno' && r.anno !== annoSel) return false;
      return true;
    });

    if (modo === 'anno') {
      document.getElementById('pivotUscite').innerHTML  = buildPivotAnno(filtered, 'Uscite', annoSel);
      document.getElementById('pivotEntrate').innerHTML = buildPivotAnno(filtered, 'Entrate', annoSel);
    } else {
      document.getElementById('pivotUscite').innerHTML  = buildPivotTutti(filtered, 'Uscite');
      document.getElementById('pivotEntrate').innerHTML = buildPivotTutti(filtered, 'Entrate');
    }
    applyPivotState();
  }

  // Pivot collassata di default: visibile solo la riga Totale; toggle per il dettaglio
  let pivotExpanded = false;

  function applyPivotState() {
    ['pivotUscite','pivotEntrate'].forEach(id => {
      document.getElementById(id).classList.toggle('pv-collapsed', !pivotExpanded);
    });
    document.getElementById('btnPivotToggle').textContent =
      pivotExpanded ? 'Nascondi dettaglio ▴' : 'Mostra dettaglio ▾';
  }

  document.getElementById('btnPivotToggle').addEventListener('click', () => {
    pivotExpanded = !pivotExpanded;
    applyPivotState();
  });

  // Mutuo: su mobile le righe del piano si espandono al tap (delegazione)
  document.getElementById('mTabella').addEventListener('click', (e) => {
    if (!window.matchMedia('(max-width: 768px)').matches) return;
    const tr = e.target.closest('tr');
    if (!tr || tr.classList.contains('m-detail') || !tr.parentElement.matches('tbody')) return;
    const det = tr.nextElementSibling;
    if (det && det.classList.contains('m-detail')) {
      tr.classList.toggle('m-open');
      det.classList.toggle('m-visible');
    }
  });

  // ── Export CSV ──
  document.getElementById('btnDownloadCsv').addEventListener('click', () => {
    const modo    = document.getElementById('tabModoSelect').value;
    const annoSel = document.getElementById('tabAnnoSelect').value;
    const activeTab = document.querySelector('.tab-btn.active')?.dataset.tab || 'uscite';

    const filtered = allRows.filter(r => {
      if (modo==='anno' && r.anno !== annoSel) return false;
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

  ['tabModoSelect','tabAnnoSelect'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderTabelle);
  });

  // Al cambio sezione i dati sono già in memoria: solo render
  function renderActiveSection(s) {
    if (s === 'elenco')   loadSheetData();
    if (s === 'annuale')  renderDashAnnuale();
    if (s === 'generale') renderDashGenerale();
    if (s === 'tabelle')  renderTabelle();
    if (s === 'utenze')   renderUtenze();
    if (s === 'mutuo')    { if (mAllRate.length) buildMutuo(mAllRate); }
    if (s === 'cashflow')   renderCashflow();
    if (s === 'patrimonio') renderPatrimonio();
    if (s === 'calendario')  { renderCalendario(); }
  }

  // ─── Sezione Utenze domestiche ────────────────────────────────
  const UTENZE = [
    { cat: 'Spese condominiali', color: '#5b9bff', label: 'Condominio',
      icon: `<svg viewBox="0 0 18 18" fill="none"><rect x="4" y="2.5" width="10" height="13" rx="1" stroke="currentColor" stroke-width="1.4"/><line x1="7" y1="6" x2="8.2" y2="6" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><line x1="9.8" y1="6" x2="11" y2="6" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><line x1="7" y1="9" x2="8.2" y2="9" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><line x1="9.8" y1="9" x2="11" y2="9" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><path d="M8 15.5v-2.5h2v2.5" stroke="currentColor" stroke-width="1.4"/></svg>` },
    { cat: 'Riscaldamento',      color: '#ff6b35', label: 'Riscaldamento',
      icon: `<svg viewBox="0 0 18 18" fill="none"><rect x="2.5" y="5" width="13" height="8" rx="1.5" stroke="currentColor" stroke-width="1.4"/><line x1="5.5" y1="5" x2="5.5" y2="13" stroke="currentColor" stroke-width="1.4"/><line x1="9" y1="5" x2="9" y2="13" stroke="currentColor" stroke-width="1.4"/><line x1="12.5" y1="5" x2="12.5" y2="13" stroke="currentColor" stroke-width="1.4"/><line x1="4" y1="15.5" x2="4" y2="13" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><line x1="14" y1="15.5" x2="14" y2="13" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>` },
    { cat: 'Bolletta luce',      color: '#ffb400', label: 'Luce',
      icon: `<svg viewBox="0 0 18 18" fill="none"><path d="M9 2a5 5 0 0 1 2.8 9.14c-.5.35-.8.9-.8 1.51V13H7v-.35c0-.61-.3-1.16-.8-1.51A5 5 0 0 1 9 2z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><line x1="7.2" y1="15" x2="10.8" y2="15" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>` },
    { cat: 'Bolletta gas',       color: '#2ecc71', label: 'Gas',
      icon: `<svg viewBox="0 0 18 18" fill="none"><path d="M9 2s4.5 4 4.5 8a4.5 4.5 0 0 1-9 0C4.5 6 9 2 9 2z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><path d="M9 8s2 1.8 2 3.4a2 2 0 0 1-4 0C7 9.8 9 8 9 8z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/></svg>` },
    { cat: 'Telefonia',          color: '#a78bfa', label: 'Telefonia',
      icon: `<svg viewBox="0 0 18 18" fill="none"><rect x="5.5" y="2" width="7" height="14" rx="1.5" stroke="currentColor" stroke-width="1.4"/><line x1="8" y1="13.5" x2="10" y2="13.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>` },
    { cat: 'Bolletta rifiuti',   color: '#888',    label: 'Rifiuti', fillArea: 'rgba(200,200,200,0.22)',
      icon: `<svg viewBox="0 0 18 18" fill="none"><path d="M3.5 5h11M7 5V3.5h4V5M5 5l.7 10h6.6L13 5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/><line x1="7.5" y1="7.5" x2="7.8" y2="12.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><line x1="10.5" y1="7.5" x2="10.2" y2="12.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>` },
    { cat: 'Bolletta wifi',      color: '#06b6d4', label: 'Wi-Fi',
      icon: `<svg viewBox="0 0 18 18" fill="none"><path d="M2.5 7a9.2 9.2 0 0 1 13 0" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><path d="M5 9.8a5.6 5.6 0 0 1 8 0" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><path d="M7.4 12.4a2.3 2.3 0 0 1 3.2 0" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/><circle cx="9" cy="14.6" r="1" fill="currentColor"/></svg>` },
  ];

  // Pill "Tutte": mostra tutte le utenze insieme nei grafici
  const U_ALL = { cat: '__all__', label: 'Tutte', color: '#5b9bff' };

  function uRgba(u, alpha) {
    if (u.fillArea && alpha <= 0.3) return u.fillArea;
    const hex = u.color.replace('#','');
    const full = hex.length === 3 ? hex.split('').map(c=>c+c).join('') : hex;
    const r = parseInt(full.slice(0,2),16), g = parseInt(full.slice(2,4),16), b = parseInt(full.slice(4,6),16);
    return `rgba(${r},${g},${b},${alpha})`;
  }

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

  ['uAnnoSelect'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderUtenze);
  });

  function renderUtenze() {
    if (!allRows.length) return;
    populateUAnnoSelect();
    destroyUtenzeCharts();

    const anno    = document.getElementById('uAnnoSelect').value;
    const annoPrec = String(parseInt(anno) - 1);

    const filterFn = (a) => allRows.filter(r => {
      if (r.anno !== a) return false;
      if (r.tipo !== 'Uscite') return false;
      return true;
    });

    // Tutte le righe (tutti gli anni) — per grafici storici
    const filteredRows = allRows;

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
          <div class="u-kpi-head">
            <span class="u-kpi-icon">${u.icon}</span>
            <span class="u-kpi-label">${u.label}</span>
          </div>
          <span class="u-kpi-value">${fmtEur(tot)}</span>
          <div class="u-kpi-foot">
            <span class="u-kpi-sub">media ${fmtEur(media)}/mese</span>
            ${deltaHtml}
          </div>
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
      if (!window._uAreaActive || (window._uAreaActive !== U_ALL.cat && !utenzeDisp.find(u=>u.cat===window._uAreaActive))) {
        window._uAreaActive = U_ALL.cat;
      }

      // Pill ("Tutte" + una per utenza)
      const pillWrap = document.getElementById('uAreaPills');
      pillWrap.innerHTML = '';
      [U_ALL, ...utenzeDisp].forEach(u => {
        const pill = document.createElement('div');
        pill.className = 'u-pill' + (u.cat===window._uAreaActive?' active':'');
        pill.textContent = u.label;
        if (u.cat===window._uAreaActive) pill.style.background = u.color;
        pill.addEventListener('click', () => {
          window._uAreaActive = u.cat;
          document.querySelectorAll('#uAreaPills .u-pill').forEach(p=>{ p.classList.remove('active'); p.style.background=''; });
          pill.classList.add('active'); pill.style.background = u.color;
          buildUAreaChart(anniTutti, u.cat===U_ALL.cat ? null : u, filteredRows, gridColor, textColor, ticksY, utenzeDisp);
        });
        pillWrap.appendChild(pill);
      });

      const selA = window._uAreaActive===U_ALL.cat ? null : utenzeDisp.find(u=>u.cat===window._uAreaActive);
      buildUAreaChart(anniTutti, selA, filteredRows, gridColor, textColor, ticksY, utenzeDisp);
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

        th += `<tr>`;
        th += `<td style="font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;color:var(--text-secondary);">${u.label}</td>`;
        mensili[u.cat].forEach(v => {
          if (v===0) {
            th += `<td style="text-align:right;font-family:'Space Mono',monospace;font-size:11px;background:transparent;"></td>`;
          } else {
            const intensity = v/maxVal;
            const alpha     = +(0.1 + intensity*0.7).toFixed(2);
            const textCol   = intensity > 0.5 ? '#fff' : 'var(--text-primary)';
            th += `<td style="text-align:right;font-family:'Space Mono',monospace;font-size:11px;background:${uRgba(u, alpha)};color:${textCol};">${v.toLocaleString('it-IT',{maximumFractionDigits:0})}</td>`;
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
      if (!window._uStackedActive || (window._uStackedActive !== U_ALL.cat && !utenzeDisp.find(u=>u.cat===window._uStackedActive))) {
        window._uStackedActive = U_ALL.cat;
      }

      const pillWrap2 = document.getElementById('uStackedPills');
      pillWrap2.innerHTML = '';
      [U_ALL, ...utenzeDisp].forEach(u => {
        const pill = document.createElement('div');
        pill.className = 'u-pill' + (u.cat===window._uStackedActive?' active':'');
        pill.textContent = u.label;
        if (u.cat===window._uStackedActive) pill.style.background = u.color;
        pill.addEventListener('click', () => {
          window._uStackedActive = u.cat;
          document.querySelectorAll('#uStackedPills .u-pill').forEach(p=>{ p.classList.remove('active'); p.style.background=''; });
          pill.classList.add('active'); pill.style.background = u.color;
          buildUMonthlyChart(anniTutti, u.cat===U_ALL.cat ? null : u, filteredRows, gridColor, textColor, ticksY, utenzeDisp);
        });
        pillWrap2.appendChild(pill);
      });

      const selS = window._uStackedActive===U_ALL.cat ? null : utenzeDisp.find(u=>u.cat===window._uStackedActive);
      buildUMonthlyChart(anniTutti, selS, filteredRows, gridColor, textColor, ticksY, utenzeDisp);
    }
  }

  function buildUAreaChart(anniTutti, u, filteredRows, gridColor, textColor, ticksY, utenzeDisp) {
    if (uChartArea) uChartArea.destroy();
    const oldA = document.getElementById('uChartArea');
    if (oldA) { const nc = document.createElement('canvas'); nc.id='uChartArea'; oldA.parentNode.replaceChild(nc,oldA); }

    // u = null → tutte le utenze insieme
    const list = u ? [u] : (utenzeDisp || []);

    const datasets = list.map(x => ({
      label: x.label,
      data: anniTutti.map(a =>
        filteredRows.filter(r=>r.tipo==='Uscite'&&r.categoria===x.cat&&r.anno===a)
                    .reduce((s,r)=>s+parseCosto(r.costo),0)
      ),
      borderColor: x.color,
      backgroundColor: uRgba(x, 0.15),
      borderWidth: 2.5, fill: true, tension: 0.35,
      pointRadius: u ? 5 : 3,
      pointBackgroundColor: x.color
    }));

    uChartArea = new Chart(document.getElementById('uChartArea'), {
      type: 'line',
      data: { labels: anniTutti, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: !u, position: 'bottom', labels: { boxWidth: 12, padding: 12 } },
          tooltip: { callbacks: { label: ctx => ' ' + ctx.dataset.label + ': ' + fmtEur(ctx.parsed.y) } }
        },
        scales: {
          x: { grid: { color: gridColor }, ticks: { color: textColor } },
          y: { grid: { color: gridColor }, ticks: ticksY }
        }
      }
    });
  }


  function buildUMonthlyChart(anniTutti, u, filteredRows, gridColor, textColor, ticksY, utenzeDisp) {
    if (uChartStacked) uChartStacked.destroy();
    const oldS = document.getElementById('uChartStacked');
    if (oldS) { const nc = document.createElement('canvas'); nc.id='uChartStacked'; oldS.parentNode.replaceChild(nc,oldS); }

    // Etichette: "Gen 2024", "Feb 2024", ... per tutti gli anni
    const labels = [];
    anniTutti.forEach(anno => MESI_LABEL.forEach(ml => labels.push(`${ml} ${anno}`)));

    // u = null → tutte le utenze insieme (barre impilate)
    const list = u ? [u] : (utenzeDisp || []);

    const datasets = list.map(x => {
      const data = [];
      anniTutti.forEach(anno => {
        MESI_LABEL.forEach((ml, mi) => {
          const v = filteredRows.filter(r=>r.tipo==='Uscite'&&r.categoria===x.cat&&r.anno===anno&&parseInt(r.mese)===mi+1)
                                .reduce((s,r)=>s+parseCosto(r.costo),0);
          data.push(v || null); // null per celle vuote (non disegna barra)
        });
      });
      return {
        label: x.label,
        data,
        backgroundColor: uRgba(x, 0.75),
        borderRadius: 3,
        borderSkipped: false,
      };
    });

    uChartStacked = new Chart(document.getElementById('uChartStacked'), {
      type: 'bar',
      data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: !u, position: 'bottom', labels: { boxWidth: 12, padding: 12 } },
          tooltip: { callbacks: { label: ctx => ctx.parsed.y ? ' ' + ctx.dataset.label + ': ' + fmtEur(ctx.parsed.y) : '' } }
        },
        scales: {
          x: {
            stacked: !u,
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
          y: { stacked: !u, grid: { color: gridColor }, ticks: ticksY }
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
      // Riga dettaglio (solo mobile, espansa al tap)
      h += `<tr class="m-detail"><td colspan="3">
        <div class="m-detail-row"><span class="m-detail-label">Rata nr</span><span class="m-detail-val">${r[0]||''}</span></div>
        <div class="m-detail-row"><span class="m-detail-label">Debito pagato</span><span class="m-detail-val pos">${r[3]?fmtEur(parseMutuoNum(r[3])):'—'}</span></div>
        <div class="m-detail-row"><span class="m-detail-label">% pagata</span><span class="m-detail-val">${r[4]?parseFloat(String(r[4]).replace(',','.')).toFixed(1)+'%':'—'}</span></div>
        <div class="m-detail-row"><span class="m-detail-label">Interesse</span><span class="m-detail-val" style="color:var(--accent-amber);">${r[6]?fmtEur(parseMutuoNum(r[6])):'—'}</span></div>
        <div class="m-detail-row"><span class="m-detail-label">Capitale</span><span class="m-detail-val pos">${r[7]?fmtEur(parseMutuoNum(r[7])):'—'}</span></div>
      </td></tr>`;
    });

    h += '</tbody>';
    document.getElementById('mTabella').innerHTML = h;
  }

  // ─── Sezione Cashflow ─────────────────────────────────────────
  let cfChartArea = null, cfChartBarUsc = null, cfChartBarEnt = null;

  function renderCashflow() {
    if (!allRows.length) return;

    const cfAnnoEl   = document.getElementById('cfAnnoSelect');

    // Popola anni
    const anniDisp = [...new Set(allRows.map(r => r.anno).filter(Boolean))].sort((a,b) => b-a);
    if (cfAnnoEl.options.length === 0) {
      anniDisp.forEach(a => { const o = document.createElement('option'); o.value = a; o.textContent = a; cfAnnoEl.appendChild(o); });
    }

    const anno = cfAnnoEl.value || anniDisp[0];
    if (!cfAnnoEl.value) cfAnnoEl.value = anno;

    const fmtEur = v => v.toLocaleString('it-IT', { style:'currency', currency:'EUR', maximumFractionDigits:0 });
    const MESI   = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];

    // Righe dell'anno selezionato (tutte, per calcolo corretto)
    let rowsAnno = allRows.filter(r => r.anno === anno);

    // ── Aggregazione per mese ──
    const entratePerMese = Array(12).fill(0);
    const uscitePerMese  = Array(12).fill(0);
    rowsAnno.forEach(r => {
      const m = parseInt(r.mese) - 1;
      if (m < 0 || m > 11) return;
      const v = parseFloat(String(r.costo).replace(',','.')) || 0;
      if (r.tipo === 'Entrate') entratePerMese[m] += v;
      else uscitePerMese[m] += v;
    });

    // Mesi con dati reali
    const oggi = new Date();
    const meseCorrente = oggi.getFullYear().toString() === anno ? oggi.getMonth() : 11;
    const mesiConDati  = entratePerMese.filter((v,i) => v > 0 || uscitePerMese[i] > 0).length || 1;

    // ── Totali reali anno ──
    const totEntrate = entratePerMese.reduce((s,v) => s+v, 0);
    const totUscite  = uscitePerMese.reduce((s,v) => s+v, 0);

    // ── Media mensile (sui mesi con dati) ──
    const mediaEntrate = totEntrate / mesiConDati;
    const mediaUscite  = totUscite  / mesiConDati;
    const mediaSaldo   = mediaEntrate - mediaUscite;

    // ── Top 5 categorie uscite ──
    const catUsc = {};
    rowsAnno.filter(r => r.tipo !== 'Entrate').forEach(r => {
      const v = parseFloat(String(r.costo).replace(',','.')) || 0;
      catUsc[r.categoria] = (catUsc[r.categoria] || 0) + v;
    });
    const top5Usc = Object.entries(catUsc).sort((a,b) => b[1]-a[1]).slice(0,5);
    const sumTop5Usc  = top5Usc.reduce((s,[,v]) => s+v, 0);
    const meanTop5Usc = sumTop5Usc / mesiConDati;

    // ── Top 5 categorie entrate ──
    const catEnt = {};
    rowsAnno.filter(r => r.tipo === 'Entrate').forEach(r => {
      const v = parseFloat(String(r.costo).replace(',','.')) || 0;
      catEnt[r.categoria] = (catEnt[r.categoria] || 0) + v;
    });
    const top5Ent = Object.entries(catEnt).sort((a,b) => b[1]-a[1]).slice(0,5);

    // ── Previsione: mesi futuri × media mensile ──
    const mesiFuturi  = Math.max(0, 11 - meseCorrente);
    const prevEntrate = mediaEntrate * mesiFuturi;
    const prevUscite  = mediaUscite  * mesiFuturi;
    const proiezione  = mediaSaldo   * mesiFuturi;

    // ── KPI ──
    document.getElementById('cfKpiEntrate').textContent   = fmtEur(mediaEntrate);
    document.getElementById('cfKpiUscite').textContent    = fmtEur(mediaUscite);
    const saldoEl = document.getElementById('cfKpiSaldo');
    saldoEl.textContent = fmtEur(mediaSaldo);
    saldoEl.className   = 'kpi-value ' + (mediaSaldo >= 0 ? 'kpi-green' : 'kpi-red');
    const projEl = document.getElementById('cfKpiProiezione');
    projEl.textContent = fmtEur(proiezione);
    projEl.className   = 'kpi-value ' + (proiezione >= 0 ? 'kpi-green' : 'kpi-red');

    const isDark    = document.documentElement.getAttribute('data-theme') !== 'light';
    const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    const textColor = isDark ? '#888' : '#666';
    Chart.defaults.color      = textColor;
    Chart.defaults.font.family = 'DM Sans';
    Chart.defaults.font.size  = 11;

    // ── Grafico area storico + previsione ──
    const dataEntStorico = entratePerMese.map((v,i) => i <= meseCorrente ? v : null);
    const dataUscStorico  = uscitePerMese.map((v,i)  => i <= meseCorrente ? v : null);
    const dataEntPrev    = entratePerMese.map((v,i) => i >= meseCorrente ? (i === meseCorrente ? (v || mediaEntrate) : mediaEntrate) : null);
    const dataUscPrev    = uscitePerMese.map((v,i)  => i >= meseCorrente ? (i === meseCorrente ? (v || mediaUscite)  : mediaUscite)  : null);

    if (cfChartArea) cfChartArea.destroy();
    const cA = document.getElementById('cfChartArea');
    const nA = document.createElement('canvas'); nA.id = 'cfChartArea';
    cA.parentNode.replaceChild(nA, cA);
    cfChartArea = new Chart(nA, {
      type: 'line',
      data: { labels: MESI, datasets: [
        { label: 'Entrate',        data: dataEntStorico, borderColor:'#2ecc71', backgroundColor:'rgba(46,204,113,0.08)', borderWidth:2, pointRadius:3, fill:true, tension:0.3, spanGaps:false },
        { label: 'Uscite',         data: dataUscStorico, borderColor:'#ff3b3b', backgroundColor:'rgba(255,59,59,0.08)',  borderWidth:2, pointRadius:3, fill:true, tension:0.3, spanGaps:false },
        { label: 'Entrate (prev)', data: dataEntPrev,    borderColor:'#2ecc71', backgroundColor:'transparent', borderWidth:2, borderDash:[6,4], pointRadius:2, fill:false, tension:0.3, spanGaps:true },
        { label: 'Uscite (prev)',  data: dataUscPrev,    borderColor:'#ff3b3b', backgroundColor:'transparent', borderWidth:2, borderDash:[6,4], pointRadius:2, fill:false, tension:0.3, spanGaps:true },
      ]},
      options: { responsive:true, maintainAspectRatio:false,
        plugins: { legend:{ position:'bottom', labels:{ boxWidth:12, padding:16 }}, tooltip:{ callbacks:{ label: ctx => ' ' + fmtEur(ctx.parsed.y||0) }}},
        scales: { x:{ grid:{color:gridColor}}, y:{ grid:{color:gridColor}, ticks:{ callback: v => fmtEur(v) }}}
      }
    });

    // ── Barre top5 uscite ──
    if (cfChartBarUsc) cfChartBarUsc.destroy();
    const cBU = document.getElementById('cfChartBarUsc');
    const nBU = document.createElement('canvas'); nBU.id = 'cfChartBarUsc';
    cBU.parentNode.replaceChild(nBU, cBU);
    cfChartBarUsc = new Chart(nBU, {
      type:'bar',
      data:{ labels: top5Usc.map(([k])=>k), datasets:[
        { label:'Totale anno', data: top5Usc.map(([,v])=>v), backgroundColor:'rgba(255,59,59,0.45)', borderRadius:6, borderSkipped:false },
        { label:'Media/mese',  data: top5Usc.map(([,v])=>v/mesiConDati), backgroundColor:'rgba(255,59,59,0.8)', borderRadius:6, borderSkipped:false },
      ]},
      options:{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
        plugins:{ legend:{ position:'bottom', labels:{ boxWidth:10, padding:12 }}, tooltip:{ callbacks:{ label: ctx => ' ' + fmtEur(ctx.parsed.x) }}},
        scales:{ x:{ grid:{color:gridColor}, ticks:{ callback: v => fmtEur(v) }}, y:{ grid:{color:gridColor} }}
      }
    });

    // ── Barre top5 entrate ──
    if (cfChartBarEnt) cfChartBarEnt.destroy();
    const cBE = document.getElementById('cfChartBarEnt');
    const nBE = document.createElement('canvas'); nBE.id = 'cfChartBarEnt';
    cBE.parentNode.replaceChild(nBE, cBE);
    cfChartBarEnt = new Chart(nBE, {
      type:'bar',
      data:{ labels: top5Ent.map(([k])=>k), datasets:[
        { label:'Totale anno', data: top5Ent.map(([,v])=>v), backgroundColor:'rgba(46,204,113,0.45)', borderRadius:6, borderSkipped:false },
        { label:'Media/mese',  data: top5Ent.map(([,v])=>v/mesiConDati), backgroundColor:'rgba(46,204,113,0.8)', borderRadius:6, borderSkipped:false },
      ]},
      options:{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
        plugins:{ legend:{ position:'bottom', labels:{ boxWidth:10, padding:12 }}, tooltip:{ callbacks:{ label: ctx => ' ' + fmtEur(ctx.parsed.x) }}},
        scales:{ x:{ grid:{color:gridColor}, ticks:{ callback: v => fmtEur(v) }}, y:{ grid:{color:gridColor} }}
      }
    });

    // ── Tabella cashflow classica ──
    let html = `<thead><tr>
      <th style="text-align:left">Mese</th>
      <th style="text-align:right">Entrate</th>
      <th style="text-align:right">Uscite</th>
      <th style="text-align:right">Saldo mese</th>
      <th style="text-align:right">Saldo cumulato</th>
      <th style="text-align:right">Nota</th>
    </tr></thead><tbody>`;

    let saldoCumulato = 0;
    MESI.forEach((label, i) => {
      const futuro = i > meseCorrente;
      const e = futuro ? mediaEntrate : entratePerMese[i];
      const u = futuro ? mediaUscite  : uscitePerMese[i];
      const vuoto = !futuro && e === 0 && u === 0;
      const saldo = e - u;
      saldoCumulato += saldo;
      const opacity = futuro ? 'opacity:0.5;font-style:italic;' : '';
      const nota = futuro
        ? '<span style="font-size:11px;color:var(--accent-blue);">previsione</span>'
        : i === meseCorrente
          ? '<span style="font-size:11px;color:var(--accent-amber);">in corso</span>'
          : '';
      html += `<tr style="${opacity}">
        <td style="font-weight:600">${label}</td>
        <td style="text-align:right;color:var(--accent-green);font-family:'Space Mono',monospace;font-size:12px">${vuoto?'—':fmtEur(e)}</td>
        <td style="text-align:right;color:var(--accent-red);font-family:'Space Mono',monospace;font-size:12px">${vuoto?'—':fmtEur(u)}</td>
        <td style="text-align:right;font-family:'Space Mono',monospace;font-size:12px" class="${vuoto?'':saldo>=0?'pos':'neg'}">${vuoto?'—':fmtEur(saldo)}</td>
        <td style="text-align:right;font-family:'Space Mono',monospace;font-size:12px" class="${saldoCumulato>=0?'pos':'neg'}">${fmtEur(saldoCumulato)}</td>
        <td style="text-align:right">${nota}</td>
      </tr>`;
    });

    // Riga riepilogo
    const totPrevE = totEntrate + prevEntrate;
    const totPrevU = totUscite  + prevUscite;
    const totPrevS = totPrevE - totPrevU;
    html += `<tr class="tot-row">
      <td>Anno completo</td>
      <td style="text-align:right;color:var(--accent-green);font-family:'Space Mono',monospace;font-size:12px">${fmtEur(totPrevE)}</td>
      <td style="text-align:right;color:var(--accent-red);font-family:'Space Mono',monospace;font-size:12px">${fmtEur(totPrevU)}</td>
      <td style="text-align:right;font-family:'Space Mono',monospace;font-size:12px" class="${totPrevS>=0?'pos':'neg'}">${fmtEur(totPrevS)}</td>
      <td style="text-align:right;font-family:'Space Mono',monospace;font-size:12px" class="${totPrevS>=0?'pos':'neg'}">${fmtEur(totPrevS)}</td>
      <td></td>
    </tr>`;
    html += `<tr style="font-size:12px;color:var(--text-secondary)">
      <td colspan="6" style="padding-top:8px">
        Top 5 uscite: ${fmtEur(sumTop5Usc)} totale, ${fmtEur(meanTop5Usc)} media/mese
      </td>
    </tr>`;
    html += '</tbody>';
    document.getElementById('cfTabella').innerHTML = html;
  }

  ['cfAnnoSelect'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderCashflow);
  });


  // ─── Sezione Patrimonio ───────────────────────────────────────
  let patChartDonut = null;

  function renderPatrimonio() {
    const rows = patrimonioRows;
    const fmtEur = v => v.toLocaleString('it-IT', { style:'currency', currency:'EUR', maximumFractionDigits:0 });

    if (!rows.length) {
      ['patKpiAttivi','patKpiPassivi','patKpiNetto'].forEach(id => document.getElementById(id).textContent = '—');
      return;
    }

    const attivi  = rows.filter(r => r.tipologia.toLowerCase().includes('attiv'));
    const passivi = rows.filter(r => r.tipologia.toLowerCase().includes('passiv'));
    const totAttivi  = attivi.reduce((s,r)  => s + r.valore, 0);
    const totPassivi = passivi.reduce((s,r) => s + r.valore, 0);
    const netto      = totAttivi - totPassivi;

    document.getElementById('patKpiAttivi').textContent  = fmtEur(totAttivi);
    document.getElementById('patKpiPassivi').textContent = fmtEur(totPassivi);
    const nettoEl = document.getElementById('patKpiNetto');
    nettoEl.textContent = fmtEur(netto);
    nettoEl.className   = 'kpi-value ' + (netto >= 0 ? 'kpi-green' : 'kpi-red');

    const isDark    = document.documentElement.getAttribute('data-theme') !== 'light';
    const textColor = isDark ? '#888' : '#666';
    Chart.defaults.color       = textColor;
    Chart.defaults.font.family = 'DM Sans';
    Chart.defaults.font.size   = 11;

    const COLORS_ACT = ['#5b9bff','#2ecc71','#f39c12','#9b59b6','#1abc9c','#e67e22','#3498db','#27ae60'];
    const COLORS_PAS = ['#ff3b3b','#e74c3c','#c0392b','#ff6b6b'];

    const donutLabels = rows.map(r => r.asset);
    const donutData   = rows.map(r => r.valore);
    const donutColors = rows.map((r,i) => r.tipologia.toLowerCase().includes('attiv') ? COLORS_ACT[i % COLORS_ACT.length] : COLORS_PAS[i % COLORS_PAS.length]);
    const totAll      = donutData.reduce((s,v) => s+v, 0);

    if (patChartDonut) patChartDonut.destroy();
    const cD = document.getElementById('patChartDonut');
    const nD = document.createElement('canvas'); nD.id = 'patChartDonut';
    cD.parentNode.replaceChild(nD, cD);

    patChartDonut = new Chart(nD, {
      type: 'doughnut',
      data: { labels: donutLabels, datasets: [{ data: donutData, backgroundColor: donutColors, borderWidth: 2, borderColor: isDark ? '#1a1a1a' : '#fff' }]},
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '60%',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ' ' + fmtEur(ctx.parsed) + '  (' + (ctx.parsed/totAll*100).toFixed(1) + '%)' }}
        }
      },
      plugins: [{
        id: 'donutPct',
        afterDraw(chart) {
          const { ctx, data } = chart;
          const tot = data.datasets[0].data.reduce((s,v) => s+v, 0);
          chart.getDatasetMeta(0).data.forEach((arc, i) => {
            const val = data.datasets[0].data[i];
            const pct = (val/tot*100).toFixed(1);
            if (parseFloat(pct) < 4) return;
            const angle = (arc.startAngle + arc.endAngle) / 2;
            const r     = (arc.innerRadius + arc.outerRadius) / 2;
            const x     = arc.x + Math.cos(angle) * r;
            const y     = arc.y + Math.sin(angle) * r;
            ctx.save();
            ctx.fillStyle = '#fff'; ctx.font = 'bold 11px DM Sans';
            ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            ctx.shadowColor = 'rgba(0,0,0,0.5)'; ctx.shadowBlur = 3;
            ctx.fillText(pct + '%', x, y);
            ctx.restore();
          });
        }
      }]
    });

    // Legenda custom a lato (sostituisce quella nativa)
    const legEl = document.getElementById('patDonutLegend');
    if (legEl) {
      let legHtml = '<div style="display:flex;flex-direction:column;gap:8px;">';
      rows.forEach((r, i) => {
        const pct = (r.valore / totAll * 100).toFixed(1);
        const isAtt = r.tipologia.toLowerCase().includes('attiv');
        legHtml += `<div style="display:flex;align-items:center;gap:8px;font-size:12px;">
          <div style="width:10px;height:10px;border-radius:50%;background:${donutColors[i]};flex-shrink:0;"></div>
          <span style="color:var(--text-secondary);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r.asset}">${r.asset}</span>
          <span style="font-family:'Space Mono',monospace;font-size:11px;color:${isAtt?'var(--accent-green)':'var(--accent-red)'};white-space:nowrap;">${pct}%</span>
          <span style="font-family:'Space Mono',monospace;font-size:11px;color:var(--text-secondary);white-space:nowrap;">${fmtEur(r.valore)}</span>
        </div>`;
      });
      legHtml += '</div>';
      legEl.innerHTML = legHtml;
    }

    // Tabelle
    const buildTable = (arr, colorClass) => {
      const tot = arr.reduce((s,r) => s+r.valore, 0);
      let h = `<thead><tr><th style="text-align:left">Asset</th><th style="text-align:right">Valore</th><th style="text-align:right">%</th></tr></thead><tbody>`;
      arr.sort((a,b) => b.valore-a.valore).forEach(r => {
        const pct = tot > 0 ? (r.valore/tot*100).toFixed(1) : '0.0';
        h += `<tr>
          <td style="font-weight:600">${r.asset}</td>
          <td style="text-align:right;font-family:'Space Mono',monospace;font-size:12px" class="${colorClass}">${fmtEur(r.valore)}</td>
          <td style="text-align:right;color:var(--text-secondary);font-size:12px">${pct}%</td>
        </tr>`;
      });
      h += `<tr class="tot-row"><td>Totale</td><td style="text-align:right;font-family:'Space Mono',monospace;font-size:12px" class="${colorClass}">${fmtEur(tot)}</td><td></td></tr></tbody>`;
      return h;
    };

    document.getElementById('patTabellaAttivi').innerHTML  = buildTable(attivi,  'pos');
    document.getElementById('patTabellaPassivi').innerHTML = buildTable(passivi, 'neg');
  }


  // ─── Sezione Calendario ───────────────────────────────────────
  let calYear  = new Date().getFullYear();
  let calMonth = new Date().getMonth(); // 0-indexed
  let calChartLine = null;
  let calMinMs = 0, calMaxMs = 0;
  let calRangeStart = null, calRangeEnd = null;

  function renderCalendario() {
    if (!allRows.length) return;
    renderCalKpi();
    renderCalGrid();
    renderProssime();
    renderHeatmap();
    initCalSlider();
    renderCalChart();
  }

  function renderCalGrid() {
    const fmtEur = v => v.toLocaleString('it-IT', { style:'currency', currency:'EUR', maximumFractionDigits:0 });

    // Aggiorna header mese
    const mesi = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
    document.getElementById('calMonthLabel').textContent = mesi[calMonth] + ' ' + calYear;

    // Spese del mese
    const speseDelMese = {};
    allRows.forEach(r => {
      if (!r.data) return;
      const d = new Date(r.data);
      if (d.getFullYear() === calYear && d.getMonth() === calMonth) {
        const day = d.getDate();
        if (!speseDelMese[day]) speseDelMese[day] = [];
        speseDelMese[day].push(r);
      }
    });

    // Costruisci griglia
    const firstDay = new Date(calYear, calMonth, 1).getDay(); // 0=dom
    const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate();
    const startOffset = (firstDay === 0) ? 6 : firstDay - 1; // lun=0

    let html = '';
    // Intestazioni giorni
    ['Lun','Mar','Mer','Gio','Ven','Sab','Dom'].forEach(g => {
      html += `<div style="text-align:center;font-size:11px;font-weight:600;color:var(--text-hint);padding:4px 0;">${g}</div>`;
    });

    // Celle vuote iniziali
    for (let i = 0; i < startOffset; i++) html += '<div></div>';

    // Giorni
    const oggi = new Date();
    for (let d = 1; d <= daysInMonth; d++) {
      const spese = speseDelMese[d] || [];
      const totUsc = spese.filter(r => r.tipo === 'Uscite').reduce((s,r) => s + (parseFloat(String(r.costo).replace(',','.')) || 0), 0);
      const totEnt = spese.filter(r => r.tipo === 'Entrate').reduce((s,r) => s + (parseFloat(String(r.costo).replace(',','.')) || 0), 0);
      const isOggi = oggi.getDate()===d && oggi.getMonth()===calMonth && oggi.getFullYear()===calYear;

      let dotHtml = '';
      if (totEnt > 0) dotHtml += `<span style="font-size:9px;color:var(--accent-green)">▲${fmtEur(totEnt)}</span> `;
      if (totUsc > 0) dotHtml += `<span style="font-size:9px;color:var(--accent-red)">▼${fmtEur(totUsc)}</span>`;

      html += `<div class="cal-day${isOggi?' cal-today':''}" onclick="calDayClick(${d})" style="cursor:pointer">
        <div style="font-size:12px;font-weight:${isOggi?'700':'500'};margin-bottom:2px">${d}</div>
        <div style="line-height:1.3">${dotHtml}</div>
      </div>`;
    }

    document.getElementById('calGrid').innerHTML = html;
  }

  window.calDayClick = function(day) {
    const date = `${calYear}-${String(calMonth+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    const spese = allRows.filter(r => r.data === date);
    const giorniSett = ['Domenica','Lunedì','Martedì','Mercoledì','Giovedì','Venerdì','Sabato'];
    const dObj = new Date(date);

    let html = `<div style="margin-bottom:14px;">
      <div style="font-weight:600;font-size:15px;">${giorniSett[dObj.getDay()]} ${day}/${calMonth+1}/${calYear}</div>
      <div style="font-size:12px;color:var(--text-hint);margin-top:2px;">${spese.length ? spese.length + (spese.length === 1 ? ' operazione — tocca per modificare' : ' operazioni — tocca per modificare') : ''}</div>
    </div>`;

    if (spese.length) {
      const totUsc = spese.filter(r => r.tipo === 'Uscite').reduce((s,r) => s + parseCosto(r.costo), 0);
      const totEnt = spese.filter(r => r.tipo === 'Entrate').reduce((s,r) => s + parseCosto(r.costo), 0);

      html += `<div style="display:flex;flex-direction:column;gap:6px;margin-bottom:12px">`;
      spese.forEach(r => {
        const v = parseCosto(r.costo);
        const isUsc = r.tipo !== 'Entrate';
        html += `<div class="calp-row" onclick="calEditSpesa(${r.rowIndex})" title="Modifica">
          <div style="min-width:0;flex:1;">
            <div class="calp-desc">${escHtml(r.descrizione || r.categoria)}</div>
            <div class="calp-cat">${escHtml(r.categoria)}</div>
          </div>
          <span class="calp-val ${isUsc ? 'neg' : 'pos'}">${isUsc ? '−' : '+'}${v.toLocaleString('it-IT',{style:'currency',currency:'EUR'})}</span>
          <svg class="calp-pen" width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M9 2l2 2-7 7H2V9L9 2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/></svg>
        </div>`;
      });
      html += '</div>';

      html += `<div style="display:flex;justify-content:space-between;font-size:12px;color:var(--text-secondary);border-top:1px solid var(--border-idle);padding-top:10px;margin-bottom:14px;">
        <span>${totEnt > 0 ? 'Entrate <b style="color:var(--accent-green)">' + totEnt.toLocaleString('it-IT',{style:'currency',currency:'EUR'}) + '</b>' : ''}</span>
        <span>${totUsc > 0 ? 'Uscite <b style="color:var(--accent-red)">' + totUsc.toLocaleString('it-IT',{style:'currency',currency:'EUR'}) + '</b>' : ''}</span>
      </div>`;
    } else {
      html += `<div style="color:var(--text-secondary);font-size:13px;margin-bottom:14px">Nessuna spesa registrata in questo giorno.</div>`;
    }

    html += `<button class="btn-submit" style="width:100%;justify-content:center;" onclick="calOpenInsert('${date}')">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 2v10M2 7h10" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>
      Nuova spesa
    </button>`;
    document.getElementById('calPopupBody').innerHTML = html;

    const popup = document.getElementById('calPopup');
    popup.style.display = 'flex';
  };

  window.calOpenInsert = function(date) {
    document.getElementById('calPopup').style.display = 'none';
    // Apri il form inserimento con data precompilata
    showSection('inserimento');
    const inp = document.getElementById('fieldData');
    if (inp) inp.value = date;
  };

  // Click su una voce del popup → vai all'elenco e apri la modifica inline
  window.calEditSpesa = function(rowIndex) {
    document.getElementById('calPopup').style.display = 'none';
    showSection('elenco');
    // Azzera i filtri così la voce è sicuramente visibile
    ['fAnno','fMese','fTipo','fCategoria'].forEach(id => { document.getElementById(id).value = ''; });
    applyFilters();
    const tr = document.querySelector(`tr[data-row-index="${rowIndex}"]`);
    if (tr) {
      tr.scrollIntoView({ block: 'center', behavior: 'smooth' });
      window.startEdit(null, rowIndex);
    }
  };

  function renderCalChart() {
    if (!allRows.length) return;

    const fmtEur    = v => v.toLocaleString('it-IT', { style:'currency', currency:'EUR', maximumFractionDigits:0 });
    const isDark    = document.documentElement.getAttribute('data-theme') !== 'light';
    const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
    Chart.defaults.color = isDark ? '#888' : '#666';
    Chart.defaults.font.family = 'DM Sans';

    const rangeStart = calRangeStart;
    const rangeEnd   = calRangeEnd;
    if (!rangeStart || !rangeEnd) return;

    const MESI_SHORT = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
    const labels = [], dataEnt = [], dataUsc = [];
    const cur = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), 1);
    const end = new Date(rangeEnd.getFullYear(),   rangeEnd.getMonth(),   1);
    while (cur <= end) {
      const y = cur.getFullYear(), m = cur.getMonth();
      labels.push(`${MESI_SHORT[m]} ${y}`);
      const mRows = allRows.filter(r => {
        if (!r.data) return false;
        const d = new Date(r.data);
        return d.getFullYear() === y && d.getMonth() === m;
      });
      dataEnt.push(mRows.filter(r => r.tipo === 'Entrate').reduce((s,r) => s + (parseFloat(String(r.costo).replace(',','.')) || 0), 0));
      dataUsc.push(mRows.filter(r => r.tipo === 'Uscite').reduce((s,r)  => s + (parseFloat(String(r.costo).replace(',','.')) || 0), 0));
      cur.setMonth(cur.getMonth() + 1);
    }

    if (calChartLine) calChartLine.destroy();
    const cL = document.getElementById('calChartLine');
    const nL = document.createElement('canvas'); nL.id = 'calChartLine';
    cL.parentNode.replaceChild(nL, cL);

    calChartLine = new Chart(nL, {
      type: 'line',
      data: { labels, datasets: [
        { label:'Entrate', data:dataEnt, borderColor:'#2ecc71', backgroundColor:'rgba(46,204,113,0.1)', borderWidth:2, pointRadius:3, fill:true, tension:0.3 },
        { label:'Uscite',  data:dataUsc, borderColor:'#ff3b3b', backgroundColor:'rgba(255,59,59,0.1)',  borderWidth:2, pointRadius:3, fill:true, tension:0.3 },
      ]},
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins: { legend:{ position:'bottom', labels:{ boxWidth:12, padding:16 }},
          tooltip:{ callbacks:{ label: ctx => ' ' + fmtEur(ctx.parsed.y) }}},
        scales: { x:{ grid:{color:gridColor}, ticks:{ maxTicksLimit:12 }},
          y:{ grid:{color:gridColor}, ticks:{ callback: v => fmtEur(v) }}}
      }
    });

    // Aggiorna label range
    const fmt = d => `${MESI_SHORT[d.getMonth()]} ${d.getFullYear()}`;
    document.getElementById('calRangeLabel').textContent = fmt(rangeStart) + ' — ' + fmt(rangeEnd);
  }

  function renderHeatmap() {
    const container = document.getElementById('calHeatmap');
    if (!container || !allRows.length) return;

    // Aggrega spese totali per giorno
    const spesaPerGiorno = {};
    allRows.filter(r => r.tipo === 'Uscite' && r.data).forEach(r => {
      const v = parseFloat(String(r.costo).replace(',','.')) || 0;
      spesaPerGiorno[r.data] = (spesaPerGiorno[r.data] || 0) + v;
    });

    const valori = Object.values(spesaPerGiorno);
    const maxVal = valori.length ? Math.max(...valori) : 1;

    // Genera ultimi 52 settimane (364 giorni) come GitHub
    const oggi = new Date();
    oggi.setHours(0,0,0,0);
    const startDate = new Date(oggi);
    startDate.setDate(startDate.getDate() - 363);
    // Allinea a lunedì
    while (startDate.getDay() !== 1) startDate.setDate(startDate.getDate() - 1);

    const fmtDate = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    const fmtEur  = v => v.toLocaleString('it-IT', { style:'currency', currency:'EUR', maximumFractionDigits:0 });

    // Settimane come colonne, giorni (lun-dom) come righe
    const weeks = [];
    const cur = new Date(startDate);
    while (cur <= oggi) {
      const week = [];
      for (let d = 0; d < 7; d++) {
        const key = fmtDate(cur);
        const val = spesaPerGiorno[key] || 0;
        week.push({ date: key, val, isFuture: cur > oggi });
        cur.setDate(cur.getDate() + 1);
      }
      weeks.push(week);
    }

    // Intensità colore: 0=trasparente, max=rosso pieno
    const getColor = val => {
      if (!val) return 'var(--bg3)';
      const t = Math.min(val / maxVal, 1);
      // da verde chiaro a rosso intenso
      const r = Math.round(80  + t * 175);
      const g = Math.round(180 - t * 150);
      const b = Math.round(80  - t * 50);
      return `rgb(${r},${g},${b})`;
    };

    const MESI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
    const GIORNI = ['L','M','M','G','V','S','D'];

    // ── Dimensione celle adattiva: tutto deve stare nella larghezza disponibile ──
    // Larghezza utile = container − colonna label giorni (~20px)
    const availW = Math.max(120, (container.clientWidth || container.parentElement.clientWidth || 320) - 24);
    let gap   = 3;
    let cellW = Math.floor((availW - gap * (weeks.length - 1)) / weeks.length);
    let shownWeeks = weeks;
    if (cellW < 6) {
      // Schermo stretto: celle minime 6px, mostra solo le ultime settimane che ci stanno
      gap = 2;
      cellW = 6;
      const fit = Math.max(4, Math.floor((availW + gap) / (cellW + gap)));
      shownWeeks = weeks.slice(-fit);
    } else if (cellW > 14) {
      cellW = 14;
    }

    // Label mesi sopra
    let monthLabels = `<div style="display:flex;gap:${gap}px;margin-bottom:3px;padding-left:20px;">`;
    let lastMonth = -1;
    shownWeeks.forEach((week, wi) => {
      const m = new Date(week[0].date).getMonth();
      if (m !== lastMonth) {
        monthLabels += `<div style="width:${cellW}px;font-size:9px;color:var(--text-hint);white-space:nowrap;overflow:visible;">${MESI[m]}</div>`;
        lastMonth = m;
      } else {
        monthLabels += `<div style="width:${cellW}px;"></div>`;
      }
    });
    monthLabels += '</div>';

    // Griglia: giorni sull'asse Y, settimane su X
    let grid = `<div style="display:flex;gap:${gap}px;">`;
    // Label giorni
    grid += `<div style="display:flex;flex-direction:column;gap:${gap}px;margin-right:4px;">`;
    GIORNI.forEach(g => grid += `<div style="height:${cellW}px;font-size:9px;color:var(--text-hint);line-height:${cellW}px;">${g}</div>`);
    grid += '</div>';

    shownWeeks.forEach(week => {
      grid += `<div style="display:flex;flex-direction:column;gap:${gap}px;">`;
      week.forEach(({ date, val, isFuture }) => {
        const color = isFuture ? 'transparent' : getColor(val);
        const tip   = val > 0 ? `${date}: ${fmtEur(val)}` : date;
        grid += `<div title="${tip}" style="width:${cellW}px;height:${cellW}px;border-radius:2px;background:${color};cursor:${val>0?'pointer':'default'}"
          ${val > 0 ? `onclick="calDayClick(${new Date(date).getDate()}); calYear=${new Date(date).getFullYear()}; calMonth=${new Date(date).getMonth()};"` : ''}></div>`;
      });
      grid += '</div>';
    });
    grid += '</div>';

    // Legenda
    const legend = `<div style="display:flex;align-items:center;gap:6px;margin-top:10px;font-size:11px;color:var(--text-secondary);">
      <span>Meno</span>
      ${[0, 0.25, 0.5, 0.75, 1].map(t => `<div style="width:14px;height:14px;border-radius:3px;background:${t===0?'var(--bg3)':`rgb(${Math.round(80+t*175)},${Math.round(180-t*150)},${Math.round(80-t*50)})`}"></div>`).join('')}
      <span>Di più</span>
    </div>`;

    container.innerHTML = monthLabels + grid + legend;
  }

  // Ridisegna la heatmap quando cambia la larghezza (rotazione, resize)
  let _hmResizeT;
  window.addEventListener('resize', () => {
    clearTimeout(_hmResizeT);
    _hmResizeT = setTimeout(() => {
      if (state.activeSection === 'calendario') renderHeatmap();
    }, 200);
  });

  // Dual-handle slider
  function initCalSlider() {
    const thumbL = document.getElementById('calThumbL');
    const thumbR = document.getElementById('calThumbR');
    const track  = document.getElementById('calTrackFill');
    const wrap   = document.getElementById('calSliderWrap');
    if (!thumbL || thumbL.dataset.init) return;
    thumbL.dataset.init = '1';

    const allDates = allRows.map(r => r.data ? new Date(r.data).getTime() : null).filter(Boolean);
    if (!allDates.length) return;

    calMinMs = Math.min(...allDates);
    calMaxMs = Date.now();

    let posL = 0, posR = 1; // 0..1

    function msToPos(ms) { return (ms - calMinMs) / (calMaxMs - calMinMs); }
    function posToMs(p)   { return calMinMs + p * (calMaxMs - calMinMs); }
    function posToDate(p) { const ms = posToMs(p); const d = new Date(ms); d.setDate(1); return d; }

    function updateFromPos() {
      calRangeStart = posToDate(posL);
      calRangeEnd   = posToDate(posR);
      const pctL = (posL * 100).toFixed(1) + '%';
      const pctR = (posR * 100).toFixed(1) + '%';
      thumbL.style.left = pctL;
      thumbR.style.left = pctR;
      track.style.left  = pctL;
      track.style.width = ((posR - posL) * 100).toFixed(1) + '%';
      renderCalChart();
    }

    // Imposta defaults: tutto il range
    posL = 0; posR = 1;
    updateFromPos();

    function makeDraggable(thumb, isLeft) {
      thumb.addEventListener('mousedown', e => {
        e.preventDefault();
        const rect = wrap.getBoundingClientRect();
        function onMove(ev) {
          const p = Math.max(0, Math.min(1, (ev.clientX - rect.left) / rect.width));
          if (isLeft)  { posL = Math.min(p, posR - 0.01); }
          else         { posR = Math.max(p, posL + 0.01); }
          updateFromPos();
        }
        function onUp() { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); }
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
      // Touch
      thumb.addEventListener('touchstart', e => {
        e.preventDefault();
        const rect = wrap.getBoundingClientRect();
        function onMove(ev) {
          const touch = ev.touches[0];
          const p = Math.max(0, Math.min(1, (touch.clientX - rect.left) / rect.width));
          if (isLeft)  { posL = Math.min(p, posR - 0.01); }
          else         { posR = Math.max(p, posL + 0.01); }
          updateFromPos();
        }
        function onUp() { thumb.removeEventListener('touchmove', onMove); thumb.removeEventListener('touchend', onUp); }
        thumb.addEventListener('touchmove', onMove, { passive: false });
        thumb.addEventListener('touchend', onUp);
      });
    }

    makeDraggable(thumbL, true);
    makeDraggable(thumbR, false);
  }

  // ── Nav mese calendario ──
  document.getElementById('calPrevMonth').addEventListener('click', () => {
    calMonth--; if (calMonth < 0) { calMonth = 11; calYear--; }
    renderCalGrid();
  });
  document.getElementById('calNextMonth').addEventListener('click', () => {
    calMonth++; if (calMonth > 11) { calMonth = 0; calYear++; }
    renderCalGrid();
  });

  // ═══ CSV utils ═══════════════════════════════════════════════
  function csvParseText(text) {
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (!lines.length) return [];
    const delim = (lines[0].match(/;/g) || []).length >= (lines[0].match(/,/g) || []).length ? ';' : ',';
    return lines.map(l => l.split(delim).map(c => c.trim().replace(/^"(.*)"$/, '$1')));
  }

  function csvDownload(filename, lines) {
    const blob = new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  // "dd/mm/yyyy" o "yyyy-mm-dd" → ISO; null se invalida
  function parseDataFlex(s) {
    s = String(s || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    return null;
  }

  // "1.234,56" | "1234.56" | "12,5" → numero
  function parseImporto(s) {
    s = String(s || '').trim().replace(/[€\s]/g, '');
    if (!s) return NaN;
    if (s.includes(',')) return parseFloat(s.replace(/\./g, '').replace(',', '.'));
    return parseFloat(s);
  }

  // ═══ Elenco spese: import / export CSV ═══════════════════════
  document.getElementById('btnCsvExport').addEventListener('click', () => {
    if (!allRows.length) { alert('Nessun dato da esportare.'); return; }
    const lines = ['Data;Costo;Descrizione;Categoria;Tipo'];
    [...allRows]
      .sort((a, b) => new Date(b.data) - new Date(a.data))
      .forEach(r => lines.push([r.data, r.costo, String(r.descrizione||'').replace(/;/g, ','), r.categoria, r.tipo].join(';')));
    csvDownload('spese.csv', lines);
  });

  // Import spese: bottone → modale con istruzioni e modello
  const csvImportModal = document.getElementById('csvImportModal');
  document.getElementById('btnCsvImport').addEventListener('click', () => { csvImportModal.style.display = 'flex'; });
  document.getElementById('csvImpAnnulla').addEventListener('click', () => { csvImportModal.style.display = 'none'; });
  document.getElementById('btnCsvChooseFile').addEventListener('click', () => {
    csvImportModal.style.display = 'none';
    document.getElementById('csvFileInput').click();
  });
  document.getElementById('btnCsvTemplate').addEventListener('click', () => {
    csvDownload('modello-spese.csv', [
      'Data;Costo;Descrizione;Categoria;Tipo',
      '15/01/2026;42,50;Spesa supermercato;Spesa;Uscite',
      '27/01/2026;1500,00;Stipendio gennaio;Stipendio;Entrate',
    ]);
  });

  document.getElementById('csvFileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    const cells = csvParseText(await file.text());
    if (!cells.length) { alert('File vuoto.'); return; }

    let start = /data/i.test(cells[0][0]) ? 1 : 0;
    const rows = [];
    let scartate = 0;
    for (let i = start; i < cells.length; i++) {
      const c = cells[i];
      const dataIso = parseDataFlex(c[0]);
      const costo   = parseImporto(c[1]);
      const desc    = (c[2] || '').trim();
      if (!dataIso || isNaN(costo) || !desc) { scartate++; continue; }
      const d = new Date(dataIso);
      const tipo = String(c[4] || '').trim().toLowerCase().startsWith('e') ? 'Entrate' : 'Uscite';
      rows.push([
        dataIso,
        costo.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
        desc,
        (c[3] || '').trim() || 'Altro',
        tipo,
        d.getMonth() + 1,
        d.getFullYear(),
      ]);
    }
    if (!rows.length) { alert('Nessuna riga valida trovata.\nColonne attese: Data;Costo;Descrizione;Categoria;Tipo'); return; }
    if (!confirm(`Importare ${rows.length} voci?` + (scartate ? ` (${scartate} righe scartate)` : ''))) return;

    const json = await asCall({ action: 'append_many', rows: JSON.stringify(rows) });
    if (json.status !== 'ok') { alert('Errore durante l\'import: ' + json.message); return; }
    allRows = [];
    sessionStorage.removeItem(CACHE_KEY);
    initApp(true).then(() => { populateAnnoFilter(); applyFilters(); });
  });

  // ═══ Mutuo: configurazione + CSV ══════════════════════════════
  const mutuoModal = document.getElementById('mutuoModal');

  function mcShowError(msg) {
    const fb = document.getElementById('mcFeedback');
    fb.textContent = msg;
    fb.className = 'form-feedback visible error';
    setTimeout(() => { fb.className = 'form-feedback'; }, 4000);
  }

  document.getElementById('btnMutuoConfig').addEventListener('click', () => {
    const now = new Date();
    const inp = document.getElementById('mcPrimaRata');
    if (!inp.value) inp.value = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
    mutuoModal.style.display = 'flex';
  });
  document.getElementById('mcAnnulla').addEventListener('click', () => { mutuoModal.style.display = 'none'; });

  document.getElementById('mcGenera').addEventListener('click', async () => {
    const debito = parseFloat(document.getElementById('mcDebito').value);
    const rata   = parseFloat(document.getElementById('mcRata').value);
    const tasso  = parseFloat(document.getElementById('mcTasso').value);
    const giorno = Math.min(28, Math.max(1, parseInt(document.getElementById('mcGiorno').value) || 1));
    const pm     = document.getElementById('mcPrimaRata').value; // YYYY-MM

    if (!(debito > 0) || !(rata > 0) || isNaN(tasso) || tasso < 0 || !pm) {
      mcShowError('Compila tutti i campi con valori validi.');
      return;
    }
    const im = tasso / 100 / 12;
    if (rata <= debito * im) {
      mcShowError('Rata troppo bassa: non copre gli interessi mensili.');
      return;
    }
    if (mAllRate.length && !confirm('Sostituire il piano di ammortamento esistente?')) return;

    const [py, pmm] = pm.split('-').map(Number);
    const rows = [];
    let residuo = debito, nr = 1;
    let d = new Date(py, pmm - 1, giorno);
    while (residuo > 0.005 && nr <= 1200) {
      const interesse = residuo * im;
      const capitale  = Math.min(rata - interesse, residuo);
      const pagato    = debito - residuo;
      const iso = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      rows.push([
        nr, iso,
        residuo.toFixed(2),
        pagato.toFixed(2),
        (pagato / debito * 100).toFixed(1),
        (interesse + capitale).toFixed(2),
        interesse.toFixed(2),
        capitale.toFixed(2),
        '',
      ]);
      residuo -= capitale;
      nr++;
      d = new Date(d.getFullYear(), d.getMonth() + 1, giorno);
    }

    const json = await asCall({ action: 'save_meta', doc: 'mutuo', json: JSON.stringify(rows) });
    if (json.status !== 'ok') { mcShowError('Errore nel salvataggio: ' + json.message); return; }
    mAllRate = rows;
    mutuoModal.style.display = 'none';
    buildMutuo(mAllRate);
  });

  document.getElementById('btnMutuoCsvTemplate').addEventListener('click', () => {
    csvDownload('modello-mutuo.csv', [
      'Nr;Data;Debito;Debito pagato;Percentuale;Rata;Interesse;Capitale;Pagata',
      '1;01/06/2019;120000.00;0.00;0.0;550.00;250.00;300.00;',
      '2;01/07/2019;119700.00;300.00;0.3;550.00;249.38;300.62;',
    ]);
  });

  // Import mutuo: bottone → modale con istruzioni e modello
  const mutuoCsvModal = document.getElementById('mutuoCsvModal');
  document.getElementById('btnMutuoCsvImport').addEventListener('click', () => { mutuoCsvModal.style.display = 'flex'; });
  document.getElementById('mutuoCsvAnnulla').addEventListener('click', () => { mutuoCsvModal.style.display = 'none'; });
  document.getElementById('btnMutuoChooseFile').addEventListener('click', () => {
    mutuoCsvModal.style.display = 'none';
    document.getElementById('mutuoCsvInput').click();
  });

  document.getElementById('mutuoCsvInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    const cells = csvParseText(await file.text());
    if (!cells.length) { alert('File vuoto.'); return; }

    const start = /^nr/i.test(cells[0][0]) ? 1 : 0;
    const rows = cells.slice(start).filter(c => c.length >= 8 && c[1]);
    if (!rows.length) {
      alert('Nessuna riga valida.\nColonne attese: Nr;Data;Debito;Debito pagato;Percentuale;Rata;Interesse;Capitale;Pagata');
      return;
    }
    if (mAllRate.length && !confirm(`Sostituire il piano esistente con ${rows.length} rate importate?`)) return;

    const json = await asCall({ action: 'save_meta', doc: 'mutuo', json: JSON.stringify(rows) });
    if (json.status !== 'ok') { alert('Errore nel salvataggio: ' + json.message); return; }
    mAllRate = rows;
    buildMutuo(mAllRate);
  });

  // ═══ Patrimonio: editor attivi/passivi ════════════════════════
  const patModal = document.getElementById('patModal');
  const patList  = document.getElementById('patEditList');

  function patAddEditorRow(item = {}) {
    const row = document.createElement('div');
    row.className = 'pat-edit-row';
    row.innerHTML = `
      <input class="form-input" type="text" placeholder="Nome asset" value="${escHtml(item.asset || '')}" data-f="asset">
      <select class="filter-select" data-f="tipologia">
        <option value="Attivo">Attivo</option>
        <option value="Passivo">Passivo</option>
      </select>
      <input class="form-input" type="number" step="0.01" placeholder="Valore €" value="${item.valore ?? ''}" data-f="valore">
      <button class="btn-row delete" title="Rimuovi">
        <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M3 3l7 7M10 3l-7 7" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
      </button>`;
    row.querySelector('[data-f="tipologia"]').value =
      String(item.tipologia || '').toLowerCase().includes('passiv') ? 'Passivo' : 'Attivo';
    row.querySelector('button').addEventListener('click', () => row.remove());
    patList.appendChild(row);
  }

  document.getElementById('btnPatEdit').addEventListener('click', () => {
    patList.innerHTML = '';
    if (patrimonioRows.length) patrimonioRows.forEach(patAddEditorRow);
    else patAddEditorRow();
    patModal.style.display = 'flex';
  });

  document.getElementById('patAddRow').addEventListener('click', () => patAddEditorRow());
  document.getElementById('patAnnulla').addEventListener('click', () => { patModal.style.display = 'none'; });

  document.getElementById('patSalva').addEventListener('click', async () => {
    const items = [...patList.children].map(row => ({
      asset:     row.querySelector('[data-f="asset"]').value.trim(),
      tipologia: row.querySelector('[data-f="tipologia"]').value,
      valore:    parseFloat(row.querySelector('[data-f="valore"]').value) || 0,
    })).filter(i => i.asset);

    const json = await asCall({ action: 'save_meta', doc: 'patrimonio', json: JSON.stringify(items) });
    if (json.status !== 'ok') { alert('Errore nel salvataggio: ' + json.message); return; }
    patrimonioRows = items;
    patModal.style.display = 'none';
    renderPatrimonio();
  });

  // ═══ Spese ricorrenti ═════════════════════════════════════════
  const RIC_FREQ = { mensile: 1, bimestrale: 2, trimestrale: 3, semestrale: 6, annuale: 12 };
  const RIC_FREQ_LABEL = { mensile: 'Mensile', bimestrale: 'Bimestrale', trimestrale: 'Trimestrale', semestrale: 'Semestrale', annuale: 'Annuale' };

  // Prossima occorrenza >= oggi, avanzando dalla data salvata secondo la frequenza
  function ricProssima(item) {
    const step = RIC_FREQ[item.ricorrenza] || 12;
    const oggi = new Date(); oggi.setHours(0,0,0,0);
    let d = new Date(item.prossima);
    if (isNaN(d)) return null;
    let guard = 0;
    while (d < oggi && guard < 600) { d = new Date(d.getFullYear(), d.getMonth() + step, d.getDate()); guard++; }
    return d;
  }

  function ricGiorniMancanti(d) {
    const oggi = new Date(); oggi.setHours(0,0,0,0);
    return Math.round((d - oggi) / 86400000);
  }

  const fmtEurRic = v => (+v).toLocaleString('it-IT', { style:'currency', currency:'EUR', maximumFractionDigits: 0 });

  // Spese in arrivo entro `giorni`: una riga per spesa (solo la prossima occorrenza)
  function ricOccorrenze(giorni) {
    return ricorrentiRows
      .map((it, idx) => ({ ...it, _idx: idx, next: ricProssima(it) }))
      .filter(it => it.next && ricGiorniMancanti(it.next) <= giorni)
      .sort((a, b) => a.next - b.next);
  }

  // Tick "pagata": avanza la ricorrenza alla prossima occorrenza e salva.
  // La voce resta in "Gestisci", sparisce da prossime spese e notifiche.
  window.ricPagata = async function(idx) {
    const it = ricorrentiRows[idx];
    if (!it) return;
    const cur = ricProssima(it);
    if (!cur) return;
    const step = RIC_FREQ[it.ricorrenza] || 12;
    const next = new Date(cur.getFullYear(), cur.getMonth() + step, cur.getDate());
    it.prossima = `${next.getFullYear()}-${String(next.getMonth()+1).padStart(2,'0')}-${String(next.getDate()).padStart(2,'0')}`;

    const json = await asCall({ action: 'save_meta', doc: 'ricorrenti', json: JSON.stringify(ricorrentiRows) });
    if (json.status !== 'ok') { alert('Errore nel salvataggio: ' + json.message); return; }
    renderProssime();
    renderCalKpi();
    updateNotifiche();
  };

  // Card "Prossime spese" con pill Settimana / Mese / Anno
  let ricRangeDays = 31;

  function renderProssime() {
    const list = document.getElementById('ricUpcomingList');
    const occ = ricOccorrenze(ricRangeDays);
    if (!occ.length) {
      list.innerHTML = ricorrentiRows.length
        ? '<p style="color:var(--text-hint);font-size:13px;">Nessuna spesa in arrivo nel periodo selezionato.</p>'
        : '<p style="color:var(--text-hint);font-size:13px;">Nessuna ricorrenza configurata. Usa "Gestisci" per aggiungerne.</p>';
      return;
    }
    list.innerHTML = occ.map(it => {
      const giorni = ricGiorniMancanti(it.next);
      const urgente = giorni <= 31;
      const quando = giorni === 0 ? 'oggi' : giorni === 1 ? 'domani' : `tra ${giorni} giorni`;
      return `
        <div class="ric-row${urgente ? ' urgente' : ''}">
          <span class="ric-nome">${escHtml(it.nome)}</span>
          <span class="ric-freq">${RIC_FREQ_LABEL[it.ricorrenza] || it.ricorrenza}</span>
          <span class="ric-importo">${it.importo ? fmtEurRic(it.importo) : ''}</span>
          <span class="ric-data">${it.next.toLocaleDateString('it-IT')}</span>
          <span class="tag ${urgente ? 'tag-amber' : 'tag-neutral'}">${quando}</span>
          <button class="btn-row save" title="Segna come pagata" onclick="ricPagata(${it._idx})">
            <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M2 7l3.5 3.5L11 3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>
          </button>
        </div>`;
    }).join('');
  }

  document.querySelectorAll('#ricRangePills .u-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('#ricRangePills .u-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      ricRangeDays = parseInt(pill.dataset.range);
      renderProssime();
    });
  });

  // KPI della sezione calendario
  function renderCalKpi() {
    const oggi = new Date();
    const anno = String(oggi.getFullYear());
    const mese = oggi.getMonth() + 1;
    const prevAnno = mese === 1 ? String(oggi.getFullYear() - 1) : anno;
    const prevMese = mese === 1 ? 12 : mese - 1;

    // Spesa in arrivo: l'occorrenza più vicina
    const prox = ricOccorrenze(365)[0];
    const kProx = document.getElementById('calKpiProssima');
    const kProxSub = document.getElementById('calKpiProssimaSub');
    if (prox) {
      const giorni = ricGiorniMancanti(prox.next);
      kProx.textContent = prox.nome + (prox.importo ? ` · ${fmtEurRic(prox.importo)}` : '');
      kProxSub.textContent = `${prox.next.toLocaleDateString('it-IT')} — ${giorni === 0 ? 'oggi' : giorni === 1 ? 'domani' : 'tra ' + giorni + ' giorni'}`;
    } else {
      kProx.textContent = '—';
      kProxSub.textContent = 'nessuna ricorrenza configurata';
    }

    const opsOf = (a, m) => allRows.filter(r => r.anno === a && parseInt(r.mese) === m);
    const cur  = opsOf(anno, mese);
    const prev = opsOf(prevAnno, prevMese);
    const uscTot = rows => rows.filter(r => r.tipo === 'Uscite').reduce((s,r) => s + parseCosto(r.costo), 0);

    document.getElementById('calKpiOpsCur').textContent = cur.length;
    document.getElementById('calKpiOpsCurSub').textContent = `uscite ${fmtEur(uscTot(cur))}`;
    document.getElementById('calKpiOpsPrev').textContent = prev.length;
    document.getElementById('calKpiOpsPrevSub').textContent = `uscite ${fmtEur(uscTot(prev))}`;
  }

  // ═══ Notifiche (campanella in header) ═════════════════════════
  function updateNotifiche() {
    const inArrivo = ricOccorrenze(31);
    const badge = document.getElementById('notifBadge');
    const list  = document.getElementById('notifList');

    badge.style.display = inArrivo.length ? 'flex' : 'none';
    badge.textContent = inArrivo.length;

    list.innerHTML = inArrivo.length
      ? inArrivo.map(it => {
          const giorni = ricGiorniMancanti(it.next);
          const quando = giorni === 0 ? 'oggi' : giorni === 1 ? 'domani' : `tra ${giorni} giorni`;
          return `<div class="notif-item">
            <div>
              <div class="notif-nome">${escHtml(it.nome)}${it.importo ? ' · ' + fmtEurRic(it.importo) : ''}</div>
              <div class="notif-sub">${it.next.toLocaleDateString('it-IT')} — ${quando}</div>
            </div>
          </div>`;
        }).join('')
      : '<p class="notif-empty">Nessuna notifica.</p>';
  }

  const notifPanel = document.getElementById('notifPanel');
  document.getElementById('notifBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    notifPanel.style.display = notifPanel.style.display === 'none' ? 'block' : 'none';
  });
  document.addEventListener('click', (e) => {
    if (notifPanel.style.display !== 'none' && !e.target.closest('.notif-wrap')) {
      notifPanel.style.display = 'none';
    }
  });

  // Editor ricorrenze (modale)
  const ricModal = document.getElementById('ricModal');
  const ricEditList = document.getElementById('ricEditList');

  function ricAddEditorRow(item = {}) {
    const row = document.createElement('div');
    row.className = 'ric-edit-row';
    const freqOpts = Object.keys(RIC_FREQ)
      .map(f => `<option value="${f}">${RIC_FREQ_LABEL[f]}</option>`).join('');
    row.innerHTML = `
      <input class="form-input" type="text" placeholder="Nome (es. Bollo auto)" value="${escHtml(item.nome || '')}" data-f="nome">
      <input class="form-input" type="number" step="0.01" placeholder="Importo €" value="${item.importo ?? ''}" data-f="importo">
      <select class="filter-select" data-f="ricorrenza">${freqOpts}</select>
      <input class="form-input" type="date" value="${item.prossima || ''}" data-f="prossima">
      <button class="btn-row delete" title="Rimuovi">
        <svg width="13" height="13" viewBox="0 0 13 13" fill="none"><path d="M3 3l7 7M10 3l-7 7" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
      </button>`;
    if (item.ricorrenza) row.querySelector('[data-f="ricorrenza"]').value = item.ricorrenza;
    row.querySelector('button').addEventListener('click', () => row.remove());
    ricEditList.appendChild(row);
  }

  document.getElementById('btnRicEdit').addEventListener('click', () => {
    ricEditList.innerHTML = '';
    if (ricorrentiRows.length) ricorrentiRows.forEach(ricAddEditorRow);
    else ricAddEditorRow();
    ricModal.style.display = 'flex';
  });

  document.getElementById('ricAddRow').addEventListener('click', () => ricAddEditorRow());
  document.getElementById('ricAnnulla').addEventListener('click', () => { ricModal.style.display = 'none'; });

  document.getElementById('ricSalva').addEventListener('click', async () => {
    const items = [...ricEditList.children].map(row => ({
      nome:       row.querySelector('[data-f="nome"]').value.trim(),
      importo:    parseFloat(row.querySelector('[data-f="importo"]').value) || 0,
      ricorrenza: row.querySelector('[data-f="ricorrenza"]').value,
      prossima:   row.querySelector('[data-f="prossima"]').value,
    })).filter(i => i.nome && i.prossima);

    const json = await asCall({ action: 'save_meta', doc: 'ricorrenti', json: JSON.stringify(items) });
    if (json.status !== 'ok') { alert('Errore nel salvataggio: ' + json.message); return; }
    ricorrentiRows = items;
    ricModal.style.display = 'none';
    renderProssime();
    renderCalKpi();
    updateNotifiche();
  });

  // ── Logout ──
  document.getElementById('logoutBtn').addEventListener('click', async () => {
    sessionStorage.removeItem(CACHE_KEY);
    await firebase.auth().signOut();
    window.location.replace('login.html');
  });

  // ── Avvio app (solo dopo auth Firebase) ──
  bmAuthReady.then(user => {
    if (!bmIsAllowed(user)) return; // redirect gestito dal guard in app.html

    // Avatar Google in header
    const av = document.getElementById('userAvatar');
    if (user.photoURL) {
      av.src = user.photoURL;
      av.style.display = 'block';
    }
    av.title = user.displayName || user.email;

    initApp();
  });
