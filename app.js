document.addEventListener('DOMContentLoaded', () => {
    // === ESTADO ===
    let db = null;
    let allSheetsData = {};
    let currentClinic = '';
    let calibrationDates = {};
    let instrumentsBank = [];
    let savedTemplates = [];
    let selectedSerieForEdit = null;

    // === SCHEMA EXACTO DEL EXCEL 2025-MP BANO ===
    // 8.1 Test de Inspección y Funcionalidad — Marcar P (Pasó) o F (Falló), filas 24-34
    const SCHEMA_81 = [
        { code: '8.1.1', label: 'Chasis', row: 24 },
        { code: '8.1.2', label: 'Montajes y Apoyos', row: 25 },
        { code: '8.1.3', label: 'Ficha de alimentación y tomacorriente', row: 26 },
        { code: '8.1.4', label: 'Cable de alimentación', row: 27 },
        { code: '8.1.5', label: 'Amarres contra tirones', row: 28 },
        { code: '8.1.6', label: 'Interruptores y fusibles', row: 29 },
        { code: '8.1.7', label: 'Escala', row: 30 },
        { code: '8.1.8', label: 'Sonda de medición', row: 31 },
        { code: '8.1.9', label: 'Controles y Teclas', row: 32 },
        { code: '8.1.10', label: 'Indicadores y Displays', row: 33 },
        { code: '8.1.11', label: 'Etiquetado', row: 34 },
    ];

    // 8.2.1 Características del termómetro patrón — valores fijos en col I
    const SCHEMA_821 = [
        { label: 'Incertidumbre de calibración (Up)', row: 38, value: 0.2 },
        { label: 'Resolución (rp)', row: 39, value: 0.1 },
    ];

    // 8.2.2 Características del termómetro a calibrar — valores fijos en col I
    const SCHEMA_822 = [
        { label: 'Resolución (rx)', row: 41, value: 1 },
        { label: 'Error máximo permitido (EMP)', row: 42, value: 'Ambiental (2 °C)' },
    ];

    // Estado de Valoración — 4 ítems
    const EVALUATION_SCHEMA = [
        { label: 'Inspección superada, el equipo es apto para el uso', row: 15 },
        { label: 'El equipo ha necesitado reparación', row: 16 },
        { label: 'El equipo no está reparado. No se puede usar', row: 17 },
        { label: 'El equipo se da DE BAJA', row: 18 },
    ];

    // Mediciones: filas 44-46 (H=Patrón, I=Termómetro a calibrar)
    const READINGS_START_ROW = 44;

    const DB_NAME = 'CalibracionesDB_Banos_v1';
    const DB_VERSION = 1;

    // === UTILIDADES ===
    function escapeHtml(str) {
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(str));
        return div.innerHTML;
    }

    function safeParseFloat(val) {
        if (val === '' || val === null || val === undefined) return null;
        const n = parseFloat(val);
        return isNaN(n) ? null : n;
    }

    // DOM
    const fileInput = document.getElementById('fileInput');
    const fileLabel = document.getElementById('fileLabel');
    const mainContent = document.getElementById('mainContent');
    const sheetSelector = document.getElementById('sheetSelector');
    const serieFilter = document.getElementById('serieFilter');
    const equiposTableBody = document.getElementById('equiposTableBody');
    const editModal = document.getElementById('editModal');
    const calibDateInput = document.getElementById('calibDateInput');
    const ordenMInput = document.getElementById('ordenMInput');
    const technicianInput = document.getElementById('technicianInput');
    const buildingInput = document.getElementById('buildingInput');
    const sectorInput = document.getElementById('sectorInput');
    const locationInput = document.getElementById('locationInput');
    const commentsInput = document.getElementById('commentsInput');
    const equipmentNameInput = document.getElementById('equipmentNameInput');
    const modalSerieInput = document.getElementById('modalSerieInput');
    const modelInput = document.getElementById('modelInput');
    const brandInput = document.getElementById('brandInput');
    const addInstrumentBtn = document.getElementById('addInstrumentBtn');
    const instrumentsContainer = document.getElementById('instrumentsContainer');
    const certFileInput = document.getElementById('certFileInput');
    const certStatus = document.getElementById('certStatus');
    const templateSelector = document.getElementById('templateSelector');
    const templateNameInput = document.getElementById('templateNameInput');
    const saveNewTemplateBtn = document.getElementById('saveNewTemplateBtn');
    const saveTemplateRow = document.getElementById('saveTemplateRow');
    const saveCalibBtn = document.getElementById('saveCalibBtn');
    const totalEquiposEl = document.getElementById('totalEquipos').querySelector('.val');
    const cercaVencerEl = document.getElementById('cercaVencer').querySelector('.val');
    const vencidosEl = document.getElementById('vencidos').querySelector('.val');

    // === INIT ===
    async function init() {
        try {
            await initDB();
            await loadSavedData();
            setupEventListeners();
            loadTemplates();
        } catch (err) {
            console.error('Init error:', err);
            alert('Error al iniciar: ' + err.message);
        }
    }

    // === INDEXEDDB ===
    function initDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(DB_NAME, DB_VERSION);
            req.onupgradeneeded = e => {
                const d = e.target.result;
                if (!d.objectStoreNames.contains('calibrations')) d.createObjectStore('calibrations', { keyPath: 'serie' });
                if (!d.objectStoreNames.contains('appData')) d.createObjectStore('appData', { keyPath: 'id' });
                if (!d.objectStoreNames.contains('templates')) d.createObjectStore('templates', { keyPath: 'id', autoIncrement: true });
            };
            req.onsuccess = e => { db = e.target.result; resolve(); };
            req.onerror = e => reject(e.target.error);
        });
    }

    async function storeCalibration(data) {
        const tx = db.transaction('calibrations', 'readwrite');
        const store = tx.objectStore('calibrations');
        if (!data.certificate) {
            const existing = await new Promise(r => { const q = store.get(data.serie); q.onsuccess = () => r(q.result); });
            if (existing && existing.certificate) {
                data.certificate = existing.certificate;
                data.certName = existing.certName;
            }
        }
        store.put(data);
    }

    function getAllCalibrations() {
        return new Promise(resolve => {
            if (!db) { resolve({}); return; }
            const map = {}, tx = db.transaction('calibrations', 'readonly');
            tx.objectStore('calibrations').openCursor().onsuccess = e => {
                const cur = e.target.result;
                if (cur) { map[cur.key] = cur.value; cur.continue(); }
                else { calibrationDates = map; updateInstrumentsBank(); resolve(map); }
            };
        });
    }

    function updateInstrumentsBank() {
        const uniq = new Map();
        Object.values(calibrationDates).forEach(c => {
            (c.instruments || []).forEach(i => {
                if (i.name && !uniq.has(i.name.toUpperCase())) uniq.set(i.name.toUpperCase(), i);
            });
        });
        instrumentsBank = Array.from(uniq.values());
        const dl = document.getElementById('instrumentsHistory');
        if (dl) { dl.innerHTML = ''; instrumentsBank.forEach(i => { const o = document.createElement('option'); o.value = i.name; dl.appendChild(o); }); }
    }

    // === EXCEL LOADING ===
    fileInput.addEventListener('change', e => {
        const file = e.target.files[0]; if (!file) return;
        const reader = new FileReader();
        reader.onload = ev => processWorkbook(XLSX.read(new Uint8Array(ev.target.result), { type: 'array' }), file.name);
        reader.readAsArrayBuffer(file);
    });

    async function processWorkbook(wb, filename) {
        allSheetsData = {};
        sheetSelector.innerHTML = '';
        wb.SheetNames.forEach(name => {
            allSheetsData[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' }).filter(r => Object.values(r).some(v => v !== ''));
            const opt = document.createElement('option'); opt.value = opt.textContent = name;
            sheetSelector.appendChild(opt);
        });
        currentClinic = wb.SheetNames[0];
        fileLabel.textContent = `✅ ${filename}`;
        db.transaction('appData', 'readwrite').objectStore('appData').put({ id: 'lastExcel', filename, allSheetsData, sheetNames: wb.SheetNames, currentClinic });
        mainContent.classList.remove('hidden');
        document.getElementById('configActions').classList.remove('hidden');
        renderTable();
    }

    async function loadSavedData() {
        const tx = db.transaction('appData', 'readonly');
        const last = await new Promise(r => { const q = tx.objectStore('appData').get('lastExcel'); q.onsuccess = () => r(q.result); });
        if (!last) return;
        allSheetsData = last.allSheetsData; currentClinic = last.currentClinic;
        sheetSelector.innerHTML = '';
        last.sheetNames.forEach(n => { const o = document.createElement('option'); o.value = o.textContent = n; if (n === currentClinic) o.selected = true; sheetSelector.appendChild(o); });
        fileLabel.textContent = `✅ ${last.filename} (Recuperado)`;
        mainContent.classList.remove('hidden');
        document.getElementById('configActions').classList.remove('hidden');
        renderTable();
    }

    document.getElementById('clearDataBtn').addEventListener('click', () => {
        if (confirm('¿Borrar datos cargados?')) { db.transaction('appData', 'readwrite').objectStore('appData').delete('lastExcel'); location.reload(); }
    });

    // === BACKUP: EXPORTAR / IMPORTAR ===
    function blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }

    function base64ToBlob(dataUrl) {
        const [header, data] = dataUrl.split(',');
        const mime = header.match(/:(.*?);/)[1];
        const bytes = atob(data);
        const arr = new Uint8Array(bytes.length);
        for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
        return new Blob([arr], { type: mime });
    }

    document.getElementById('exportBackupBtn').addEventListener('click', async () => {
        try {
            const backup = { version: 1, exportDate: new Date().toISOString(), calibrations: {}, templates: [] };
            const calTx = db.transaction('calibrations', 'readonly');
            const allCals = await new Promise(r => { const q = calTx.objectStore('calibrations').getAll(); q.onsuccess = () => r(q.result); });
            for (const cal of allCals) {
                const entry = { ...cal };
                if (entry.certificate instanceof Blob) {
                    entry._certBase64 = await blobToBase64(entry.certificate);
                    delete entry.certificate;
                }
                backup.calibrations[cal.serie] = entry;
            }
            const tmplTx = db.transaction('templates', 'readonly');
            const allTmpls = await new Promise(r => { const q = tmplTx.objectStore('templates').getAll(); q.onsuccess = () => r(q.result); });
            for (const tmpl of allTmpls) {
                const entry = { ...tmpl };
                if (entry.blob instanceof Blob) {
                    entry._blobBase64 = await blobToBase64(entry.blob);
                    delete entry.blob;
                }
                backup.templates.push(entry);
            }
            const json = JSON.stringify(backup, null, 2);
            const blob = new Blob([json], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `backup_banos_${new Date().toISOString().slice(0, 10)}.json`;
            a.click();
            URL.revokeObjectURL(url);
            alert('✅ Backup exportado correctamente.');
        } catch (err) {
            console.error('Error al exportar:', err);
            alert('Error al exportar: ' + err.message);
        }
    });

    document.getElementById('importBackupBtn').addEventListener('click', () => {
        document.getElementById('importBackupFile').click();
    });

    document.getElementById('importBackupFile').addEventListener('change', async e => {
        const file = e.target.files[0];
        if (!file) return;
        try {
            const text = await file.text();
            const backup = JSON.parse(text);
            if (!backup.version || !backup.calibrations) throw new Error('Formato de backup inválido');
            const count = { cals: 0, tmpls: 0 };
            const calTx = db.transaction('calibrations', 'readwrite');
            const calStore = calTx.objectStore('calibrations');
            for (const [serie, cal] of Object.entries(backup.calibrations)) {
                const entry = { ...cal, serie };
                if (entry._certBase64) {
                    entry.certificate = base64ToBlob(entry._certBase64);
                    delete entry._certBase64;
                }
                calStore.put(entry);
                count.cals++;
            }
            if (backup.templates && backup.templates.length > 0) {
                const tmplTx = db.transaction('templates', 'readwrite');
                const tmplStore = tmplTx.objectStore('templates');
                for (const tmpl of backup.templates) {
                    const entry = { ...tmpl };
                    delete entry.id;
                    if (entry._blobBase64) {
                        entry.blob = base64ToBlob(entry._blobBase64);
                        delete entry._blobBase64;
                    }
                    tmplStore.add(entry);
                    count.tmpls++;
                }
            }
            alert(`✅ Backup importado: ${count.cals} calibraciones, ${count.tmpls} plantillas.`);
            location.reload();
        } catch (err) {
            console.error('Error al importar:', err);
            alert('Error al importar: ' + err.message);
        }
        e.target.value = '';
    });

    // === TABLA ===
    async function renderTable() {
        if (!currentClinic || !allSheetsData[currentClinic]) return;
        await getAllCalibrations();
        const rows = allSheetsData[currentClinic];
        const search = serieFilter.value.trim().toUpperCase();
        equiposTableBody.innerHTML = '';
        let stats = { total: 0, warn: 0, danger: 0 };

        rows.forEach(row => {
            const keys = Object.keys(row);
            const serieKey = keys.find(k => k.toLowerCase().includes('serie') || k.toLowerCase().includes('n°') || k.toLowerCase().includes('sensor'));
            const nombreKey = keys.find(k => k.toLowerCase().includes('equipo') || k.toLowerCase().includes('nombre') || k.toLowerCase().includes('ubicacion') || k.toLowerCase().includes('ubicación'));

            const serie = String(row[serieKey] || '').toUpperCase().trim();
            if (!serie || serie === '') return;
            if (search && !serie.includes(search)) return;

            stats.total++;
            const cal = calibrationDates[serie] || null;
            const status = getStatus(cal?.date);
            if (status.class === 'status-warning') stats.warn++;
            if (status.class === 'status-danger') stats.danger++;

            const displayName = cal?.editedName || (nombreKey ? row[nombreKey] : 'N/A');
            const displaySerie = cal?.editedSerie || serie;
            const safeSerie = escapeHtml(serie);
            const safeDisplayName = escapeHtml(String(displayName));
            const safeDisplaySerie = escapeHtml(String(displaySerie));

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${safeDisplayName}</td>
                <td>${safeDisplaySerie}</td>
                <td>${cal ? formatDate(cal.date) : '<span style="color:#aaa">Sin registrar</span>'}</td>
                <td>${escapeHtml(cal?.technician || '-')}</td>
                <td>${cal?.certificate ? `<button class="btn btn-small" data-action="viewCert" data-serie="${safeSerie}">📄</button>` : '-'}</td>
                <td><span class="status-badge ${status.class}">${escapeHtml(status.text)}</span></td>
                <td><button class="btn btn-secondary btn-small" data-action="openEdit" data-serie="${safeSerie}">📅 Registrar</button></td>
            `;
            equiposTableBody.appendChild(tr);
        });

        totalEquiposEl.textContent = stats.total;
        cercaVencerEl.textContent = stats.warn;
        vencidosEl.textContent = stats.danger;
    }

    function getStatus(dateStr) {
        if (!dateStr) return { text: 'Pendiente', class: '' };
        const next = new Date(dateStr); next.setFullYear(next.getFullYear() + 1);
        const diff = Math.ceil((next - new Date()) / 86400000);
        if (diff < 0) return { text: 'Vencido', class: 'status-danger' };
        if (diff <= 30) return { text: `Vence ${diff}d`, class: 'status-warning' };
        return { text: 'Vigente', class: 'status-ok' };
    }

    // === INSTRUMENTAL ===
    function createInstrumentRow(data = {}) {
        const div = document.createElement('div'); div.className = 'instrument-item';
        const dateVal = formatDateForInput(data.date);
        div.innerHTML = `
            <button type="button" class="remove-instrument">×</button>
            <div class="field-group full-width"><label>Nombre del Instrumental</label><input type="text" class="inst-name" list="instrumentsHistory" value="${escapeHtml(data.name || '')}"></div>
            <div class="field-group"><label>Marca</label><input type="text" class="inst-brand" value="${escapeHtml(data.brand || '')}"></div>
            <div class="field-group"><label>Modelo</label><input type="text" class="inst-model" value="${escapeHtml(data.model || '')}"></div>
            <div class="field-group"><label>N° Serie</label><input type="text" class="inst-serie" value="${escapeHtml(data.serie || '')}"></div>
            <div class="field-group"><label>Últ. Calibración</label><input type="date" class="inst-date" value="${dateVal}"></div>
        `;
        div.querySelector('.remove-instrument').onclick = () => div.remove();
        instrumentsContainer.appendChild(div);
    }
    addInstrumentBtn.onclick = () => createInstrumentRow();

    function getInstrumentsData() {
        return Array.from(instrumentsContainer.querySelectorAll('.instrument-item')).map(div => ({
            name: div.querySelector('.inst-name').value,
            brand: div.querySelector('.inst-brand').value,
            model: div.querySelector('.inst-model').value,
            serie: div.querySelector('.inst-serie').value,
            date: div.querySelector('.inst-date').value,
        }));
    }

    // === INSPECCIÓN UI ===
    function renderInspectionPoints(saved = {}) {
        const container = document.getElementById('inspectionPointsContainer');
        container.innerHTML = '';

        // ── Sección 8.1 ──────────────────────────────────────────────────
        addSectionHeader(container, '8.1 Test de Inspección y Funcionalidad');
        addSubLabel(container, 'Marcar P (Pasó) o F (Falló) según corresponda');
        SCHEMA_81.forEach(item => {
            const saved_val = saved[item.label] || 'na';
            const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
            rowEl.innerHTML = `
                <div class="inspection-label"><strong>${item.code}</strong> ${item.label}</div>
                <div class="inspection-options" data-label="${item.label}" data-type="choice" data-row="${item.row}">
                    <div class="inspection-opt ${saved_val === 'P' ? 'selected' : ''}" data-val="P">P</div>
                    <div class="inspection-opt ${saved_val === 'F' ? 'selected' : ''}" data-val="F" style="background:${saved_val === 'F' ? '#e53e3e' : ''}">F</div>
                    <div class="inspection-opt ${saved_val === 'na' ? 'selected' : ''}" data-val="na">NA</div>
                </div>`;
            wireChoiceOpts(rowEl);
            container.appendChild(rowEl);
        });

        // ── Sección 8.2.1 (valores fijos) ────────────────────────────────
        addSectionHeader(container, '8.2.1 Características del Termómetro Patrón');
        SCHEMA_821.forEach(item => {
            const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
            rowEl.innerHTML = `
                <div class="inspection-label">${item.label}</div>
                <div class="inspection-options" style="justify-content:flex-end;padding:6px 12px;">
                    <span style="color:var(--accent);font-weight:700;">${item.value} °C</span>
                </div>`;
            container.appendChild(rowEl);
        });

        // ── Sección 8.2.2 (valores fijos) ────────────────────────────────
        addSectionHeader(container, '8.2.2 Características del Termómetro a Calibrar');
        SCHEMA_822.forEach(item => {
            const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
            const unit = typeof item.value === 'number' ? '°C' : '';
            rowEl.innerHTML = `
                <div class="inspection-label">${item.label}</div>
                <div class="inspection-options" style="justify-content:flex-end;padding:6px 12px;">
                    <span style="color:var(--accent);font-weight:700;">${item.value} ${unit}</span>
                </div>`;
            container.appendChild(rowEl);
        });

        // ── Sección 8.2.3 Mediciones (3 lecturas: Patrón H44-46 | Termómetro I44-46) ──
        addSectionHeader(container, '8.2.3 Mediciones');

        const savedReadings = saved['_readings'] || {};
        const readingsDiv = document.createElement('div');
        readingsDiv.id = 'readingsContainer';
        readingsDiv.style.cssText = 'border:1px solid #3a3a5c;border-radius:8px;padding:12px;margin-bottom:12px;background:#1a1a2e;';

        const r1 = READINGS_START_ROW;
        const ePt1 = escapeHtml(savedReadings.pt1 || '');
        const ePt2 = escapeHtml(savedReadings.pt2 || '');
        const ePt3 = escapeHtml(savedReadings.pt3 || '');
        const eTc1 = escapeHtml(savedReadings.tc1 || '');
        const eTc2 = escapeHtml(savedReadings.tc2 || '');
        const eTc3 = escapeHtml(savedReadings.tc3 || '');

        readingsDiv.innerHTML = `
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:4px;background:#2a2a4a;padding:6px;border-radius:4px;">
                <div class="field-group"><label>Patrón - Lectura 1</label><input type="number" step="any" class="r-pt1" value="${ePt1}"></div>
                <div class="field-group"><label>Termómetro - Lectura 1</label><input type="number" step="any" class="r-tc1" value="${eTc1}"></div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:4px;background:#2a2a4a;padding:6px;border-radius:4px;">
                <div class="field-group"><label>Patrón - Lectura 2</label><input type="number" step="any" class="r-pt2" value="${ePt2}"></div>
                <div class="field-group"><label>Termómetro - Lectura 2</label><input type="number" step="any" class="r-tc2" value="${eTc2}"></div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;background:#2a2a4a;padding:6px;border-radius:4px;">
                <div class="field-group"><label>Patrón - Lectura 3</label><input type="number" step="any" class="r-pt3" value="${ePt3}"></div>
                <div class="field-group"><label>Termómetro - Lectura 3</label><input type="number" step="any" class="r-tc3" value="${eTc3}"></div>
            </div>
        `;
        container.appendChild(readingsDiv);

        // ── Sección 8.2.4 Calibración (fórmulas estadísticas se calculan al exportar Excel) ──
        addSectionHeader(container, '8.2.4 Calibración');
        const statsInfo = document.createElement('div');
        statsInfo.style.cssText = 'padding:10px 14px;background:#2a2a4a;border-radius:6px;font-size:0.85em;color:#ccc;line-height:1.8;';
        statsInfo.innerHTML = `
            <div><strong style="color:#7c83fd;">8.2.4.1</strong> Incertidumbre combinada expandida (Uc)</div>
            <div><strong style="color:#7c83fd;">8.2.4.2</strong> Error sistemático (ES)</div>
            <div><strong style="color:#7c83fd;">8.2.4.3</strong> Error total en la medida del termómetro a calibrar (ETtx)</div>
        `;
        container.appendChild(statsInfo);
    }

    function addSectionHeader(container, text) {
        const h = document.createElement('div'); h.className = 'inspection-category'; h.textContent = text;
        container.appendChild(h);
    }
    function addSubLabel(container, text) {
        const p = document.createElement('p'); p.style.cssText = 'font-size:0.78em;color:#aaa;margin:2px 0 8px 0;'; p.textContent = text;
        container.appendChild(p);
    }

    function wireChoiceOpts(rowEl) {
        rowEl.querySelectorAll('.inspection-opt').forEach(opt => {
            opt.onclick = () => {
                rowEl.querySelectorAll('.inspection-opt').forEach(o => {
                    o.classList.remove('selected');
                    if (o.dataset.val === 'F') o.style.background = '';
                });
                opt.classList.add('selected');
                if (opt.dataset.val === 'F') opt.style.background = '#e53e3e';
            };
        });
    }

    function getInspectionsData() {
        const data = {};
        document.querySelectorAll('.inspection-options').forEach(g => {
            const label = g.dataset.label;
            if (g.dataset.type === 'choice') {
                const sel = g.querySelector('.inspection-opt.selected');
                data[label] = sel ? sel.dataset.val : 'na';
            }
        });
        // Lecturas de medición (8.2.3)
        const rc = document.getElementById('readingsContainer');
        if (rc) {
            data['_readings'] = {
                pt1: rc.querySelector('.r-pt1')?.value || '',
                pt2: rc.querySelector('.r-pt2')?.value || '',
                pt3: rc.querySelector('.r-pt3')?.value || '',
                tc1: rc.querySelector('.r-tc1')?.value || '',
                tc2: rc.querySelector('.r-tc2')?.value || '',
                tc3: rc.querySelector('.r-tc3')?.value || '',
            };
        }
        return data;
    }

    function getEvaluationsData() {
        const data = {};
        document.querySelectorAll('#evaluationStatusContainer .inspection-options').forEach(g => {
            const sel = g.querySelector('.inspection-opt.selected');
            data[g.dataset.label] = sel ? sel.dataset.val : '';
        });
        return data;
    }

    // === TEMPLATES ===
    async function loadTemplates() {
        if (!db) return;
        const tx = db.transaction('templates', 'readonly');
        savedTemplates = await new Promise(r => { const q = tx.objectStore('templates').getAll(); q.onsuccess = () => r(q.result); });
        templateSelector.innerHTML = '<option value="">-- Seleccionar Plantilla --</option>';
        savedTemplates.forEach(t => { const o = document.createElement('option'); o.value = t.id; o.textContent = t.name; templateSelector.appendChild(o); });
    }

    // === EVENT LISTENERS ===
    function setupEventListeners() {
        document.getElementById('dropZone').addEventListener('click', () => fileInput.click());
        sheetSelector.addEventListener('change', e => { currentClinic = e.target.value; renderTable(); });
        serieFilter.addEventListener('input', renderTable);

        certFileInput.addEventListener('change', async e => {
            const file = e.target.files[0]; if (!file) return;
            const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
            saveTemplateRow.classList.toggle('hidden', !isExcel);
            if (isExcel) {
                try {
                    const extracted = await extractInstrumentsFromExcel(file);
                    if (extracted && extracted.length > 0) {
                        const currentInstruments = getInstrumentsData();
                        if (currentInstruments.length > 0) {
                            if (confirm(`Se detectaron ${extracted.length} instrumentos en el Excel. ¿Deseas añadirlos a la lista actual? (Cancelar para limpiar primero)`)) {
                                extracted.forEach(inst => createInstrumentRow(inst));
                            } else {
                                instrumentsContainer.innerHTML = '';
                                extracted.forEach(inst => createInstrumentRow(inst));
                            }
                        } else {
                            extracted.forEach(inst => createInstrumentRow(inst));
                        }
                    }
                } catch (err) {
                    console.error("Error al extraer instrumentos del archivo subido:", err);
                }
            }
        });

        saveNewTemplateBtn.addEventListener('click', async () => {
            const file = certFileInput.files[0], name = templateNameInput.value.trim();
            if (!file || !name) { alert('Falta archivo o nombre'); return; }
            const tx = db.transaction('templates', 'readwrite');
            tx.objectStore('templates').add({ name, blob: file });
            tx.oncomplete = () => { alert('Plantilla guardada'); templateNameInput.value = ''; saveTemplateRow.classList.add('hidden'); loadTemplates(); };
        });

        document.getElementById('closeModalBtn').onclick = () => editModal.classList.add('hidden');

        equiposTableBody.addEventListener('click', e => {
            const btn = e.target.closest('[data-action]');
            if (!btn) return;
            const serie = btn.dataset.serie;
            if (btn.dataset.action === 'viewCert') {
                const c = calibrationDates[serie];
                if (c?.certificate) window.open(URL.createObjectURL(c.certificate), '_blank');
            } else if (btn.dataset.action === 'openEdit') {
                openEditModal(serie);
            }
        });

        function openEditModal(serie) {
            selectedSerieForEdit = serie;
            const existing = calibrationDates[serie] || {};
            const eqRow = (allSheetsData[currentClinic] || [])
                .find(r => String(r[Object.keys(r).find(k => k.toLowerCase().includes('serie') || k.toLowerCase().includes('n°'))] || '').toUpperCase() === serie) || {};

            calibDateInput.value = existing.date || '';
            ordenMInput.value = existing.ordenM || '';
            technicianInput.value = existing.technician || '';
            buildingInput.value = existing.building || eqRow.edificio || '';
            sectorInput.value = existing.sector || eqRow.sector || '';
            locationInput.value = existing.location || eqRow.ubicacion || eqRow['ubicación'] || '';
            equipmentNameInput.value = existing.editedName || eqRow.equipo || '';
            modalSerieInput.value = existing.editedSerie || serie;
            modelInput.value = existing.model || eqRow.modelo || '';
            brandInput.value = existing.brand || eqRow.marca || '';
            commentsInput.value = existing.comments || '';

            instrumentsContainer.innerHTML = '';
            if (existing.instruments && existing.instruments.length > 0) {
                existing.instruments.forEach(i => createInstrumentRow(i));
            } else if (existing.certificate && existing.certName && (existing.certName.toLowerCase().endsWith('.xlsx') || existing.certName.toLowerCase().endsWith('.xls'))) {
                extractInstrumentsFromExcel(existing.certificate).then(extracted => {
                    if (extracted && extracted.length > 0) {
                        instrumentsContainer.innerHTML = '';
                        extracted.forEach(inst => createInstrumentRow(inst));
                    }
                }).catch(err => console.error("Error al extraer instrumentos iniciales:", err));
            }

            // Estado de Valoración
            const evalContainer = document.getElementById('evaluationStatusContainer');
            evalContainer.innerHTML = '';
            EVALUATION_SCHEMA.forEach(item => {
                const cur = (existing.evaluations || {})[item.label] || '';
                const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
                rowEl.innerHTML = `
                    <div class="inspection-label">${item.label}</div>
                    <div class="inspection-options" data-label="${item.label}" data-type="evaluation">
                        <div class="inspection-opt ${cur === 'SI' ? 'selected' : ''}" data-val="SI">SI</div>
                        <div class="inspection-opt ${cur === 'NA' ? 'selected' : ''}" data-val="NA">NA</div>
                    </div>`;
                rowEl.querySelectorAll('.inspection-opt').forEach(o => o.onclick = () => {
                    rowEl.querySelectorAll('.inspection-opt').forEach(x => x.classList.remove('selected'));
                    o.classList.add('selected');
                });
                evalContainer.appendChild(rowEl);
            });

            certStatus.textContent = existing.certName ? `Certificado: ${existing.certName}` : 'Sin certificado';
            renderInspectionPoints(existing.inspections || {});
            editModal.classList.remove('hidden');
        }

        // Resetear equipo
        document.getElementById('resetCalibBtn').onclick = async () => {
            if (!selectedSerieForEdit) return;
            if (!confirm(`¿Estás seguro de resetear todos los datos de calibración para "${selectedSerieForEdit}"? Esta acción no se puede deshacer.`)) return;
            try {
                const tx = db.transaction('calibrations', 'readwrite');
                tx.objectStore('calibrations').delete(selectedSerieForEdit);
                tx.oncomplete = () => {
                    editModal.classList.add('hidden');
                    renderTable();
                    alert('Equipo reseteado correctamente.');
                };
            } catch (err) {
                console.error(err);
                alert('Error al resetear: ' + err.message);
            }
        };

        saveCalibBtn.onclick = async () => {
            if (!selectedSerieForEdit) return;
            if (!calibDateInput.value) { alert('Fecha requerida'); return; }
            const inspections = getInspectionsData();
            const evaluations = getEvaluationsData();
            const instruments = getInstrumentsData();
            const selectedTmplId = templateSelector.value;
            const tmpl = savedTemplates.find(t => String(t.id) === String(selectedTmplId));
            let blob = certFileInput.files[0] || (tmpl ? tmpl.blob : null);

            try {
                let finalCert = blob;
                if (blob && (blob.name?.endsWith('.xlsx') || blob.name?.endsWith('.xls'))) {
                    finalCert = await updateExcelCertificate(blob, {
                        editedName: equipmentNameInput.value,
                        editedSerie: modalSerieInput.value,
                        model: modelInput.value,
                        brand: brandInput.value,
                        building: buildingInput.value,
                        sector: sectorInput.value,
                        location: locationInput.value,
                        date: calibDateInput.value,
                        ordenM: ordenMInput.value,
                        technician: technicianInput.value,
                        instruments,
                        inspections,
                        evaluations,
                    });
                }

                await storeCalibration({
                    serie: selectedSerieForEdit,
                    date: calibDateInput.value,
                    technician: technicianInput.value,
                    ordenM: ordenMInput.value,
                    building: buildingInput.value,
                    sector: sectorInput.value,
                    location: locationInput.value,
                    brand: brandInput.value,
                    model: modelInput.value,
                    comments: commentsInput.value,
                    editedName: equipmentNameInput.value,
                    editedSerie: modalSerieInput.value,
                    instruments,
                    inspections,
                    evaluations,
                    certificate: finalCert,
                    certName: finalCert?.name,
                });
                editModal.classList.add('hidden');
                renderTable();
                alert('✅ Calibración guardada exitosamente.');
            } catch (err) {
                console.error(err);
                alert('Error al guardar: ' + err.message);
            }
        };
    }

    // === EXCEL CERTIFICATE UPDATE ===
    async function updateExcelCertificate(blob, d) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await blob.arrayBuffer());
        const ws = wb.getWorksheet('Certificado') || wb.worksheets[0];
        if (!ws) throw new Error('Hoja "Certificado" no encontrada');

        // ── Cabecera ──
        ws.getCell('A5').value = `Equipo: ${d.editedName}`;
        ws.getCell('E5').value = `Modelo: ${d.model}`;
        ws.getCell('A8').value = `N° serie: ${d.editedSerie}`;
        ws.getCell('E8').value = `Marca: ${d.brand}`;
        ws.getCell('H5').value = d.building;
        ws.getCell('H6').value = d.sector;
        ws.getCell('H7').value = d.location;
        ws.getCell('H8').value = d.date ? formatDate(d.date) : '';
        ws.getCell('H9').value = d.ordenM;
        ws.getCell('H10').value = d.technician;

        // ── Instrumental (fila 12, col A=nombre B=marca D=modelo E=serie F=fecha) ──
        (d.instruments || []).forEach((inst, i) => {
            if (i >= 4) return;
            const r = 12 + i;
            ws.getCell(`A${r}`).value = inst.name || '';
            ws.getCell(`B${r}`).value = inst.brand || '';
            ws.getCell(`D${r}`).value = inst.model || '';
            ws.getCell(`E${r}`).value = inst.serie || '';
            ws.getCell(`F${r}`).value = inst.date ? formatDate(inst.date) : '';
        });

        // ── Estado de Valoración (col H, filas 15-18) ──
        EVALUATION_SCHEMA.forEach(item => {
            const val = (d.evaluations || {})[item.label] || '';
            ws.getCell(`H${item.row}`).value = val === 'SI' ? 'x' : (val || 'NA');
        });

        // ── 8.1 Test de Inspección — resultado en col H, filas 24-34 ──
        SCHEMA_81.forEach(item => {
            const val = (d.inspections || {})[item.label] || '';
            const colTarget = (val === 'P') ? 'H' : (val === 'F') ? 'I' : 'H';
            if (val === 'P') {
                ws.getCell(`H${item.row}`).value = 'Pasó';
            } else if (val === 'F') {
                const cell = ws.getCell(`I${item.row}`);
                cell.value = 'Falló';
                cell.font = { color: { argb: 'FFFF0000' }, bold: true };
            } else if (val === 'na') {
                ws.getCell(`H${item.row}`).value = 'N/A';
            }
        });

        // ── 8.2.1 Características Patrón — col I, filas 38-39 (valores fijos) ──
        SCHEMA_821.forEach(item => {
            ws.getCell(`I${item.row}`).value = item.value;
        });

        // ── 8.2.2 Características Calibrar — col I, filas 41-42 (valores fijos) ──
        SCHEMA_822.forEach(item => {
            ws.getCell(`I${item.row}`).value = item.value;
        });

        // ── 8.2.3 Mediciones — H44-46 (Patrón), I44-46 (Termómetro a calibrar) ──
        const readings = (d.inspections || {})['_readings'] || {};
        const r = READINGS_START_ROW; // 44

        const pt1 = safeParseFloat(readings.pt1);
        const pt2 = safeParseFloat(readings.pt2);
        const pt3 = safeParseFloat(readings.pt3);
        const tc1 = safeParseFloat(readings.tc1);
        const tc2 = safeParseFloat(readings.tc2);
        const tc3 = safeParseFloat(readings.tc3);

        if (pt1 !== null) ws.getCell(`H${r}`).value = pt1;
        if (pt2 !== null) ws.getCell(`H${r + 1}`).value = pt2;
        if (pt3 !== null) ws.getCell(`H${r + 2}`).value = pt3;
        if (tc1 !== null) ws.getCell(`I${r}`).value = tc1;
        if (tc2 !== null) ws.getCell(`I${r + 1}`).value = tc2;
        if (tc3 !== null) ws.getCell(`I${r + 2}`).value = tc3;

        // ── 8.2.4 Calibración — Fórmulas estadísticas ──
        // 8.2.4.1 Uc (Incertidumbre combinada expandida)
        ws.getCell('I48').value = { formula: '2*SQRT(_xlfn.STDEV.S(H44:H46)^2+_xlfn.STDEV.S(I44:I46)^2+0.084*I39^2+0.084*I41^2+0.25*I38^2)' };
        // 8.2.4.2 ES (Error sistemático)
        ws.getCell('I49').value = { formula: 'ABS(AVERAGE(H44:H46)-AVERAGE(I44:I46))' };
        // 8.2.4.3 ETtx (Error total en la medida)
        ws.getCell('I50').value = { formula: 'ABS(I48+I49)' };

        const out = await wb.xlsx.writeBuffer();
        return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }

    function formatDate(s) {
        if (!s) return '';
        const d = new Date(s + 'T12:00:00');
        if (isNaN(d.getTime())) return s;
        return d.toLocaleDateString('es-ES');
    }

    function formatDateForInput(s) {
        if (!s) return '';
        if (s instanceof Date) return s.toISOString().split('T')[0];
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
        const d = new Date(s);
        if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];
        return '';
    }

    async function extractInstrumentsFromExcel(blob) {
        try {
            const arrayBuffer = await blob.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const worksheet = workbook.getWorksheet('Certificado') || workbook.worksheets[0];
            if (!worksheet) return [];

            const instruments = [];
            const startRow = 12;
            const maxRows = 4;

            for (let i = 0; i < maxRows; i++) {
                const rowIdx = startRow + i;
                const name = worksheet.getCell(`A${rowIdx}`).value;
                if (!name || (typeof name === 'string' && name.trim() === '')) break;
                const lowerName = String(name).toLowerCase().trim();
                if (lowerName.includes('estado de') || lowerName.includes('comentarios')) break;
                if (lowerName.includes('instrumental') || lowerName.includes('patrón')) continue;

                const instName = String(name).trim();
                const brand = String(worksheet.getCell(`B${rowIdx}`).value || '').trim();
                const model = String(worksheet.getCell(`D${rowIdx}`).value || '').trim();
                const serie = String(worksheet.getCell(`E${rowIdx}`).value || '').trim();
                const dateCell = worksheet.getCell(`F${rowIdx}`).value;
                let dateStr = '';
                if (dateCell instanceof Date) {
                    dateStr = dateCell.toISOString().split('T')[0];
                } else {
                    dateStr = String(dateCell || '').trim();
                }

                if (instName.length > 0 && !instName.startsWith('N/A')) {
                    instruments.push({ name: instName, brand, model, serie, date: dateStr });
                }
            }
            return instruments;
        } catch (err) {
            console.error("Error en extractInstrumentsFromExcel:", err);
            return [];
        }
    }

    init();
});
