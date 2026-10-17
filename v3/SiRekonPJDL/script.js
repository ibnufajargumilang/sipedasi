// ==========================================================================
// SiRekon PJDL — V3 Executive Engine Script
// ==========================================================================

// --- Theme Management ---
const themeToggle = document.getElementById('themeToggle');
const themeIcon = document.getElementById('themeIcon');
const html = document.documentElement;

function updateThemeUI() {
    const isDark = html.classList.contains('dark');
    if (themeIcon) {
        if (isDark) {
            themeIcon.className = 'fa-solid fa-sun text-amber-300 text-sm';
            themeToggle?.setAttribute('title', 'Ganti ke Mode Terang');
        } else {
            themeIcon.className = 'fa-solid fa-moon text-indigo-600 text-sm';
            themeToggle?.setAttribute('title', 'Ganti ke Mode Gelap');
        }
    }
}

function toggleTheme() {
    const isDark = html.classList.contains('dark');
    if (isDark) {
        html.classList.remove('dark');
        localStorage.theme = 'light';
    } else {
        html.classList.add('dark');
        localStorage.theme = 'dark';
    }
    updateThemeUI();
}

if (themeToggle) {
    themeToggle.addEventListener('click', toggleTheme);
}
updateThemeUI();

// --- Live Clock & Copyright ---
function updateRealtimeClock() {
    const clockEl = document.getElementById('liveClock');
    if (clockEl) {
        const now = new Date();
        const timeString = now.toLocaleTimeString('id-ID', {
            timeZone: 'Asia/Jakarta',
            hour12: false,
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        });
        clockEl.textContent = `${timeString} WIB`;
    }
}
setInterval(updateRealtimeClock, 1000);
updateRealtimeClock();

const copyrightEl = document.getElementById('copyrightYear');
if (copyrightEl) copyrightEl.textContent = new Date().getFullYear();

// --- UI Elements ---
const btnOpenUpload = document.getElementById('btnOpenUpload');
const fabUpload = document.getElementById('fabUpload');
const uploadModal = document.getElementById('uploadModal');
const closeModal = document.getElementById('closeModal');
const btnCancelModal = document.getElementById('btnCancelModal');
const savedDataModal = document.getElementById('savedDataModal');
const closeSavedModal = document.getElementById('closeSavedModal');
const btnSaveResult = document.getElementById('btnSaveResult');
const btnDeleteSelected = document.getElementById('btnDeleteSelected');
const openSavedDataBtn = document.getElementById('openSavedDataBtn');

// --- Global State ---
const STORAGE_KEY = 'sipedasi:pjdl_saved_data';
let lastFilteredDataPJDL = [];
let excelData = [], headers = [], isNewFormat = false, currentCols = {}, unfilteredData = [];
let masterData = []; // Immutable backup for restoration
let currentRenderData = []; // [headers, ...rows]
let currentPage = 1;
const rowsPerPage = 50;
let selectedIds = new Set(); // Stores _uid of selected rows
let tableSearchQuery = '';

// --- Column Mappings ---
const oldFormatCols = { 'Jenis Pajak': 1, 'Uraian Pajak': 3, 'Pokok (Rp)': 13, 'Denda (Rp)': 14, 'Pokok Opsen (Rp)': 15, 'Denda Opsen (Rp)': 16, 'Jumlah Bayar (Rp)': 17, 'Tanggal Bayar': 20, 'Tempat Bayar': 22, 'Keterangan': 23 };
const newFormatCols = { 'Jenis Pajak': 1, 'Uraian Pajak': 3, 'Pokok (Rp)': 15, 'Denda (Rp)': 16, 'Pokok Opsen (Rp)': 17, 'Denda Opsen (Rp)': 18, 'Jumlah Bayar (Rp)': 19, 'Tanggal Bayar': 22, 'Tempat Bayar': 24, 'Keterangan': 25 };

// --- Category Logic ---
const kategoriMap = {
    'PJDLU': ['Bank BJB - Internet Banking', 'Bank BJB - Mobile Banking', 'Bank BJB - TELER', 'Bank BJB - Lainnya', 'Bank BJB - POS', 'Bank BJB - ATM'],
    'PJDLQRIS': ['Bank BJB - QRIS'],
    'PJDLVA': ['Bank BJB - Virtual Account'],
    'POS': ['PT. POS']
};
let tomSelectInstance = null;

// --- Utilities ---
const num = (v) => parseFloat(String(v || '0').replace(/[^0-9.-]/g, '')) || 0;
const fmt = (v) => (v || 0).toLocaleString('id-ID');
function normalizeJenis(j) { 
    if (typeof j !== 'string') j = String(j || '').trim(); 
    return j.replace(/\s*\(NonReg\)$/i, '').trim(); 
}
function generateUID() { 
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 6); 
}

// Fallback Safe Copy
function copyToClipboard(text, successMsg = 'Teks berhasil disalin!') {
    if (!text) return;
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(() => {
            showToast(successMsg, 'success');
        }).catch(() => fallbackCopy(text, successMsg));
    } else {
        fallbackCopy(text, successMsg);
    }
}

function fallbackCopy(text, successMsg) {
    try {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.left = "-999999px";
        textArea.style.top = "-999999px";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        showToast(successMsg, 'success');
    } catch (err) {
        showToast('Gagal menyalin teks.', 'error');
    }
}

function copyText(text) {
    copyToClipboard(text, `Uraian "${text}" disalin!`);
}

function copyTotal(id) {
    const el = document.getElementById(id);
    if (!el) return;
    const val = el.textContent.replace(/\./g, '').replace(/,/g, '.');
    copyToClipboard(val, `Nominal Rp ${el.textContent} disalin!`);
}

// --- SweetAlert Theme Helper ---
function swalTheme() {
    const isDark = document.documentElement.classList.contains('dark');
    return {
        background: isDark ? '#0f172a' : '#ffffff',
        color: isDark ? '#f8fafc' : '#0f172a'
    };
}

// --- Modal Helper ---
function toggleModal(modal, show) {
    if (!modal) return;
    if (show) {
        modal.classList.remove('hidden');
        requestAnimationFrame(() => {
            modal.classList.remove('opacity-0');
            if (modal.children[0]) {
                modal.children[0].classList.remove('scale-95');
                modal.children[0].classList.add('scale-100');
            }
        });
    } else {
        modal.classList.add('opacity-0');
        if (modal.children[0]) {
            modal.children[0].classList.remove('scale-100');
            modal.children[0].classList.add('scale-95');
        }
        setTimeout(() => modal.classList.add('hidden'), 200);
    }
}

// Modal Listeners
if (btnOpenUpload) btnOpenUpload.addEventListener('click', () => toggleModal(uploadModal, true));
if (fabUpload) fabUpload.addEventListener('click', () => toggleModal(uploadModal, true));
if (closeModal) closeModal.addEventListener('click', () => toggleModal(uploadModal, false));
if (btnCancelModal) btnCancelModal.addEventListener('click', () => toggleModal(uploadModal, false));
if (uploadModal) {
    uploadModal.addEventListener('click', (e) => { 
        if (e.target === uploadModal) toggleModal(uploadModal, false); 
    });
}

// --- Multi-File Upload Logic ---
const inputExcel = document.getElementById('input-excel');
if (inputExcel) {
    inputExcel.addEventListener('change', async function (e) {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        const labelDiv = document.getElementById('uploadLabel');
        if (labelDiv) {
            labelDiv.innerHTML = `
                <div class="w-10 h-10 border-3 border-brand-500 border-t-transparent rounded-full animate-spin"></div>
                <div>
                    <p class="text-sm font-bold text-slate-800 dark:text-white truncate max-w-[240px]">Memproses ${files.length} file...</p>
                    <p class="text-xs text-brand-600 font-medium animate-pulse mt-1">Menggabungkan data & ekstraksi baris...</p>
                </div>
            `;
        }

        // Reset Working Sets
        excelData = [];
        unfilteredData = [];
        headers = [];
        selectedIds.clear();

        let failedCount = 0;

        const readExcel = (file) => {
            return new Promise((resolve, reject) => {
                const r = new FileReader();
                r.onload = (ev) => {
                    try {
                        const data = new Uint8Array(ev.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        const sheet = workbook.Sheets[workbook.SheetNames[0]];
                        const rawJson = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                        if (!rawJson || rawJson.length < 2) {
                            resolve([]);
                            return;
                        }

                        if (headers.length === 0) {
                            headers = rawJson[0] || [];
                            isNewFormat = headers.includes('Kecamatan OP') && headers.includes('Kelurahan OP');
                            currentCols = isNewFormat ? newFormatCols : oldFormatCols;
                        }

                        let rows = rawJson.slice(1);
                        // Clean empty rows at top & bottom
                        while (rows.length > 0 && (!rows[0][0] || String(rows[0][0]).trim() === '' || isNaN(parseInt(String(rows[0][0]).trim())))) rows.shift();
                        while (rows.length > 0 && (!rows[rows.length - 1][0] || String(rows[rows.length - 1][0]).trim() === '' || isNaN(parseInt(String(rows[rows.length - 1][0]).trim())))) rows.pop();

                        resolve(rows);
                    } catch (err) {
                        console.error(`Error parsing ${file.name}:`, err);
                        reject(err);
                    }
                };
                r.onerror = (err) => reject(err);
                r.readAsArrayBuffer(file);
            });
        };

        try {
            const results = await Promise.all(files.map(file => readExcel(file).catch(() => { failedCount++; return []; })));

            let combinedRows = [];
            results.forEach(rows => {
                if (rows && rows.length > 0) combinedRows = combinedRows.concat(rows);
            });

            if (combinedRows.length === 0) {
                showToast('Tidak ada data valid yang ditemukan pada file.', 'error');
                setTimeout(() => toggleModal(uploadModal, false), 1200);
                return;
            }

            // Assign UIDs & Process Blank Tempat Bayar
            const colTempat = currentCols['Tempat Bayar'];
            combinedRows.forEach(row => {
                Object.defineProperty(row, '_uid', { value: generateUID(), writable: true, enumerable: false });

                const rawTempat = row[colTempat];
                if (!rawTempat || String(rawTempat).trim() === '') {
                    row[colTempat] = 'BLANK';
                }
            });

            const sorted = sortByTanggalWaktu(combinedRows);
            masterData = sorted;
            unfilteredData = [...masterData];
            excelData = [headers].concat(unfilteredData);

            initTomSelect();
            renderJenisPajakOptions(); 
            renderTanggalOptions(); 
            renderUraianPajakOptions(); 
            filterAndRender();

            const successMsg = failedCount > 0
                ? `Berhasil memproses ${files.length - failedCount} file (${failedCount} gagal).`
                : `${files.length} file Excel berhasil dimuat! (${unfilteredData.length} baris)`;

            showToast(successMsg, 'success');
            toggleInputs(true);

            if (labelDiv) {
                labelDiv.innerHTML = `
                    <div class="w-14 h-14 rounded-2xl bg-emerald-500/10 text-emerald-600 flex items-center justify-center text-2xl mb-1">
                        <i class="fa-solid fa-check"></i>
                    </div>
                    <div>
                        <p class="text-sm font-bold text-slate-800 dark:text-white">File Berhasil Dimuat!</p>
                        <p class="text-xs text-emerald-600 font-bold mt-0.5">${unfilteredData.length} baris transaksi aktif</p>
                    </div>
                `;
            }

            setTimeout(() => toggleModal(uploadModal, false), 1200);

        } catch (err) {
            console.error("Global Error:", err);
            showToast('Terjadi kesalahan saat memproses data file.', 'error');
            setTimeout(() => toggleModal(uploadModal, false), 1000);
        }
    });
}

// --- TomSelect Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    initTomSelect();
    updateSavedDataCounter();
});

function initTomSelect() {
    if (tomSelectInstance) {
        tomSelectInstance.destroy();
        tomSelectInstance = null;
    }

    const uniquePlaces = unfilteredData && unfilteredData.length > 0
        ? [...new Set(unfilteredData.map(r => String(r[currentCols['Tempat Bayar']] || '').trim()))].filter(Boolean).sort()
        : [];

    const selectEl = document.getElementById('detail-tempat-select');
    if (!selectEl) return;

    tomSelectInstance = new TomSelect(selectEl, {
        plugins: ['remove_button', 'dropdown_input'],
        create: false,
        placeholder: "Pilih detail tempat bayar...",
        onItemAdd: handlePlaceChange,
        onItemRemove: handlePlaceChange,
        onClear: handlePlaceChange
    });

    tomSelectInstance.clearOptions();
    uniquePlaces.forEach(p => tomSelectInstance.addOption({ value: p, text: p }));

    if (uniquePlaces.length > 0) {
        tomSelectInstance.enable();
    } else {
        tomSelectInstance.disable();
    }
}

function handlePlaceChange() {
    renderTanggalOptions();
    renderUraianPajakOptions();
    debouncedFilterAndRender();
}

const tempatBayarSelect = document.getElementById('tempatbayar-select');
if (tempatBayarSelect) {
    tempatBayarSelect.addEventListener('change', function (e) {
        if (!tomSelectInstance) return;
        const cat = e.target.value;
        const targets = kategoriMap[cat] || [];
        if (targets.length > 0) {
            tomSelectInstance.setValue(targets);
            showToast(`Kategori ${cat} diterapkan.`);
        } else {
            tomSelectInstance.clear();
        }
    });
}

// --- Time String Extraction ---
function extractTimeString(raw) {
    if (!raw) return '';
    const s = String(raw).trim();
    const m = s.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
    if (!m) return '';
    const hh = String(m[1]).padStart(2, '0');
    const mm = String(m[2]).padStart(2, '0');
    return `${hh}:${mm}`;
}

// --- Date Normalization ---
function extractStandardDate(raw) {
    if (!raw) return '';
    const s = String(raw).trim();
    const tdate = s.substr(0, 10);
    if (/^\d{8}$/.test(tdate)) {
        return `${tdate.slice(4)}-${tdate.slice(2, 4)}-${tdate.slice(0, 2)}`;
    } else if (/^\d{4}-\d{2}-\d{2}$/.test(tdate)) {
        return tdate;
    } else if (/^\d{2}-\d{2}-\d{4}$/.test(tdate)) {
        const [d, m, y] = tdate.split('-');
        return `${y}-${m}-${d}`;
    }
    return '';
}

// --- Filtering Logic ---
function getFilteredData(ignoreWaktu = false) {
    const jenis = document.getElementById('jenispajak-select')?.value || '';
    const mblbOpsen = document.querySelector('input[name="mblbOpsenRadio"]:checked')?.value || '';
    const uraian = document.getElementById('uraianspajak-select')?.value || '';
    const waktu = document.getElementById('waktu-select')?.value || '';
    const tanggalSel = document.getElementById('tanggal-select')?.value || '';
    const isRange = document.querySelector('input[name="tanggalMode"]:checked')?.value === 'range';
    const fromDate = document.getElementById('tanggal-from')?.value || '';
    const toDate = document.getElementById('tanggal-to')?.value || '';

    let selectedDetails = [];
    if (tomSelectInstance) selectedDetails = tomSelectInstance.getValue();

    const colPokokOpsen = currentCols['Pokok Opsen (Rp)'];
    const colDendaOpsen = currentCols['Denda Opsen (Rp)'];

    return unfilteredData.filter(row => {
        // Jenis Pajak Filter (Menggabungkan reguler & NonReg menjadi 1 jenis)
        if (jenis) {
            const rowJenis = normalizeJenis(row[currentCols['Jenis Pajak']]);
            if (rowJenis !== jenis) return false;
        }

        // Kategori Opsen MBLB Filter (Bukan Opsen vs Khusus Opsen)
        if (mblbOpsen && colPokokOpsen !== undefined) {
            const pokokOpsenVal = num(row[colPokokOpsen]);
            const dendaOpsenVal = colDendaOpsen !== undefined ? num(row[colDendaOpsen]) : 0;
            const hasRowOpsen = pokokOpsenVal > 0 || dendaOpsenVal > 0;

            if (mblbOpsen === 'non_opsen' && hasRowOpsen) return false;
            if (mblbOpsen === 'opsen' && !hasRowOpsen) return false;
        }

        // Tempat Bayar Filter
        if (selectedDetails.length > 0) {
            const rawTempat = String(row[currentCols['Tempat Bayar']] || '').trim();
            if (!selectedDetails.includes(rawTempat)) return false;
        }

        // Uraian Pajak Filter
        if (uraian) {
            const u = String(row[currentCols['Uraian Pajak']] || '').trim();
            if (u !== uraian) return false;
        }

        // Tanggal Filter
        const rawDate = row[currentCols['Tanggal Bayar']];
        const stdDate = extractStandardDate(rawDate);

        if (isRange) {
            if (fromDate && toDate) {
                if (!stdDate || stdDate < fromDate || stdDate > toDate) return false;
            } else if (fromDate && (!stdDate || stdDate < fromDate)) {
                return false;
            } else if (toDate && (!stdDate || stdDate > toDate)) {
                return false;
            }
        } else if (tanggalSel) {
            if (stdDate) {
                const [y, m, d] = stdDate.split('-');
                const currentDMY = `${d}-${m}-${y}`;
                if (currentDMY !== tanggalSel) return false;
            } else {
                return false;
            }
        }

        // Waktu Cut-Off Filter (< 17:00 vs >= 17:00)
        if (!ignoreWaktu && waktu) {
            const hm = extractTimeString(rawDate);
            if (!hm) return false;
            return waktu === 'akhir' ? hm < '17:00' : hm >= '17:00';
        }

        return true;
    });
}

function filterAndRender() {
    const waktu = document.getElementById('waktu-select')?.value || '';
    const mode = document.querySelector('input[name="fitur"]:checked')?.value || 'tambah';
    let tambahan = Math.max(0, parseInt(document.getElementById('tambahanCount')?.value) || 0);
    const badge = document.getElementById('shiftAdjusterBadge');

    let filtered = [];

    // Ambil SEMUA data yang sesuai kriteria (Jenis Pajak, Tempat Bayar, Tanggal, MBLB)
    const baseData = getFilteredData(true);
    const dataAkhir = sortByTanggalWaktu(baseData.filter(row => {
        const hm = extractTimeString(row[currentCols['Tanggal Bayar']]);
        return hm && hm < '17:00';
    }));
    const dataPelimpahan = sortByTanggalWaktu(baseData.filter(row => {
        const hm = extractTimeString(row[currentCols['Tanggal Bayar']]);
        return hm && hm >= '17:00';
    }));

    if (waktu === 'akhir') {
        // Mode: Akhir Pembayaran (< 17:00)
        if (mode === 'tambah') {
            // Hanya bisa menambah dari sisa transaksi pelimpahan (>= 17:00) pada jenis & tanggal yang sama
            const maxAdd = dataPelimpahan.length;
            const actualAdd = Math.min(tambahan, maxAdd);
            const countEl = document.getElementById('tambahanCount');
            if (countEl && tambahan > maxAdd) countEl.value = maxAdd;
            tambahan = actualAdd;

            const addedFromPelimpahan = dataPelimpahan.slice(0, actualAdd);
            filtered = [...dataAkhir, ...addedFromPelimpahan];
        } else {
            // Kurangi dari transaksi akhir pembayaran
            const actualCut = Math.min(tambahan, dataAkhir.length);
            filtered = dataAkhir.slice(0, Math.max(0, dataAkhir.length - actualCut));
        }
    } else if (waktu === 'pelimpahan') {
        // Mode: Tgl Pelimpahan (>= 17:00)
        if (mode === 'tambah') {
            // Hanya bisa menambah dari transaksi akhir pembayaran (< 17:00) pada jenis & tanggal yang sama
            const maxAdd = dataAkhir.length;
            const actualAdd = Math.min(tambahan, maxAdd);
            const countEl = document.getElementById('tambahanCount');
            if (countEl && tambahan > maxAdd) countEl.value = maxAdd;
            tambahan = actualAdd;

            const addedFromAkhir = dataAkhir.slice(Math.max(0, dataAkhir.length - actualAdd));
            filtered = [...addedFromAkhir, ...dataPelimpahan];
        } else {
            // Kurangi dari transaksi pelimpahan
            const actualCut = Math.min(tambahan, dataPelimpahan.length);
            filtered = dataPelimpahan.slice(actualCut);
        }
    } else {
        // Mode: Semua Waktu Transaksi (Semua data pada filter tanggal & jenis tersebut)
        if (mode === 'kurangi') {
            const actualCut = Math.min(tambahan, baseData.length);
            filtered = baseData.slice(0, Math.max(0, baseData.length - actualCut));
        } else {
            // Pada mode Semua Transaksi, seluruh data jenis & tanggal tersebut sudah tampil penuh (tidak ada data luar yang boleh masuk)
            tambahan = 0;
            const countEl = document.getElementById('tambahanCount');
            if (countEl && parseInt(countEl.value) > 0) countEl.value = 0;
            filtered = baseData;
        }
    }

    if (tambahan > 0) {
        if (badge) {
            badge.classList.remove('hidden');
            badge.textContent = mode === 'tambah' ? `+${tambahan} Baris` : `-${tambahan} Baris`;
            badge.className = mode === 'tambah' 
                ? "text-[11px] font-bold px-2.5 py-1 rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20"
                : "text-[11px] font-bold px-2.5 py-1 rounded-lg bg-rose-500/10 text-rose-600 dark:text-rose-400 border border-rose-500/20";
        }
    } else {
        if (badge) badge.classList.add('hidden');
    }

    filtered = sortByTanggalWaktu(filtered);
    lastFilteredDataPJDL = filtered.slice();

    // Apply quick in-table search if query exists
    let displayRows = filtered;
    if (tableSearchQuery) {
        const q = tableSearchQuery.toLowerCase();
        displayRows = filtered.filter(r => {
            return r.some(cell => String(cell || '').toLowerCase().includes(q));
        });
    }

    currentRenderData = [headers].concat(displayRows);
    currentPage = 1;
    renderPage();

    updateSummaryCards(filtered);
    const jumlahDataEl = document.getElementById('jumlahData');
    if (jumlahDataEl) {
        jumlahDataEl.textContent = tableSearchQuery 
            ? `Data: ${displayRows.length} / ${filtered.length}` 
            : `Data: ${filtered.length}`;
    }

    // Update Action Button States
    const hasData = filtered.length > 0;
    const exportBtn = document.getElementById('exportPjdlBtn');
    if (exportBtn) exportBtn.disabled = !hasData;
    if (btnSaveResult) btnSaveResult.disabled = !hasData;

    updateDeleteButton();
}

function handleTableSearch(val) {
    tableSearchQuery = (val || '').trim();
    filterAndRender();
}

function updateDeleteButton() {
    if (btnDeleteSelected) {
        if (selectedIds.size > 0) {
            btnDeleteSelected.disabled = false;
            btnDeleteSelected.classList.remove('hidden');
            btnDeleteSelected.classList.add('flex');
            btnDeleteSelected.innerHTML = `<i class="fa-solid fa-trash text-xs"></i> Hapus (${selectedIds.size})`;
        } else {
            btnDeleteSelected.disabled = true;
            btnDeleteSelected.classList.add('hidden');
            btnDeleteSelected.classList.remove('flex');
        }
    }
}

// --- Table Renderer ---
function renderPage() {
    const dataRows = currentRenderData.slice(1);
    const container = document.getElementById('excel-table');
    const emptyState = document.getElementById('empty-state');
    const controls = document.getElementById('paginationControls');

    if (!container) return;

    if (dataRows.length === 0) {
        container.innerHTML = '';
        if (emptyState) emptyState.classList.remove('hidden');
        if (controls) controls.classList.add('hidden');
        return;
    } else {
        if (emptyState) emptyState.classList.add('hidden');
        if (controls) controls.classList.remove('hidden');
    }

    const totalPages = Math.ceil(dataRows.length / rowsPerPage) || 1;
    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const pageData = dataRows.slice(startIdx, endIdx);

    const curPageEl = document.getElementById("currentPageDisplay");
    const totPageEl = document.getElementById("totalPagesDisplay");
    const prevBtn = document.getElementById("prevPageBtn");
    const nextBtn = document.getElementById("nextPageBtn");

    if (curPageEl) curPageEl.textContent = currentPage;
    if (totPageEl) totPageEl.textContent = totalPages;
    if (prevBtn) prevBtn.disabled = currentPage === 1;
    if (nextBtn) nextBtn.disabled = currentPage === totalPages;

    let displayHeaders = ['Jenis Pajak', 'Uraian Pajak', 'Pokok (Rp)', 'Denda (Rp)', 'Jumlah Bayar (Rp)', 'Tanggal Bayar', 'Tempat Bayar'];
    const jenisSel = document.getElementById('jenispajak-select')?.value || '';
    const mblbOpsenMode = document.querySelector('input[name="mblbOpsenRadio"]:checked')?.value || '';
    const isMBLB = jenisSel && (jenisSel.toLowerCase().includes('mineral') || jenisSel.toLowerCase().includes('opsen'));
    
    // Tampilkan kolom Pokok Opsen & Denda Opsen jika MBLB dan bukan mode 'non_opsen'
    if (isMBLB && mblbOpsenMode !== 'non_opsen') {
        displayHeaders = ['Jenis Pajak', 'Uraian Pajak', 'Pokok (Rp)', 'Denda (Rp)', 'Pokok Opsen (Rp)', 'Denda Opsen (Rp)', 'Jumlah Bayar (Rp)', 'Tanggal Bayar', 'Tempat Bayar'];
    }

    let html = `<thead class="text-[11px] text-slate-500 dark:text-slate-400 uppercase bg-slate-100/80 dark:bg-slate-800/80 sticky top-0 z-10 backdrop-blur-md"><tr>`;
    
    // Checkbox Column
    html += `<th scope="col" class="px-4 py-3 font-extrabold border-b border-slate-200 dark:border-white/5 w-10 text-center">
                <input type="checkbox" id="selectAll" class="rounded-md border-slate-300 dark:border-slate-600 text-brand-600 focus:ring-brand-500 h-4 w-4 cursor-pointer">
             </th>`;

    html += `<th scope="col" class="px-4 py-3 font-extrabold border-b border-slate-200 dark:border-white/5 text-center w-12">No</th>`;
    
    displayHeaders.forEach(h => {
        const align = h.includes('(Rp)') ? 'text-right' : 'text-left';
        html += `<th scope="col" class="px-4 py-3 font-extrabold border-b border-slate-200 dark:border-white/5 ${align}">${h.replace('(Rp)', '')}</th>`;
    });
    html += `</tr></thead><tbody class="divide-y divide-slate-100 dark:divide-white/5">`;

    pageData.forEach((row, i) => {
        const globalIndex = startIdx + i + 1;
        const isAlt = i % 2 === 0 ? 'bg-white dark:bg-transparent' : 'bg-slate-50/50 dark:bg-white/[0.015]';
        const isChecked = selectedIds.has(row._uid) ? 'checked' : '';

        html += `<tr class="${isAlt} hover:bg-brand-50/50 dark:hover:bg-white/[0.04] transition-colors">`;

        // Row Checkbox
        html += `<td class="px-4 py-2.5 text-center">
                    <input type="checkbox" class="row-checkbox rounded-md border-slate-300 dark:border-slate-600 text-brand-600 focus:ring-brand-500 h-4 w-4 cursor-pointer" value="${row._uid}" ${isChecked}>
                 </td>`;

        html += `<td class="px-4 py-2.5 font-medium text-slate-400 text-center">${globalIndex}</td>`;

        displayHeaders.forEach(h => {
            const align = h.includes('(Rp)') ? 'text-right font-mono' : 'text-left';
            let v = row[currentCols[h]] !== undefined ? row[currentCols[h]] : '';
            let content = String(v || '');
            let colorClass = "text-slate-600 dark:text-slate-400";

            if (h.includes('(Rp)')) {
                content = fmt(num(v));
                colorClass = h.includes('Pokok') ? "text-slate-800 dark:text-slate-200 font-semibold" :
                    h.includes('Denda') ? "text-amber-600 dark:text-amber-400 font-semibold" :
                        h.includes('Jumlah') ? "text-emerald-600 dark:text-emerald-400 font-bold" : colorClass;
            } else if (h === 'Uraian Pajak') {
                const rawContent = String(v || '').trim().replace(/\s*-\s*/g, '-');
                const safeContent = rawContent.replace(/'/g, "\\'").replace(/"/g, "&quot;");
                content = `
                    <div class="flex items-center justify-between gap-2 group min-w-[200px] cursor-pointer hover:bg-brand-50 dark:hover:bg-white/5 -mx-2 px-2 py-1 rounded-lg transition-colors select-none" onclick="copyText('${safeContent}')" title="Klik untuk menyalin">
                         <span class="whitespace-normal text-slate-800 dark:text-slate-200 font-medium">${rawContent}</span>
                         <div class="w-5 h-5 rounded flex items-center justify-center text-brand-600 bg-brand-50 dark:bg-brand-900/30 opacity-0 group-hover:opacity-100 transition-all shrink-0">
                             <i class="fa-regular fa-copy text-[10px]"></i>
                         </div>
                    </div>
                `;
            } else if (h === 'Tanggal Bayar') {
                colorClass = "text-slate-500 font-mono text-xs";
            }

            html += `<td class="px-4 py-2.5 ${align} ${colorClass}">${content}</td>`;
        });
        html += `</tr>`;
    });
    html += `</tbody>`;
    container.innerHTML = html;

    const sc = document.getElementById('table-scroll-container');
    if (sc) sc.scrollTop = 0;

    // Attach Listeners for Checkboxes
    const selectAll = document.getElementById('selectAll');
    const rowCheckboxes = document.querySelectorAll('.row-checkbox');

    const allOnPageSelected = Array.from(rowCheckboxes).every(cb => cb.checked) && rowCheckboxes.length > 0;
    if (selectAll) {
        selectAll.checked = allOnPageSelected;
        selectAll.addEventListener('change', (e) => {
            const checked = e.target.checked;
            rowCheckboxes.forEach(cb => {
                cb.checked = checked;
                if (checked) selectedIds.add(cb.value);
                else selectedIds.delete(cb.value);
            });
            updateDeleteButton();
        });
    }

    rowCheckboxes.forEach(cb => {
        cb.addEventListener('change', (e) => {
            if (e.target.checked) selectedIds.add(e.target.value);
            else selectedIds.delete(e.target.value);

            if (selectAll) selectAll.checked = Array.from(rowCheckboxes).every(c => c.checked);
            updateDeleteButton();
        });
    });
}

// --- Delete Selected Rows Handling ---
if (btnDeleteSelected) {
    btnDeleteSelected.addEventListener('click', () => {
        if (selectedIds.size === 0) return;

        Swal.fire({
            title: `Hapus ${selectedIds.size} Data?`,
            text: "Data yang dihapus akan dikeluarkan dari daftar rekonsiliasi aktif.",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#e11d48',
            cancelButtonColor: '#64748b',
            confirmButtonText: 'Ya, Hapus',
            cancelButtonText: 'Batal',
            ...swalTheme()
        }).then((result) => {
            if (result.isConfirmed) {
                const initialCount = unfilteredData.length;
                unfilteredData = unfilteredData.filter(row => !selectedIds.has(row._uid));
                const deletedCount = initialCount - unfilteredData.length;

                selectedIds.clear();
                updateDeleteButton();

                excelData = [headers].concat(unfilteredData);
                renderJenisPajakOptions(); 
                renderTanggalOptions(); 
                renderUraianPajakOptions();
                filterAndRender();

                showToast(`${deletedCount} data berhasil dihapus.`, 'success');
            }
        });
    });
}

// --- Pagination Events ---
document.getElementById("prevPageBtn")?.addEventListener("click", () => {
    if (currentPage > 1) { currentPage--; renderPage(); }
});
document.getElementById("nextPageBtn")?.addEventListener("click", () => {
    const totalPages = Math.ceil((currentRenderData.length - 1) / rowsPerPage);
    if (currentPage < totalPages) { currentPage++; renderPage(); }
});

// --- Save & Saved Data Modal Handlers ---
if (openSavedDataBtn) openSavedDataBtn.addEventListener('click', () => {
    renderSavedDataList();
    toggleModal(savedDataModal, true);
});
if (closeSavedModal) closeSavedModal.addEventListener('click', () => toggleModal(savedDataModal, false));
if (savedDataModal) {
    savedDataModal.addEventListener('click', (e) => { 
        if (e.target === savedDataModal) toggleModal(savedDataModal, false); 
    });
}
if (btnSaveResult) btnSaveResult.addEventListener('click', saveFilteredData);

async function saveFilteredData() {
    if (!lastFilteredDataPJDL || lastFilteredDataPJDL.length === 0) return;

    const { value: name } = await Swal.fire({
        title: 'Simpan Hasil Rekonsiliasi',
        input: 'text',
        inputLabel: 'Beri nama untuk arsip data ini',
        inputPlaceholder: 'Contoh: Rekon PJDL Januari 2026 Shift Pagi',
        showCancelButton: true,
        confirmButtonText: 'Simpan Data',
        cancelButtonText: 'Batal',
        confirmButtonColor: '#4f46e5',
        inputValidator: (value) => {
            if (!value) return 'Nama arsip tidak boleh kosong!';
        },
        ...swalTheme()
    });

    if (name) {
        const saveData = {
            id: Date.now(),
            name: name,
            date: new Date().toISOString(),
            headers: headers,
            rows: lastFilteredDataPJDL,
            currentCols: currentCols,
            isNewFormat: isNewFormat
        };

        const existing = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
        existing.push(saveData);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));

        updateSavedDataCounter();
        showToast('Data berhasil disimpan ke arsip!', 'success');
    }
}

function updateSavedDataCounter() {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    const badge = document.getElementById('savedCountBadge');
    if (badge) {
        if (saved.length > 0) {
            badge.textContent = saved.length;
            badge.classList.remove('hidden');
        } else {
            badge.classList.add('hidden');
        }
    }
}

function renderSavedDataList() {
    const list = document.getElementById('savedDataList');
    const empty = document.getElementById('emptySaved');
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');

    if (!list) return;
    list.innerHTML = '';

    if (saved.length === 0) {
        list.classList.add('hidden');
        if (empty) empty.classList.remove('hidden');
        return;
    }

    list.classList.remove('hidden');
    if (empty) empty.classList.add('hidden');

    saved.sort((a, b) => new Date(b.date) - new Date(a.date)).forEach(item => {
        const dateObj = new Date(item.date);
        const dateStr = dateObj.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
        const timeStr = dateObj.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/70 dark:hover:bg-white/[0.02] transition-colors group";
        tr.innerHTML = `
            <td class="px-5 py-3 font-semibold text-slate-800 dark:text-white">
                ${item.name}
            </td>
            <td class="px-5 py-3 text-center text-slate-500 font-mono text-xs">
                ${dateStr} <span class="text-slate-400">(${timeStr})</span>
            </td>
            <td class="px-5 py-3 text-center">
                <span class="px-2.5 py-1 bg-brand-50 dark:bg-brand-900/30 text-brand-700 dark:text-brand-300 font-bold rounded-lg text-xs">${item.rows.length} baris</span>
            </td>
            <td class="px-5 py-3 text-right">
                <div class="flex items-center justify-end gap-1.5 opacity-90 group-hover:opacity-100 transition-opacity">
                    <button onclick="loadSavedData(${item.id})" class="px-2.5 py-1.5 bg-brand-50 hover:bg-brand-100 text-brand-700 dark:bg-brand-900/30 dark:text-brand-300 rounded-lg text-xs font-bold transition-colors flex items-center gap-1" title="Buka Data">
                        <i class="fa-solid fa-folder-open text-[11px]"></i> Buka
                    </button>
                    <button onclick="deleteSavedData(${item.id})" class="p-1.5 text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 rounded-lg transition-colors" title="Hapus">
                        <i class="fa-solid fa-trash text-xs"></i>
                    </button>
                </div>
            </td>
        `;
        list.appendChild(tr);
    });
}

window.loadSavedData = function (id) {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    const item = saved.find(x => x.id === id);
    if (!item) return;

    headers = item.headers;
    const rowsWithUid = item.rows.map(r => {
        if (!r._uid) Object.defineProperty(r, '_uid', { value: generateUID(), writable: true, enumerable: false });
        return r;
    });

    masterData = rowsWithUid;
    unfilteredData = [...masterData];
    excelData = [headers].concat(unfilteredData);
    currentCols = item.currentCols || (item.isNewFormat ? newFormatCols : oldFormatCols);
    selectedIds.clear();

    initTomSelect();

    ['jenispajak-select', 'tempatbayar-select', 'waktu-select', 'tanggal-select'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    const tambahanEl = document.getElementById('tambahanCount');
    if (tambahanEl) tambahanEl.value = 0;

    renderJenisPajakOptions();
    renderTanggalOptions();
    renderUraianPajakOptions();
    filterAndRender();

    toggleModal(savedDataModal, false);
    showToast(`Arsip "${item.name}" berhasil dimuat.`, 'success');
    toggleInputs(true);
};

window.deleteSavedData = function (id) {
    Swal.fire({
        title: 'Hapus arsip ini?',
        text: "Data yang dihapus tidak dapat dikembalikan!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#f43f5e',
        cancelButtonColor: '#64748b',
        confirmButtonText: 'Ya, Hapus',
        cancelButtonText: 'Batal',
        ...swalTheme()
    }).then((result) => {
        if (result.isConfirmed) {
            let saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
            saved = saved.filter(x => x.id !== id);
            localStorage.setItem(STORAGE_KEY, JSON.stringify(saved));
            renderSavedDataList();
            updateSavedDataCounter();
            showToast('Arsip data berhasil dihapus.', 'success');
        }
    });
};

document.getElementById('btnClearAll')?.addEventListener('click', () => {
    Swal.fire({
        title: 'Hapus SEMUA data tersimpan?',
        text: "Seluruh arsip riwayat rekonsiliasi akan dikosongkan!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#f43f5e',
        cancelButtonColor: '#64748b',
        confirmButtonText: 'Ya, Bersihkan Semua',
        cancelButtonText: 'Batal',
        ...swalTheme()
    }).then((result) => {
        if (result.isConfirmed) {
            localStorage.removeItem(STORAGE_KEY);
            renderSavedDataList();
            updateSavedDataCounter();
            showToast('Semua arsip tersimpan telah dibersihkan.', 'success');
        }
    });
});

// --- Single Sheet Merge Download ---
document.getElementById('btnDownloadAll')?.addEventListener('click', () => {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    if (saved.length === 0) { 
        showToast('Tidak ada data tersimpan untuk diunduh.', 'error'); 
        return; 
    }

    try {
        const originalHeaders = saved[0].headers || [];
        const masterHeaders = ["Nama Laporan", "Waktu Simpan", ...originalHeaders];
        let allRows = [masterHeaders];

        saved.forEach(item => {
            const savedDate = new Date(item.date);
            const timeStr = savedDate.toLocaleDateString('id-ID') + ' ' + savedDate.toLocaleTimeString('id-ID');

            const processedRows = item.rows.map(row => {
                const r = Array.isArray(row) ? row : Object.values(row);
                return [item.name, timeStr, ...r];
            });
            allRows = allRows.concat(processedRows);
        });

        const ws = XLSX.utils.aoa_to_sheet(allRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Gabungan Data PJDL");

        const now = new Date();
        const fname = `SiRekonPJDL_Gabungan_Arsip_${now.toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fname);

        showToast(`Berhasil mengekspor gabungan ${saved.length} arsip ke Excel!`, 'success');
    } catch (err) {
        console.error("Bulk Export Error:", err);
        showToast('Gagal membuat file gabungan Excel.', 'error');
    }
});

// --- Export PJDL To Excel (Current Filter) ---
function exportPJDLToExcel() {
    if (!lastFilteredDataPJDL || lastFilteredDataPJDL.length === 0) { 
        showToast('Tidak ada data aktif untuk diekspor.', 'error'); 
        return; 
    }

    try {
        const wsData = [headers].concat(lastFilteredDataPJDL);
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'PJDL Rekon');

        const now = new Date();
        const fname = `SiRekonPJDL_Export_${now.toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fname);
        showToast(`Berhasil mengekspor ${lastFilteredDataPJDL.length} baris data ke Excel.`, 'success');
    } catch (e) { 
        console.error(e); 
        showToast('Gagal mengekspor data.', 'error'); 
    }
}
document.getElementById('exportPjdlBtn')?.addEventListener('click', exportPJDLToExcel);

// --- Toast Function ---
function showToast(msg, type = 'success') {
    const toast = document.getElementById("toast");
    const icon = document.getElementById("toastIcon");
    const txt = document.getElementById("toastMsg");
    if (!toast || !txt) return;

    txt.textContent = msg;
    toast.className = "fixed bottom-6 right-6 px-5 py-3.5 rounded-2xl shadow-2xl transition-all duration-300 z-[70] transform flex items-center gap-3 font-medium text-xs sm:text-sm text-white backdrop-blur-md";
    
    if (type === 'success') { 
        toast.classList.add('bg-emerald-600/90'); 
        if (icon) icon.innerHTML = '<i class="fa-solid fa-circle-check text-base"></i>'; 
    } else { 
        toast.classList.add('bg-rose-600/90'); 
        if (icon) icon.innerHTML = '<i class="fa-solid fa-triangle-exclamation text-base"></i>'; 
    }

    requestAnimationFrame(() => toast.classList.remove('translate-y-10', 'opacity-0'));
    setTimeout(() => toast.classList.add('translate-y-10', 'opacity-0'), 3200);
}

// --- Sorting Helper ---
function sortByTanggalWaktu(data) {
    const idx = currentCols['Tanggal Bayar'];
    return data.filter(r => r && r.some(c => c !== undefined && String(c).trim() !== '')).sort((a, b) => {
        const ta = String(a[idx] || '').trim(), tb = String(b[idx] || '').trim();
        const norm = s => { 
            const [d, t] = String(s || '').split(' '); 
            if (/^\d{2}-\d{2}-\d{4}$/.test(d)) { 
                const [dd, mm, yy] = d.split('-'); 
                return `${yy}-${mm}-${dd}${t ? ' ' + t : ''}`; 
            } 
            return `${d || ''}${t ? ' ' + t : ''}`; 
        };
        return norm(ta).localeCompare(norm(tb));
    });
}

function debounce(fn, ms = 300) { 
    let t; 
    return function (...a) { 
        clearTimeout(t); 
        t = setTimeout(() => fn.apply(this, a), ms); 
    }; 
}
const debouncedFilterAndRender = debounce(filterAndRender, 300);

// Filter Input Events
['change', 'input'].forEach(evt => {
    document.getElementById('jenispajak-select')?.addEventListener('change', (e) => { 
        const jenisVal = e.target.value || '';
        const isMBLB = jenisVal.toLowerCase().includes('mineral') || jenisVal.toLowerCase().includes('opsen');
        const mblbToolbar = document.getElementById('mblb-opsen-toolbar');
        
        if (mblbToolbar) {
            if (isMBLB) {
                mblbToolbar.classList.remove('hidden');
                mblbToolbar.classList.add('inline-flex');
            } else {
                mblbToolbar.classList.add('hidden');
                mblbToolbar.classList.remove('inline-flex');
                const allRadio = document.querySelector('input[name="mblbOpsenRadio"][value=""]');
                if (allRadio) allRadio.checked = true;
            }
        }

        renderTanggalOptions(); 
        renderUraianPajakOptions(); 
        debouncedFilterAndRender(); 
    });

    document.getElementById('tanggal-select')?.addEventListener('change', () => { 
        renderUraianPajakOptions(); 
        debouncedFilterAndRender(); 
    });
    document.getElementById('waktu-select')?.addEventListener('change', () => { 
        renderTanggalOptions(); 
        renderUraianPajakOptions(); 
        debouncedFilterAndRender(); 
    });
    document.getElementById('uraianspajak-select')?.addEventListener('change', debouncedFilterAndRender);
    document.getElementById('tambahanCount')?.addEventListener('input', () => {
        filterAndRender();
    });
});

// Shift Adjuster Stepper Buttons
const countInputEl = document.getElementById('tambahanCount');
document.getElementById('btnStepUp')?.addEventListener('click', () => {
    if (!countInputEl) return;
    const waktu = document.getElementById('waktu-select')?.value || '';
    const mode = document.querySelector('input[name="fitur"]:checked')?.value || 'tambah';
    
    if (mode === 'tambah' && !waktu) {
        showToast('Pilih filter waktu (Akhir Pembayaran / Tgl Pelimpahan) untuk menambah baris pelimpahan shift.', 'error');
        return;
    }

    const cur = parseInt(countInputEl.value) || 0;
    countInputEl.value = cur + 1;
    filterAndRender();
});

document.getElementById('btnStepDown')?.addEventListener('click', () => {
    if (!countInputEl) return;
    const cur = parseInt(countInputEl.value) || 0;
    countInputEl.value = Math.max(0, cur - 1);
    filterAndRender();
});

document.querySelectorAll('input[name="fitur"]').forEach(r => {
    r.addEventListener('change', () => {
        filterAndRender();
    });
});

// MBLB Opsen Toolbar Radio Listeners
document.querySelectorAll('input[name="mblbOpsenRadio"]').forEach(r => {
    r.addEventListener('change', () => {
        debouncedFilterAndRender();
    });
});

// Date Range Mode UI
function setTanggalModeUI(mode) {
    const rangeDiv = document.getElementById('tanggal-range-inputs');
    const singleContainer = document.getElementById('tanggal-single-container');
    const fromInput = document.getElementById('tanggal-from');
    const toInput = document.getElementById('tanggal-to');
    const singleSel = document.getElementById('tanggal-select');

    if (mode === 'range') {
        rangeDiv?.classList.remove('hidden'); 
        singleContainer?.classList.add('hidden');
        if (singleSel) singleSel.value = '';
    } else {
        rangeDiv?.classList.add('hidden'); 
        singleContainer?.classList.remove('hidden');
        if (fromInput) fromInput.value = '';
        if (toInput) toInput.value = '';
    }
}

document.querySelectorAll('input[name="tanggalMode"]').forEach(r => {
    r.addEventListener('change', function () {
        setTanggalModeUI(this.value);
        debouncedFilterAndRender();
    });
});

document.getElementById('tanggal-from')?.addEventListener('change', () => { 
    if (document.querySelector('input[name="tanggalMode"]:checked')?.value === 'range') debouncedFilterAndRender(); 
});
document.getElementById('tanggal-to')?.addEventListener('change', () => { 
    if (document.querySelector('input[name="tanggalMode"]:checked')?.value === 'range') debouncedFilterAndRender(); 
});

// Reset Button Logic
document.getElementById('btnResetPJDL')?.addEventListener('click', function () {
    ['jenispajak-select', 'tempatbayar-select', 'waktu-select', 'tanggal-select', 'uraianspajak-select'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });

    const allRadio = document.querySelector('input[name="mblbOpsenRadio"][value=""]');
    if (allRadio) allRadio.checked = true;

    const mblbToolbar = document.getElementById('mblb-opsen-toolbar');
    if (mblbToolbar) {
        mblbToolbar.classList.add('hidden');
        mblbToolbar.classList.remove('inline-flex');
    }
    
    const countEl = document.getElementById('tambahanCount');
    if (countEl) countEl.value = 0;

    const shiftBadge = document.getElementById('shiftAdjusterBadge');
    if (shiftBadge) shiftBadge.classList.add('hidden');

    const tableInput = document.getElementById('tableFilterInput');
    if (tableInput) tableInput.value = '';
    tableSearchQuery = '';

    // Reset Date Mode
    const singleMode = document.querySelector('input[name="tanggalMode"][value="single"]');
    if (singleMode) { 
        singleMode.checked = true; 
        setTanggalModeUI('single'); 
    }

    // Reset Fitur
    document.querySelectorAll('input[name="fitur"]').forEach(r => r.checked = (r.value === 'tambah'));

    // Reset Tom Select
    if (tomSelectInstance) tomSelectInstance.clear();

    // RESTORE DATA FROM MASTER
    if (masterData.length > 0) {
        unfilteredData = [...masterData];
        excelData = [headers].concat(unfilteredData);
    }

    renderJenisPajakOptions(); 
    renderTanggalOptions(); 
    renderUraianPajakOptions();
    selectedIds.clear();

    filterAndRender();
    showToast('Filter dan data berhasil direset.', 'success');
});

// Dropdown Renders
function renderJenisPajakOptions() {
    const sel = document.getElementById('jenispajak-select'); 
    if (!sel) return;
    const prev = sel.value;
    sel.innerHTML = '<option value="">-- Semua Jenis Pajak --</option>';
    const s = new Set(); 
    unfilteredData.forEach(r => { 
        const v = normalizeJenis(r[currentCols['Jenis Pajak']]); 
        if (v) s.add(v); 
    });
    Array.from(s).sort().forEach(v => { 
        const o = document.createElement('option'); 
        o.value = v; 
        o.textContent = v; 
        sel.appendChild(o); 
    });
    sel.value = Array.from(sel.options).some(o => o.value === prev) ? prev : '';
}

function renderTanggalOptions() {
    const sel = document.getElementById('tanggal-select'); 
    if (!sel) return;
    const prev = sel.value; 
    sel.innerHTML = '<option value="">-- Pilih Tanggal --</option>';
    const jenis = document.getElementById('jenispajak-select')?.value || '';
    const uraian = document.getElementById('uraianspajak-select')?.value || '';
    const tanggalCounts = {}; 
    const set = new Set();

    let selectedDetails = [];
    if (tomSelectInstance) selectedDetails = tomSelectInstance.getValue();

    unfilteredData.forEach(row => {
        if (jenis && normalizeJenis(row[currentCols['Jenis Pajak']]) !== jenis) return;

        if (selectedDetails.length > 0) {
            const rawTempat = String(row[currentCols['Tempat Bayar']] || '').trim();
            if (!selectedDetails.includes(rawTempat)) return;
        }

        if (uraian) {
            const u = (typeof row[currentCols['Uraian Pajak']] === 'string') ? row[currentCols['Uraian Pajak']].trim() : String(row[currentCols['Uraian Pajak']] || '').trim();
            if (u !== uraian) return;
        }
        const t = String(row[currentCols['Tanggal Bayar']] || '').trim().substr(0, 10); 
        if (!t) return;

        let normalizedTgl = '';
        if (/^\d{8}$/.test(t)) normalizedTgl = `${t.slice(0, 2)}-${t.slice(2, 4)}-${t.slice(4)}`;
        else if (/^\d{4}-\d{2}-\d{2}$/.test(t)) { 
            const [y, m, d] = t.split('-'); 
            normalizedTgl = `${d}-${m}-${y}`; 
        }
        else if (/^\d{2}-\d{2}-\d{4}$/.test(t)) normalizedTgl = t;

        if (normalizedTgl) { 
            tanggalCounts[normalizedTgl] = (tanggalCounts[normalizedTgl] || 0) + 1; 
            set.add(normalizedTgl); 
        }
    });

    const sortedDates = Array.from(set).sort((a, b) => { 
        const ta = a.split('-').reverse().join('-'); 
        const tb = b.split('-').reverse().join('-'); 
        return ta.localeCompare(tb); 
    });

    const fromInput = document.getElementById('tanggal-from');
    const toInput = document.getElementById('tanggal-to');
    function convertToYYYYMMDD(dateStr) {
        if (/^\d{2}-\d{2}-\d{4}$/.test(dateStr)) { 
            const [d, m, y] = dateStr.split('-'); 
            return `${y}-${m}-${d}`; 
        }
        return '';
    }

    if (sortedDates.length > 0) {
        const minDate = convertToYYYYMMDD(sortedDates[0]);
        const maxDate = convertToYYYYMMDD(sortedDates[sortedDates.length - 1]);
        if (fromInput && toInput) { 
            fromInput.min = minDate; 
            fromInput.max = maxDate; 
            toInput.min = minDate; 
            toInput.max = maxDate; 
            fromInput.disabled = false; 
            toInput.disabled = false; 
        }
    } else {
        if (fromInput && toInput) { 
            fromInput.value = ''; 
            fromInput.disabled = true; 
            toInput.value = ''; 
            toInput.disabled = true; 
        }
    }

    sortedDates.forEach(t => { 
        const o = document.createElement('option'); 
        o.value = t; 
        o.textContent = `${t} (${tanggalCounts[t]} transaksi)`; 
        sel.appendChild(o); 
    });
    sel.value = Array.from(sel.options).some(o => o.value === prev) ? prev : '';
}

function renderUraianPajakOptions() {
    const sel = document.getElementById('uraianspajak-select'); 
    if (!sel) return;
    sel.innerHTML = '<option value="">-- Semua Uraian Pajak --</option>';
    const jenis = document.getElementById('jenispajak-select')?.value || '';
    const tgl = document.getElementById('tanggal-select')?.value || '';

    let arr = unfilteredData.slice();
    if (jenis) arr = arr.filter(r => normalizeJenis(r[currentCols['Jenis Pajak']]) === jenis);
    if (tgl) arr = arr.filter(r => String(r[currentCols['Tanggal Bayar']] || '').substr(0, 10) === tgl);

    if (tomSelectInstance) {
        let selectedDetails = tomSelectInstance.getValue();
        if (selectedDetails.length > 0) {
            arr = arr.filter(r => selectedDetails.includes(String(r[currentCols['Tempat Bayar']] || '').trim()));
        }
    }

    const s = new Set(); 
    arr.forEach(r => { 
        let u = r[currentCols['Uraian Pajak']]; 
        if (typeof u !== 'string') u = String(u || '').trim(); 
        if (u) s.add(u); 
    });
    Array.from(s).sort().forEach(u => { 
        const o = document.createElement('option'); 
        o.value = u; 
        o.textContent = String(u).replace(/\s*-\s*/g, '-'); 
        sel.appendChild(o); 
    });
}

function updateSummaryCards(data) {
    let totalPokok = 0, totalDenda = 0, totalJumlah = 0;
    let totalPokokOpsen = 0, totalDendaOpsen = 0, totalOpsen = 0;
    const colPokokOpsen = currentCols['Pokok Opsen (Rp)'];
    const colDendaOpsen = currentCols['Denda Opsen (Rp)'];

    data.forEach(r => {
        const p = num(r[currentCols['Pokok (Rp)']]);
        const d = num(r[currentCols['Denda (Rp)']]);
        const j = num(r[currentCols['Jumlah Bayar (Rp)']]);
        const po = colPokokOpsen !== undefined ? num(r[colPokokOpsen]) : 0;
        const doVal = colDendaOpsen !== undefined ? num(r[colDendaOpsen]) : 0;

        totalPokok += p;
        totalDenda += d;
        totalJumlah += j;
        totalPokokOpsen += po;
        totalDendaOpsen += doVal;
        totalOpsen += (po + doVal);
    });

    const mblbOpsenMode = document.querySelector('input[name="mblbOpsenRadio"]:checked')?.value || '';
    const isMBLBOpsenMode = mblbOpsenMode === 'opsen';

    const cardPokok = document.getElementById('cardPokok');
    const cardDenda = document.getElementById('cardDenda');
    const cardOpsen = document.getElementById('cardOpsen');
    const cardJumlah = document.getElementById('cardJumlah');

    const cardPokokDendaMBLB = document.getElementById('cardPokokDendaMBLB');
    const cardPokokOpsenMBLB = document.getElementById('cardPokokOpsenMBLB');
    const cardDendaOpsenMBLB = document.getElementById('cardDendaOpsenMBLB');

    const pokokEl = document.getElementById('total-pokok');
    const dendaEl = document.getElementById('total-denda');
    const jumlahEl = document.getElementById('total-jumlah');
    const opsenEl = document.getElementById('total-opsen');

    const pokokDendaMBLBEl = document.getElementById('total-pokok-denda-gabungan');
    const pokokOpsenEl = document.getElementById('total-pokok-opsen');
    const dendaOpsenEl = document.getElementById('total-denda-opsen');

    if (jumlahEl) jumlahEl.textContent = fmt(totalJumlah);

    if (isMBLBOpsenMode) {
        // Mode Khusus Opsen: 
        // 1 card Total Pokok + Denda MBLB
        // 2 card Opsen (Pokok Opsen & Denda Opsen)
        // 1 card Total Penerimaan
        if (cardPokok) cardPokok.classList.add('hidden');
        if (cardDenda) cardDenda.classList.add('hidden');
        if (cardOpsen) cardOpsen.classList.add('hidden');

        if (cardPokokDendaMBLB) cardPokokDendaMBLB.classList.remove('hidden');
        if (cardPokokOpsenMBLB) cardPokokOpsenMBLB.classList.remove('hidden');
        if (cardDendaOpsenMBLB) cardDendaOpsenMBLB.classList.remove('hidden');

        if (pokokDendaMBLBEl) pokokDendaMBLBEl.textContent = fmt(totalPokok + totalDenda);
        if (pokokOpsenEl) pokokOpsenEl.textContent = fmt(totalPokokOpsen);
        if (dendaOpsenEl) dendaOpsenEl.textContent = fmt(totalDendaOpsen);
    } else {
        // Mode Normal / Bukan Opsen / Default:
        if (cardPokokDendaMBLB) cardPokokDendaMBLB.classList.add('hidden');
        if (cardPokokOpsenMBLB) cardPokokOpsenMBLB.classList.add('hidden');
        if (cardDendaOpsenMBLB) cardDendaOpsenMBLB.classList.add('hidden');

        if (cardPokok) cardPokok.classList.remove('hidden');
        if (cardDenda) cardDenda.classList.remove('hidden');

        if (pokokEl) pokokEl.textContent = fmt(totalPokok);
        if (dendaEl) dendaEl.textContent = fmt(totalDenda);

        if (totalOpsen > 0 && cardOpsen && opsenEl) {
            opsenEl.textContent = fmt(totalOpsen);
            cardOpsen.classList.remove('hidden');
        } else if (cardOpsen) {
            cardOpsen.classList.add('hidden');
        }
    }
}

// --- Toggle Inputs State ---
function toggleInputs(enable) {
    const ids = [
        'btnResetPJDL',
        'jenispajak-select',
        'tempatbayar-select',
        'waktu-select',
        'uraianspajak-select',
        'tanggal-select',
        'tambahanCount',
        'btnStepUp',
        'btnStepDown',
        'tableFilterInput'
    ];

    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.disabled = !enable;
    });

    document.querySelectorAll('input[name="fitur"]').forEach(r => r.disabled = !enable);

    if (tomSelectInstance) {
        if (enable) tomSelectInstance.enable();
        else tomSelectInstance.disable();
    }
}
