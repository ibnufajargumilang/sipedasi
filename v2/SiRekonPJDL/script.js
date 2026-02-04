const themeToggle = document.getElementById('themeToggle');
const themeIcon = document.getElementById('themeIcon');
const html = document.documentElement;

// --- Theme Logic ---
function setTheme(isDark) {
    if (isDark) {
        html.classList.add('dark');
        themeIcon.classList.remove('bi-moon-stars-fill');
        themeIcon.classList.add('bi-sun-fill');
        localStorage.setItem('sipedasi:theme', 'dark');
    } else {
        html.classList.remove('dark');
        themeIcon.classList.remove('bi-sun-fill');
        themeIcon.classList.add('bi-moon-stars-fill');
        localStorage.setItem('sipedasi:theme', 'light');
    }
}
const savedTheme = localStorage.getItem('sipedasi:theme');
const systemDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
setTheme(savedTheme === 'dark' || (!savedTheme && systemDark));
themeToggle.addEventListener('click', () => setTheme(!html.classList.contains('dark')));

// --- UI Elements ---
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
let excelData = [], headers = [], isNewFormat = false, currentCols = {}, unfilteredData = []; // Working set
let masterData = []; // Immutable backup for restoration
let currentRenderData = []; // [headers, ...rows]
let currentPage = 1;
const rowsPerPage = 50;
let selectedIds = new Set(); // Stores _uid of selected rows

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

// --- Utils ---
const num = (v) => parseFloat(String(v || '0').replace(/[^0-9.-]/g, '')) || 0;
const fmt = (v) => (v || 0).toLocaleString('id-ID');
function normalizeJenis(j) { if (typeof j !== 'string') j = String(j || '').trim(); return j.replace(/\s*\(NonReg\)$/i, '').trim(); }
function generateUID() { return Date.now().toString(36) + Math.random().toString(36).substr(2); }

// --- Modal Helpers & Theme ---
function swalTheme() {
    const isDark = document.documentElement.classList.contains('dark');
    return {
        background: isDark ? '#1e293b' : '#fff',
        color: isDark ? '#f8fafc' : '#1e293b'
    };
}

function toggleModal(modal, show) {
    if (!modal) return;
    if (show) {
        modal.classList.remove('hidden');
        setTimeout(() => {
            modal.classList.remove('opacity-0');
            modal.children[0].classList.remove('scale-95');
            modal.children[0].classList.add('scale-100');
        }, 10);
    } else {
        modal.classList.add('opacity-0');
        modal.children[0].classList.remove('scale-100');
        modal.children[0].classList.add('scale-95');
        setTimeout(() => modal.classList.add('hidden'), 200);
    }
}

// --- Upload Logic ---
fabUpload.addEventListener('click', () => toggleModal(uploadModal, true));
closeModal.addEventListener('click', () => toggleModal(uploadModal, false));
btnCancelModal.addEventListener('click', () => toggleModal(uploadModal, false));
uploadModal.addEventListener('click', (e) => { if (e.target === uploadModal) toggleModal(uploadModal, false); });

document.getElementById('input-excel').addEventListener('change', async function (e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    const labelDiv = document.getElementById('uploadLabel');
    labelDiv.innerHTML = `
        <div class="spinner border-indigo-500 border-t-transparent w-8 h-8 mb-2 rounded-full animate-spin"></div>
        <div>
            <p class="text-sm font-bold text-slate-800 dark:text-white truncate max-w-[200px]">Memproses ${files.length} file...</p>
            <p class="text-xs text-indigo-600 font-medium animate-pulse">Mohon tunggu...</p>
        </div>
    `;

    // Reset Data
    excelData = [];
    unfilteredData = [];
    headers = [];
    selectedIds.clear(); // Reset selections on new file

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
        const results = await Promise.all(files.map(file => readExcel(file).catch(e => { failedCount++; return []; })));

        let combinedRows = [];
        results.forEach(rows => {
            if (rows && rows.length > 0) combinedRows = combinedRows.concat(rows);
        });

        if (combinedRows.length === 0) {
            showToast('Tidak ada data valid yang ditemukan.', 'error');
            setTimeout(() => toggleModal(uploadModal, false), 1000);
            return;
        }

        // Processing & UID Assignment & Blank Handling
        const colTempat = currentCols['Tempat Bayar'];
        combinedRows.forEach(row => {
            Object.defineProperty(row, '_uid', { value: generateUID(), writable: true, enumerable: false });

            // Handle Blank Tempat Bayar
            const rawTempat = row[colTempat];
            if (!rawTempat || String(rawTempat).trim() === '') {
                row[colTempat] = 'BLANK';
            }
        });

        const sorted = sortByTanggalWaktu(combinedRows);
        masterData = sorted; // Backup original
        unfilteredData = [...masterData]; // Working copy
        excelData = [headers].concat(unfilteredData); // Legacy wrapper

        initTomSelect();
        renderJenisPajakOptions(); renderTanggalOptions(); renderUraianPajakOptions(); filterAndRender();

        const successMsg = failedCount > 0
            ? `Berhasil: ${files.length - failedCount} file. Gagal: ${failedCount}.`
            : `${files.length} file berhasil digabung!`;

        showToast(successMsg, 'success');

        // Enable Filters
        toggleInputs(true);

        labelDiv.innerHTML = `
            <div class="w-14 h-14 rounded-full bg-green-100 text-green-600 flex items-center justify-center text-xl mb-1"><i class="bi bi-check-lg"></i></div>
            <div><p class="text-sm font-bold text-slate-800 dark:text-white">Selesai!</p><p class="text-xs text-green-600 font-medium">${unfilteredData.length} baris data.</p></div>
        `;

        setTimeout(() => toggleModal(uploadModal, false), 1500);

    } catch (err) {
        console.error("Global Error:", err);
        showToast('Terjadi kesalahan saat memproses file.', 'error');
        setTimeout(() => toggleModal(uploadModal, false), 1000);
    }
});

// Initialize immediately for UI consistency
document.addEventListener('DOMContentLoaded', () => {
    initTomSelect();
});

function initTomSelect() {
    if (tomSelectInstance) tomSelectInstance.destroy();

    // Get unique tempat bayar (if data exists)
    const uniquePlaces = unfilteredData && unfilteredData.length > 0
        ? [...new Set(unfilteredData.map(r => String(r[currentCols['Tempat Bayar']] || '').trim()))].filter(Boolean).sort()
        : [];

    tomSelectInstance = new TomSelect('#detail-tempat-select', {
        plugins: ['remove_button', 'dropdown_input'],
        create: false,
        placeholder: "Cari detail tempat bayar...",
        onItemAdd: handlePlaceChange,
        onItemRemove: handlePlaceChange,
        onClear: handlePlaceChange
    });

    tomSelectInstance.clearOptions();
    uniquePlaces.forEach(p => tomSelectInstance.addOption({ value: p, text: p }));

    // Enable only if we have data and options
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

document.getElementById('tempatbayar-select').addEventListener('change', function (e) {
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


// --- Filtering Logic ---
function getFilteredData() {
    const jenis = document.getElementById('jenispajak-select').value;
    const uraian = document.getElementById('uraianspajak-select').value;
    const waktu = document.getElementById('waktu-select').value;
    const tanggalSel = document.getElementById('tanggal-select').value;
    const isRange = document.querySelector('input[name="tanggalMode"]:checked')?.value === 'range';
    const fromDate = document.getElementById('tanggal-from').value;
    const toDate = document.getElementById('tanggal-to').value;

    let selectedDetails = [];
    if (tomSelectInstance) selectedDetails = tomSelectInstance.getValue();

    return unfilteredData.filter(row => {
        if (jenis && normalizeJenis(row[currentCols['Jenis Pajak']]) !== jenis) return false;

        if (selectedDetails.length > 0) {
            const rawTempat = String(row[currentCols['Tempat Bayar']] || '').trim();
            if (!selectedDetails.includes(rawTempat)) return false;
        }

        if (uraian) {
            const u = (typeof row[currentCols['Uraian Pajak']] === 'string') ? row[currentCols['Uraian Pajak']].trim() : String(row[currentCols['Uraian Pajak']] || '').trim();
            if (u !== uraian) return false;
        }

        const raw = String(row[currentCols['Tanggal Bayar']] || '').trim();
        let tdate = raw.substr(0, 10);
        let stdDate = "";
        if (/^\d{8}$/.test(tdate)) {
            stdDate = `${tdate.slice(4)}-${tdate.slice(2, 4)}-${tdate.slice(0, 2)}`;
        } else if (/^\d{4}-\d{2}-\d{2}$/.test(tdate)) {
            stdDate = tdate;
        } else if (/^\d{2}-\d{2}-\d{4}$/.test(tdate)) {
            const [d, m, y] = tdate.split('-');
            stdDate = `${y}-${m}-${d}`;
        }

        if (isRange && fromDate && toDate) {
            if (!stdDate) return false;
            return stdDate >= fromDate && stdDate <= toDate;
        }
        if (tanggalSel) {
            let currentDMY = "";
            if (stdDate) {
                const [y, m, d] = stdDate.split('-');
                currentDMY = `${d}-${m}-${y}`;
            }
            return currentDMY === tanggalSel;
        }

        if (waktu) {
            const m = raw.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
            if (!m) return false;
            const hh = String(m[1]).padStart(2, '0') + ':' + m[2];
            return waktu === 'akhir' ? hh < '17:00' : hh >= '17:00';
        }
        return true;
    });
}

function filterAndRender() {
    let filtered = getFilteredData();
    const mode = document.querySelector('input[name="fitur"]:checked')?.value || 'tambah';
    const tambahan = parseInt(document.getElementById('tambahanCount').value) || 0;
    const waktu = document.getElementById('waktu-select').value;

    if (waktu) {
        const dataWaktu = filtered.filter(row => {
            const t = String(row[currentCols['Tanggal Bayar']] || '').trim();
            const m = t.match(/\b(\d{2}:\d{2}:\d{2})|\b(\d{2}:\d{2})/);
            const jm = m ? m[1] || m[2] : '';
            if (!jm) return false;
            return waktu === 'akhir' ? jm < '17:00' : jm >= '17:00';
        });

        if (tambahan > 0) {
            if (mode === 'tambah') {
                const lawan = getFilteredData().filter(row => {
                    const t = String(row[currentCols['Tanggal Bayar']] || '').trim();
                    const m = t.match(/\b(\d{2}:\d{2}:\d{2})|\b(\d{2}:\d{2})/);
                    const jm = m ? m[1] || m[2] : '';
                    if (!jm) return false;
                    const tglSel = document.getElementById('tanggal-select').value;
                    return waktu === 'akhir' ? jm >= '17:00' : jm < '17:00';
                });
                const added = waktu === 'akhir' ? sortByTanggalWaktu(lawan).slice(0, tambahan) : sortByTanggalWaktu(lawan).reverse().slice(0, tambahan).reverse();
                filtered = [...dataWaktu, ...added];
            } else {
                filtered = waktu === 'akhir' ? dataWaktu.slice(0, Math.max(0, dataWaktu.length - tambahan)) : dataWaktu.slice(Math.min(dataWaktu.length, tambahan));
            }
        } else filtered = dataWaktu;
    }

    filtered = sortByTanggalWaktu(filtered);
    lastFilteredDataPJDL = filtered.slice();
    currentRenderData = [headers].concat(filtered);

    currentPage = 1;
    renderPage();

    updateSummaryCards(filtered);
    document.getElementById('jumlahData').textContent = `Data: ${filtered.length}`;

    // Update Button States
    const hasData = filtered.length > 0;
    if (document.getElementById('exportPjdlBtn')) document.getElementById('exportPjdlBtn').disabled = !hasData;
    if (btnSaveResult) btnSaveResult.disabled = !hasData;

    // Reset Delete Button on re-filter (optional, usually clear selection?)
    // Actually we keep selection if IDs persist, but filtered out rows aren't visible
    updateDeleteButton();
}

function updateDeleteButton() {
    if (btnDeleteSelected) {
        if (selectedIds.size > 0) {
            btnDeleteSelected.disabled = false;
            btnDeleteSelected.classList.remove('hidden');
            btnDeleteSelected.innerHTML = `<i class="bi bi-trash"></i> Hapus (${selectedIds.size})`;
        } else {
            btnDeleteSelected.disabled = true;
            btnDeleteSelected.classList.add('hidden');
        }
    }
}

function renderPage() {
    const dataRows = currentRenderData.slice(1);
    const container = document.getElementById('excel-table');
    const emptyState = document.getElementById('empty-state');
    const controls = document.getElementById('paginationControls');

    if (dataRows.length === 0) {
        container.innerHTML = '';
        emptyState.classList.remove('hidden');
        controls.classList.add('hidden');
        return;
    } else {
        emptyState.classList.add('hidden');
        controls.classList.remove('hidden');
    }

    const totalPages = Math.ceil(dataRows.length / rowsPerPage);
    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const pageData = dataRows.slice(startIdx, endIdx);

    document.getElementById("currentPageDisplay").textContent = currentPage;
    document.getElementById("totalPagesDisplay").textContent = totalPages;
    document.getElementById("prevPageBtn").disabled = currentPage === 1;
    document.getElementById("nextPageBtn").disabled = currentPage === totalPages;

    let displayHeaders = ['Jenis Pajak', 'Uraian Pajak', 'Pokok (Rp)', 'Denda (Rp)', 'Jumlah Bayar (Rp)', 'Tanggal Bayar', 'Tempat Bayar'];
    const jenisSel = document.getElementById('jenispajak-select').value;
    if (jenisSel && jenisSel.toLowerCase().includes('mineral')) displayHeaders = ['Jenis Pajak', 'Uraian Pajak', 'Pokok (Rp)', 'Denda (Rp)', 'Pokok Opsen (Rp)', 'Denda Opsen (Rp)', 'Jumlah Bayar (Rp)', 'Tanggal Bayar', 'Tempat Bayar'];

    let html = `<thead class="text-xs text-slate-500 dark:text-slate-400 uppercase bg-slate-50 dark:bg-slate-800 sticky top-0 z-10 shadow-sm"><tr>`;
    // Select All Checkbox
    html += `<th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700 w-10">
                <input type="checkbox" id="selectAll" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 h-4 w-4">
             </th>`;

    html += `<th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700">No</th>`;
    displayHeaders.forEach(h => {
        const align = h.includes('(Rp)') ? 'text-right' : 'text-left';
        html += `<th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700 ${align}">${h.replace('(Rp)', '')}</th>`;
    });
    html += `</tr></thead><tbody class="bg-white dark:bg-slate-900/50 divide-y divide-slate-100 dark:divide-slate-800">`;

    pageData.forEach((row, i) => {
        const globalIndex = startIdx + i + 1;
        const isAlt = i % 2 === 0 ? 'bg-white dark:bg-transparent' : 'bg-slate-50/50 dark:bg-slate-800/30';
        const isChecked = selectedIds.has(row._uid) ? 'checked' : '';

        html += `<tr class="${isAlt} hover:bg-blue-50/50 dark:hover:bg-slate-800 transition-colors">`;

        // Row Checkbox
        html += `<td class="px-6 py-3 border-b border-slate-100 dark:border-slate-800">
                    <input type="checkbox" class="row-checkbox rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 h-4 w-4" value="${row._uid}" ${isChecked}>
                 </td>`;

        html += `<td class="px-6 py-3 font-medium text-slate-400 text-xs">${globalIndex}</td>`;

        displayHeaders.forEach(h => {
            const align = h.includes('(Rp)') ? 'text-right' : 'text-left';
            let v = row[currentCols[h]] !== undefined ? row[currentCols[h]] : '';
            let content = String(v || '');
            let colorClass = "text-slate-600 dark:text-slate-400";

            if (h.includes('(Rp)')) {
                content = fmt(num(v));
                colorClass = h.includes('Pokok') ? "text-slate-700 dark:text-slate-300 font-medium" :
                    h.includes('Denda') ? "text-amber-600 dark:text-amber-500" :
                        h.includes('Jumlah') ? "text-emerald-600 dark:text-emerald-400 font-bold" : colorClass;
            } else if (h === 'Uraian Pajak') {
                const rawContent = String(v || '').trim().replace(/\s*-\s*/g, '-');
                // Escape quotes for onclick
                const safeContent = rawContent.replace(/'/g, "\\'").replace(/"/g, "&quot;");
                content = `
                    <div class="flex items-center justify-between gap-2 group min-w-[200px] cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800/50 -mx-3 px-3 py-1.5 rounded-lg transition-colors select-none" onclick="copyText('${safeContent}')" title="Klik untuk menyalin">
                         <span class="whitespace-normal text-slate-700 dark:text-slate-300 font-medium relative top-[1px]">${rawContent}</span>
                         <div class="w-6 h-6 rounded flex items-center justify-center text-indigo-500 bg-indigo-50 dark:bg-indigo-900/30 opacity-0 group-hover:opacity-100 transition-all transform scale-90 group-hover:scale-100 shadow-sm border border-indigo-100 dark:border-indigo-800">
                             <i class="bi bi-copy text-xs"></i>
                         </div>
                    </div>
                `;
            } else if (h === 'Tanggal Bayar') {
                colorClass = "text-slate-500 dark:text-slate-500 font-mono text-xs";
            }

            html += `<td class="px-6 py-3 ${align} ${colorClass}">${content}</td>`;
        });
        html += `</tr>`;
    });
    html += `</tbody>`;
    container.innerHTML = html;

    const sc = document.getElementById('table-scroll-container');
    if (sc) sc.scrollTop = 0;

    // Attach Event Listeners for Checkboxes
    const selectAll = document.getElementById('selectAll');
    const rowCheckboxes = document.querySelectorAll('.row-checkbox');

    // Update Select All state based on page selection
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

// --- Delete Selected Handling ---
if (btnDeleteSelected) {
    btnDeleteSelected.addEventListener('click', () => {
        if (selectedIds.size === 0) return;

        Swal.fire({
            title: `Hapus ${selectedIds.size} Data?`,
            text: "Data yang dihapus tidak dapat dipulihkan dari tampilan ini.",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#e11d48',
            cancelButtonColor: '#64748b',
            confirmButtonText: 'Ya, Hapus',
            cancelButtonText: 'Batal',
            ...swalTheme()
        }).then((result) => {
            if (result.isConfirmed) {
                // Perform deletion
                const initialCount = unfilteredData.length;
                unfilteredData = unfilteredData.filter(row => !selectedIds.has(row._uid));
                const deletedCount = initialCount - unfilteredData.length;

                selectedIds.clear();
                updateDeleteButton();

                // Re-calculate derived data
                excelData = [headers].concat(unfilteredData);

                // Refresh UI
                renderJenisPajakOptions(); renderTanggalOptions(); renderUraianPajakOptions();
                filterAndRender();

                showToast(`${deletedCount} data berhasil dihapus.`, 'success');
            }
        });
    });
}

// --- Pagination Events ---
document.getElementById("prevPageBtn").addEventListener("click", () => {
    if (currentPage > 1) { currentPage--; renderPage(); }
});
document.getElementById("nextPageBtn").addEventListener("click", () => {
    const totalPages = Math.ceil((currentRenderData.length - 1) / rowsPerPage);
    if (currentPage < totalPages) { currentPage++; renderPage(); }
});

// --- Save & Load Handlers ---
if (openSavedDataBtn) openSavedDataBtn.addEventListener('click', () => {
    renderSavedDataList();
    toggleModal(savedDataModal, true);
});
if (closeSavedModal) closeSavedModal.addEventListener('click', () => toggleModal(savedDataModal, false));
if (savedDataModal) savedDataModal.addEventListener('click', (e) => { if (e.target === savedDataModal) toggleModal(savedDataModal, false); });
if (btnSaveResult) btnSaveResult.addEventListener('click', saveFilteredData);

async function saveFilteredData() {
    if (!lastFilteredDataPJDL || lastFilteredDataPJDL.length === 0) return;

    const { value: name } = await Swal.fire({
        title: 'Simpan Hasil Filter',
        input: 'text',
        inputLabel: 'Beri nama untuk data ini',
        inputPlaceholder: 'Contoh: Laporan Januari 2025',
        showCancelButton: true,
        confirmButtonText: 'Simpan',
        cancelButtonText: 'Batal',
        confirmButtonColor: '#3b82f6',
        inputValidator: (value) => {
            if (!value) return 'Nama tidak boleh kosong!';
        },
        ...swalTheme()
    });

    if (name) {
        const saveData = {
            id: Date.now(),
            name: name,
            date: new Date().toISOString(),
            headers: headers,
            rows: lastFilteredDataPJDL, // Just the filtered rows array
            currentCols: currentCols,
            isNewFormat: isNewFormat
        };

        const existing = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
        existing.push(saveData);
        localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));

        showToast('Data berhasil disimpan!', 'success');
    }
}

function renderSavedDataList() {
    const list = document.getElementById('savedDataList');
    const empty = document.getElementById('emptySaved');
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');

    list.innerHTML = '';
    if (saved.length === 0) {
        list.classList.add('hidden');
        empty.classList.remove('hidden');
        return;
    }

    list.classList.remove('hidden');
    empty.classList.add('hidden');

    saved.sort((a, b) => new Date(b.date) - new Date(a.date)).forEach(item => {
        const dateObj = new Date(item.date);
        const dateStr = dateObj.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
        const timeStr = dateObj.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors group";
        tr.innerHTML = `
            <td class="px-6 py-3 font-medium text-slate-700 dark:text-white">
                ${item.name}
            </td>
            <td class="px-6 py-3 text-center text-slate-500 text-xs">
                ${dateStr} <span class="text-slate-400">(${timeStr})</span>
            </td>
            <td class="px-6 py-3 text-center">
                <span class="px-2 py-1 bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 rounded-lg text-xs font-bold">${item.rows.length}</span>
            </td>
            <td class="px-6 py-3 text-right">
                <div class="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button onclick="loadSavedData(${item.id})" class="p-1.5 text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors" title="Buka">
                        <i class="bi bi-folder2-open"></i>
                    </button>
                    <button onclick="deleteSavedData(${item.id})" class="p-1.5 text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/30 rounded-lg transition-colors" title="Hapus">
                        <i class="bi bi-trash"></i>
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
    // Restore UIDs if missing (for older saves) or use existing
    const rowsWithUid = item.rows.map(r => {
        if (!r._uid) Object.defineProperty(r, '_uid', { value: generateUID(), writable: true, enumerable: false });
        return r;
    });

    masterData = rowsWithUid; // Backup
    unfilteredData = [...masterData]; // Working copy
    excelData = [headers].concat(unfilteredData);
    currentCols = item.currentCols || (item.isNewFormat ? newFormatCols : oldFormatCols);
    selectedIds.clear();

    initTomSelect();

    ['jenispajak-select', 'tempatbayar-select', 'waktu-select', 'tanggal-select'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('tambahanCount').value = 0;

    renderJenisPajakOptions();
    renderTanggalOptions();
    renderUraianPajakOptions();

    filterAndRender();

    toggleModal(savedDataModal, false);
    showToast(`Data "${item.name}" dimuat.`, 'success');

    // Enable Filters
    toggleInputs(true);
};

window.deleteSavedData = function (id) {
    Swal.fire({
        title: 'Hapus data ini?',
        text: "Data tidak dapat dikembalikan!",
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
            showToast('Data dihapus.', 'success');
        }
    });
};

document.getElementById('btnClearAll')?.addEventListener('click', () => {
    Swal.fire({
        title: 'Hapus SEMUA data?',
        text: "Seluruh riwayat penyimpanan akan hilang!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#f43f5e',
        cancelButtonColor: '#64748b',
        confirmButtonText: 'Ya, Bersihkan',
        cancelButtonText: 'Batal',
        ...swalTheme()
    }).then((result) => {
        if (result.isConfirmed) {
            localStorage.removeItem(STORAGE_KEY);
            renderSavedDataList();
            showToast('Semua data dihapus.', 'success');
        }
    });
});

// --- SINGLE SHEET MERGE LOGIC ---
// --- SINGLE SHEET MERGE LOGIC (Updated Format) ---
document.getElementById('btnDownloadAll')?.addEventListener('click', () => {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    if (saved.length === 0) { showToast('Tidak ada data tersimpan.', 'error'); return; }

    try {
        // Master Headers: Prepend Metadata Columns
        const originalHeaders = saved[0].headers || [];
        const masterHeaders = ["Nama Laporan", "Waktu Simpan", ...originalHeaders];

        let allRows = [masterHeaders];

        saved.forEach(item => {
            // Format Date for the row
            const savedDate = new Date(item.date);
            const timeStr = savedDate.toLocaleDateString('id-ID') + ' ' + savedDate.toLocaleTimeString('id-ID');

            const processedRows = item.rows.map(row => {
                // Ensure row is array
                const r = Array.isArray(row) ? row : Object.values(row);
                // Remove internal _uid if present (it's usually a non-enumerable prop, but just in case of spread)
                // Actually the spread [...r] handles iterable properties. _uid is non-enumerable.

                // Prepend: [Name, Time, ...Data]
                return [item.name, timeStr, ...r];
            });
            allRows = allRows.concat(processedRows);
        });

        const ws = XLSX.utils.aoa_to_sheet(allRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Gabungan Data");

        const now = new Date();
        const fname = `Backup_SiRekonPJDL_Merged_${now.toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fname);

        showToast(`${saved.length} dataset berhasil digabung (Format Baru)!`, 'success');

    } catch (err) {
        console.error("Bulk Export Error:", err);
        showToast('Gagal membuat file backup.', 'error');
    }
});

// --- Other Helpers ---
function copyText(text) {
    if (!text) return;
    navigator.clipboard.writeText(text).then(() => {
        showToast(`Uraian "${text}" disalin!`, 'success');
    }).catch(err => {
        console.error('Failed to copy: ', err);
        showToast('Gagal menyalin teks.', 'error');
    });
}
function copyTotal(id) {
    const el = document.getElementById(id); if (!el) return;
    const val = el.textContent.replace(/\./g, '').replace(/,/g, '.');
    navigator.clipboard.writeText(val).then(() => showToast(`Nominal "${el.textContent}" disalin!`, 'success'));
}

function showToast(msg, type = 'success') {
    const toast = document.getElementById("toast");
    const icon = document.getElementById("toastIcon");
    const txt = document.getElementById("toastMsg");
    txt.textContent = msg;
    toast.className = "fixed bottom-5 right-5 px-6 py-4 rounded-xl shadow-xl transition-all duration-300 z-[60] transform flex items-center gap-3 font-medium text-sm glass text-white backdrop-blur-md";
    if (type === 'success') { toast.classList.add('bg-emerald-600/90'); icon.innerHTML = '<i class="bi bi-check-circle-fill text-lg"></i>'; }
    else { toast.classList.add('bg-rose-600/90'); icon.innerHTML = '<i class="bi bi-x-circle-fill text-lg"></i>'; }
    requestAnimationFrame(() => toast.classList.remove('translate-y-10', 'opacity-0'));
    setTimeout(() => toast.classList.add('translate-y-10', 'opacity-0'), 3000);
}

function sortByTanggalWaktu(data) {
    const idx = currentCols['Tanggal Bayar'];
    return data.filter(r => r && r.some(c => c !== undefined && String(c).trim() !== '')).sort((a, b) => {
        const ta = String(a[idx] || '').trim(), tb = String(b[idx] || '').trim();
        const norm = s => { const [d, t] = String(s || '').split(' '); if (/^\d{2}-\d{2}-\d{4}$/.test(d)) { const [dd, mm, yy] = d.split('-'); return `${yy}-${mm}-${dd}${t ? ' ' + t : ''}`; } return `${d || ''}${t ? ' ' + t : ''}`; };
        return norm(ta).localeCompare(norm(tb));
    });
}

function debounce(fn, ms = 300) { let t; return function (...a) { clearTimeout(t); t = setTimeout(() => fn.apply(this, a), ms); }; }
const debouncedFilterAndRender = debounce(filterAndRender, 300);

['change', 'input'].forEach(evt => {
    document.getElementById('jenispajak-select').addEventListener('change', () => { renderTanggalOptions(); renderUraianPajakOptions(); debouncedFilterAndRender(); });
    document.getElementById('tanggal-select').addEventListener('change', () => { renderUraianPajakOptions(); debouncedFilterAndRender(); });
    // Tempat bayar listener already handled specifically
    document.getElementById('waktu-select').addEventListener('change', () => { renderTanggalOptions(); renderUraianPajakOptions(); debouncedFilterAndRender(); });
    document.getElementById('uraianspajak-select').addEventListener('change', debouncedFilterAndRender);
    document.getElementById('tambahanCount').addEventListener('input', debouncedFilterAndRender);
});

// Date Range UI
function setTanggalModeUI(mode) {
    const rangeDiv = document.getElementById('tanggal-range-inputs');
    const singleContainer = document.getElementById('tanggal-single-container');
    if (mode === 'range') {
        rangeDiv.classList.remove('hidden'); singleContainer.classList.add('hidden');
    } else {
        rangeDiv.classList.add('hidden'); singleContainer.classList.remove('hidden');
        document.getElementById('tanggal-from').value = '';
        document.getElementById('tanggal-to').value = '';
    }
}
document.querySelectorAll('input[name="tanggalMode"]').forEach(r => {
    r.addEventListener('change', function () {
        setTanggalModeUI(this.value);
        if (this.value === 'single') {
            document.getElementById('tanggal-from').value = '';
            document.getElementById('tanggal-to').value = '';
        } else {
            document.getElementById('tanggal-select').value = '';
        }
        debouncedFilterAndRender();
    });
});
document.getElementById('tanggal-from').addEventListener('change', () => { if (document.querySelector('input[name="tanggalMode"]:checked')?.value === 'range') debouncedFilterAndRender(); });
document.getElementById('tanggal-to').addEventListener('change', () => { if (document.querySelector('input[name="tanggalMode"]:checked')?.value === 'range') debouncedFilterAndRender(); });

// Reset Button Logic
document.getElementById('btnResetPJDL').addEventListener('click', function () {
    ['jenispajak-select', 'tempatbayar-select', 'waktu-select', 'tanggal-select', 'uraianspajak-select'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('tambahanCount').value = 0;

    // Reset Date Mode
    const singleMode = document.querySelector('input[name="tanggalMode"][value="single"]');
    if (singleMode) { singleMode.checked = true; setTanggalModeUI('single'); }

    // Reset Fitur
    document.querySelectorAll('input[name="fitur"]').forEach(r => r.checked = (r.value === 'tambah'));

    // Reset Tom Select
    if (tomSelectInstance) tomSelectInstance.clear();

    // RESTORE DATA FROM MASTER
    if (masterData.length > 0) {
        unfilteredData = [...masterData];
        excelData = [headers].concat(unfilteredData);
    }

    // Reset Options & Data
    renderJenisPajakOptions(); renderTanggalOptions(); renderUraianPajakOptions();

    // Clear selection on reset
    selectedIds.clear();

    filterAndRender();
    showToast('Filter & Data direset (Data dihapus dikembalikan).', 'success');
});

// Dropdown Renders
function renderJenisPajakOptions() {
    const sel = document.getElementById('jenispajak-select'); sel.innerHTML = '<option value="">-- Semua Jenis Pajak --</option>';
    const s = new Set(); unfilteredData.forEach(r => { const v = normalizeJenis(r[currentCols['Jenis Pajak']]); if (v) s.add(v); });
    Array.from(s).sort().forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
}

function renderTanggalOptions() {
    const sel = document.getElementById('tanggal-select'); const prev = sel.value; sel.innerHTML = '<option value="">-- Pilih Tanggal --</option>';
    const jenis = document.getElementById('jenispajak-select').value;
    const uraian = document.getElementById('uraianspajak-select').value;
    const tanggalCounts = {}; const set = new Set();

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
        const t = String(row[currentCols['Tanggal Bayar']] || '').trim().substr(0, 10); if (!t) return;

        let normalizedTgl = '';
        if (/^\d{8}$/.test(t)) normalizedTgl = `${t.slice(0, 2)}-${t.slice(2, 4)}-${t.slice(4)}`;
        else if (/^\d{4}-\d{2}-\d{2}$/.test(t)) { const [y, m, d] = t.split('-'); normalizedTgl = `${d}-${m}-${y}`; }
        else if (/^\d{2}-\d{2}-\d{4}$/.test(t)) normalizedTgl = t;

        if (normalizedTgl) { tanggalCounts[normalizedTgl] = (tanggalCounts[normalizedTgl] || 0) + 1; set.add(normalizedTgl); }
    });

    const sortedDates = Array.from(set).sort((a, b) => { const ta = a.split('-').reverse().join('-'); const tb = b.split('-').reverse().join('-'); return ta.localeCompare(tb); });

    const fromInput = document.getElementById('tanggal-from');
    const toInput = document.getElementById('tanggal-to');
    function convertToYYYYMMDD(dateStr) {
        if (/^\d{2}-\d{2}-\d{4}$/.test(dateStr)) { const [d, m, y] = dateStr.split('-'); return `${y}-${m}-${d}`; }
        return '';
    }
    if (sortedDates.length > 0) {
        const minDate = convertToYYYYMMDD(sortedDates[0]);
        const maxDate = convertToYYYYMMDD(sortedDates[sortedDates.length - 1]);
        if (fromInput && toInput) { fromInput.min = minDate; fromInput.max = maxDate; toInput.min = minDate; toInput.max = maxDate; fromInput.disabled = false; toInput.disabled = false; }
    } else {
        if (fromInput && toInput) { fromInput.value = ''; fromInput.disabled = true; toInput.value = ''; toInput.disabled = true; }
    }

    sortedDates.forEach(t => { const o = document.createElement('option'); o.value = t; o.textContent = `${t} (${tanggalCounts[t]} data)`; sel.appendChild(o); });
    sel.value = Array.from(sel.options).some(o => o.value === prev) ? prev : '';
}

function renderUraianPajakOptions() {
    const sel = document.getElementById('uraianspajak-select'); sel.innerHTML = '<option value="">-- Semua Uraian Pajak --</option>';
    const jenis = document.getElementById('jenispajak-select').value;
    const tgl = document.getElementById('tanggal-select').value;

    let arr = unfilteredData.slice();
    if (jenis) arr = arr.filter(r => normalizeJenis(r[currentCols['Jenis Pajak']]) === jenis);
    if (tgl) arr = arr.filter(r => String(r[currentCols['Tanggal Bayar']] || '').substr(0, 10) === tgl);

    if (tomSelectInstance) {
        let selectedDetails = tomSelectInstance.getValue();
        if (selectedDetails.length > 0) {
            arr = arr.filter(r => selectedDetails.includes(String(r[currentCols['Tempat Bayar']] || '').trim()));
        }
    }

    const s = new Set(); arr.forEach(r => { let u = r[currentCols['Uraian Pajak']]; if (typeof u !== 'string') u = String(u || '').trim(); if (u) s.add(u); });
    Array.from(s).sort().forEach(u => { const o = document.createElement('option'); o.value = u; o.textContent = String(u).replace(/\s*-\s*/g, '-'); sel.appendChild(o); });
}

function updateSummaryCards(data) {
    let totalPokok = 0, totalDenda = 0, totalJumlah = 0;
    data.forEach(r => {
        totalPokok += num(r[currentCols['Pokok (Rp)']]);
        totalDenda += num(r[currentCols['Denda (Rp)']]);
        totalJumlah += num(r[currentCols['Jumlah Bayar (Rp)']]);
    });
    document.getElementById('total-pokok').textContent = fmt(totalPokok);
    document.getElementById('total-denda').textContent = fmt(totalDenda);
    document.getElementById('total-jumlah').textContent = fmt(totalJumlah);

    const bars = document.querySelectorAll('#summaryCards .bg-slate-100, #summaryCards .dark\\:bg-slate-800 > div');
    bars.forEach(bar => { bar.style.width = '0%'; setTimeout(() => bar.style.width = '100%', 100); });
}

function exportPJDLToExcel() {
    if (!lastFilteredDataPJDL || lastFilteredDataPJDL.length === 0) { showToast('Tidak ada data untuk diekspor.', 'error'); return; }

    try {
        const wsData = [headers].concat(lastFilteredDataPJDL);
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'PJDL');

        const now = new Date();
        const fname = `SiRekonPJDL_Export_Full_${now.toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fname);
        showToast(`Berhasil mengekspor ${lastFilteredDataPJDL.length} baris (Full Data).`, 'success');
    } catch (e) { console.error(e); showToast('Gagal mengekspor data.', 'error'); }
}
if (document.getElementById('exportPjdlBtn')) document.getElementById('exportPjdlBtn').addEventListener('click', exportPJDLToExcel);

// --- Helper to Toggle Inputs ---
function toggleInputs(enable) {
    const ids = [
        'btnResetPJDL',
        'jenispajak-select',
        'tempatbayar-select',
        'detail-tempat-select',
        'waktu-select',
        'uraianspajak-select',
        'tanggal-select',
        'tambahanCount'
    ];

    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.disabled = !enable;
            if (enable) {
                el.classList.remove('opacity-60', 'cursor-not-allowed');
            } else {
                // Class list handling usually done by CSS :disabled selector
            }
        }
    });

    // Radios
    document.querySelectorAll('input[name="fitur"]').forEach(r => r.disabled = !enable);

    // Tom Select
    if (tomSelectInstance) {
        if (enable) tomSelectInstance.enable();
        else tomSelectInstance.disable();
    }
}
