// --- Security ---
function escapeHtml(text) {
    if (text === null || text === undefined) return text;
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

const themeToggle = document.getElementById('themeToggle');
const themeIcon = document.getElementById('themeIcon');
const html = document.documentElement;

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

const fabUpload = document.getElementById('fabUpload');
const uploadModal = document.getElementById('uploadModal');
const closeModal = document.getElementById('closeModal');
const btnCancelModal = document.getElementById('btnCancelModal');

function toggleModal(show) {
    if (show) {
        uploadModal.classList.remove('hidden');
        setTimeout(() => {
            uploadModal.classList.remove('opacity-0');
            uploadModal.children[0].classList.remove('scale-95');
            uploadModal.children[0].classList.add('scale-100');
        }, 10);
    } else {
        uploadModal.classList.add('opacity-0');
        uploadModal.children[0].classList.remove('scale-100');
        uploadModal.children[0].classList.add('scale-95');
        setTimeout(() => uploadModal.classList.add('hidden'), 200);
    }
}
fabUpload.addEventListener('click', () => toggleModal(true));
closeModal.addEventListener('click', () => toggleModal(false));
btnCancelModal.addEventListener('click', () => toggleModal(false));
uploadModal.addEventListener('click', (e) => { if (e.target === uploadModal) toggleModal(false); });

let lastFilteredData = [];
let dataMonitoringList = [];
let dataKeduaList = [];
let originalDataBackup = [];
let fullCombinedData = [];
let currentPage = 1;
const rowsPerPage = 50;
let currentRenderData = [];
const STORAGE_KEY = 'sipedasi:bphtb_saved';

function generateUID() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

function parseNumber(v) {
    if (v == null) return 0;
    if (typeof v === 'number') return v;
    const s = String(v).replace(/\./g, '').replace(/,/g, '.').replace(/[^\d\.\-]/g, '');
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : 0;
}

function processFile(file, labelElement, callback) {
    const originalHTML = labelElement.innerHTML;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (event) {
        try {
            const data = new Uint8Array(event.target.result);
            let workbook;
            try {
                workbook = XLSX.read(data, { type: "array" });
            } catch (e) {
                // Fallback: Try reading as text (HTML/XML masquerading as XLS)
                console.warn("Binary read failed, trying text mode...", e);
                const str = new TextDecoder("utf-8").decode(data);
                workbook = XLSX.read(str, { type: "string" });
            }

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let resultData;

            if (labelElement.id.includes('Monitoring')) {
                let domSuccess = false;
                try {
                    const htmlContent = new TextDecoder("utf-8").decode(data);
                    if (htmlContent.includes('<tr')) {
                        console.log("Using DOMParser for Monitoring...");
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(htmlContent, 'text/html');
                        const tables = doc.querySelectorAll('table');
                        for (const table of tables) {
                            const rows = Array.from(table.rows);
                            let hIdx = -1;
                            let colMap = { tgl: -1, tmpt: -1 };
                            for (let i = 0; i < Math.min(rows.length, 50); i++) {
                                const cells = Array.from(rows[i].cells).map(c => c.textContent ? c.textContent.trim().toUpperCase() : '');
                                const idx = cells.findIndex(c => c.includes('TANGGAL PEMBAYARAN'));
                                if (idx > -1) {
                                    hIdx = i; colMap.tgl = idx;
                                    colMap.tmpt = cells.findIndex(c => c.includes('TEMPAT PEMBAYARAN') || c.includes('KECAMATAN') || c.includes('BANK'));
                                    break;
                                }
                            }
                            if (hIdx > -1) {
                                const res = rows.slice(hIdx + 1).map(tr => {
                                    const c = tr.cells;
                                    const get = (i) => i > -1 && c[i] ? c[i].textContent.trim() : '';
                                    const t = get(colMap.tgl);
                                    return t ? { tgl: excelSerialToDateTimeString(t), tmpt: get(colMap.tmpt) } : null;
                                }).filter(x => x && x.tgl);
                                if (res.length) {
                                    resultData = res;
                                    domSuccess = true;
                                    document.getElementById('countMonitoring').textContent = `${resultData.length} Data`;
                                    break;
                                }
                            }
                        }
                    }
                } catch (e) { console.warn(e); }

                if (!domSuccess) {
                    // Use array array mode to find header row manually, just like we do for Data Ke-2
                    const rowsAsArray = XLSX.utils.sheet_to_json(worksheet, { defval: null, header: 1 });
                    let headerRowIndex = -1, headerRow = [];

                    // Search for header in first 20 rows
                    for (let i = 0; i < Math.min(rowsAsArray.length, 20); i++) {
                        const row = rowsAsArray[i];
                        if (row) {
                            // Check for key columns using partial match to handle tags like <b>
                            const hasTgl = row.some(h => String(h).toUpperCase().includes('TANGGAL PEMBAYARAN'));
                            const hasTmpt = row.some(h => String(h).toUpperCase().includes('TEMPAT PEMBAYARAN') || String(h).toUpperCase().includes('KECAMATAN'));

                            if (hasTgl && hasTmpt) {
                                headerRowIndex = i;
                                headerRow = row.map(h => String(h || '').trim());
                                break;
                            }
                        }
                    }

                    if (headerRowIndex === -1) {
                        // Fallback: try to see if it was parsed as object with keys already (if simple table)
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null, raw: true });
                        if (jsonData.length > 0) {
                            // Look for key in first row properties
                            const keys = Object.keys(jsonData[0]);
                            const tglKey = keys.find(k => k.toUpperCase().includes('TANGGAL PEMBAYARAN'));

                            if (tglKey) {
                                // Use existing logic as fallback
                                resultData = jsonData.map(row => {
                                    // Normalize keys
                                    const newRow = {};
                                    for (const key in row) newRow[key.trim()] = row[key];

                                    // Find values regardless of exact case/trim in keys
                                    const tglVal = row[tglKey];
                                    // Try to find place or bank
                                    const tmptKey = keys.find(k => k.toUpperCase().includes('TEMPAT PEMBAYARAN') || k.toUpperCase().includes('BANK'));
                                    const tmptVal = tmptKey ? row[tmptKey] : '';

                                    return {
                                        tgl: excelSerialToDateTimeString(tglVal),
                                        tmpt: tmptVal || ''
                                    };
                                }).filter(item => item.tgl);
                            } else {
                                throw new Error("Header 'Tanggal Pembayaran' tidak ditemukan pada file Monitoring.");
                            }
                        } else {
                            throw new Error("Data tidak ditemukan atau kosong.");
                        }
                    } else {
                        // Map using indices with partial matching
                        const colIndices = {
                            tgl: headerRow.findIndex(h => h.toUpperCase().includes('TANGGAL PEMBAYARAN')),
                            tmpt: headerRow.findIndex(h => h.toUpperCase().includes('TEMPAT PEMBAYARAN')),
                            bank: headerRow.findIndex(h => h.toUpperCase().includes('BANK'))
                        };

                        resultData = rowsAsArray.slice(headerRowIndex + 1).map(row => {
                            const tglVal = colIndices.tgl > -1 ? row[colIndices.tgl] : null;
                            let tmptVal = '';
                            if (colIndices.tmpt > -1) tmptVal = row[colIndices.tmpt];
                            else if (colIndices.bank > -1) tmptVal = row[colIndices.bank];

                            return {
                                tgl: excelSerialToDateTimeString(tglVal),
                                tmpt: tmptVal || ''
                            };
                        }).filter(item => item.tgl);
                    }

                    document.getElementById('countMonitoring').textContent = `${resultData.length} Data`;
                } // End fallback
            } else {
                const rowsAsArray = XLSX.utils.sheet_to_json(worksheet, { defval: null, header: 1 });
                let headerRowIndex = -1, headerRow = [];

                // Robust Header Search
                for (let i = 0; i < Math.min(rowsAsArray.length, 20); i++) { // Increased search depth
                    const row = rowsAsArray[i];
                    if (row) {
                        const rowStr = row.map(h => String(h || '').toUpperCase().trim());
                        if (rowStr.includes('NO') && (rowStr.includes('NAMA WP') || rowStr.includes('NAMA WAJIB PAJAK'))) {
                            headerRowIndex = i;
                            headerRow = row.map(h => String(h || '').trim());
                            break;
                        }
                    }
                }
                if (headerRowIndex === -1) throw new Error("Header tidak ditemukan.");

                // PERUBAHAN: Menambahkan kolom tambahan
                const colIndices = {
                    no: headerRow.findIndex(h => h.toUpperCase() === 'NO'),
                    noStbp: headerRow.findIndex(h => h.toUpperCase().includes('STBP')),
                    nama: headerRow.findIndex(h => h.toUpperCase().includes('NAMA WP')),
                    alamat: headerRow.findIndex(h => h.toUpperCase().includes('ALAMAT')),
                    nop: headerRow.findIndex(h => h.toUpperCase() === 'NOP'),
                    pembayaran: headerRow.findIndex(h => h.toUpperCase() === 'PEMBAYARAN' || h.toUpperCase() === 'JUMLAH BAYAR'),
                    denda: headerRow.findIndex(h => h.toUpperCase() === 'DENDA'),
                    notaris: headerRow.findIndex(h => h.toUpperCase() === 'NOTARIS'),
                    ket: headerRow.findIndex(h => h.toUpperCase() === 'KET' || h.toUpperCase() === 'KETERANGAN')
                };

                resultData = rowsAsArray.slice(headerRowIndex + 1).map(row => ({
                    'No': row[colIndices.no],
                    'NoSTBP': row[colIndices.noStbp],
                    'Nama': row[colIndices.nama],
                    'Alamat': row[colIndices.alamat],
                    'NOP': row[colIndices.nop],
                    'Pembayaran': row[colIndices.pembayaran],
                    'Denda': row[colIndices.denda],
                    'Notaris': row[colIndices.notaris],
                    'Keterangan': row[colIndices.ket]
                })).filter(row => row.No != null && String(row.No).trim() !== "");
                document.getElementById('countData2').textContent = `${resultData.length} Data`;
            }
            showToast(`File "${file.name}" berhasil diproses.`, 'success');
            callback(resultData);
        } catch (error) {
            console.error(error);
            showToast("Gagal memuat file: " + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

document.getElementById("fileMonitoring").addEventListener("change", function (e) {
    const labelElement = document.getElementById('uploadLabelMonitoring');
    processFile(e.target.files[0], labelElement, (result) => {
        dataMonitoringList = result.sort((a, b) => (b.tgl || '').localeCompare(a.tgl || ''));
        combineAndRenderData();
    });
});

document.getElementById("fileData2").addEventListener("change", function (e) {
    const labelElement = document.getElementById('uploadLabelData2');
    processFile(e.target.files[0], labelElement, (result) => {
        dataKeduaList = result;
        combineAndRenderData();
    });
});

function excelSerialToDateTimeString(serial) {
    if (!serial) return null;

    // Handle Excel Serial Number
    if (typeof serial === 'number' && serial > 1) {
        const days = Math.floor(serial - 25569);
        const date = new Date(days * 86400 * 1000);
        const totalSeconds = Math.floor((serial - Math.floor(serial)) * 86400 + 0.5);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        const pad = (num) => String(num).padStart(2, '0');
        return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())} ${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
    }

    // Handle String Date (e.g. "2024-05-20" or "20/05/2024" or "05/20/2024")
    const s = String(serial).trim();

    // Try ISO format first
    let d = new Date(s);
    if (!isNaN(d.getTime())) {
        // Format to "YYYY-MM-DD HH:mm:ss"
        const pad = (num) => String(num).padStart(2, '0');
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
    }

    // Manual parse for DD/MM/YYYY or DD-MM-YYYY (Indonesian format common)
    // Regex for DD/MM/YYYY or DD-MM-YYYY HH:mm:ss
    const match = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
    if (match) {
        // Assume DD-MM-YYYY
        const day = parseInt(match[1]);
        const month = parseInt(match[2]);
        const year = parseInt(match[3]);
        const hour = match[4] ? parseInt(match[4]) : 0;
        const min = match[5] ? parseInt(match[5]) : 0;
        const sec = match[6] ? parseInt(match[6]) : 0;

        const d2 = new Date(year, month - 1, day, hour, min, sec);
        if (!isNaN(d2.getTime())) {
            const pad = (num) => String(num).padStart(2, '0');
            return `${d2.getFullYear()}-${pad(d2.getMonth() + 1)}-${pad(d2.getDate())} ${pad(d2.getHours())}:${pad(d2.getMinutes())}:${pad(d2.getSeconds())}`;
        }
    }

    // If we are desperate, return original string if it looks like a date?
    // No, consistency is key. Return null if failed.
    return null;
}

function combineAndRenderData() {
    if (dataMonitoringList.length === 0 || dataKeduaList.length === 0) return;
    fullCombinedData = dataKeduaList.map((data2Row, index) => {
        const monitor = dataMonitoringList[index] || {};
        return {
            _uid: generateUID(),
            ...data2Row,
            TanggalPembayaran: monitor.tgl || '',
            TempatPembayaran: monitor.tmpt || ''
        };
    });

    // Generate UIDs for tracking (if not already done or to ensure uniqueness/format)
    // Note: The map above already generates _uid using generateUID().
    // This forEach loop will overwrite those UIDs with a new format.
    fullCombinedData.forEach((row, index) => {
        row._uid = `row-${Date.now()}-${index}`;
    });

    // Save backup for Restore/Undelete
    originalDataBackup = [...fullCombinedData];

    populateDateFilter(fullCombinedData);
    populateKeteranganFilter(fullCombinedData);
    applyFilters();

    document.getElementById('tglSelect').disabled = false;
    document.getElementById('waktuSelect').disabled = false;
    document.getElementById('kategoriSelect').disabled = false;
    // document.getElementById('detailSelect').disabled = false; 

    document.getElementById('btnReset').disabled = false;
    document.querySelectorAll('input[name="fitur"]').forEach(i => i.disabled = false);
    document.getElementById('tambahanCount').disabled = false;
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('btnSaveResult').disabled = false;

    // Enable Tom Selects
    if (tomSelectDetail) tomSelectDetail.enable();

    setTimeout(() => toggleModal(false), 800);
}

// --- Tom Select Logic ---
let tomSelectDetail = null;

function initTomSelects() {
    if (tomSelectDetail) tomSelectDetail.destroy();

    tomSelectDetail = new TomSelect('#detailSelect', {
        create: false,
        sortField: { field: "text", direction: "asc" },
        placeholder: "Pilih Detail Keterangan...",
        plugins: ['dropdown_input', 'remove_button'],
        valueField: 'value',
        labelField: 'text',
        searchField: 'text',
        maxItems: null // Multi-select
    });

    // Disable initially until data loaded
    tomSelectDetail.disable();
}

// Initialize immediately so UI looks good
document.addEventListener('DOMContentLoaded', () => {
    initTomSelects();
});

function updateDetailOptions(data) {
    if (!tomSelectDetail) initTomSelects();

    // Populate with ALL unique options
    tomSelectDetail.clear();
    tomSelectDetail.clearOptions();
    // We don't add "Semua Detail" option in multi-select usually, 
    // empty selection implies "All".

    const unique = [...new Set(data.map(r => (r.Keterangan || '').trim()).filter(x => x))];
    unique.sort();

    unique.forEach(ket => {
        tomSelectDetail.addOption({ value: ket, text: ket });
    });

    tomSelectDetail.refreshOptions(false);
}

function applyFilters() {
    const tanggalMode = document.querySelector('input[name="tanggalMode"]:checked')?.value || 'single';
    const selectedDate = document.getElementById("tglSelect").value;
    const startDate = document.getElementById('tglStart')?.value || '';
    const endDate = document.getElementById('tglEnd')?.value || '';
    const selectedTime = document.getElementById("waktuSelect").value;

    // Get Multi-Select Values
    const selectedDetails = tomSelectDetail ? tomSelectDetail.getValue() : []; // Returns array

    const fiturMode = document.querySelector('input[name="fitur"]:checked')?.value || 'tambah';
    let tambahanCount = parseInt(document.getElementById('tambahanCount')?.value || '0', 10);
    if (tambahanCount < 0) tambahanCount = 0;

    let baseFiltered = (fullCombinedData || []).slice().sort((a, b) => (b.TanggalPembayaran || '').localeCompare(a.TanggalPembayaran || ''));

    // Date Filter
    if (tanggalMode === 'single') {
        if (selectedDate) baseFiltered = baseFiltered.filter(row => (row.TanggalPembayaran || '').startsWith(selectedDate));
    } else {
        if (startDate && endDate) {
            baseFiltered = baseFiltered.filter(row => {
                const d = (row.TanggalPembayaran || '').split(' ')[0];
                return d && d >= startDate && d <= endDate;
            });
        } else if (startDate) {
            baseFiltered = baseFiltered.filter(row => (row.TanggalPembayaran || '').startsWith(startDate));
        }
    }

    // Detail Filter (Multi-select)
    // If specific details are selected, filter to ONLY those.
    // If NOTHING is selected, show ALL data (standard filter behavior).
    if (selectedDetails.length > 0) {
        baseFiltered = baseFiltered.filter(row => selectedDetails.includes((row.Keterangan || "").trim()));
    }

    if (selectedTime === 'semua') {
        finalizeRender(baseFiltered);
        return;
    }

    const matchFn = (row) => {
        const timePart = (row.TanggalPembayaran || '').split(' ')[1];
        if (!timePart) return false;
        return selectedTime === 'kurang' ? timePart < '17:00:00' : timePart >= '17:00:00';
    };
    const groupMatch = baseFiltered.filter(matchFn);
    const groupOther = baseFiltered.filter(r => !matchFn(r));

    let finalData = [];
    if (fiturMode === 'tambah') {
        const extra = groupOther.slice(0, tambahanCount);
        finalData = groupMatch.concat(extra).sort((a, b) => (b.TanggalPembayaran || '').localeCompare(a.TanggalPembayaran || ''));
    } else {
        finalData = groupMatch.slice(tambahanCount).sort((a, b) => (b.TanggalPembayaran || '').localeCompare(a.TanggalPembayaran || ''));
    }
    finalizeRender(finalData);
}

function finalizeRender(data) {
    lastFilteredData = data.slice();
    currentRenderData = data;
    currentPage = 1;
    renderPage();
    updateSummaryCards(data);
    document.getElementById('jumlahData').textContent = `Data: ${data.length}`;
}

function isPemberianHakBaru(keterangan) {
    if (!keterangan) return false;
    const s = String(keterangan).trim().toLowerCase();
    // Specific match as requested by user logic
    return s.includes('pemberian hak baru');
}

// --- Selection & Bulk Delete Logic ---
let selectedRowIds = new Set();

function toggleRowSelection(id) {
    if (selectedRowIds.has(id)) selectedRowIds.delete(id);
    else selectedRowIds.add(id);
    renderPage(); // Re-render to update checkbox UI (or just update toggle, but re-render is safer for syncing)
    updateDeleteButtonState();
}

function toggleSelectAll(checked) {
    const pageData = getCurrentPageData();
    pageData.forEach(row => {
        if (checked) selectedRowIds.add(row._uid);
        else selectedRowIds.delete(row._uid);
    });
    renderPage();
    updateDeleteButtonState();
}

function updateDeleteButtonState() {
    const btn = document.getElementById('btnDeleteSelected');
    const countSpan = document.getElementById('selectedCount');
    if (selectedRowIds.size > 0) {
        btn.classList.remove('hidden');
        countSpan.textContent = selectedRowIds.size;
    } else {
        btn.classList.add('hidden');
    }
}

document.getElementById('btnDeleteSelected')?.addEventListener('click', () => {
    if (selectedRowIds.size === 0) return;
    if (!confirm(`Hapus ${selectedRowIds.size} data terpilih?`)) return;

    // Remove from fullCombinedData
    fullCombinedData = fullCombinedData.filter(row => !selectedRowIds.has(row._uid));

    // Clear selection
    selectedRowIds.clear();

    // Re-apply filters to update view
    applyFilters();
    showToast('Data berhasil dihapus.', 'success');
    updateDeleteButtonState();
});

function getCurrentPageData() {
    const dataRows = currentRenderData;
    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    return dataRows.slice(startIdx, endIdx);
}

function renderPage() {
    const dataRows = currentRenderData;
    const container = document.getElementById('excel-table');
    const placeholder = document.getElementById('table-placeholder');
    const controls = document.getElementById('paginationControls');
    const wrapper = document.getElementById('table-scroll-container');

    if (dataRows.length === 0) {
        container.innerHTML = '';
        placeholder.classList.remove('hidden');
        placeholder.innerHTML = '<i class="bi bi-search text-4xl mb-3 opacity-50"></i><p class="text-sm">Tidak ada data yang cocok dengan filter.</p>';
        controls.classList.add('hidden');
        return;
    } else {
        placeholder.classList.add('hidden');
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

    // Check if all on current page are selected
    const allSelected = pageData.length > 0 && pageData.every(r => selectedRowIds.has(r._uid));

    let html = `<thead class="bg-gray-100 dark:bg-slate-800 sticky top-0 z-10 shadow-sm text-xs text-slate-500 dark:text-slate-400 uppercase"><tr>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-center w-10">
                    <input type="checkbox" onclick="toggleSelectAll(this.checked)" ${allSelected ? 'checked' : ''} class="rounded border-slate-300 text-pink-500 focus:ring-pink-500 cursor-pointer">
                </th>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">No</th>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">Nama WP</th>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Pembayaran</th>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Denda</th>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">Keterangan</th>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">Tempat Pembayaran</th>
                <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">Tanggal Bayar</th>
            </tr></thead><tbody class="divide-y divide-gray-100 dark:divide-slate-800 text-sm text-slate-600 dark:text-slate-300">`;

    pageData.forEach((row) => {
        const isSelected = selectedRowIds.has(row._uid);
        const pembayaran = typeof row.Pembayaran === 'number' ? row.Pembayaran.toLocaleString('id-ID') : (row.Pembayaran || 0);
        const denda = typeof row.Denda === 'number' ? row.Denda.toLocaleString('id-ID') : (row.Denda || 0);

        html += `<tr class="hover:bg-blue-50/50 dark:hover:bg-slate-800 transition-colors ${isSelected ? 'bg-blue-50 dark:bg-slate-800/50' : ''}">
                    <td class="p-3 text-center">
                        <input type="checkbox" onclick="toggleRowSelection('${row._uid}')" ${isSelected ? 'checked' : ''} class="rounded border-slate-300 text-pink-500 focus:ring-pink-500 cursor-pointer">
                    </td>
                    <td class="p-3 font-medium text-slate-400 text-xs">${escapeHtml(row.No || '')}</td>
                    <td class="p-3 font-medium">${escapeHtml(row.Nama || '')}</td>
                    <td class="p-3 text-right font-medium text-slate-700 dark:text-slate-200">${pembayaran}</td>
                    <td class="p-3 text-right text-amber-600 dark:text-amber-500">${denda}</td>
                    <td class="p-3">${escapeHtml(row.Keterangan || '')}</td>
                    <td class="p-3 text-slate-500 dark:text-slate-400">${escapeHtml(row.TempatPembayaran || '-')}</td>
                    <td class="p-3 font-mono text-xs text-slate-500">${escapeHtml(row.TanggalPembayaran || '')}</td>
                </tr>`;
    });
    html += `</tbody>`;
    container.innerHTML = html;
    wrapper.scrollTop = 0;
}

document.getElementById("prevPageBtn").addEventListener("click", () => { if (currentPage > 1) { currentPage--; renderPage(); } });
document.getElementById("nextPageBtn").addEventListener("click", () => { const totalPages = Math.ceil(currentRenderData.length / rowsPerPage); if (currentPage < totalPages) { currentPage++; renderPage(); } });

function updateSummaryCards(data) {
    let totalPembayaran = 0, totalDenda = 0;
    (data || []).forEach(row => {
        totalPembayaran += parseNumber(row.Pembayaran);
        totalDenda += parseNumber(row.Denda);
    });
    document.getElementById('totalPembayaran').textContent = totalPembayaran.toLocaleString('id-ID');
    document.getElementById('totalDenda').textContent = totalDenda.toLocaleString('id-ID');
    document.getElementById('totalMutasi').textContent = (totalPembayaran + totalDenda).toLocaleString('id-ID');

    const bars = document.querySelectorAll('#summaryCards .bg-slate-100, #summaryCards .dark\\:bg-slate-800 > div');
    bars.forEach(bar => { bar.style.width = '0%'; setTimeout(() => bar.style.width = '100%', 100); });
}

function populateDateFilter(data) {
    const sel = document.getElementById('tglSelect');
    sel.innerHTML = '<option value="">Pilih Tanggal</option>';
    const counts = {};
    (data || []).forEach(r => {
        const tp = (r.TanggalPembayaran || '').toString().split(' ')[0];
        if (tp) counts[tp] = (counts[tp] || 0) + 1;
    });
    Object.keys(counts).sort((a, b) => b.localeCompare(a)).forEach(date => {
        const opt = document.createElement('option');
        opt.value = date;
        opt.textContent = `${date} (${counts[date]})`;
        sel.appendChild(opt);
    });

    const allDates = Object.keys(counts).sort();
    const startInput = document.getElementById('tglStart');
    const endInput = document.getElementById('tglEnd');
    if (allDates.length) {
        startInput.min = allDates[0]; startInput.max = allDates[allDates.length - 1];
        endInput.min = allDates[0]; endInput.max = allDates[allDates.length - 1];
        startInput.disabled = false; endInput.disabled = false;
    } else {
        startInput.disabled = true; endInput.disabled = true;
    }
}

function populateKeteranganFilter(data) {
    // Populate Detail options with ALL unique data
    updateDetailOptions(data);
}

function copyNominal(elementId) {
    const element = document.getElementById(elementId);
    if (!element) return;
    const rawValue = parseNumber(element.textContent).toString();

    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(rawValue).then(() => {
            showToast(`Nominal ${element.textContent} disalin.`, 'success');
        }).catch(err => {
            fallbackCopyTextToClipboard(rawValue, element.textContent);
        });
    } else {
        fallbackCopyTextToClipboard(rawValue, element.textContent);
    }
}

function fallbackCopyTextToClipboard(text, formattedText) {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.top = "0";
    textArea.style.left = "0";
    textArea.style.position = "fixed";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    try {
        const successful = document.execCommand('copy');
        if (successful) showToast(`Nominal ${formattedText} disalin.`, 'success');
        else showToast("Gagal menyalin nominal.", 'error');
    } catch (err) {
        showToast("Gagal menyalin nominal.", 'error');
    }
    document.body.removeChild(textArea);
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

document.getElementById('exportBtn')?.addEventListener('click', function () {
    if (!lastFilteredData || lastFilteredData.length === 0) { showToast('Tidak ada data.', 'error'); return; }

    // PERUBAHAN: Ekspor data lengkap
    const rows = lastFilteredData.map(r => ({
        'No': r.No || '',
        'No STBP': r.NoSTBP || '',
        'Nama Wajib Pajak': r.Nama || '',
        'Alamat Wajib Pajak': r.Alamat || '',
        'NOP': r.NOP || '',
        'Pembayaran': parseNumber(r.Pembayaran),
        'Denda': parseNumber(r.Denda),
        'Notaris': r.Notaris || '',
        'Keterangan': r.Keterangan || '',
        'Tempat Pembayaran': r.TempatPembayaran || '',
        'Tanggal Pembayaran': r.TanggalPembayaran || ''
    }));

    try {
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'DataGabungan');
        const now = new Date();
        const fname = `SiRekonBPHTB_${now.toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fname);
        showToast('Data berhasil diekspor.', 'success');
    } catch (error) { console.error(error); showToast('Gagal ekspor.', 'error'); }
});

document.getElementById("tglSelect").addEventListener("change", applyFilters);
document.querySelectorAll('input[name="tanggalMode"]').forEach(r => {
    r.addEventListener('change', () => {
        const mode = document.querySelector('input[name="tanggalMode"]:checked').value;
        document.getElementById('tanggalSingle').classList.toggle('hidden', mode !== 'single');
        document.getElementById('tanggalRange').classList.toggle('hidden', mode !== 'range');
        if (mode === 'single') { document.getElementById('tglStart').value = ''; document.getElementById('tglEnd').value = ''; }
        else { document.getElementById('tglSelect').value = ''; }
        applyFilters();
    });
});
document.getElementById('tglStart')?.addEventListener('change', applyFilters);
document.getElementById('tglEnd')?.addEventListener('change', applyFilters);
document.getElementById("waktuSelect").addEventListener("change", applyFilters);

// New Event Listeners
document.getElementById("kategoriSelect").addEventListener("change", (e) => {
    const category = e.target.value;
    if (!category || !tomSelectDetail || !fullCombinedData) return;

    // Find options matching category
    // Note: This logic assumes 'fullCombinedData' has all data.
    const uniqueMatches = [...new Set(fullCombinedData.map(r => (r.Keterangan || '').trim()).filter(x => {
        if (!x) return false;
        const isHakBaru = isPemberianHakBaru(x);
        if (category === 'hak_baru') return isHakBaru;
        if (category === 'pemindahan') return !isHakBaru;
        return false;
    }))];

    tomSelectDetail.setValue(uniqueMatches);
});

document.getElementById("detailSelect").addEventListener("change", applyFilters);

document.getElementById('fiturRadioGroup')?.addEventListener('change', () => applyFilters());
document.getElementById('tambahanCount')?.addEventListener('input', () => applyFilters());

document.getElementById("btnReset").addEventListener("click", function () {
    // Reset Filters ONLY
    document.getElementById("tglSelect").value = "";
    document.getElementById("waktuSelect").value = "semua";

    // Reset Inputs
    document.getElementById("kategoriSelect").value = "";
    if (tomSelectDetail) {
        tomSelectDetail.clear();
        tomSelectDetail.clearOptions();

        // Re-populate options with all data (since we are resetting)
        if (fullCombinedData && fullCombinedData.length > 0) {
            updateDetailOptions(fullCombinedData);
            tomSelectDetail.enable(); // Ensure it's enabled if data exists
        } else {
            tomSelectDetail.disable(); // Disable if no data
        }
    }

    document.getElementById('tglStart').value = '';
    document.getElementById('tglEnd').value = '';

    // Clear Selection
    selectedRowIds.clear();
    updateDeleteButtonState();

    // Check if we have data to keep enabled
    if (fullCombinedData && fullCombinedData.length > 0) {
        // RESTORE DATA from backup (Undelete)
        if (originalDataBackup && originalDataBackup.length > 0) {
            fullCombinedData = [...originalDataBackup];
            showToast('Data dikembalikan ke kondisi awal (undelete).', 'success');
        }

        // Re-apply filters (which creates 'baseFiltered' from fullCombinedData)
        // Since filters are empty, it will show ALL data.
        applyFilters();
        showToast('Filter direset.', 'success');
        // Re-enable filter controls if data exists
        ['tglSelect', 'tglStart', 'tglEnd', 'waktuSelect', 'kategoriSelect', 'tambahanCount', 'exportBtn', 'btnSaveResult'].forEach(id => {
            const el = document.getElementById(id); if (el) el.disabled = false;
        });
        document.querySelectorAll('input[name="fitur"]').forEach(i => i.disabled = false);
        document.getElementById('btnReset').disabled = false; // Always enabled if data is present
    } else {
        // If no data was loaded, then we do full reset
        document.getElementById("fileMonitoring").value = '';
        document.getElementById("fileData2").value = '';
        document.getElementById("uploadLabelMonitoring").innerHTML = '<div class="w-10 h-10 rounded-full bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 flex items-center justify-center text-lg flex-shrink-0"><i class="bi bi-file-earmark-bar-graph"></i></div><div class="flex-1 min-w-0"><p class="text-sm font-bold text-slate-700 dark:text-slate-200 truncate">File Monitoring</p><p class="text-xs text-slate-400 mt-0.5">Klik untuk pilih file</p></div><div class="text-xs font-medium px-2 py-1 bg-slate-200 dark:bg-slate-700 rounded text-slate-600 dark:text-slate-300" id="countMonitoring">0 Data</div>';
        document.getElementById("uploadLabelData2").innerHTML = '<div class="w-10 h-10 rounded-full bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400 flex items-center justify-center text-lg flex-shrink-0"><i class="bi bi-file-earmark-spreadsheet"></i></div><div class="flex-1 min-w-0"><p class="text-sm font-bold text-slate-700 dark:text-slate-200 truncate">File Data Ke-2</p><p class="text-xs text-slate-400 mt-0.5">Klik untuk pilih file</p></div><div class="text-xs font-medium px-2 py-1 bg-slate-200 dark:bg-slate-700 rounded text-slate-600 dark:text-slate-300" id="countData2">0 Data</div>';

        dataMonitoringList = []; dataKeduaList = []; fullCombinedData = []; currentRenderData = []; originalDataBackup = [];
        document.getElementById("totalMonitoring").textContent = '0';
        document.getElementById("totalData2").textContent = '0';

        ['tglSelect', 'tglStart', 'tglEnd', 'waktuSelect', 'kategoriSelect', 'btnReset', 'tambahanCount', 'exportBtn'].forEach(id => {
            const el = document.getElementById(id); if (el) el.disabled = true;
        });

        document.querySelectorAll('input[name="fitur"]').forEach(i => i.disabled = true);

        document.getElementById('table-placeholder').classList.remove('hidden');
        document.getElementById('table-placeholder').innerHTML = '<i class="bi bi-arrow-counterclockwise text-4xl mb-3 opacity-50"></i><p class="text-sm">Filter telah direset.</p>';
        document.getElementById('paginationControls').classList.add('hidden');
        document.getElementById('excel-table').innerHTML = '';

        updateSummaryCards([]);
        document.getElementById('jumlahData').textContent = 'Data: 0';
        document.getElementById('btnSaveResult').disabled = true;
    }
});

// --- Save / Load Logic ---
const savedDataModal = document.getElementById('savedDataModal');
const openSavedDataBtn = document.getElementById('openSavedDataBtn');
const closeSavedModal = document.getElementById('closeSavedModal');

if (openSavedDataBtn) openSavedDataBtn.addEventListener('click', () => {
    renderSavedDataList();
    savedDataModal.classList.remove('hidden');
    setTimeout(() => {
        savedDataModal.classList.remove('opacity-0');
        savedDataModal.children[0].classList.remove('scale-95');
        savedDataModal.children[0].classList.add('scale-100');
    }, 10);
});

if (closeSavedModal) closeSavedModal.addEventListener('click', () => {
    savedDataModal.classList.add('opacity-0');
    savedDataModal.children[0].classList.remove('scale-100');
    savedDataModal.children[0].classList.add('scale-95');
    setTimeout(() => savedDataModal.classList.add('hidden'), 200);
});

document.getElementById('btnSaveResult')?.addEventListener('click', async () => {
    if (!lastFilteredData || lastFilteredData.length === 0) return;
    // Using prompt for simplicity or Swal if available. Index.html does not seem to include Swal in head?
    // Actually PBB used Swal. BPHTB index.html (checked earlier) shows no Swal script tag in head (only tailwind, fonts, xlsx, bootstrap-icons).
    // I'll use native prompt to be safe, or check if I can add Swal. 
    // User request said "sama seperti sirekonpbb". SiRekonPBB uses Swal.
    // But if I don't add Swal library, it will break. 
    // I'll stick to prompt for safety unless I see Swal included.
    // Checked head: lines 9-51. No Swal.
    // I will use `prompt`.

    const name = prompt("Beri nama untuk data ini:", "Hasil BPHTB " + new Date().toLocaleDateString('id-ID'));
    if (!name) return;

    const saveData = {
        id: Date.now(),
        name: name,
        date: new Date().toISOString(),
        rows: lastFilteredData // Saving just the filtered rows
    };

    const existing = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    existing.push(saveData);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));
    showToast('Data berhasil disimpan!', 'success');
});

function renderSavedDataList() {
    const list = document.getElementById('savedDataList');
    const empty = document.getElementById('emptySaved');
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');

    list.innerHTML = '';
    if (saved.length === 0) {
        empty.classList.remove('hidden');
        return;
    }
    empty.classList.add('hidden');

    saved.sort((a, b) => new Date(b.date) - new Date(a.date)).forEach(item => {
        const d = new Date(item.date);
        const dateStr = d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors group";
        tr.innerHTML = `
                    <td class="px-6 py-3 font-medium text-slate-700 dark:text-white">${escapeHtml(item.name)}</td>
                    <td class="px-6 py-3 text-center text-slate-500 text-xs">${dateStr}</td>
                    <td class="px-6 py-3 text-center"><span class="px-2 py-1 bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 rounded-lg text-xs font-bold">${item.rows.length}</span></td>
                    <td class="px-6 py-3 text-right">
                        <div class="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                            <button onclick="loadSavedData(${item.id})" class="p-1.5 text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors" title="Buka"><i class="bi bi-folder2-open"></i></button>
                            <button onclick="deleteSavedData(${item.id})" class="p-1.5 text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/30 rounded-lg transition-colors" title="Hapus"><i class="bi bi-trash"></i></button>
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

    fullCombinedData = item.rows.map(r => ({
        ...r,
        _uid: r._uid || generateUID() // Ensure UID exists
    }));

    // Populate filters based on loaded data
    populateDateFilter(fullCombinedData);
    populateKeteranganFilter(fullCombinedData);
    applyFilters();

    document.getElementById('tglSelect').disabled = false;
    document.getElementById('waktuSelect').disabled = false;
    document.getElementById('keteranganSelect').disabled = false;
    document.getElementById('btnReset').disabled = false;
    document.querySelectorAll('input[name="fitur"]').forEach(i => i.disabled = false);
    document.getElementById('tambahanCount').disabled = false;
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('btnSaveResult').disabled = false; // Loaded data can be saved again as new

    // Close modal
    savedDataModal.classList.add('opacity-0');
    setTimeout(() => savedDataModal.classList.add('hidden'), 200);

    showToast(`Data "${item.name}" dimuat.`, 'success');
};

window.deleteSavedData = function (id) {
    if (!confirm("Hapus data tersimpan ini?")) return;
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    const newSaved = saved.filter(x => x.id !== id);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(newSaved));
    renderSavedDataList();
    showToast('Data dihapus.', 'success');
};

window.deleteRow = function (uid) {
    if (!confirm("Hapus baris ini dari tampilan?")) return;
    // Remove from fullCombinedData
    // Note: If we just remove from fullCombinedData and re-apply filters, it works.
    const idx = fullCombinedData.findIndex(r => r._uid === uid);
    if (idx > -1) {
        fullCombinedData.splice(idx, 1);
        applyFilters();
        showToast('Baris dihapus.', 'success');
    }
};

document.getElementById('btnClearAll')?.addEventListener('click', () => {
    if (!confirm("Hapus SEMUA riwayat data tersimpan?")) return;
    localStorage.removeItem(STORAGE_KEY);
    renderSavedDataList();
    showToast('Semua data dihapus.', 'success');
});

document.getElementById('btnDownloadAll')?.addEventListener('click', () => {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    if (saved.length === 0) { showToast('Tidak ada data.', 'error'); return; }

    let allRows = [];
    saved.forEach(item => {
        item.rows.forEach(r => {
            allRows.push({
                'Nama Laporan': item.name,
                'Waktu Simpan': new Date(item.date).toLocaleString('id-ID'),
                'No': r.No,
                'Nama WP': r.Nama,
                'Pembayaran': parseNumber(r.Pembayaran),
                'Denda': parseNumber(r.Denda),
                'Keterangan': r.Keterangan,
                'Tempat Pembayaran': r.TempatPembayaran,
                'Tanggal Pembayaran': r.TanggalPembayaran
            });
        });
    });

    const ws = XLSX.utils.json_to_sheet(allRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'AllSavedData');
    XLSX.writeFile(wb, `Backup_SiRekonBPHTB_Merged.xlsx`);
});
