let allData = [];
let currentTableData = []; // Data yang sedang difilter
let currentPage = 1;
const rowsPerPage = 50; // Pagination limit untuk performa ringan
let originalDataBackup = []; // Backup for undelete
let selectedRowIds = new Set(); // Selection tracking

// Key for LocalStorage
const STORAGE_KEY = 'sipedasi:saved_filter_data';

const kategoriBankMap = {
    PBBU: [
        "Bank BJB - Lainnya", "Bank BJB - Mobile Banking", "Bank BJB - POS",
        "Bank BJB - Teller", "Bank BJB - Internet Banking", "Bank BJB - ATM"
    ],
    PBBQRIS: ["Bank BJB - QRIS"],
    PBBVA: ["Bank BJB - Virtual Account"],
    PBBKOLEKTIF: ["Loket Bank BJB - Kolektif"],
    EPBB: [
        "Bank Mandiri - Internet Banking", "Bank Mandiri - Mobile Banking",
        "Bank Mandiri - Teller", "Bank Mandiri - POS"
    ],
    POS: ["PT. POS"]
};

// --- TOM SELECT LOGIC ---
let tomSelectInstance = null;

function initTomSelect(options = []) {
    if (tomSelectInstance) {
        tomSelectInstance.destroy();
        tomSelectInstance = null;
    }

    tomSelectInstance = new TomSelect("#bankDetailSelect", {
        plugins: ['remove_button', 'dropdown_input'],
        create: false,
        placeholder: "Detail bank akan terisi otomatis...",
        onItemAdd: function () { refreshAll(); updateDeleteButton(); },
        onItemRemove: function () { refreshAll(); updateDeleteButton(); },
        onClear: function () { refreshAll(); updateDeleteButton(); }
    });

    if (options && options.length > 0) {
        tomSelectInstance.clearOptions();
        options.forEach(bank => {
            tomSelectInstance.addOption({ value: bank, text: bank });
        });
        tomSelectInstance.refreshOptions(false);
        tomSelectInstance.enable();
    } else {
        tomSelectInstance.disable();
    }
}



// --- Event Listeners & Core Logic ---
document.addEventListener('DOMContentLoaded', () => {
    // Init Tom Select Immediately
    initTomSelect([]);


    // Theme Logic
    const themeToggle = document.getElementById('themeToggle');
    const themeIcon = document.getElementById('themeIcon');
    const html = document.documentElement;

    const savedTheme = localStorage.getItem('sipedasi:theme');
    const systemDark = window.matchMedia('(prefers-color-scheme: dark)').matches;

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

    setTheme(savedTheme === 'dark' || (!savedTheme && systemDark));

    themeToggle.addEventListener('click', () => {
        setTheme(!html.classList.contains('dark'));
    });

    // Upload Modal Logic
    const fabUpload = document.getElementById('fabUpload');
    const uploadModal = document.getElementById('uploadModal');
    const closeModal = document.getElementById('closeModal');
    const btnCancelModal = document.getElementById('btnCancelModal');
    const fileInput = document.getElementById("fileExcel");
    const uploadLabel = document.getElementById("uploadLabel");

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

    if (fabUpload) fabUpload.addEventListener('click', () => toggleModal(uploadModal, true));
    if (closeModal) closeModal.addEventListener('click', () => toggleModal(uploadModal, false));
    if (btnCancelModal) btnCancelModal.addEventListener('click', () => toggleModal(uploadModal, false));
    if (uploadModal) uploadModal.addEventListener('click', (e) => {
        if (e.target === uploadModal) toggleModal(uploadModal, false);
    });

    // Saved Data Modal Logic
    const openSavedDataBtn = document.getElementById('openSavedDataBtn');
    const savedDataModal = document.getElementById('savedDataModal');
    const closeSavedModal = document.getElementById('closeSavedModal');

    if (openSavedDataBtn) openSavedDataBtn.addEventListener('click', () => {
        renderSavedDataList();
        toggleModal(savedDataModal, true);
    });
    if (closeSavedModal) closeSavedModal.addEventListener('click', () => toggleModal(savedDataModal, false));
    if (savedDataModal) savedDataModal.addEventListener('click', (e) => {
        if (e.target === savedDataModal) toggleModal(savedDataModal, false);
    });


    // File Upload Handler
    if (fileInput) fileInput.addEventListener("change", function (e) {
        const file = e.target.files[0];

        if (!file) {
            resetUploadLabel();
            return;
        }

        // Show loading state
        uploadLabel.innerHTML = `
            <div class="spinner border-blue-500 border-t-transparent w-8 h-8 mb-2"></div>
            <span class="text-sm font-medium text-blue-600 animate-pulse">Memproses Data...</span>
        `;

        const reader = new FileReader();
        reader.onload = function (event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: "array" });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "", cellDates: true });

                if (jsonData.length === 0) {
                    throw new Error("File kosong");
                }

                // User Specific Column Mapping
                // POKOK -> Header "POKOK PEMBAYARAN" (Col S usually)
                // DENDA -> Header "DENDA" (Col U usually)
                // TOTAL -> Header "TOTAL" (Col V usually)

                const keys = Object.keys(jsonData[0]);
                const findDirectKey = (name) => keys.find(k => k.toUpperCase().trim() === name.toUpperCase());
                const findPartialKey = (keywords) => keys.find(k => keywords.some(kw => k.toUpperCase().includes(kw)));

                const keyPokok = findDirectKey("POKOK PEMBAYARAN") || findPartialKey(["POKOK", "TAGIHAN", "NOMINAL"]);
                const keyDenda = findDirectKey("DENDA") || findPartialKey(["DENDA", "SANKSI", "ADM", "BIAYA"]);
                const keyTotal = findDirectKey("TOTAL") || findPartialKey(["TOTAL", "JUMLAH", "BAYAR", "SETORAN"]);
                // keyBank matches user request context or fallback
                const keyBank = findPartialKey(["BANK", "TEMPAT", "VIA", "CHANNEL"]);

                const keyTgl = keys.find(k => {
                    const upper = k.toUpperCase();
                    if ((upper.includes("TANGGAL") || upper.includes("TGL")) && upper.includes("BAYAR")) return true;
                    if ((upper.includes("TANGGAL") || upper.includes("TGL")) && !upper.includes("JATUH") && !upper.includes("TEMPO")) return true;
                    return false;
                });

                console.log("Detected Keys:", { keyPokok, keyDenda, keyTotal, keyBank, keyTgl });

                if (!keyPokok || !keyBank) {
                    console.warn("Could not detect crucial columns.");
                }

                // Robust Number Parser: handles "10.000", "10,000", "Rp 10.000"
                const parseNumber = (val) => {
                    if (typeof val === 'number') return val;
                    if (val instanceof Date) return 0; // Should not happen for number cols
                    if (!val) return 0;
                    // Remove currency symbols, cleanup whitespace
                    let clean = String(val).replace(/[^0-9.,-]/g, '').trim();
                    if (!clean) return 0;

                    // Detect format: 
                    // If it has commas but no dots (10,000) -> assume comma is thousand sep? OR decimal?
                    // In Indonesia context: dot is thousand, comma is decimal.
                    // Let's assume Indonesia standard "10.000" = 10000.
                    // Replace dot with nothing, replace comma with dot.

                    // Simple heuristic: if dot is present, remove it. If comma is present, replace with dot.
                    // BUT be careful with "1000.50" (english) vs "1.000,50" (indo).
                    // Safe bet: if value contains points and the LAST point is before the LAST comma -> Indo
                    // If value contains commas and the LAST comma is before the LAST point -> US

                    // Simplified for this context (PBB data usually Indo/Excel format):
                    // Remove all points (thousand sep), replace comma with dot (decimal).
                    clean = clean.replace(/\./g, '').replace(/,/g, '.');

                    return parseFloat(clean) || 0;
                };

                // Date Formatter Helper
                const formatDate = (val) => {
                    if (!val) return "";
                    if (val instanceof Date) {
                        // Return YYYY-MM-DD HH:mm:ss
                        const pad = (n) => String(n).padStart(2, '0');
                        const y = val.getFullYear();
                        const m = pad(val.getMonth() + 1);
                        const d = pad(val.getDate());
                        const h = pad(val.getHours());
                        const min = pad(val.getMinutes());
                        const s = pad(val.getSeconds());
                        // If time is 00:00:00, maybe omit? The filter expects string.
                        // Current format usually: "2023-10-25 14:30:00"
                        return `${y}-${m}-${d} ${h}:${min}:${s}`;
                    }
                    return String(val);
                };

                allData = jsonData.map((item, index) => {
                    // Check if first column (often used for numbering or 'Total') indicates a footer row
                    const firstVal = Object.values(item)[0];
                    if (firstVal && String(firstVal).toUpperCase().includes("TOTAL")) {
                        return null; // Flag for filter
                    }

                    const pokok = parseNumber(item[keyPokok]);
                    const denda = parseNumber(item[keyDenda]);
                    // If total column missing or invalid, sum manual
                    let total = parseNumber(item[keyTotal]);
                    if (!keyTotal || total === 0) {
                        total = pokok + denda;
                    }

                    return {
                        _id: Date.now() + '_' + index, // Unique ID for row management
                        ...item, // Preserve all original fields
                        pokok: pokok,
                        denda: denda,
                        total: total,
                        pembayaran: String(item[keyBank] || "").replace(/\s*-\s*/g, ' - ').replace(/\s+/g, ' ').trim(),
                        tglpembayaran: formatDate(item[keyTgl])
                    };
                }).filter(item => item !== null); // Removing flagged rows

                // Success UI Update
                uploadLabel.innerHTML = `
                    <div class="w-12 h-12 rounded-full bg-green-100 text-green-600 flex items-center justify-center text-xl mb-1">
                        <i class="bi bi-check-lg"></i>
                    </div>
                    <span class="text-sm font-bold text-slate-800 dark:text-white truncate max-w-[200px]">${file.name}</span>
                    <span class="text-xs text-green-600 font-medium">Berhasil Dimuat</span>
                `;
                showToast(`File "${file.name}" berhasil diunggah.`, 'success');

                // --- SAVE BACKUP FOR UNDELETE ---
                originalDataBackup = [...allData];
                // --------------------------------

                document.getElementById("categorySelect").disabled = false;
                document.getElementById("bankDetailSelect").disabled = false;
                document.getElementById("tglSelect").disabled = false;
                document.getElementById("waktuSelect").disabled = false;
                document.querySelectorAll('input[name="fitur"]').forEach(r => r.disabled = false);
                document.getElementById("tambahanCount").disabled = false;
                document.getElementById("btnReset").disabled = false;

                // --- TOM SELECT INITIALIZATION ---
                // Get unique banks
                const uniqueBanks = [...new Set(allData.map(item => item.pembayaran))].filter(Boolean).sort();

                // Re-initialize with data
                initTomSelect(uniqueBanks);
                // ---------------------------------
                // ---------------------------------

                // --- POPULATE DATES WITH COUNTS ---
                populateDateDropdown();

                const uniqueDates = allData
                    .map(item => item.tglpembayaran?.substring(0, 10))
                    .filter((v, i, a) => a.indexOf(v) === i && v !== undefined && v !== "")
                    .sort();

                // Init Range Inputs
                const tglStart = document.getElementById("tglStart");
                const tglEnd = document.getElementById("tglEnd");

                if (uniqueDates.length > 0) {
                    const minDateObj = parseDateObj(uniqueDates[0]);
                    const maxDateObj = parseDateObj(uniqueDates[uniqueDates.length - 1]);
                    const toISODate = (d) => d ? d.toISOString().split('T')[0] : '';

                    if (minDateObj && maxDateObj) {
                        tglStart.min = toISODate(minDateObj);
                        tglStart.max = toISODate(maxDateObj);
                        tglEnd.min = toISODate(minDateObj);
                        tglEnd.max = toISODate(maxDateObj);

                        tglStart.disabled = false;
                        tglEnd.disabled = false;
                    }
                }
                // ----------------------------------

                // --- AUTO DISPLAY LOGIC ---
                // Immediately show all data
                currentTableData = allData;
                currentPage = 1;

                // Enable Export & Save since we have data
                document.getElementById("exportPbbBtn").disabled = false;
                document.getElementById("btnSaveResult").disabled = false;

                // Render Table & Summary
                renderPage();

                // Show toast for totals
                const totalAmount = allData.reduce((acc, curr) => acc + (curr.total || 0), 0);
                showToast(`Total Data: ${allData.length}, Nominal: Rp ${totalAmount.toLocaleString('id-ID')}`, 'info');
                // --------------------------

                // Close modal after short delay
                setTimeout(() => toggleModal(uploadModal, false), 1000);

            } catch (error) {
                console.error("Error processing file:", error);
                resetUploadLabel();
                showToast("Format file tidak valid atau kolom tidak dikenali.", 'error');
            }
        };

        reader.onerror = function () {
            resetUploadLabel();
            showToast("Gagal membaca file.", 'error');
        };

        reader.readAsArrayBuffer(file);
    });

    function resetUploadLabel() {
        const label = document.getElementById("uploadLabel");
        if (label) {
            label.innerHTML = `
                <div class="w-14 h-14 rounded-full bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 flex items-center justify-center text-2xl">
                    <i class="bi bi-file-earmark-spreadsheet"></i>
                </div>
                <div>
                    <p class="text-sm font-bold text-slate-700 dark:text-slate-200">Klik untuk pilih file</p>
                    <p class="text-xs text-slate-400 mt-1">Format .xlsx atau .xls</p>
                </div>
            `;
        }
    }

    // Category Selection Logic (The "Trigger")
    const categorySelect = document.getElementById("categorySelect");
    if (categorySelect) categorySelect.addEventListener("change", (e) => {
        if (!tomSelectInstance) return;

        const category = e.target.value;
        const targetBanks = kategoriBankMap[category] || [];

        // Ensure we only set banks that exist in the loaded data (avoid phantom selections)
        // Actually user wants "terisi otomatis sesuai bank yang terdaftar", assuming strict map.
        // Tom Select ignores values not in options, so passing the full list is safe.

        if (targetBanks.length > 0) {
            tomSelectInstance.setValue(targetBanks);
            showToast(`Kategori ${category} diterapkan.`, 'success');
        } else {
            // If "Semua Template" or empty
            tomSelectInstance.clear();
        }
        // Refresh happens automatically via Tom Select callbacks
    });



    // Input Event Listeners
    document.getElementById("tglSelect")?.addEventListener("change", refreshTable);
    document.getElementById("tglStart")?.addEventListener("change", refreshTable);
    document.getElementById("tglEnd")?.addEventListener("change", refreshTable);
    document.getElementById("waktuSelect")?.addEventListener("change", refreshTable);
    document.getElementById("tambahanCount")?.addEventListener("input", refreshTable);

    document.querySelectorAll('input[name="tanggalMode"]').forEach(r => {
        r.addEventListener('change', (e) => {
            const mode = e.target.value;
            document.getElementById('tanggalSingle').classList.toggle('hidden', mode !== 'single');
            document.getElementById('tanggalRange').classList.toggle('hidden', mode !== 'range');

            // Clear values when switching
            if (mode === 'single') {
                document.getElementById('tglStart').value = '';
                document.getElementById('tglEnd').value = '';
            } else {
                document.getElementById('tglSelect').value = '';
            }
            refreshTable();
        });
    });

    document.querySelectorAll('input[name="fitur"]').forEach(r => {
        r.addEventListener('change', () => {
            document.getElementById("tambahanCount").value = 0;
            refreshTable();
        });
    });

    // Pagination Controls
    document.getElementById("prevPageBtn")?.addEventListener("click", () => {
        if (currentPage > 1) {
            currentPage--;
            renderPage();
        }
    });
    document.getElementById("nextPageBtn")?.addEventListener("click", () => {
        const totalPages = Math.ceil(currentTableData.length / rowsPerPage);
        if (currentPage < totalPages) {
            currentPage++;
            renderPage();
        }
    });



    // Reset Button
    document.getElementById("btnReset")?.addEventListener("click", () => {
        // Undelete / Restore Data
        if (originalDataBackup && originalDataBackup.length > 0) {
            allData = [...originalDataBackup];
        }

        // Reset Inputs
        document.getElementById("categorySelect").value = "";
        if (tomSelectInstance) {
            tomSelectInstance.clear();
        }

        document.getElementById("tglSelect").value = "";
        document.getElementById("waktuSelect").value = "semua";
        document.getElementById("tambahanCount").value = 0;

        // Reset Date Mode to Single
        const singleMode = document.querySelector('input[name="tanggalMode"][value="single"]');
        if (singleMode) singleMode.click(); // Triggers change event to hide range inputs

        document.getElementById('tglStart').value = '';
        document.getElementById('tglEnd').value = '';

        // Clear Selection
        selectedRowIds.clear();
        updateDeleteButton();

        // Reset Data View (Show All)
        currentTableData = allData;
        currentPage = 1;

        // Refresh Counts
        populateDateDropdown(); // Reset to "All Banks" context

        // Refresh Table
        renderPage();

        showToast("Filter direset & Data dipulihkan.", 'success');
    });

    // Export Button
    document.getElementById('exportPbbBtn')?.addEventListener('click', () => exportDataToExcel(currentTableData));

    // Save Result Button
    document.getElementById('btnSaveResult')?.addEventListener('click', saveFilteredData);
});

// --- Core Functions ---

function parseDateObj(dateStr) {
    if (!dateStr) return null;
    let match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}):?(\d{2})?)?/);
    if (match) {
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]), match[4] ? parseInt(match[4]) : 0, match[5] ? parseInt(match[5]) : 0, match[6] ? parseInt(match[6]) : 0);
    }
    match = dateStr.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:[ T](\d{2}):(\d{2}):?(\d{2})?)?/);
    if (match) {
        return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]), match[4] ? parseInt(match[4]) : 0, match[5] ? parseInt(match[5]) : 0, match[6] ? parseInt(match[6]) : 0);
    }
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? null : d;
}

function populateDateDropdown() {
    const tglSelect = document.getElementById("tglSelect");
    if (!tglSelect) return;

    // Get selected banks from Tom Select
    let selectedBanks = [];
    if (tomSelectInstance) {
        selectedBanks = tomSelectInstance.getValue();
    }

    // Get all unique dates
    const uniqueDates = allData
        .map(item => item.tglpembayaran?.substring(0, 10))
        .filter((v, i, a) => a.indexOf(v) === i && v !== undefined && v !== "")
        .sort();

    const currentVal = tglSelect.value;
    tglSelect.innerHTML = `<option value="">Semua Tanggal</option>`;

    uniqueDates.forEach(date => {
        let count = 0;
        if (selectedBanks.length > 0) {
            count = allData.filter(item =>
                item.tglpembayaran?.substring(0, 10) === date &&
                selectedBanks.includes(item.pembayaran)
            ).length;
        } else {
            count = allData.filter(item => item.tglpembayaran?.substring(0, 10) === date).length;
        }

        const option = document.createElement("option");
        option.value = date;
        option.textContent = `${date} (${count} Data)`;
        tglSelect.appendChild(option);
    });

    if (currentVal && uniqueDates.includes(currentVal)) {
        tglSelect.value = currentVal;
    }
}

function refreshAll() {
    populateDateDropdown();
    refreshTable();
}

function refreshTable() {
    if (!allData || allData.length === 0) return;

    // const bankSelect = document.getElementById("bankSelect"); // Legacy removed
    // const selectedKategori = bankSelect ? bankSelect.value : ""; // Legacy removed

    const tanggalMode = document.querySelector('input[name="tanggalMode"]:checked').value;
    const selectedDate = document.getElementById("tglSelect").value;
    const startDate = document.getElementById("tglStart").value;
    const endDate = document.getElementById("tglEnd").value;

    if (tanggalMode === 'single' && !selectedDate) {
        // If single mode and no date, passed empty to processData implies "All Dates"
        // No clearTable return here anymore
    }
    if (tanggalMode === 'range' && (!startDate || !endDate)) {
        // Range mode still needs start/end
        // clearTable("Pilih Rentang Tanggal yang valid."); // Optional: Keep range strict or default?
        // Let's keep range strict for now to avoid confusion, but single mode allows "All"
        if (!startDate && !endDate) {
            // If both empty in range mode, show all? Or wait? 
            // User said "tidak harus menunggu pilih satu tanggal". 
            // Logic in bankSelect sets "Semua Tanggal" in single mode.
            // Let's assume range needs explicit input.
            // So if single mode is active (default), pass through.
        } else if (!startDate || !endDate) {
            clearTable("Pilih Rentang Tanggal yang valid.");
            return;
        }
    }

    const waktuFilter = document.getElementById("waktuSelect").value;
    const fitur = document.querySelector('input[name="fitur"]:checked').value;

    processData(selectedDate, waktuFilter, startDate, endDate, tanggalMode, fitur);
}

function clearTable(msg) {
    document.getElementById("tabelData").innerHTML = `
        <div class="flex flex-col items-center justify-center h-full text-slate-400 dark:text-slate-600 py-20">
             <i class="bi bi-info-circle text-3xl mb-3 opacity-50"></i>
            <p class="text-sm">${msg}</p>
        </div>
    `;
    document.getElementById("jumlahData").textContent = "Data: 0";
    document.getElementById("paginationControls").classList.add("hidden");
    updateSummary(0, 0, 0);
    document.getElementById("exportPbbBtn").disabled = true;
    document.getElementById("btnSaveResult").disabled = true;
    currentTableData = [];
}

function processData(selectedDate, waktuFilter, startDate, endDate, tanggalMode, fitur) {
    // Filter by Tom Select Values (Detail)
    let selectedBanks = [];
    if (tomSelectInstance) {
        selectedBanks = tomSelectInstance.getValue();
    }

    let startTs = 0, endTs = 0;
    if (tanggalMode === 'range') {
        const sObj = new Date(startDate);
        const eObj = new Date(endDate);
        sObj.setHours(0, 0, 0, 0);
        eObj.setHours(23, 59, 59, 999);
        startTs = sObj.getTime();
        endTs = eObj.getTime();
    }

    let baseData = allData.filter(item => {
        if (!item.tglpembayaran) return false;

        let isBankMatch = true;
        if (selectedBanks.length > 0) {
            isBankMatch = selectedBanks.includes(item.pembayaran);
        } else {
            isBankMatch = true;
        }

        const dStr = item.tglpembayaran.substring(0, 10);

        let isDateMatch = false;
        if (tanggalMode === 'single') {
            // Allow all dates if selectedDate is empty (default option)
            isDateMatch = !selectedDate || dStr === selectedDate;
        } else {
            const dObj = parseDateObj(dStr);
            if (dObj) {
                const dTs = dObj.getTime();
                isDateMatch = dTs >= startTs && dTs <= endTs;
            }
        }
        return isBankMatch && isDateMatch;
    });

    let finalData = [];
    const tambahanCount = parseInt(document.getElementById("tambahanCount").value) || 0;

    if (fitur === "kurangi") {
        let filtered = baseData.slice();
        if (waktuFilter !== "semua") {
            filtered = filtered.filter(item => checkTime(item.tglpembayaran, waktuFilter));
        }
        filtered.sort((a, b) => {
            const da = parseDateObj(a.tglpembayaran);
            const db = parseDateObj(b.tglpembayaran);
            return (da && db) ? da - db : 0;
        });

        if (waktuFilter === "lebih") {
            finalData = filtered.slice(Math.min(tambahanCount, filtered.length));
        } else {
            finalData = filtered.slice(0, Math.max(filtered.length - tambahanCount, 0));
        }

    } else {
        let mainList = baseData.slice();
        if (waktuFilter !== "semua") {
            mainList = baseData.filter(item => checkTime(item.tglpembayaran, waktuFilter));
        }

        let adjustmentList = [];
        if (tambahanCount > 0 && waktuFilter !== "semua") {
            const inverseFilter = waktuFilter === "kurang" ? "lebih" : "kurang";
            let candidates = baseData.filter(item => checkTime(item.tglpembayaran, inverseFilter));

            candidates.sort((a, b) => {
                const da = parseDateObj(a.tglpembayaran);
                const db = parseDateObj(b.tglpembayaran);
                return (da && db) ? da - db : 0;
            });

            if (waktuFilter === "lebih") {
                candidates.reverse();
            }
            adjustmentList = candidates.slice(0, tambahanCount);
        }
        finalData = [...mainList, ...adjustmentList];
    }

    finalData.sort((a, b) => {
        const da = parseDateObj(a.tglpembayaran);
        const db = parseDateObj(b.tglpembayaran);
        return (da && db) ? da - db : 0;
    });

    currentTableData = finalData;
    document.getElementById("exportPbbBtn").disabled = finalData.length === 0;
    document.getElementById("btnSaveResult").disabled = finalData.length === 0;

    currentPage = 1;
    renderPage();
}

function checkTime(dateStr, mode) {
    if (!dateStr) return false;
    let timePart = "";
    const parts = dateStr.split(" ");
    if (parts.length > 1) {
        timePart = parts[1];
    } else {
        return false;
    }

    const [h, m, s] = timePart.split(":").map(Number);
    const seconds = (h || 0) * 3600 + (m || 0) * 60 + (s || 0);
    const limit = 17 * 3600;

    if (mode === "kurang") return seconds < limit;
    if (mode === "lebih") return seconds >= limit;
    return true;
}

// --- Selection & Bulk Delete Logic ---
function toggleRowSelection(id) {
    if (selectedRowIds.has(id)) selectedRowIds.delete(id);
    else selectedRowIds.add(id);
    renderPage();
    updateDeleteButton();
}

function toggleSelectAll(checked) {
    const pageData = currentTableData.slice((currentPage - 1) * rowsPerPage, currentPage * rowsPerPage);
    pageData.forEach(row => {
        if (checked) selectedRowIds.add(row._id);
        else selectedRowIds.delete(row._id);
    });
    renderPage();
    updateDeleteButton();
}

function updateDeleteButton() {
    const btn = document.getElementById('btnDeleteSelected');
    const countSpan = document.getElementById('selectedCount');
    if (!btn) return;

    if (selectedRowIds.size > 0) {
        btn.classList.remove('hidden');
        btn.disabled = false;
        if (countSpan) countSpan.textContent = selectedRowIds.size;
    } else {
        btn.classList.add('hidden');
        btn.disabled = true;
    }
}

// Bulk Delete Listener
document.getElementById('btnDeleteSelected')?.addEventListener('click', () => {
    if (selectedRowIds.size === 0) return;

    Swal.fire({
        title: `Hapus ${selectedRowIds.size} Data?`,
        text: "Data yang dihapus akan hilang dari tampilan ini (bisa dikembalikan dengan Reset).",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#e11d48',
        cancelButtonColor: '#64748b',
        confirmButtonText: 'Ya, Hapus',
        cancelButtonText: 'Batal'
    }).then((result) => {
        if (result.isConfirmed) {
            // Remove from allData
            const initialCount = allData.length;
            allData = allData.filter(row => !selectedRowIds.has(String(row._id)));
            const deletedCount = initialCount - allData.length;

            selectedRowIds.clear();
            updateDeleteButton();

            // Re-apply filters
            refreshTable();

            showToast(`${deletedCount} data berhasil dihapus.`, 'success');
        }
    });
});

function renderPage() {
    const data = currentTableData;

    let tPokok = 0, tDenda = 0, tTotal = 0;
    data.forEach(d => {
        tPokok += d.pokok;
        tDenda += d.denda;
        tTotal += d.total;
    });
    updateSummary(tPokok, tDenda, tTotal);
    document.getElementById("jumlahData").textContent = `Data: ${data.length}`;

    const container = document.getElementById("tabelData");
    const controls = document.getElementById("paginationControls");

    if (data.length === 0) {
        clearTable("Tidak ada data yang cocok dengan filter.");
        return;
    }

    controls.classList.remove("hidden");

    const totalPages = Math.ceil(data.length / rowsPerPage);
    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = startIdx + rowsPerPage;
    const pageData = data.slice(startIdx, endIdx);

    document.getElementById("currentPageDisplay").textContent = currentPage;
    document.getElementById("totalPagesDisplay").textContent = totalPages;
    document.getElementById("prevPageBtn").disabled = currentPage === 1;
    document.getElementById("nextPageBtn").disabled = currentPage === totalPages;

    // Check if all on current page are selected
    const allSelected = pageData.length > 0 && pageData.every(r => selectedRowIds.has(String(r._id)));

    let html = `
        <table class="w-full text-sm text-left text-slate-600 dark:text-slate-300">
            <thead class="text-xs text-slate-500 dark:text-slate-400 uppercase bg-slate-50 dark:bg-slate-800 sticky top-0 z-10 shadow-sm">
                <tr>
                    <th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700 w-10 text-center">
                        <input type="checkbox" onclick="toggleSelectAll(this.checked)" ${allSelected ? 'checked' : ''} class="rounded border-slate-300 text-rose-600 focus:ring-rose-500 cursor-pointer h-4 w-4">
                    </th>
                    <th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700">No</th>
                    <th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Pokok</th>
                    <th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Denda</th>
                    <th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Total</th>
                    <th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700">Bank</th>
                    <th scope="col" class="px-6 py-3 font-semibold border-b border-slate-200 dark:border-slate-700">Waktu Transaksi</th>
                </tr>
            </thead>
            <tbody class="bg-white dark:bg-slate-900/50 divide-y divide-slate-100 dark:divide-slate-800">
    `;

    pageData.forEach((row, i) => {
        const globalIndex = startIdx + i + 1;
        const isSelected = selectedRowIds.has(String(row._id));
        const isAlt = i % 2 === 0 ? 'bg-white dark:bg-transparent' : 'bg-slate-50/50 dark:bg-slate-800/30';

        html += `
            <tr class="${isAlt} hover:bg-blue-50/50 dark:hover:bg-slate-800 transition-colors ${isSelected ? 'bg-blue-50 dark:bg-slate-800/50' : ''}">
                <td class="px-6 py-3 text-center border-b border-slate-100 dark:border-slate-800">
                    <input type="checkbox" onclick="toggleRowSelection('${row._id}')" ${isSelected ? 'checked' : ''} class="rounded border-slate-300 text-rose-600 focus:ring-rose-500 cursor-pointer h-4 w-4">
                </td>
                <td class="px-6 py-3 font-medium text-slate-400 text-xs">${globalIndex}</td>
                <td class="px-6 py-3 text-right font-medium text-slate-700 dark:text-slate-300">${row.pokok.toLocaleString('id-ID')}</td>
                <td class="px-6 py-3 text-right text-amber-600 dark:text-amber-500">${row.denda.toLocaleString('id-ID')}</td>
                <td class="px-6 py-3 text-right font-bold text-emerald-600 dark:text-emerald-400">${row.total.toLocaleString('id-ID')}</td>
                <td class="px-6 py-3 text-slate-600 dark:text-slate-400">${row.pembayaran}</td>
                <td class="px-6 py-3 text-slate-500 dark:text-slate-500 font-mono text-xs">${row.tglpembayaran}</td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    container.innerHTML = html;

    const scrollContainer = document.getElementById("tabelContainer");
    if (scrollContainer) scrollContainer.scrollTop = 0;
}

function updateSummary(pokok, denda, total) {
    const format = (n) => n.toLocaleString('id-ID');
    document.getElementById("totalPbbText").textContent = format(pokok);
    document.getElementById("dendaText").textContent = format(denda);
    document.getElementById("mutasiText").textContent = format(total);

    const bars = document.querySelectorAll('#summaryCards .bg-slate-100, #summaryCards .dark\\:bg-slate-800 > div');
    bars.forEach(bar => {
        bar.style.width = '0%';
        setTimeout(() => bar.style.width = '100%', 100);
    });
}



// Template Logic


function showToast(msg, type = 'success') {
    const toast = document.getElementById("toast");
    const icon = document.getElementById("toastIcon");
    const txt = document.getElementById("toastMsg");

    if (!toast || !icon || !txt) return;

    txt.textContent = msg;

    toast.className = "fixed bottom-5 right-5 px-6 py-4 rounded-xl shadow-xl transition-all duration-300 z-[60] transform flex items-center gap-3 font-medium text-sm glass text-white backdrop-blur-md";

    if (type === 'success') {
        toast.classList.add('bg-emerald-600/90');
        icon.innerHTML = '<i class="bi bi-check-circle-fill text-lg"></i>';
    } else {
        toast.classList.add('bg-rose-600/90');
        icon.innerHTML = '<i class="bi bi-x-circle-fill text-lg"></i>';
    }

    requestAnimationFrame(() => {
        toast.classList.remove('translate-y-10', 'opacity-0');
    });

    setTimeout(() => {
        toast.classList.add('translate-y-10', 'opacity-0');
    }, 3000);
}

// --- New Features: Save / Preview / Download ---

function saveFilteredData() {
    if (!currentTableData || currentTableData.length === 0) {
        showToast("Tidak ada data untuk disimpan.", 'error');
        return;
    }

    const name = prompt("Masukkan nama untuk data ini:");
    if (!name) return;

    const newItem = {
        id: Date.now(),
        name: name,
        date: new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
        data: currentTableData,
        count: currentTableData.length
    };

    let savedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    savedData.unshift(newItem); // Add to top
    localStorage.setItem(STORAGE_KEY, JSON.stringify(savedData));

    showToast("Data hasil filter berhasil disimpan!", 'success');
}

function renderSavedDataList() {
    const listContainer = document.getElementById('savedDataList');
    const emptyState = document.getElementById('emptySaved');
    if (!listContainer) return;

    const savedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");

    if (savedData.length === 0) {
        listContainer.innerHTML = '';
        emptyState.classList.remove('hidden');
        return;
    }

    emptyState.classList.add('hidden');
    let html = '';

    savedData.forEach(item => {
        html += `
            <tr class="hover:bg-slate-50 dark:hover:bg-slate-800 transition-colors group">
                <td class="px-6 py-4 font-medium text-slate-700 dark:text-slate-200">
                    <div class="flex flex-col">
                        <span>${item.name}</span>
                        <span class="text-[10px] text-slate-400 font-normal">ID: ${item.id}</span>
                    </div>
                </td>
                <td class="px-6 py-4 text-center text-xs text-slate-500">${item.date}</td>
                <td class="px-6 py-4 text-center text-xs font-bold text-blue-600 dark:text-blue-400">${item.count}</td>
                <td class="px-6 py-4 text-right">
                    <div class="flex justify-end gap-2">
                        <button onclick="previewSavedData(${item.id})" 
                            class="p-2 rounded-lg bg-blue-50 text-blue-600 hover:bg-blue-100 dark:bg-blue-900/30 dark:text-blue-300 dark:hover:bg-blue-900/50 transition-colors"
                            title="Preview / Muat">
                            <i class="bi bi-eye"></i>
                        </button>
                        <button onclick="downloadSavedData(${item.id})"
                            class="p-2 rounded-lg bg-emerald-50 text-emerald-600 hover:bg-emerald-100 dark:bg-emerald-900/30 dark:text-emerald-300 dark:hover:bg-emerald-900/50 transition-colors"
                            title="Download Excel">
                            <i class="bi bi-download"></i>
                        </button>
                        <button onclick="deleteSavedData(${item.id})"
                            class="p-2 rounded-lg bg-rose-50 text-rose-600 hover:bg-rose-100 dark:bg-rose-900/30 dark:text-rose-300 dark:hover:bg-rose-900/50 transition-colors"
                            title="Hapus">
                            <i class="bi bi-trash"></i>
                        </button>
                    </div>
                </td>
            </tr>
        `;
    });

    listContainer.innerHTML = html;
}

// Made global for onclick handlers in HTML string
window.previewSavedData = (id) => {
    const savedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    const item = savedData.find(i => i.id === id);
    if (!item) return;

    currentTableData = item.data;
    currentPage = 1;

    // Update UI to reflect we are in "Saved View" mode
    renderPage();
    document.getElementById("btnSaveResult").disabled = true; // No need to save what is already saved
    document.getElementById("exportPbbBtn").disabled = false;

    // Close modal
    const savedDataModal = document.getElementById('savedDataModal');
    savedDataModal.classList.add('opacity-0', 'hidden'); // Simple close
    // Using the toggle logic would be better but direct manipulation works for now

    showToast(`Memuat data tersimpan: ${item.name}`, 'success');

    // Optional: Reset filters UI to indicate custom data loaded?
    document.getElementById("tabelData").scrollIntoView({ behavior: 'smooth' });
    document.getElementById("btnReset").disabled = false;
};

window.downloadSavedData = (id) => {
    const savedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    const item = savedData.find(i => i.id === id);
    if (!item) return;

    exportDataToExcel(item.data, `SiRekonPBB_Saved_${item.name.replace(/\s+/g, '_')}`);
};

window.deleteSavedData = (id) => {
    if (!confirm("Hapus data tersimpan ini?")) return;

    let savedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    savedData = savedData.filter(i => i.id !== id);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(savedData));

    renderSavedDataList();
    showToast("Data terhapus.", 'success');
};

document.getElementById('btnDownloadAll')?.addEventListener('click', () => {
    const savedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    if (savedData.length === 0) {
        showToast("Tidak ada data tersimpan untuk didownload.", 'warning');
        return;
    }

    let allRows = [];

    // Merge all data
    savedData.forEach(item => {
        // Add 'Nama Laporan' field to identify source
        const reportRows = item.data.map(r => ({
            'Nama Laporan': item.name,
            'Waktu Simpan': item.date,
            ...r // Spread original data
        }));
        allRows = allRows.concat(reportRows);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(allRows);
    XLSX.utils.book_append_sheet(wb, ws, "Semua Data Tersimpan");

    const fname = `SiRekonPBB_Merged_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fname);
    showToast("Download gabungan berhasil!", 'success');
});

document.getElementById('btnClearAll')?.addEventListener('click', () => {
    const savedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    if (savedData.length === 0) {
        showToast("Tidak ada data untuk dibersihkan.", 'warning');
        return;
    }

    Swal.fire({
        title: 'Bersihkan Semua Data?',
        text: "Tindakan ini akan menghapus SELURUH riwayat data tersimpan. Data tidak dapat dikembalikan!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#e11d48', // Rose-600
        cancelButtonColor: '#64748b', // Slate-500
        confirmButtonText: 'Ya, Hapus Semua!',
        cancelButtonText: 'Batal',
        background: document.documentElement.classList.contains('dark') ? '#0f172a' : '#fff',
        color: document.documentElement.classList.contains('dark') ? '#f1f5f9' : '#1e293b'
    }).then((result) => {
        if (result.isConfirmed) {
            localStorage.removeItem(STORAGE_KEY);
            renderSavedDataList();

            Swal.fire({
                title: 'Terhapus!',
                text: 'Semua data tersimpan telah dibersihkan.',
                icon: 'success',
                confirmButtonColor: '#3b82f6',
                background: document.documentElement.classList.contains('dark') ? '#0f172a' : '#fff',
                color: document.documentElement.classList.contains('dark') ? '#f1f5f9' : '#1e293b'
            });
        }
    });
});

// Make globally accessible
window.deleteRow = (id) => {
    if (!confirm("Hapus baris ini? Nominal total akan berubah.")) return;

    // Remove from main data
    const initialCount = allData.length;
    allData = allData.filter(item => item._id !== id);

    if (allData.length < initialCount) {
        // Re-apply filters to update currentTableData
        processData();
        showToast("Baris berhasil dihapus.", 'success');
    } else {
        console.error("ID not found:", id);
    }
};

function exportDataToExcel(data, filenamePrefix = "SiRekonPBB") {
    if (!data || data.length === 0) return;

    // If data contains original fields (from ...item), use them.
    // Otherwise fallback to the mapped view.
    // Since we now merge ...item, the objects in 'data' have all keys.
    // We should try to export EVERYTHING.

    // However, we want to keep the important columns first or formatted nicely?
    // User asked for "lengkap seperti pada file excel yang diimport".
    // So we should just dump the object, but maybe ensure formatting for key columns?

    // Let's just pass the data object directly to json_to_sheet?
    // But data also contains our parsed 'pokok', 'denda', etc.
    // If we just dump it, it might have duplicates (mapped vs original).
    // But usually original keys are uppercase or different.

    // Better approach: Use the keys from the first item to define headers?
    // Or just use the data as is.

    const rows = data.map(r => {
        // Create a new object to control order if desired, or just return r
        // User wants "like original".
        return r;
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data PBB");
    const fname = `${filenamePrefix}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fname);
    showToast("Data berhasil diekspor.", 'success');
}
