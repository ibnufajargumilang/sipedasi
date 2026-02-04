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

const utamaFileInput = document.getElementById("utamaFile");
const konsolidasiFileInput = document.getElementById("konsolidasiFile");
const processButton = document.getElementById("processButton");
const exportButton = document.getElementById("exportButton");
const monthSelector = document.getElementById("monthSelector");
const rowCountInput = document.getElementById("rowCount");
const summaryCards = document.getElementById("summaryCards");

let utamaData = null;
let konsolidasiData = null;
let processedDataForExport = [];

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

const checkFiles = () => {
    const ready = utamaData && konsolidasiData;
    processButton.disabled = !ready;
    document.getElementById('resultStatus').textContent = ready ? "Siap Diproses" : "Menunggu File";
};

const handleFileUpload = (event, type, statusId) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        if (type === "utama") {
            const sheetName = "LRA";
            if (!workbook.SheetNames.includes(sheetName)) {
                showToast(`File utama harus punya sheet "${sheetName}".`, 'error');
                return;
            }
            const worksheet = workbook.Sheets[sheetName];
            utamaData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            document.getElementById(statusId).textContent = "OK";
            document.getElementById(statusId).classList.replace("bg-slate-200", "bg-green-100");
            document.getElementById(statusId).classList.replace("text-slate-600", "text-green-600");
            document.getElementById(statusId).classList.replace("dark:bg-slate-700", "dark:bg-green-900/30");
            document.getElementById(statusId).classList.replace("dark:text-slate-300", "dark:text-green-400");
        } else {
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            konsolidasiData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            document.getElementById(statusId).textContent = "OK";
            document.getElementById(statusId).classList.replace("bg-slate-200", "bg-green-100");
            document.getElementById(statusId).classList.replace("text-slate-600", "text-green-600");
            document.getElementById(statusId).classList.replace("dark:bg-slate-700", "dark:bg-green-900/30");
            document.getElementById(statusId).classList.replace("dark:text-slate-300", "dark:text-green-400");
        }
        checkFiles();
        showToast(`File ${type} berhasil dimuat.`, 'success');
    };
    reader.readAsArrayBuffer(file);
};

utamaFileInput.addEventListener("change", (e) => handleFileUpload(e, "utama", "statusUtama"));
konsolidasiFileInput.addEventListener("change", (e) => handleFileUpload(e, "konsolidasi", "statusKonsol"));

processButton.addEventListener("click", () => {
    if (utamaData && konsolidasiData) {
        document.getElementById('table-placeholder').classList.add('hidden');
        setTimeout(() => { processData(utamaData, konsolidasiData); toggleModal(false); }, 50);
    }
});

const processData = (dataUtama, dataKonsolidasi) => {
    const findStartRow = (data, keyword = "KODE REKENING") => {
        for (let i = 0; i < data.length; i++) {
            if (data[i] && data[i].some(cell => typeof cell === 'string' && cell.toUpperCase().includes(keyword))) {
                return i + 1;
            }
        }
        return 0;
    };

    const startRowUtama = findStartRow(dataUtama);
    const startRowKonsolidasi = findStartRow(dataKonsolidasi);
    const selectedMonthIndex = parseInt(monthSelector.value);

    const cleanAndMap = (data, startRow, type) => {
        return data.slice(startRow)
            .map(row => {
                if (!row || !row[0] || !row[1]) return null;
                let realisasiValue = 0;
                if (type === 'konsolidasi') {
                    realisasiValue = row[3] || 0;
                } else {
                    realisasiValue = row[selectedMonthIndex] || 0;
                }
                return {
                    kode: String(row[0]).trim(),
                    uraian: String(row[1]).trim(),
                    realisasi: realisasiValue,
                };
            })
            .filter(item => item && item.kode && item.uraian);
    };

    const mappedUtama = cleanAndMap(dataUtama, startRowUtama, 'utama');
    const mappedKonsolidasi = cleanAndMap(dataKonsolidasi, startRowKonsolidasi, 'konsolidasi');

    const dataMap = new Map();

    mappedKonsolidasi.forEach(item => {
        dataMap.set(item.kode, {
            kode: item.kode,
            uraian: item.uraian,
            realisasiUtama: 0,
            realisasiKonsolidasi: item.realisasi,
        });
    });

    mappedUtama.forEach(item => {
        if (dataMap.has(item.kode)) {
            dataMap.get(item.kode).realisasiUtama = item.realisasi;
        } else {
            dataMap.set(item.kode, {
                kode: item.kode,
                uraian: item.uraian,
                realisasiUtama: item.realisasi,
                realisasiKonsolidasi: 0,
            });
        }
    });

    const mergedData = Array.from(dataMap.values());
    const detailAccountRegex = /^\d\.\d\.\d{2}\.\d{2}\.\d{2}\.\d{4}$/;
    const filteredData = mergedData.filter(item => detailAccountRegex.test(item.kode));

    processedDataForExport = filteredData;
    displayResult(filteredData);
};

const displayResult = (data) => {
    const formatRupiah = (val) => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(val);
    const container = document.getElementById('excel-table');
    const placeholder = document.getElementById('table-placeholder');

    if (data.length === 0) {
        container.innerHTML = '';
        placeholder.classList.remove('hidden');
        placeholder.innerHTML = '<i class="bi bi-search text-4xl mb-3 opacity-50"></i><p class="text-sm">Tidak ada data rekening detail yang cocok.</p>';
        summaryCards.classList.add('hidden');
        exportButton.style.display = 'none';
        return;
    }

    placeholder.classList.add('hidden');
    summaryCards.classList.remove('hidden');

    const rowCount = parseInt(rowCountInput.value) || data.length;
    const slicedData = data.slice(0, rowCount);

    const mismatchCount = slicedData.filter(item => (Number(item.realisasiUtama) || 0) !== (Number(item.realisasiKonsolidasi) || 0)).length;
    exportButton.style.display = mismatchCount > 0 ? 'inline-flex' : 'none';
    if (mismatchCount > 0) exportButton.disabled = false;

    const totals = slicedData.reduce((acc, item) => {
        acc.realisasiUtama += Number(item.realisasiUtama) || 0;
        acc.realisasiKonsolidasi += Number(item.realisasiKonsolidasi) || 0;
        return acc;
    }, { realisasiUtama: 0, realisasiKonsolidasi: 0 });
    totals.selisih = totals.realisasiUtama - totals.realisasiKonsolidasi;

    document.getElementById("totalRealisasiUtama").textContent = formatRupiah(totals.realisasiUtama);
    document.getElementById("totalRealisasiKonsolidasi").textContent = formatRupiah(totals.realisasiKonsolidasi);
    document.getElementById("totalSelisih").textContent = formatRupiah(totals.selisih);
    document.getElementById("resultStatus").innerHTML = `Total: <b>${data.length}</b> | Selisih: <b class="text-red-500">${mismatchCount}</b>`;

    let html = `<thead class="bg-gray-100 dark:bg-slate-800 sticky top-0 z-10 shadow-sm text-xs text-slate-500 dark:text-slate-400 uppercase"><tr>
        <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">No</th>
        <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">Kode Rekening</th>
        <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700">Uraian</th>
        <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Realisasi Utama</th>
        <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Realisasi Konsolidasi</th>
        <th class="p-3 font-semibold border-b border-slate-200 dark:border-slate-700 text-right">Selisih</th>
    </tr></thead><tbody class="divide-y divide-gray-100 dark:divide-slate-800 text-sm text-slate-600 dark:text-slate-300">`;

    slicedData.forEach((item, i) => {
        const rUtama = Number(item.realisasiUtama) || 0;
        const rKonsol = Number(item.realisasiKonsolidasi) || 0;
        const selisih = rUtama - rKonsol;
        const mismatch = selisih !== 0;
        const rowClass = mismatch ? "row-mismatch" : "hover:bg-blue-50/50 dark:hover:bg-slate-800 transition-colors";

        let selisihClass = "selisih-zero";
        if (selisih > 0) selisihClass = "selisih-positive";
        else if (selisih < 0) selisihClass = "selisih-negative";

        html += `<tr class="${rowClass}">
            <td class="p-3 font-medium text-slate-400 text-xs">${i + 1}</td>
            <td class="p-3 font-mono text-xs">${item.kode}</td>
            <td class="p-3">${item.uraian}</td>
            <td class="p-3 text-right font-medium text-teal-600 dark:text-teal-400">${formatRupiah(rUtama)}</td>
            <td class="p-3 text-right font-medium text-indigo-600 dark:text-indigo-400">${formatRupiah(rKonsol)}</td>
            <td class="p-3 text-right ${selisihClass}">${formatRupiah(selisih)}</td>
        </tr>`;
    });
    html += `</tbody>`;
    container.innerHTML = html;
    document.getElementById('table-scroll-container').scrollTop = 0;
}

exportButton.addEventListener("click", () => {
    if (processedDataForExport.length === 0) { showToast("Tidak ada data.", 'error'); return; }

    const rowCount = parseInt(rowCountInput.value) || processedDataForExport.length;
    const slicedData = processedDataForExport.slice(0, rowCount);
    const dataToExport = slicedData.filter(item => {
        const rU = Number(item.realisasiUtama) || 0;
        const rK = Number(item.realisasiKonsolidasi) || 0;
        return rU !== rK;
    });

    if (dataToExport.length === 0) { showToast(`Tidak ada selisih dalam ${rowCount} baris pertama.`, 'error'); return; }

    const headers = ["Kode Rekening", "Uraian", "Realisasi Utama", "Realisasi Konsolidasi", "Selisih"];
    const sheetData = dataToExport.map(item => {
        const rU = Number(item.realisasiUtama) || 0;
        const rK = Number(item.realisasiKonsolidasi) || 0;
        return [item.kode, item.uraian, rU, rK, rU - rK];
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...sheetData]);
    worksheet['!cols'] = [{ wch: 25 }, { wch: 60 }, { wch: 20 }, { wch: 20 }, { wch: 20 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data Selisih");
    XLSX.writeFile(workbook, `Selisih_LRA_${rowCount}_Baris.xlsx`);
    showToast('Data selisih berhasil diekspor.', 'success');
});
