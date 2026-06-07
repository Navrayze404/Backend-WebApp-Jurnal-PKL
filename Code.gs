/**
 * Backend API Google Apps Script untuk upload foto jurnal PKL ke Google Drive.
 * Deploy sebagai Web App dengan akses: Anyone.
 * Method yang digunakan website: POST.
 */

const CONFIG = {
  ROOT_FOLDER_NAME: 'LaporanJurnalPKL',
  TIMEZONE: 'Asia/Jakarta',
  DEFAULT_DESCRIPTION: '(tidak ada deskripsi)'
};

const HARI_INDONESIA = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const BULAN_INDONESIA = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

/**
 * Fungsi utama untuk menerima request POST dari website.
 * Data yang diterima harus berupa JSON: { foto: base64, deskripsi: string, timestamp: string }.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({
        status: 'error',
        message: 'Request POST tidak berisi body JSON.'
      });
    }

    const payload = parsePayload(e);
    validatePayload(payload);

    const uploadDate = parseTimestamp(payload.timestamp);
    const info = getDateInfo(uploadDate);

    const rootFolder = getOrCreateFolder_(DriveApp, CONFIG.ROOT_FOLDER_NAME);
    const monthFolderName = 'Bulan' + info.bulan + info.tahun;
    const monthFolder = getOrCreateFolder_(rootFolder, monthFolderName);
    const weekFolderName = 'Minggu' + info.minggu;
    const weekFolder = getOrCreateFolder_(monthFolder, weekFolderName);

    const fileName = buildPhotoFileName(info);
    const photoBlob = createJpegBlobFromBase64(payload.foto, fileName);
    weekFolder.createFile(photoBlob);

    const description = normalizeDescription(payload.deskripsi);
    updateWeeklyRecap(weekFolder, info, description);

    return jsonResponse({
      status: 'success',
      fileName: fileName,
      folderPath: CONFIG.ROOT_FOLDER_NAME + '/' + monthFolderName + '/' + weekFolderName,
      rekapUpdated: true
    });
  } catch (error) {
    return jsonResponse({
      status: 'error',
      message: error && error.message ? error.message : String(error)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

/**
 * Fungsi GET hanya untuk memberi informasi bahwa API harus dipanggil menggunakan POST.
 */
function doGet() {
  return jsonResponse({
    status: 'error',
    message: 'Endpoint ini hanya menerima request POST.'
  });
}

/**
 * Mengubah request body JSON menjadi object JavaScript.
 */
function parsePayload(e) {
  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('Body request harus berupa JSON yang valid.');
  }
}

/**
 * Memvalidasi field wajib pada payload, terutama foto base64.
 */
function validatePayload(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Payload tidak valid.');
  }

  if (!payload.foto || typeof payload.foto !== 'string' || payload.foto.trim() === '') {
    throw new Error('Field foto wajib diisi dalam format base64.');
  }

  if (payload.deskripsi !== undefined && typeof payload.deskripsi !== 'string') {
    throw new Error('Field deskripsi harus berupa string.');
  }

  if (payload.timestamp !== undefined && typeof payload.timestamp !== 'string') {
    throw new Error('Field timestamp harus berupa string.');
  }
}

/**
 * Mengubah timestamp dari website menjadi Date.
 * Jika timestamp kosong atau tidak valid, sistem memakai waktu upload saat ini.
 */
function parseTimestamp(timestamp) {
  if (!timestamp || String(timestamp).trim() === '') {
    return new Date();
  }

  const parsedDate = new Date(timestamp);
  if (isNaN(parsedDate.getTime())) {
    return new Date();
  }

  return parsedDate;
}

/**
 * Membuat semua informasi tanggal dan waktu dalam timezone Asia/Jakarta.
 */
function getDateInfo(date) {
  const tahun = Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy'));
  const bulanIndex = Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'M')) - 1;
  const tanggal = Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'd'));
  const hariIndex = Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'u')) % 7;
  const jamMenitDetik = Utilities.formatDate(date, CONFIG.TIMEZONE, 'HHmmss');
  const jamMenit = Utilities.formatDate(date, CONFIG.TIMEZONE, 'HH:mm');

  return {
    date: date,
    tahun: tahun,
    bulan: BULAN_INDONESIA[bulanIndex],
    bulanIndex: bulanIndex,
    tanggal: tanggal,
    hari: HARI_INDONESIA[hariIndex],
    minggu: getWeekNumberInMonth(tanggal),
    jamMenitDetik: jamMenitDetik,
    jamMenit: jamMenit,
    sortKey: Number(
      Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyyMMddHHmmss')
    )
  };
}

/**
 * Menentukan nomor minggu berdasarkan tanggal kalender dalam bulan.
 * Minggu1: 1-7, Minggu2: 8-14, Minggu3: 15-21, Minggu4: 22-28, Minggu5: 29-31.
 */
function getWeekNumberInMonth(tanggal) {
  return Math.ceil(tanggal / 7);
}

/**
 * Membuat nama file foto sesuai format: foto-[NamaHari][Tanggal][NamaBulan][Tahun]-[HHMMSS].jpg.
 */
function buildPhotoFileName(info) {
  return 'foto-' + info.hari + info.tanggal + info.bulan + info.tahun + '-' + info.jamMenitDetik + '.jpg';
}

/**
 * Mengambil folder berdasarkan nama. Jika belum ada, folder akan dibuat otomatis.
 */
function getOrCreateFolder_(parent, folderName) {
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(folderName);
}

/**
 * Membersihkan base64, melakukan decode, lalu membuat Blob JPG untuk disimpan ke Drive.
 */
function createJpegBlobFromBase64(base64Input, fileName) {
  const cleanedBase64 = String(base64Input)
    .replace(/^data:image\/[a-zA-Z0-9.+-]+;base64,/, '')
    .replace(/\s/g, '');

  try {
    const bytes = Utilities.base64Decode(cleanedBase64);
    return Utilities.newBlob(bytes, MimeType.JPEG, fileName);
  } catch (error) {
    throw new Error('Foto base64 tidak valid atau gagal didecode.');
  }
}

/**
 * Mengubah deskripsi kosong menjadi teks default agar rekap tetap rapi.
 */
function normalizeDescription(description) {
  const text = description === undefined || description === null ? '' : String(description).trim();
  return text === '' ? CONFIG.DEFAULT_DESCRIPTION : text;
}

/**
 * Membuat atau memperbarui file Rekap-MingguX-Bulan.txt.
 * Entry lama dipertahankan, entry baru ditambahkan, lalu seluruh entry diurutkan berdasarkan hari dan jam.
 */
function updateWeeklyRecap(weekFolder, info, description) {
  const recapFileName = 'Rekap-Minggu' + info.minggu + '-' + info.bulan + '.txt';
  const files = weekFolder.getFilesByName(recapFileName);

  let recapFile = null;
  let existingContent = '';

  if (files.hasNext()) {
    recapFile = files.next();
    existingContent = recapFile.getBlob().getDataAsString('UTF-8');
  }

  const entries = parseExistingRecap(existingContent, info);
  entries.push({
    sortKey: info.sortKey,
    hari: info.hari,
    tanggal: info.tanggal,
    bulan: info.bulan,
    tahun: info.tahun,
    jamMenit: info.jamMenit,
    deskripsi: description
  });

  const newContent = buildRecapContent(entries, info);

  if (recapFile) {
    recapFile.setContent(newContent);
  } else {
    weekFolder.createFile(recapFileName, newContent, MimeType.PLAIN_TEXT);
  }
}

/**
 * Membaca isi rekap lama dan mengubahnya menjadi daftar entry agar bisa ditambah dan diurutkan kembali.
 */
function parseExistingRecap(content, fallbackInfo) {
  if (!content || String(content).trim() === '') {
    return [];
  }

  const entries = [];
  const lines = String(content).split(/\r?\n/);
  let currentDateInfo = null;

  lines.forEach(function(line) {
    const sectionMatch = line.match(/^([A-Za-zÀ-ÿ]+),\s+(\d{1,2})\s+([A-Za-zÀ-ÿ]+)\s+(\d{4}):$/);
    if (sectionMatch) {
      currentDateInfo = {
        hari: sectionMatch[1],
        tanggal: Number(sectionMatch[2]),
        bulan: sectionMatch[3],
        tahun: Number(sectionMatch[4])
      };
      return;
    }

    const entryMatch = line.match(/^\s{2}(\d{2}:\d{2})\s+-\s+(.+)$/);
    if (entryMatch && currentDateInfo) {
      entries.push({
        sortKey: buildSortKeyFromParts(currentDateInfo, entryMatch[1]),
        hari: currentDateInfo.hari,
        tanggal: currentDateInfo.tanggal,
        bulan: currentDateInfo.bulan,
        tahun: currentDateInfo.tahun,
        jamMenit: entryMatch[1],
        deskripsi: entryMatch[2]
      });
    }
  });

  return entries.filter(function(entry) {
    return entry.bulan === fallbackInfo.bulan && entry.tahun === fallbackInfo.tahun;
  });
}

/**
 * Membuat angka sortKey dari bagian tanggal rekap lama supaya urutan entry tetap konsisten.
 */
function buildSortKeyFromParts(dateInfo, jamMenit) {
  const bulanIndex = BULAN_INDONESIA.indexOf(dateInfo.bulan) + 1;
  const tanggalText = String(dateInfo.tanggal).padStart(2, '0');
  const bulanText = String(bulanIndex).padStart(2, '0');
  const jamMenitText = String(jamMenit).replace(':', '');
  return Number(String(dateInfo.tahun) + bulanText + tanggalText + jamMenitText + '00');
}

/**
 * Menyusun ulang isi rekap sesuai format yang diminta dan mengelompokkan entry per hari.
 */
function buildRecapContent(entries, info) {
  const separator = '═══════════════════════════════════════';
  const sortedEntries = entries.slice().sort(function(a, b) {
    return a.sortKey - b.sortKey;
  });

  const lines = [
    separator,
    'REKAP JURNAL PKL - MINGGU ' + info.minggu + ' ' + info.bulan + ' ' + info.tahun,
    separator
  ];

  let currentSectionKey = '';

  sortedEntries.forEach(function(entry) {
    const sectionKey = entry.tahun + '-' + entry.bulan + '-' + entry.tanggal;
    if (sectionKey !== currentSectionKey) {
      if (currentSectionKey !== '') {
        lines.push('');
      }
      lines.push(entry.hari + ', ' + entry.tanggal + ' ' + entry.bulan + ' ' + entry.tahun + ':');
      currentSectionKey = sectionKey;
    }

    lines.push('  ' + entry.jamMenit + ' - ' + entry.deskripsi);
  });

  lines.push(separator);
  return lines.join('\n');
}

/**
 * Membuat response JSON standar untuk sukses maupun gagal.
 */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

