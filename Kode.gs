// ===== GOOGLE APPS SCRIPT (Paste di Editor Apps Script) =====

const SHEET_NAME = 'SafeTrack Reports';
const HEADER_ROW = 1;
const SPREADSHEET_ID = '1sq5leW45LVQjUzmojpadHKpu1RYFXymAg7LEuRpJDFM';

function getSpreadsheet() {
  // Prioritaskan Spreadsheet ID agar web app konsisten lintas deployment
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Inisialisasi sheet jika belum ada
function initializeSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
    const headers = [
      'ID Backend',
      'Tanggal',
      'Tipe Laporan',
      'Kategori',
      'Pelapor',
      'Departemen',
      'Lokasi',
      'Status',
      'Prioritas',
      'Detail',
      'Deskripsi',
      'Tipe Box P3K',
      'Status Item P3K',
      'Catatan Admin',
      'Tanggal Selesai',
      'Timestamp Sinkronisasi'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1e40af')
              .setFontColor('#ffffff')
              .setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 120);  // ID Backend
    sheet.setColumnWidth(2, 100);  // Tanggal
    sheet.setColumnWidth(3, 130);  // Tipe
    sheet.setColumnWidth(4, 120);  // Kategori
    sheet.setColumnWidth(5, 100);  // Pelapor
    sheet.setColumnWidth(6, 120);  // Departemen
    sheet.setColumnWidth(7, 150);  // Lokasi
    sheet.setColumnWidth(8, 140);  // Status
    sheet.setColumnWidth(9, 100);  // Prioritas
    sheet.setColumnWidth(10, 150); // Detail
    sheet.setColumnWidth(11, 200); // Deskripsi
    sheet.setColumnWidth(12, 130); // Tipe Box P3K
    sheet.setColumnWidth(13, 200); // Status Item P3K
    sheet.setColumnWidth(14, 200); // Catatan Admin
    sheet.setColumnWidth(15, 120); // Tanggal Selesai
    sheet.setColumnWidth(16, 150); // Timestamp
  }
  
  return sheet;
}

function createApiOutput(payload, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// Main handler untuk POST request dari Canva
function doPost(e) {
  try {
    const sheet = initializeSheet();
    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(rawBody);

    if (data.action === 'deleteReport' && data.backendId) {
      const deleted = deleteReportById(sheet, data.backendId);
      return createApiOutput({
        status: deleted ? 'success' : 'error',
        message: deleted ? 'Laporan berhasil dihapus dari spreadsheet' : 'Laporan tidak ditemukan di spreadsheet'
      });
    }
    
    if (data.action === 'syncAll' && data.reports && Array.isArray(data.reports)) {
      // Sinkronisasi seluruh laporan
      syncAllReports(sheet, data.reports);
      
      return createApiOutput({ status: 'success', message: `${data.reports.length} laporan tersinkronisasi` });
    }
    
    return createApiOutput({ status: 'error', message: 'Action tidak dikenali' });
    
  } catch (error) {
    Logger.log('Error di doPost: ' + error.toString());
    return createApiOutput({ status: 'error', message: error.toString() });
  }
}

// Sinkronisasi semua laporan
function syncAllReports(sheet, reports) {
  if (!reports || reports.length === 0) return;
  
  // Ambil semua data yang sudah ada
  const existingData = sheet.getDataRange().getValues();
  const existingIds = new Set(existingData.slice(1).map(row => row[0])); // Kolom ID Backend
  
  // Pisahkan laporan baru dan update
  const newReports = [];
  const updateIndices = [];
  
  reports.forEach(report => {
    const rowIndex = existingData.findIndex(row => row[0] === report.backendId);
    if (rowIndex === -1) {
      // Laporan baru
      newReports.push(report);
    } else {
      // Laporan update (existing)
      updateIndices.push({ report, rowIndex: rowIndex + 1 }); // +1 karena Google Sheets 1-indexed
    }
  });
  
  // Tambah laporan baru
  if (newReports.length > 0) {
    const newRows = newReports.map(r => [
      r.backendId,
      r.date,
      r.type,
      r.category,
      r.reporter,
      r.department,
      r.location,
      r.status,
      r.priority,
      r.details,
      r.description,
      r.boxType,
      r.itemsStatus,
      r.adminNotes,
      r.completedDate,
      new Date().toLocaleString('id-ID')
    ]);
    
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 16).setValues(newRows);
  }
  
  // Update laporan yang sudah ada
  updateIndices.forEach(({ report, rowIndex }) => {
    const values = [
      report.backendId,
      report.date,
      report.type,
      report.category,
      report.reporter,
      report.department,
      report.location,
      report.status,
      report.priority,
      report.details,
      report.description,
      report.boxType,
      report.itemsStatus,
      report.adminNotes,
      report.completedDate,
      new Date().toLocaleString('id-ID')
    ];
    
    sheet.getRange(rowIndex, 1, 1, 16).setValues([values]);
  });
  
  // Hapus duplikat terakhir
  removeDuplicates(sheet);
}

// Hapus duplikat berdasarkan ID Backend
function removeDuplicates(sheet) {
  const data = sheet.getDataRange().getValues();
  const seen = new Set();
  const rowsToDelete = [];
  
  // Iterasi dari bawah ke atas agar index tetap valid saat delete
  for (let i = data.length - 1; i > 0; i--) {
    const id = data[i][0];
    if (seen.has(id)) {
      rowsToDelete.push(i + 1); // +1 karena Google Sheets 1-indexed
    } else {
      seen.add(id);
    }
  }
  
  // Delete rows dari yang tertinggi ke terendah
  rowsToDelete.forEach(row => {
    sheet.deleteRow(row);
  });
}

function deleteReportById(sheet, backendId) {
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][0] === backendId) {
      sheet.deleteRow(i + 1); // +1 karena Google Sheets 1-indexed
      return true;
    }
  }

  return false;
}
function doGet(e) {
  try {
    const sheet = initializeSheet();
    const action = e && e.parameter && e.parameter.action ? e.parameter.action : "summary";
    const callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : "";

    // action=getReports -> kirim semua data laporan agar bisa sinkron lintas device
    if (action === "getReports") {
      const data = sheet.getDataRange().getValues();
      const reports = data.slice(1).map(row => ({
        __backendId: row[0] || "",
        report_date: row[1] || "",
        type: row[2] || "",
        category: row[3] || "",
        reporter_name: row[4] || "",
        department: row[5] || "",
        location: row[6] || "",
        status: row[7] || "",
        priority: row[8] || "Normal",
        details: row[9] || "",
        description: row[10] || "",
        box_type: row[11] || "",
        items_status: row[12] || "",
        admin_notes: row[13] || "",
        completed_date: row[14] || "",
        synced_at: row[15] || ""
      }));

      return createApiOutput({ status: "success", reports, total: reports.length }, callback);
    }

    // default: ringkasan total
    const total = Math.max(sheet.getLastRow() - HEADER_ROW, 0);
    return createApiOutput({ status: "success", total }, callback);
  } catch (error) {
    Logger.log("Error di doGet: " + error.toString());
    const callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : "";
    return createApiOutput({ status: "error", message: error.toString() }, callback);
  }
}
// Fungsi untuk test (jalankan dari Editor)
function testSync() {
  const sheet = initializeSheet();
  const testData = {
    action: 'syncAll',
    reports: [
      {
        backendId: 'test-001',
        date: '2024-01-15',
        type: 'Inspeksi APAR',
        category: 'Powder',
        reporter: 'Budi Santoso',
        department: 'Dyeing',
        location: 'Lantai 2 Area Produksi',
        status: 'Selesai',
        priority: 'Normal',
        details: 'APAR-001, Exp: 2025-06-30',
        description: 'Semua checklist baik',
        boxType: '',
        itemsStatus: '',
        adminNotes: '',
        completedDate: '2024-01-15'
      }
    ]
  };
  
  syncAllReports(sheet, testData.reports);
  Logger.log('Test sync berhasil!');
}
