/**
 * Google Apps Script untuk Supplier Gathering PLN
 * Backend + Frontend Hosting (HTML Service)
 */

// KONFIGURASI - Ganti dengan Spreadsheet ID Anda
const SPREADSHEET_ID = "1pjAvTf8ug7lHTOKBUDR-R1E_8Eyom-s5pR7LaRGSshQ";

// Nama-nama sheet
const SHEET_NAMES = {
  KEHADIRAN: "Data Kehadiran",
  PAITON: "Kuesioner UP Paiton",
  BRANTAS: "Kuesioner UP Brantas",
  PACITAN: "Kuesioner UP Pacitan",
  REKAP: "Rekap Kehadiran",
  BUKU_TAMU: "Buku Tamu",
};

// DAFTAR PERUSAHAAN - Edit sesuai kebutuhan
const COMPANY_LIST = [
  "PT Astra International Tbk",
  "PT Bank Central Asia Tbk",
  "PT Bank Mandiri (Persero) Tbk",
  "PT Bank Rakyat Indonesia (Persero) Tbk",
  "PT Gudang Garam Tbk",
  "PT HM Sampoerna Tbk",
  "PT Indofood CBP Sukses Makmur Tbk",
  "PT Indofood Sukses Makmur Tbk",
  "PT Telkom Indonesia (Persero) Tbk",
  "PT Unilever Indonesia Tbk",
  "PT Vale Indonesia Tbk",
  "PT Wijaya Karya (Persero) Tbk",
  "PT Adaro Energy Indonesia Tbk",
  "PT Bukit Asam Tbk",
  "PT Chandra Asri Petrochemical Tbk",
  "PT Kalbe Farma Tbk",
  "PT Pertamina (Persero)",
  "PT Perusahaan Gas Negara Tbk",
  "PT Semen Indonesia (Persero) Tbk",
  "PT United Tractors Tbk",
  "PT Aneka Tambang (Persero) Tbk",
  "PT Jasa Marga (Persero) Tbk",
  "PT Kimia Farma (Persero) Tbk",
  "PT Bio Farma (Persero)",
  "PT PLN (Persero)",
  // Tambahkan perusahaan lain sesuai kebutuhan
];

/**
 * Fungsi ini dipanggil saat user akses Web App URL di browser
 * Serve HTML form (Step 1) langsung dari Google
 */
function doGet(e) {
  // Route berdasarkan parameter page
  const page = e.parameter.page || "step1";

  // Logging untuk debugging
  Logger.log("doGet called with page: " + page);

  let htmlOutput;

  if (page === "step2") {
    Logger.log("Loading step2.html");
    htmlOutput = HtmlService.createHtmlOutputFromFile("step2").setTitle(
      "Kuesioner Supplier - PT PLN Nusantara Power 2025"
    );
  } else if (page === "thankyou") {
    Logger.log("Loading thankyou.html");
    htmlOutput = HtmlService.createHtmlOutputFromFile("thankyou").setTitle(
      "Terima Kasih - Supplier Gathering 2025"
    );
  } else if (page === "scanner") {
    Logger.log("Loading scanner.html");
    htmlOutput = HtmlService.createHtmlOutputFromFile("scanner").setTitle(
      "Scanner QR Code - Buku Tamu Supplier Gathering 2025"
    );
  } else {
    Logger.log("Loading step1.html (default)");
    htmlOutput = HtmlService.createHtmlOutputFromFile("step1").setTitle(
      "Supplier Gathering 2025 - PT PLN Nusantara Power"
    );
  }

  // Force mobile-friendly rendering (PENTING untuk responsiveness!)
  return htmlOutput
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes"
    );
}

/**
 * Fungsi untuk menyimpan data dari client-side (dipanggil via google.script.run)
 */
function saveFormData(data) {
  try {
    // Validasi data
    if (!data.step1 || !data.kuesioner) {
      throw new Error("Data tidak lengkap");
    }

    // Validasi: Cek apakah perusahaan sudah submit
    const companyName = data.step1.namaPerusahaan;
    if (isCompanySubmitted(companyName)) {
      throw new Error(
        "Perusahaan '" +
          companyName +
          "' sudah mengisi kuesioner. Setiap perusahaan hanya dapat mengisi 1 kali."
      );
    }

    // Buka spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Simpan data kehadiran (Step 1)
    saveKehadiranData(ss, data.step1);

    // 2. Simpan data kuesioner per unit
    if (data.kuesioner["UP PAITON"]) {
      saveKuesionerData(
        ss,
        SHEET_NAMES.PAITON,
        data.step1,
        data.kuesioner["UP PAITON"]
      );
    }

    if (data.kuesioner["UP BRANTAS"]) {
      saveKuesionerData(
        ss,
        SHEET_NAMES.BRANTAS,
        data.step1,
        data.kuesioner["UP BRANTAS"]
      );
    }

    if (data.kuesioner["UP PACITAN"]) {
      saveKuesionerData(
        ss,
        SHEET_NAMES.PACITAN,
        data.step1,
        data.kuesioner["UP PACITAN"]
      );
    }

    // Generate rekap otomatis setelah data tersimpan
    try {
      generateRekapKehadiran();
    } catch (rekapError) {
      Logger.log("Error generating rekap: " + rekapError.toString());
      // Jangan throw error, karena data sudah tersimpan
    }

    // Kirim email dengan QR code untuk konfirmasi kehadiran
    try {
      if (data.step1.kehadiran === "Ya") {
        sendQRCodeEmail(data.step1);
        Logger.log("✅ Email QR code berhasil dikirim ke: " + data.step1.email);
      }
    } catch (emailError) {
      Logger.log("⚠️ Error sending email: " + emailError.toString());
      // Jangan throw error, karena data sudah tersimpan
    }

    return { success: true, message: "Data berhasil disimpan" };
  } catch (error) {
    Logger.log("Error saveFormData: " + error.toString());
    throw new Error("Terjadi kesalahan: " + error.toString());
  }
}

/**
 * Fungsi ini dipanggil saat form di-submit (HTTP POST)
 * Kept for backward compatibility if needed
 */
function doPost(e) {
  try {
    // Parse JSON data dari request
    const data = JSON.parse(e.postData.contents);

    // Validasi data
    if (!data.step1 || !data.kuesioner) {
      return createResponse(false, "Data tidak lengkap");
    }

    // Buka spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Simpan data kehadiran (Step 1)
    saveKehadiranData(ss, data.step1);

    // 2. Simpan data kuesioner per unit
    if (data.kuesioner["UP PAITON"]) {
      saveKuesionerData(
        ss,
        SHEET_NAMES.PAITON,
        data.step1,
        data.kuesioner["UP PAITON"]
      );
    }

    if (data.kuesioner["UP BRANTAS"]) {
      saveKuesionerData(
        ss,
        SHEET_NAMES.BRANTAS,
        data.step1,
        data.kuesioner["UP BRANTAS"]
      );
    }

    if (data.kuesioner["UP PACITAN"]) {
      saveKuesionerData(
        ss,
        SHEET_NAMES.PACITAN,
        data.step1,
        data.kuesioner["UP PACITAN"]
      );
    }

    return createResponse(true, "Data berhasil disimpan");
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return createResponse(false, "Terjadi kesalahan: " + error.toString());
  }
}

/**
 * Simpan data kehadiran ke Sheet 1
 */
function saveKehadiranData(ss, step1Data) {
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.KEHADIRAN);

  // Header columns (jika belum ada)
  if (sheet.getLastRow() === 0) {
    const headers = [
      "Timestamp",
      "Nama Perusahaan",
      "Email",
      "Nama Peserta",
      "Jabatan",
      "Unit Pembangkit",
      "Konfirmasi Kehadiran",
    ];
    sheet.appendRow(headers);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#667eea");
    headerRange.setFontColor("#ffffff");
  }

  // Prepare data row
  const timestamp = new Date();
  const units = step1Data.units.join(", ");

  const rowData = [
    timestamp,
    step1Data.namaPerusahaan || "",
    step1Data.email,
    step1Data.nama,
    step1Data.jabatan,
    units,
    step1Data.kehadiran,
  ];

  sheet.appendRow(rowData);
}

/**
 * Simpan data kuesioner ke sheet yang sesuai
 */
function saveKuesionerData(ss, sheetName, step1Data, kuesionerData) {
  const sheet = getOrCreateSheet(ss, sheetName);

  // Header columns (jika belum ada)
  if (sheet.getLastRow() === 0) {
    const headers = [
      "Timestamp",
      "Nama Perusahaan",
      "Email",
      "Nama Peserta",
      "Jabatan",
      "Nomor HP/WA (Kuesioner)",
      // Aktivitas Pengadaan (q1-q15)
      "Durasi pengumuman lelang",
      "Waktu penyampaian Permintaan Penawaran",
      "Isi dan kejelasan RKS",
      "Spesifikasi barang/jasa",
      "Keterbukaan/transparansi tender",
      "Akses informasi spesifikasi",
      "Kewajaran Nilai HPS",
      "Proses Penerbitan PO/Kontrak",
      "Spesifikasi dalam PO/Kontrak",
      "Persyaratan dokumen pengiriman",
      "Ketersediaan sarana bongkar muat",
      "Pemeriksaan Barang 14 hari",
      "Pengurusan Surat Ijin Kerja",
      "Pembayaran",
      "Sistem Pengadaan",
      // SDM (q16-q22)
      "Profesionalisme Staff",
      "Integritas Staff",
      "Sikap/Attitude Staff",
      "Kecepatan/Responsif",
      "Kemampuan Komunikasi",
      "Respon Panitia Sanggah",
      "Hubungan Kerjasama",
      // E-Proc dan Sarana (q23-q26)
      "Proses E-Proc",
      "Pengoperasian E-Proc/E-Auction",
      "Sarana ruangan",
      "Sarana parkir",
      // Tanggapan Tambahan
      "Hubungan kerjasama berjalan baik",
      "Kekurangan dan Kelebihan",
      "Penerapan PLN NP BERSIH",
      // Kritik dan Saran
      "Kritik",
      "Saran",
    ];
    sheet.appendRow(headers);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1e3c72");
    headerRange.setFontColor("#ffffff");
    headerRange.setWrap(true);

    // Set column widths
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 120);

    for (let i = 7; i <= 32; i++) {
      sheet.setColumnWidth(i, 80);
    }

    sheet.setColumnWidth(33, 150);
    sheet.setColumnWidth(34, 300);
    sheet.setColumnWidth(35, 300);
    sheet.setColumnWidth(36, 300);
    sheet.setColumnWidth(37, 300);
  }

  // Prepare data row
  const timestamp = new Date();

  const rowData = [
    timestamp,
    step1Data.namaPerusahaan || "",
    step1Data.email,
    step1Data.nama,
    step1Data.jabatan,
    "'" + kuesionerData.nomorHp, // Prepend dengan apostrophe untuk preserve leading zero
    kuesionerData.q1,
    kuesionerData.q2,
    kuesionerData.q3,
    kuesionerData.q4,
    kuesionerData.q5,
    kuesionerData.q6,
    kuesionerData.q7,
    kuesionerData.q8,
    kuesionerData.q9,
    kuesionerData.q10,
    kuesionerData.q11,
    kuesionerData.q12,
    kuesionerData.q13,
    kuesionerData.q14,
    kuesionerData.q15,
    kuesionerData.q16,
    kuesionerData.q17,
    kuesionerData.q18,
    kuesionerData.q19,
    kuesionerData.q20,
    kuesionerData.q21,
    kuesionerData.q22,
    kuesionerData.q23,
    kuesionerData.q24,
    kuesionerData.q25,
    kuesionerData.q26,
    kuesionerData.kerjasamaBaik,
    kuesionerData.kekuranganKelebihan,
    kuesionerData.programBersih,
    kuesionerData.kritik,
    kuesionerData.saran,
  ];

  sheet.appendRow(rowData);
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Get or create sheet at the end (rightmost position)
 */
function getOrCreateSheetAtEnd(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Insert sheet at the end (rightmost position)
    const numSheets = ss.getSheets().length;
    sheet = ss.insertSheet(sheetName, numSheets);
    Logger.log("✅ Sheet '" + sheetName + "' dibuat di posisi paling kanan");
  }
  return sheet;
}

function createResponse(success, message, data = null) {
  const response = {
    success: success,
    message: message,
  };

  if (data) {
    response.data = data;
  }

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON
  );
}

/**
 * Fungsi untuk mendapatkan list perusahaan yang sudah submit
 */
function getSubmittedCompanies() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.KEHADIRAN);

    if (!sheet || sheet.getLastRow() <= 1) {
      return []; // Belum ada data
    }

    const dataRange = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
    const values = dataRange.getValues();

    const submittedCompanies = values
      .map((row) => row[0])
      .filter((company) => company && company.toString().trim() !== "");

    Logger.log("Submitted companies: " + submittedCompanies.join(", "));
    return submittedCompanies;
  } catch (error) {
    Logger.log("Error getSubmittedCompanies: " + error.toString());
    return [];
  }
}

/**
 * Fungsi untuk mendapatkan list perusahaan yang masih available
 */
function getAvailableCompanies() {
  try {
    const submittedCompanies = getSubmittedCompanies();

    const availableCompanies = COMPANY_LIST.filter((company) => {
      return !submittedCompanies.includes(company);
    });

    Logger.log(
      "Available companies: " +
        availableCompanies.length +
        "/" +
        COMPANY_LIST.length
    );
    return availableCompanies;
  } catch (error) {
    Logger.log("Error getAvailableCompanies: " + error.toString());
    return COMPANY_LIST;
  }
}

/**
 * Fungsi untuk cek apakah perusahaan sudah submit
 */
function isCompanySubmitted(companyName) {
  try {
    const submittedCompanies = getSubmittedCompanies();
    return submittedCompanies.includes(companyName);
  } catch (error) {
    Logger.log("Error isCompanySubmitted: " + error.toString());
    return false;
  }
}

/**
 * Test connection
 */
function testConnection() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("✅ Berhasil terhubung ke spreadsheet: " + ss.getName());
    return true;
  } catch (error) {
    Logger.log("❌ Error: " + error.toString());
    return false;
  }
}

/**
 * Generate rekap kehadiran otomatis
 */
function generateRekapKehadiran() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetKehadiran = ss.getSheetByName(SHEET_NAMES.KEHADIRAN);
    if (!sheetKehadiran || sheetKehadiran.getLastRow() <= 1) {
      Logger.log("Tidak ada data kehadiran untuk direkap");
      return;
    }

    // Get or create sheet Rekap Kehadiran at correct position (after Pacitan, before Buku Tamu)
    let rekapSheet = ss.getSheetByName(SHEET_NAMES.REKAP);
    const pacitanSheet = ss.getSheetByName(SHEET_NAMES.PACITAN);

    if (!rekapSheet) {
      // Create new sheet at correct position (after Pacitan)
      if (pacitanSheet) {
        const pacitanIndex = pacitanSheet.getIndex(); // 1-based index
        // insertSheet uses 0-based index, so we use pacitanIndex directly (which equals 0-based position after Pacitan)
        rekapSheet = ss.insertSheet(SHEET_NAMES.REKAP, pacitanIndex);
        Logger.log(
          "✅ Sheet 'Rekap Kehadiran' dibuat di posisi " +
            pacitanIndex +
            " (setelah Kuesioner UP Pacitan)"
        );
      } else {
        rekapSheet = ss.insertSheet(SHEET_NAMES.REKAP);
        Logger.log(
          "⚠️ Sheet Pacitan tidak ditemukan, Rekap dibuat di posisi default"
        );
      }
    } else {
      // Sheet exists, move to correct position if needed
      if (pacitanSheet) {
        const pacitanIndex = pacitanSheet.getIndex(); // 1-based index
        const currentRekapIndex = rekapSheet.getIndex(); // 1-based index
        // moveActiveSheet uses 0-based index
        const targetIndex = pacitanIndex; // 0-based target position (after Pacitan)

        // Only move if not already in correct position
        // Compare 1-based indices: if Rekap is not right after Pacitan
        if (currentRekapIndex !== pacitanIndex + 1) {
          ss.setActiveSheet(rekapSheet);
          ss.moveActiveSheet(targetIndex);
          Logger.log(
            "✅ Sheet 'Rekap Kehadiran' dipindah ke posisi 0-based " +
              targetIndex +
              " (setelah Kuesioner UP Pacitan)"
          );
        }
      }
      rekapSheet.clear();
    }
    const data = sheetKehadiran.getDataRange().getValues();
    const vendorMap = {};
    const vendorPerUnit = {
      PAITON: new Set(),
      BRANTAS: new Set(),
      PACITAN: new Set(),
    };
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const namaVendor = row[1];
      const unitString = row[5];
      const konfirmasiKehadiran = row[6]; // Kolom "Konfirmasi Kehadiran"

      // Skip vendor yang tidak konfirmasi "Ya"
      if (konfirmasiKehadiran !== "Ya") continue;

      if (!namaVendor || !unitString) continue;
      const units = unitString
        .toString()
        .split(",")
        .map((s) => s.trim().toUpperCase());
      if (!vendorMap[namaVendor]) {
        vendorMap[namaVendor] = {
          jumlahSubmit: 0,
          units: [],
        };
      }
      units.forEach((unit) => {
        if (!vendorMap[namaVendor].units.includes(unit)) {
          vendorMap[namaVendor].units.push(unit);
          vendorMap[namaVendor].jumlahSubmit++;
        }
        if (unit.includes("PAITON")) vendorPerUnit["PAITON"].add(namaVendor);
        if (unit.includes("BRANTAS")) vendorPerUnit["BRANTAS"].add(namaVendor);
        if (unit.includes("PACITAN")) vendorPerUnit["PACITAN"].add(namaVendor);
      });
    }
    const vendorPaiton = vendorPerUnit["PAITON"].size;
    const vendorBrantas = vendorPerUnit["BRANTAS"].size;
    const vendorPacitan = vendorPerUnit["PACITAN"].size;
    const totalVendorHadir = Object.keys(vendorMap).length;

    // Judul
    rekapSheet.appendRow(["REKAP KEHADIRAN SUPPLIER GATHERING 2025"]);
    rekapSheet.appendRow([""]); // Baris kosong
    rekapSheet.appendRow(["Jumlah vendor yang hadir:", totalVendorHadir]);
    rekapSheet.appendRow(["Vendor UP Paiton:", vendorPaiton]);
    rekapSheet.appendRow(["Vendor UP Brantas:", vendorBrantas]);
    rekapSheet.appendRow(["Vendor UP Pacitan:", vendorPacitan]);
    // Header tabel: SEKARANG di baris 6
    rekapSheet.appendRow(["No", "Nama Vendor Hadir", "Jumlah Submit"]);

    let no = 1;
    Object.entries(vendorMap)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([nama, data]) => {
        rekapSheet.appendRow([no++, nama, data.jumlahSubmit]);
      });

    // === FORMAT STYLING ===
    rekapSheet.getRange(1, 1, 1, 3).merge();
    rekapSheet
      .getRange(1, 1)
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground("#1e3c72")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");

    // Summary (Baris 3-6, empat baris: total dan 3 unit)
    rekapSheet
      .getRange(3, 1, 4, 2)
      .setFontWeight("bold")
      .setBackground("#eef2ff");

    // Table header biru di baris 6
    const headerRow = 6;
    rekapSheet
      .getRange(headerRow, 1, 1, 3)
      .setFontWeight("bold")
      .setBackground("#667eea")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");

    // Column widths
    rekapSheet.setColumnWidth(1, 250); // Diperlebar untuk text seperti "Jumlah vendor yang hadir:"
    rekapSheet.setColumnWidth(2, 300);
    rekapSheet.setColumnWidth(3, 120);

    // Border tabel
    if (no > 1) {
      const tableRange = rekapSheet.getRange(headerRow, 1, no, 3);
      tableRange.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID
      );
    }
  } catch (error) {
    Logger.log("❌ Error generateRekapKehadiran: " + error.toString());
    throw error;
  }
}

/**
 * Setup initial sheets
 */
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  Object.values(SHEET_NAMES).forEach((sheetName) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log("✅ Sheet dibuat: " + sheetName);
    } else {
      Logger.log("ℹ️ Sheet sudah ada: " + sheetName);
    }
  });

  Logger.log("✅ Setup sheets selesai!");
}

// ============================================================
// QR CODE & EMAIL AUTOMATION FUNCTIONS
// ============================================================

/**
 * Generate QR Code URL menggunakan API yang lebih reliable untuk email
 * Menggunakan qrserver.com API (gratis, tidak butuh API key, reliable untuk email)
 */
function generateQRCodeUrl(supplierData) {
  // Data yang akan di-encode dalam QR Code (tanpa units)
  const qrData = JSON.stringify({
    nama: supplierData.nama,
    perusahaan: supplierData.namaPerusahaan,
    email: supplierData.email,
    timestamp: new Date().getTime(),
  });

  // Generate QR Code menggunakan qrserver.com API
  // Lebih reliable untuk email daripada Google Charts API
  // Size: 300x300 pixels, format PNG
  const qrCodeUrl =
    "https://api.qrserver.com/v1/create-qr-code/?size=300x300&format=png&data=" +
    encodeURIComponent(qrData);

  return qrCodeUrl;
}

/**
 * Generate QR Code sebagai inline attachment (Base64)
 * Alternatif method yang lebih reliable untuk semua email client
 */
function generateQRCodeInline(supplierData) {
  try {
    // Data yang akan di-encode dalam QR Code (tanpa units)
    const qrData = JSON.stringify({
      nama: supplierData.nama,
      perusahaan: supplierData.namaPerusahaan,
      email: supplierData.email,
      timestamp: new Date().getTime(),
    });

    // Fetch QR code image dari qrserver.com
    const qrCodeUrl =
      "https://api.qrserver.com/v1/create-qr-code/?size=300x300&format=png&data=" +
      encodeURIComponent(qrData);

    // Download QR code as blob
    const response = UrlFetchApp.fetch(qrCodeUrl);
    const qrBlob = response.getBlob();
    qrBlob.setName("qrcode_kehadiran.png");

    return qrBlob;
  } catch (error) {
    Logger.log("Error generating inline QR code: " + error.toString());
    return null;
  }
}

/**
 * Kirim email dengan QR Code ke supplier
 */
function sendQRCodeEmail(supplierData) {
  try {
    // Generate QR Code sebagai inline image (blob)
    const qrCodeBlob = generateQRCodeInline(supplierData);

    if (!qrCodeBlob) {
      throw new Error("Gagal generate QR code");
    }

    // Email subject
    const subject =
      "QR Code Kehadiran - Supplier Gathering 2025 | PT PLN Nusantara Power";

    // Email body (HTML)
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
          }
          .header {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 30px;
            text-align: center;
            border-radius: 8px 8px 0 0;
          }
          .content {
            background: #f8f9fa;
            padding: 30px;
            border: 1px solid #e2e8f0;
          }
          .qr-container {
            text-align: center;
            background: white;
            padding: 30px;
            margin: 20px 0;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .qr-code {
            max-width: 100%;
            width: 250px;
            height: auto;
            display: block;
            margin: 0 auto;
          }
          @media only screen and (max-width: 600px) {
            .qr-code {
              width: 200px;
            }
            .qr-container {
              padding: 20px 10px;
            }
          }
          .info-box {
            background: white;
            padding: 20px;
            margin: 20px 0;
            border-left: 4px solid #667eea;
            border-radius: 4px;
          }
          .info-item {
            margin: 10px 0;
          }
          .info-label {
            font-weight: bold;
            color: #1e3c72;
          }
          .footer {
            text-align: center;
            padding: 20px;
            color: #718096;
            font-size: 12px;
          }
          .important-note {
            background: #fff3cd;
            border: 1px solid #ffc107;
            padding: 15px;
            margin: 20px 0;
            border-radius: 4px;
            color: #856404;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>SUPPLIER GATHERING 2025</h1>
          <p>PT PLN NUSANTARA POWER</p>
        </div>
        
        <div class="content">
          <h2>Terima Kasih atas Konfirmasi Kehadiran Anda!</h2>
          
          <p>Yth. <strong>${supplierData.nama}</strong>,</p>
          
          <p>Terima kasih telah mengisi kuesioner dan mengkonfirmasi kehadiran untuk acara <strong>Supplier Gathering 2025</strong>.</p>
          
          <div class="info-box">
            <div class="info-item">
              <span class="info-label">Nama:</span> ${supplierData.nama}
            </div>
            <div class="info-item">
              <span class="info-label">Perusahaan:</span> ${
                supplierData.namaPerusahaan
              }
            </div>
            <div class="info-item">
              <span class="info-label">Unit:</span> ${supplierData.units.join(
                ", "
              )}
            </div>
          </div>
          
          <div class="important-note">
            <strong>⚠️ PENTING:</strong> Simpan QR Code di bawah ini dan tunjukkan kepada panitia saat acara berlangsung untuk keperluan absensi.
          </div>
          
          <div class="qr-container">
            <h3>QR Code Kehadiran Anda</h3>
            <img src="cid:qrcode" alt="QR Code" class="qr-code" />
            <p style="color: #718096; font-size: 12px; margin-top: 15px;">
              Simpan atau screenshot QR code ini
            </p>
          </div>
          
          <div class="info-box">
            <h3 style="color: #1e3c72;">Detail Acara</h3>
            <div class="info-item">
              <span class="info-label">📅 Tanggal:</span> Kamis, 30 Oktober 2025
            </div>
            <div class="info-item">
              <span class="info-label">🕐 Waktu:</span> 08.00 WIB - Selesai
            </div>
            <div class="info-item">
              <span class="info-label">📍 Lokasi:</span> Novotel Samator Surabaya<br>
              <span style="padding-left: 30px;">Jl. Raya Kedung Baruk No.26-28, Kedung Baruk, Kec. Rungkut, Surabaya, Jawa Timur 60298</span>
            </div>
            <div class="info-item">
              <span class="info-label">📞 Kontak:</span><br>
              <span style="padding-left: 30px;">0822-5781-9900 (Fahmi Achmad)</span><br>
              <span style="padding-left: 30px;">0853-3545-8328 (Aminudin)</span>
            </div>
          </div>
          
          <p>Kami menunggu kehadiran Anda. Sampai jumpa di acara!</p>
          
          <p style="margin-top: 30px;">
            Hormat kami,<br>
            <strong>Tim PT PLN Nusantara Power</strong>
          </p>
        </div>
        
        <div class="footer">
          <p>Email ini dikirim secara otomatis oleh sistem. Mohon tidak membalas email ini.</p>
          <p>&copy; 2025 PT PLN Nusantara Power. All rights reserved.</p>
        </div>
      </body>
      </html>
    `;

    // Plain text version (fallback)
    const plainBody = `
SUPPLIER GATHERING 2025 - PT PLN NUSANTARA POWER

Terima Kasih atas Konfirmasi Kehadiran Anda!

Yth. ${supplierData.nama},

Terima kasih telah mengisi kuesioner dan mengkonfirmasi kehadiran untuk acara Supplier Gathering 2025.

Data Anda:
- Nama: ${supplierData.nama}
- Perusahaan: ${supplierData.namaPerusahaan}
- Unit: ${supplierData.units.join(", ")}

PENTING: Buka email ini di perangkat yang mendukung HTML untuk melihat QR Code Anda.
QR Code telah dilampirkan sebagai gambar inline di email ini.

Detail Acara:
- Tanggal: Kamis, 30 Oktober 2025
- Waktu: 08.00 WIB - Selesai
- Lokasi: Novotel Samator Surabaya
  Jl. Raya Kedung Baruk No.26-28, Kedung Baruk, Kec. Rungkut, Surabaya, Jawa Timur 60298
- Kontak:
  0822-5781-9900 (Fahmi Achmad)
  0853-3545-8328 (Aminudin)

Sampai jumpa di acara!

Hormat kami,
Tim PT PLN Nusantara Power
    `;

    // Send email menggunakan MailApp dengan inline QR code image
    MailApp.sendEmail({
      to: supplierData.email,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: "PT PLN Nusantara Power - Supplier Gathering 2025",
      inlineImages: {
        qrcode: qrCodeBlob, // Embed QR code sebagai inline image dengan Content-ID "qrcode"
      },
    });

    Logger.log("✅ Email berhasil dikirim ke: " + supplierData.email);
    return true;
  } catch (error) {
    Logger.log("❌ Error sendQRCodeEmail: " + error.toString());
    throw error;
  }
}

/**
 * Save data hasil scan QR ke sheet Buku Tamu
 */
function saveAttendanceRecord(qrData) {
  try {
    // Parse QR data
    const data = typeof qrData === "string" ? JSON.parse(qrData) : qrData;

    // Validasi data
    if (!data.nama || !data.perusahaan || !data.email) {
      throw new Error("Data QR tidak valid");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateSheetAtEnd(ss, SHEET_NAMES.BUKU_TAMU);

    // Setup header jika belum ada
    if (sheet.getLastRow() === 0) {
      const headers = [
        "Timestamp Scan",
        "Nama Peserta",
        "Nama Perusahaan",
        "Email",
        "Status",
      ];
      sheet.appendRow(headers);

      // Format header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1e3c72");
      headerRange.setFontColor("#ffffff");

      // Set column widths
      sheet.setColumnWidth(1, 180);
      sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 250);
      sheet.setColumnWidth(4, 200);
      sheet.setColumnWidth(5, 100);
    }

    // Cek duplicate - apakah supplier sudah scan sebelumnya
    const isDuplicate = checkDuplicateAttendance(sheet, data.email);
    if (isDuplicate) {
      return {
        success: false,
        message: "Supplier ini sudah melakukan absensi sebelumnya",
        data: data,
      };
    }

    // Append data ke sheet
    const timestamp = new Date();
    const rowData = [
      timestamp,
      data.nama,
      data.perusahaan,
      data.email,
      "Hadir",
    ];

    sheet.appendRow(rowData);

    Logger.log("✅ Data kehadiran berhasil disimpan: " + data.nama);

    return {
      success: true,
      message: "Absensi berhasil! Selamat datang " + data.nama,
      data: data,
    };
  } catch (error) {
    Logger.log("❌ Error saveAttendanceRecord: " + error.toString());
    return {
      success: false,
      message: "Error: " + error.toString(),
    };
  }
}

/**
 * Cek apakah supplier sudah melakukan absensi sebelumnya
 */
function checkDuplicateAttendance(sheet, email) {
  try {
    if (sheet.getLastRow() <= 1) {
      return false; // Belum ada data
    }

    // Get all emails from column 4 (Email column)
    const dataRange = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1);
    const emails = dataRange.getValues().flat();

    // Check if email exists
    return emails.includes(email);
  } catch (error) {
    Logger.log("Error checkDuplicateAttendance: " + error.toString());
    return false;
  }
}
