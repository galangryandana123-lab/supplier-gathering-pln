/**
 * Google Apps Script untuk Supplier Gathering PLN
 * Backend + Frontend Hosting (HTML Service)
 */

// KONFIGURASI - Ganti dengan Spreadsheet ID Anda
const SPREADSHEET_ID = "175u9HZZOxpGCHS8GwurWbivY2i6gZn-DRgroehkx9MY";

// ============================================================
// KONFIGURASI RESEND API - UBAH HANYA API KEY DI BAWAH INI!
// ============================================================
// Dapatkan API key dari: https://resend.com/api-keys
const RESEND_API_KEY = "re_xxxxxxxxxxxxxxxxxxxxxxxxxx"; // ⬅️ GANTI DENGAN API KEY ANDA

// ============================================================
// KONFIGURASI LAINNYA (JANGAN DIUBAH KECUALI PERLU)
// ============================================================
const RESEND_FROM_EMAIL = "PLN Supplier Gathering <noreply@galangproject.my.id>";
// Domain galangproject.my.id sudah verified di Resend

// Nama-nama sheet
const SHEET_NAMES = {
  KEHADIRAN: "Data Kehadiran",
  PAITON: "Kuesioner UP Paiton",
  BRANTAS: "Kuesioner UP Brantas",
  PACITAN: "Kuesioner UP Pacitan",
  REKAP: "Rekap Kehadiran",
  BUKU_TAMU: "Buku Tamu",
  EMAIL_QUEUE: "Email Queue",
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
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    // Validasi data
    if (!data.step1 || !data.kuesioner) {
      throw new Error("Data tidak lengkap");
    }

    // Validasi: Cek apakah perusahaan sudah submit (di bawah lock untuk mencegah race condition)
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

    // Update cache perusahaan yang sudah submit agar getAvailableCompanies lebih cepat
    try {
      updateSubmittedCompaniesCache(companyName);
    } catch (cacheErr) {
      Logger.log("Cache update error: " + cacheErr);
    }

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

    // Generate rekap DEFER ke background (dijalankan oleh time-driven trigger)
    // Tidak perlu dipanggil di sini untuk menghindari blocking response

    // Kirim email ASYNC dengan queue (tidak blocking response)
    try {
      if (data.step1.kehadiran === "Ya") {
        addEmailToQueue(data.step1);
        Logger.log(
          "✅ Email ditambahkan ke antrian untuk: " + data.step1.email
        );
      }
    } catch (emailError) {
      Logger.log("⚠️ Error adding email to queue: " + emailError.toString());
      // Jangan throw error, karena data sudah tersimpan
    }

    return { success: true, message: "Data berhasil disimpan" };
  } catch (error) {
    Logger.log("Error saveFormData: " + error.toString());
    throw new Error("Terjadi kesalahan: " + error.toString());
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

// doPost removed - not used (client uses google.script.run, not HTTP POST)

/**
 * Simpan data kehadiran ke Sheet 1 (optimized: batch write)
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
    // Tulis header via setValues (lebih cepat)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#667eea");
    headerRange.setFontColor("#ffffff");
    
    // Set timestamp column format to show date and time
    sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm:ss");
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(2, 300); // Nama Perusahaan
    sheet.setColumnWidth(3, 220); // Email
    sheet.setColumnWidth(4, 200); // Nama Peserta
    sheet.setColumnWidth(5, 150); // Jabatan
    sheet.setColumnWidth(6, 200); // Unit Pembangkit
    sheet.setColumnWidth(7, 150); // Konfirmasi Kehadiran
  }

  // Prepare data row
  const timestamp = new Date();
  const units = step1Data.units.join(", ");

  const rowData = [
    timestamp,
    step1Data.namaPerusahaan || "",
    (step1Data.email || "").toLowerCase().trim(), // Normalize email
    step1Data.nama,
    step1Data.jabatan,
    units,
    step1Data.kehadiran,
  ];

  // Tulis data via setValues di row berikutnya (lebih cepat daripada appendRow)
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
}

/**
 * Simpan data kuesioner ke sheet yang sesuai (optimized: batch write)
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
    // Tulis header via setValues (lebih cepat)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1e3c72");
    headerRange.setFontColor("#ffffff");
    headerRange.setWrap(true);

    // Set column widths
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(2, 250); // Nama Perusahaan
    sheet.setColumnWidth(3, 200); // Email
    sheet.setColumnWidth(4, 180); // Nama Peserta
    sheet.setColumnWidth(5, 150); // Jabatan
    sheet.setColumnWidth(6, 130); // Nomor HP/WA

    // Rating columns (q1-q26)
    for (let i = 7; i <= 32; i++) {
      sheet.setColumnWidth(i, 110);
    }

    // Text response columns
    sheet.setColumnWidth(33, 200); // Hubungan kerjasama berjalan baik
    sheet.setColumnWidth(34, 350); // Kekurangan dan Kelebihan
    sheet.setColumnWidth(35, 350); // Penerapan PLN NP BERSIH
    sheet.setColumnWidth(36, 350); // Kritik
    sheet.setColumnWidth(37, 350); // Saran
    
    // Set timestamp column format to show date and time
    sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm:ss");
  }

  // Prepare data row
  const timestamp = new Date();

  const rowData = [
    timestamp,
    step1Data.namaPerusahaan || "",
    (step1Data.email || "").toLowerCase().trim(), // Normalize email
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

  // Tulis data via setValues di row berikutnya (lebih cepat daripada appendRow)
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
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
    // Coba ambil dari cache dulu untuk mengurangi read ke spreadsheet
    const cache = CacheService.getScriptCache();
    const cached = cache.get("submitted_companies");
    if (cached) {
      const arr = JSON.parse(cached);
      Logger.log("Submitted companies (cache): " + arr.length);
      return arr;
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.KEHADIRAN);

    if (!sheet || sheet.getLastRow() <= 1) {
      cache.put("submitted_companies", JSON.stringify([]), 300); // 5 menit
      return []; // Belum ada data
    }

    const dataRange = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
    const values = dataRange.getValues();

    const submittedCompanies = values
      .map((row) => row[0])
      .filter((company) => company && company.toString().trim() !== "");

    cache.put("submitted_companies", JSON.stringify(submittedCompanies), 300);
    Logger.log("Submitted companies (fresh): " + submittedCompanies.length);
    return submittedCompanies;
  } catch (error) {
    Logger.log("Error getSubmittedCompanies: " + error.toString());
    return [];
  }
}

/**
 * Update cache daftar perusahaan yang sudah submit
 */
function updateSubmittedCompaniesCache(companyName) {
  try {
    if (!companyName) return;
    const cache = CacheService.getScriptCache();
    const key = "submitted_companies";
    const cached = cache.get(key);
    let arr = [];
    if (cached) {
      arr = JSON.parse(cached);
    }
    if (!arr.includes(companyName)) {
      arr.push(companyName);
    }
    cache.put(key, JSON.stringify(arr), 300); // 5 menit
  } catch (e) {
    Logger.log("Error updateSubmittedCompaniesCache: " + e.toString());
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

// ============================================================
// EMAIL QUEUE & ASYNC PROCESSING
// ============================================================

/**
 * Tambahkan email ke antrian untuk dikirim secara async
 */
function addEmailToQueue(supplierData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const queueSheet = getOrCreateSheet(ss, SHEET_NAMES.EMAIL_QUEUE);

    // Setup header jika belum ada
    if (queueSheet.getLastRow() === 0) {
      const headers = [
        "Timestamp",
        "Email",
        "Nama",
        "Nama Perusahaan",
        "Units",
        "Status",
        "Retry Count",
        "Last Attempt",
        "Error Message",
      ];
      queueSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // Format header
      const headerRange = queueSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1e3c72");
      headerRange.setFontColor("#ffffff");
      
      // Set timestamp column format to show date and time
      queueSheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      queueSheet.getRange("H:H").setNumberFormat("dd/mm/yyyy hh:mm:ss"); // Last Attempt
      
      // Set column widths
      queueSheet.setColumnWidth(1, 160); // Timestamp
      queueSheet.setColumnWidth(2, 220); // Email
      queueSheet.setColumnWidth(3, 200); // Nama
      queueSheet.setColumnWidth(4, 280); // Nama Perusahaan
      queueSheet.setColumnWidth(5, 200); // Units
      queueSheet.setColumnWidth(6, 100); // Status
      queueSheet.setColumnWidth(7, 100); // Retry Count
      queueSheet.setColumnWidth(8, 160); // Last Attempt
      queueSheet.setColumnWidth(9, 300); // Error Message
    }

    // Tambahkan ke antrian
    const timestamp = new Date();
    const rowData = [
      timestamp,
      (supplierData.email || "").toLowerCase().trim(),
      supplierData.nama,
      supplierData.namaPerusahaan || "",
      supplierData.units.join(", "),
      "pending",
      0,
      "",
      "",
    ];

    const nextRow = queueSheet.getLastRow() + 1;
    queueSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    Logger.log("✅ Email added to queue: " + supplierData.email);
  } catch (error) {
    Logger.log("❌ Error addEmailToQueue: " + error.toString());
    throw error;
  }
}

/**
 * Proses email queue secara batch (dipanggil oleh time-driven trigger)
 * Rate limiting: 90 email per run untuk safety (buffer 10 dari limit 100/hari)
 */
function processEmailQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("⚠️ processEmailQueue already running, skipping...");
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const queueSheet = ss.getSheetByName(SHEET_NAMES.EMAIL_QUEUE);

    if (!queueSheet || queueSheet.getLastRow() <= 1) {
      Logger.log("No emails in queue");
      return;
    }

    // Ambil semua data
    const data = queueSheet.getDataRange().getValues();
    let emailsSent = 0;
    const MAX_EMAILS_PER_RUN = 90; // Safety buffer dari limit 100/hari

    for (let i = 1; i < data.length && emailsSent < MAX_EMAILS_PER_RUN; i++) {
      const row = data[i];
      const status = row[5]; // Status column
      const retryCount = row[6]; // Retry Count column

      // Skip jika sudah sent atau failed (retry > 3)
      if (status === "sent" || retryCount >= 3) {
        continue;
      }

      try {
        // Rekonstruksi supplierData
        const supplierData = {
          email: row[1],
          nama: row[2],
          namaPerusahaan: row[3],
          units: row[4].split(", "),
        };

        // Kirim email
        sendQRCodeEmail(supplierData);

        // Update status jadi "sent"
        queueSheet.getRange(i + 1, 6).setValue("sent");
        queueSheet.getRange(i + 1, 8).setValue(new Date());
        emailsSent++;

        Logger.log("✅ Email sent to: " + supplierData.email);

        // Sleep 1 detik antar email untuk menghindari rate limit spike
        Utilities.sleep(1000);
      } catch (error) {
        // Update retry count dan error message
        const newRetryCount = retryCount + 1;
        queueSheet.getRange(i + 1, 7).setValue(newRetryCount);
        queueSheet.getRange(i + 1, 8).setValue(new Date());
        queueSheet
          .getRange(i + 1, 9)
          .setValue(error.toString().substring(0, 200));

        if (newRetryCount >= 3) {
          queueSheet.getRange(i + 1, 6).setValue("failed");
          Logger.log("❌ Email failed after 3 retries: " + row[1]);
        } else {
          Logger.log("⚠️ Email retry " + newRetryCount + "/3: " + row[1]);
        }
      }
    }

    Logger.log(
      "📧 processEmailQueue completed: " + emailsSent + " emails sent"
    );
  } catch (error) {
    Logger.log("❌ Error processEmailQueue: " + error.toString());
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// TRIGGER SETUP FUNCTIONS
// ============================================================

/**
 * Setup semua time-driven triggers untuk proses background
 * JALANKAN MANUAL 1 KALI SAJA setelah deploy
 */
function setupTriggers() {
  // Hapus semua trigger lama terlebih dahulu
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    const funcName = trigger.getHandlerFunction();
    if (
      funcName === "processEmailQueue" ||
      funcName === "generateRekapKehadiran"
    ) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("🗑️ Deleted old trigger: " + funcName);
    }
  });

  // Setup trigger untuk processEmailQueue (tiap 1 menit)
  ScriptApp.newTrigger("processEmailQueue")
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("✅ Created trigger: processEmailQueue (every 1 minute)");

  // Setup trigger untuk generateRekapKehadiran (tiap 1 menit)
  ScriptApp.newTrigger("generateRekapKehadiran")
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log("✅ Created trigger: generateRekapKehadiran (every 1 minute)");

  Logger.log(
    "\n🎉 All triggers setup complete!\n" +
      "- Email queue: processed every 1 minute (max 90 emails/run)\n" +
      "- Rekap kehadiran: updated every 1 minute\n" +
      "- Estimated: 90 emails/min = 5400 emails/hour\n" +
      "- Daily capacity: ~5000-8000 emails (with Resend API)"
  );
}

/**
 * Hapus semua triggers (untuk cleanup/reset)
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    ScriptApp.deleteTrigger(trigger);
    Logger.log("🗑️ Deleted trigger: " + trigger.getHandlerFunction());
  });
  Logger.log("✅ All triggers removed");
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
        .map((s) => s.trim().toUpperCase())
        .filter(Boolean);
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

    // Susun output sekali, lalu tulis via setValues (lebih cepat daripada appendRow berulang)
    const output = [];
    output.push(["REKAP KEHADIRAN SUPPLIER GATHERING 2025", "", ""]);
    output.push(["", "", ""]); // Baris kosong
    output.push(["Jumlah vendor yang hadir:", totalVendorHadir, ""]);
    output.push(["Vendor UP Paiton:", vendorPaiton, ""]);
    output.push(["Vendor UP Brantas:", vendorBrantas, ""]);
    output.push(["Vendor UP Pacitan:", vendorPacitan, ""]);
    const headerRowIndex = output.length + 1; // 1-based index di sheet untuk header tabel
    output.push(["No", "Nama Vendor Hadir", "Jumlah Submit"]);

    let no = 1;
    Object.entries(vendorMap)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([nama, data]) => {
        output.push([no++, nama, data.jumlahSubmit]);
      });

    // Tulis ke sheet sekaligus
    rekapSheet.getRange(1, 1, output.length, 3).setValues(output);

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

    // Table header biru
    rekapSheet
      .getRange(headerRowIndex, 1, 1, 3)
      .setFontWeight("bold")
      .setBackground("#667eea")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");

    // Column widths
    rekapSheet.setColumnWidth(1, 280); // Kolom 1: Label summary + No (perlu lebar untuk label)
    rekapSheet.setColumnWidth(2, 350); // Kolom 2: Nilai + Nama Vendor (diperlebar untuk nama perusahaan panjang)
    rekapSheet.setColumnWidth(3, 140); // Kolom 3: Jumlah Submit

    // Border tabel
    const totalRows = output.length - headerRowIndex + 1; // termasuk baris header
    if (totalRows > 1) {
      const tableRange = rekapSheet.getRange(headerRowIndex, 1, totalRows, 3);
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

/**
 * Reorder sheets - move Email Queue to before Buku Tamu
 * Urutan akhir:
 * 1. Data Kehadiran
 * 2. Kuesioner UP Paiton
 * 3. Kuesioner UP Brantas
 * 4. Kuesioner UP Pacitan
 * 5. Rekap Kehadiran
 * 6. Email Queue
 * 7. Buku Tamu
 * 
 * JALANKAN MANUAL 1 KALI untuk mengatur ulang urutan sheet
 */
function reorderSheets() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Get all sheets
    const kehadiranSheet = ss.getSheetByName(SHEET_NAMES.KEHADIRAN);
    const paitonSheet = ss.getSheetByName(SHEET_NAMES.PAITON);
    const brantasSheet = ss.getSheetByName(SHEET_NAMES.BRANTAS);
    const pacitanSheet = ss.getSheetByName(SHEET_NAMES.PACITAN);
    const rekapSheet = ss.getSheetByName(SHEET_NAMES.REKAP);
    const emailQueueSheet = ss.getSheetByName(SHEET_NAMES.EMAIL_QUEUE);
    const bukuTamuSheet = ss.getSheetByName(SHEET_NAMES.BUKU_TAMU);
    
    if (!emailQueueSheet || !bukuTamuSheet) {
      Logger.log("⚠️ Sheet Email Queue atau Buku Tamu tidak ditemukan");
      return;
    }
    
    // Set desired order (1-based index)
    let position = 1;
    
    if (kehadiranSheet) {
      ss.setActiveSheet(kehadiranSheet);
      ss.moveActiveSheet(position++);
      Logger.log("✅ Moved: Data Kehadiran to position " + (position - 1));
    }
    
    if (paitonSheet) {
      ss.setActiveSheet(paitonSheet);
      ss.moveActiveSheet(position++);
      Logger.log("✅ Moved: Kuesioner UP Paiton to position " + (position - 1));
    }
    
    if (brantasSheet) {
      ss.setActiveSheet(brantasSheet);
      ss.moveActiveSheet(position++);
      Logger.log("✅ Moved: Kuesioner UP Brantas to position " + (position - 1));
    }
    
    if (pacitanSheet) {
      ss.setActiveSheet(pacitanSheet);
      ss.moveActiveSheet(position++);
      Logger.log("✅ Moved: Kuesioner UP Pacitan to position " + (position - 1));
    }
    
    if (rekapSheet) {
      ss.setActiveSheet(rekapSheet);
      ss.moveActiveSheet(position++);
      Logger.log("✅ Moved: Rekap Kehadiran to position " + (position - 1));
    }
    
    // Move Email Queue before Buku Tamu
    ss.setActiveSheet(emailQueueSheet);
    ss.moveActiveSheet(position++);
    Logger.log("✅ Moved: Email Queue to position " + (position - 1));
    
    if (bukuTamuSheet) {
      ss.setActiveSheet(bukuTamuSheet);
      ss.moveActiveSheet(position++);
      Logger.log("✅ Moved: Buku Tamu to position " + (position - 1));
    }
    
    Logger.log("\n🎉 Sheet reordering complete!");
    Logger.log("Final order:");
    Logger.log("1. Data Kehadiran");
    Logger.log("2. Kuesioner UP Paiton");
    Logger.log("3. Kuesioner UP Brantas");
    Logger.log("4. Kuesioner UP Pacitan");
    Logger.log("5. Rekap Kehadiran");
    Logger.log("6. Email Queue");
    Logger.log("7. Buku Tamu");
    
  } catch (error) {
    Logger.log("❌ Error reorderSheets: " + error.toString());
    throw error;
  }
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
 * Generate QR Code sebagai base64 string untuk Resend API
 */
function generateQRCodeBase64(supplierData) {
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
    
    // Convert blob to base64 string untuk Resend API
    const qrBase64 = Utilities.base64Encode(qrBlob.getBytes());

    return qrBase64;
  } catch (error) {
    Logger.log("Error generating QR code base64: " + error.toString());
    return null;
  }
}

/**
 * Kirim email dengan QR Code ke supplier menggunakan Resend API
 */
function sendQRCodeEmail(supplierData) {
  try {
    // Validasi Resend API Key
    if (!RESEND_API_KEY || RESEND_API_KEY === "re_xxxxxxxxxxxxxxxxxxxxxxxxxx") {
      throw new Error("RESEND_API_KEY belum dikonfigurasi. Set di baris 9-10 Code.gs");
    }

    // Generate QR Code sebagai base64
    const qrCodeBase64 = generateQRCodeBase64(supplierData);

    if (!qrCodeBase64) {
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

    // Send email menggunakan Resend API
    const payload = {
      from: RESEND_FROM_EMAIL,
      to: [supplierData.email],
      subject: subject,
      html: htmlBody,
      attachments: [
        {
          content: qrCodeBase64,
          filename: "qrcode_kehadiran.png",
          content_id: "qrcode", // Content-ID untuk inline display (reference: cid:qrcode)
        },
      ],
    };

    const options = {
      method: "post",
      headers: {
        Authorization: "Bearer " + RESEND_API_KEY,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(
      "https://api.resend.com/emails",
      options
    );
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode !== 200) {
      const errorData = JSON.parse(responseBody);
      throw new Error(
        "Resend API error: " + (errorData.message || responseBody)
      );
    }

    const result = JSON.parse(responseBody);
    Logger.log(
      "✅ Email berhasil dikirim ke: " +
        supplierData.email +
        " (ID: " +
        result.id +
        ")"
    );
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
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
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
      sheet.setColumnWidth(1, 160); // Timestamp Scan
      sheet.setColumnWidth(2, 220); // Nama Peserta
      sheet.setColumnWidth(3, 300); // Nama Perusahaan
      sheet.setColumnWidth(4, 220); // Email
      sheet.setColumnWidth(5, 110); // Status
      
      // Set timestamp column format to show date and time
      sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm:ss");
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
      (data.email || "").toLowerCase().trim(), // Normalize email
      "Hadir",
    ];

    sheet.appendRow(rowData);

    // Tandai di cache agar pengecekan duplicate berikutnya cepat
    try {
      const cache = CacheService.getScriptCache();
      const emailKey = "attendance_" + data.email.toLowerCase().trim();
      cache.put(emailKey, "1", 3600); // 1 jam
    } catch (e) {}

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
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/**
 * Cek apakah supplier sudah melakukan absensi sebelumnya
 * ALWAYS check sheet first to ensure data consistency (cache bisa outdated jika data dihapus manual)
 */
function checkDuplicateAttendance(sheet, email) {
  try {
    const normalizedEmail = (email || "").toLowerCase().trim();
    const cache = CacheService.getScriptCache();
    const key = "attendance_" + normalizedEmail;

    // Check sheet kosong
    if (sheet.getLastRow() <= 1) {
      // Sheet kosong, remove cache jika ada
      try {
        cache.remove(key);
      } catch (e) {}
      return false; // Belum ada data
    }

    // ALWAYS check sheet first (source of truth) - jangan percaya cache
    // Karena data bisa dihapus manual dari sheet tapi cache masih ada
    const found = sheet
      .createTextFinder(normalizedEmail)
      .matchCase(false)
      .matchEntireCell(true)
      .findNext();

    if (found) {
      // Data ada di sheet, update cache
      try {
        cache.put(key, "1", 3600); // 1 jam
      } catch (e) {}
      return true;
    } else {
      // Data tidak ada di sheet, remove cache jika ada (mungkin data dihapus)
      try {
        cache.remove(key);
      } catch (e) {}
      return false;
    }
  } catch (error) {
    Logger.log("Error checkDuplicateAttendance: " + error.toString());
    return false;
  }
}
