# Supplier Gathering PLN - Apps Script Documentation

## 📋 Overview
Web-based questionnaire system untuk Supplier Gathering 2025 PT PLN Nusantara Power dengan fitur:
- Multi-step form (Step 1: Data perusahaan, Step 2: Kuesioner per unit)
- **QR Code INSTANT di browser** (primary) + Email backup (async)
- Email async dengan antrian untuk handle 500+ supplier
- Scanner QR Code untuk buku tamu
- Auto-generate rekap kehadiran

---

## 🚀 Setup Awal (Deployment)

### 1. Deploy Web App
1. Buka Google Apps Script Editor
2. Deploy → New deployment → Web app
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Copy Web App URL

### 2. **PENTING: Setup Resend API**
Lihat section "Setup Resend API" di bawah untuk mendapatkan API key dan konfigurasi Code.gs.

### 3. **PENTING: Setup Triggers (Jalankan 1x saja)**
Setelah deploy dan konfigurasi Resend, jalankan fungsi ini **MANUAL** dari Apps Script Editor:

```javascript
setupTriggers()
```

Fungsi ini akan membuat 2 time-driven triggers:
- `processEmailQueue`: Jalan tiap **1 menit** (kirim email dari antrian via Resend API)
- `generateRekapKehadiran`: Jalan tiap **1 menit** (update rekap otomatis)

**Verifikasi triggers sudah jalan:**
- Apps Script Editor → Triggers (⏰ icon di sidebar kiri)
- Harusnya ada 2 triggers aktif

---

## 📱 QR Code Instant (Primary Method)

### User Experience:
1. **User submit form** → Loading 2-3 detik
2. **QR Code langsung muncul di browser** ✅
3. User bisa:
   - **Screenshot** QR code
   - **Download PNG** (button tersedia)
   - Simpan untuk absensi saat acara

### Keuntungan:
- ✅ **Zero delay** - tidak perlu tunggu email
- ✅ **No dependency** pada email quota
- ✅ **Reliable** - user langsung dapat QR
- ✅ **Self-service** - user kontrol penuh

---

## 📧 Email Queue System (Backup Method)

### ⚙️ Setup Resend API (PENTING - Lakukan Sebelum Deploy!)

**Resend** adalah email API service untuk mengirim transactional email tanpa batasan MailApp.

#### Langkah Setup:

1. **Daftar Resend** (GRATIS)
   - Buka https://resend.com
   - Sign up dengan email atau GitHub
   - Free tier: **3,000 email/bulan** + **100 email/hari**

2. **Dapatkan API Key**
   - Setelah login → https://resend.com/api-keys
   - Klik "Create API Key"
   - Copy API key (format: `re_xxxxxxxxxxxx`)
   - ⚠️ **Simpan baik-baik**, key hanya ditampilkan 1x!

3. **Setup Domain (Opsional tapi Recommended)**
   - Resend Dashboard → Domains → Add Domain
   - Masukkan domain Anda (contoh: `yourdomain.com`)
   - Tambahkan DNS records (SPF, DKIM, DMARC) ke domain provider
   - Tunggu verifikasi (~5-10 menit)
   - Setelah verified, gunakan format: `Nama Pengirim <noreply@yourdomain.com>`
   
   **Alternatif tanpa domain custom:**
   - Gunakan domain Resend: `onboarding@resend.dev` (untuk testing)
   - ⚠️ **Catatan**: Dengan `onboarding@resend.dev`, Resend hanya mengizinkan mengirim ke email yang terdaftar di akun Resend Anda
   - Untuk kirim ke semua recipient, WAJIB verifikasi domain sendiri
   - Tidak recommended untuk production

4. **Konfigurasi di Code.gs**
   - Buka file `Code.gs` baris 11-12
   - Ganti nilai berikut:
   ```javascript
   const RESEND_API_KEY = "re_xxxxxxxxxxxx"; // API key dari step 2
   const RESEND_FROM_EMAIL = "PLN Supplier Gathering <noreply@yourdomain.com>"; // Email pengirim
   ```

#### Contoh Konfigurasi:
```javascript
// Dengan domain terverifikasi (RECOMMENDED)
const RESEND_API_KEY = "re_AbCdEfGh123456789";
const RESEND_FROM_EMAIL = "PLN Supplier Gathering <noreply@plnsurabaya.com>";

// Atau tanpa domain custom (TESTING ONLY)
const RESEND_API_KEY = "re_AbCdEfGh123456789";
const RESEND_FROM_EMAIL = "onboarding@resend.dev";
```

---

### Cara Kerja:
1. **User submit form** → Email **TIDAK langsung terkirim** (instant response)
2. **QR Code langsung ditampilkan di browser** (primary)
3. Data email masuk ke **Sheet "Email Queue"** dengan status `pending` (backup)
4. Trigger `processEmailQueue` jalan tiap **1 menit**:
   - Ambil max 90 email dengan status `pending`
   - Kirim via **Resend API** + update status jadi `sent`
   - Retry 3x jika gagal, lalu status jadi `failed`

### Rate Limiting:
- **Max 90 email per run** (safety buffer)
- **Delay 1 detik** antar email untuk menghindari spike
- **Estimasi**: 90 email/menit = **~5400 email/jam**
- **Free tier limit**: 100 email/hari, 3000 email/bulan
- ⚠️ **Testing dengan `onboarding@resend.dev`**: Hanya bisa kirim ke email akun Resend Anda sendiri

### Monitoring:
Cek Sheet **"Email Queue"**:
| Status | Arti |
|--------|------|
| `pending` | Menunggu dikirim |
| `sent` | Sudah terkirim ✅ |
| `failed` | Gagal setelah 3x retry ❌ |

### Manual Trigger (jika perlu):
Jika ingin kirim email sekarang (tidak tunggu 5 menit):
```javascript
processEmailQueue()
```

---

## 📊 Rekap Kehadiran Auto-Update

- Trigger `generateRekapKehadiran` jalan tiap **1 menit**
- Sheet **"Rekap Kehadiran"** otomatis ter-update
- **Tidak perlu manual refresh** lagi

### Manual Trigger (jika perlu):
```javascript
generateRekapKehadiran()
```

---

## 🛠️ Maintenance

### Reset Triggers (jika bermasalah):
```javascript
removeTriggers()  // Hapus semua trigger
setupTriggers()   // Buat ulang
```

### Reorder Sheet Position:
Jika ingin mengatur ulang urutan sheet (Email Queue sebelum Buku Tamu):
```javascript
reorderSheets()  // Jalankan 1x manual
```

Urutan sheet setelah reorder:
1. Data Kehadiran
2. Kuesioner UP Paiton
3. Kuesioner UP Brantas
4. Kuesioner UP Pacitan
5. Rekap Kehadiran
6. **Email Queue**
7. Buku Tamu

### Cek Email Queue Status:
```javascript
// Di Apps Script Editor → View → Logs
// Atau cek Sheet "Email Queue" kolom Status
```

### Troubleshooting:

**Q: Email tidak terkirim?**
- Cek Sheet "Email Queue" → kolom "Error Message"
- Cek Logs: View → Executions (lihat error processEmailQueue)
- Pastikan trigger `processEmailQueue` aktif

**Q: Rekap tidak update?**
- Tunggu 5 menit (trigger jalan otomatis)
- Atau manual run: `generateRekapKehadiran()`
- Cek trigger `generateRekapKehadiran` aktif

**Q: Resend API limit tercapai?**
- Free tier: 100 email/hari, 3000/bulan
- Email queue akan pending sampai reset (daily/monthly)
- Upgrade Resend plan jika perlu kapasitas lebih besar

**Q: Error "RESEND_API_KEY belum dikonfigurasi"?**
- Pastikan sudah mengisi `RESEND_API_KEY` di Code.gs baris 11
- Pastikan API key valid (cek di https://resend.com/api-keys)
- Deploy ulang setelah update konfigurasi

---

## 📁 Sheet Structure

| Sheet Name | Fungsi |
|------------|--------|
| Data Kehadiran | Data submit step 1 (konfirmasi kehadiran) |
| Kuesioner UP Paiton | Data kuesioner unit Paiton |
| Kuesioner UP Brantas | Data kuesioner unit Brantas |
| Kuesioner UP Pacitan | Data kuesioner unit Pacitan |
|| Rekap Kehadiran | Auto-generated summary (updated tiap 1 menit) |
| **Email Queue** | Antrian email (status: pending/sent/failed) |
| Buku Tamu | Data scan QR Code kehadiran |

---

## 🔧 Performance Optimizations

### Backend:
- ✅ CacheService untuk daftar perusahaan (TTL 5 menit)
- ✅ LockService untuk race condition prevention
- ✅ Batch write (setValues) bukan appendRow
- ✅ TextFinder untuk duplicate check
- ✅ Email async dengan queue
- ✅ Rekap via time-driven trigger

### Frontend:
- ✅ localStorage cache companies (TTL 5 menit)
- ✅ Debounce search input (200ms)

### Result:
- **Submit response**: 2-3 detik (dari ~5-7 detik)
- **Email capacity**: ~5400 email/jam dengan Resend API (trigger 1 menit)
- **Page load**: <100ms jika cache hit

---

## 📞 Support

Jika ada error atau pertanyaan:
1. Cek **Logs**: Apps Script Editor → View → Executions
2. Cek **Sheet "Email Queue"** kolom Error Message
3. Manual trigger: `processEmailQueue()` atau `generateRekapKehadiran()`

---

## 📝 Notes

- **Primary QR delivery**: Instant di browser (100% reliable, no limit)
- **Backup QR delivery**: Email via Resend API (delay ~1 menit)
- **Resend Free Tier**: 100 email/hari, 3000 email/bulan
- **Testing mode** (`onboarding@resend.dev`): Hanya bisa kirim ke email akun Resend Anda
- **Production mode**: WAJIB verifikasi domain sendiri untuk kirim ke semua recipient
- **Estimasi user**: 500 supplier → email butuh ~10 menit (jika trigger 1 menit), tapi semua langsung dapat QR di browser ✅
- **Production ready**: Resend API lebih reliable daripada MailApp
- **Upgrade option**: Resend paid plans untuk kapasitas lebih besar (opsional)
