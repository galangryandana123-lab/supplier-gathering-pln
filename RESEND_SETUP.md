# Setup Resend API - Panduan Lengkap

## 📋 Checklist Setup

- [ ] Daftar akun Resend.com
- [ ] Dapatkan API Key
- [ ] Setup domain (opsional)
- [ ] Konfigurasi Code.gs
- [ ] Test kirim email
- [ ] Deploy dan jalankan setupTriggers()

---

## 🚀 Step-by-Step Setup

### 1. Daftar Akun Resend (5 menit)

1. Buka https://resend.com
2. Klik **Sign Up**
3. Pilih salah satu:
   - Sign up with GitHub (recommended - tercepat)
   - Sign up with Email
4. Verifikasi email jika diperlukan
5. Selesai! ✅

**Free Tier Benefits:**
- 3,000 emails/bulan
- 100 emails/hari
- No credit card required
- Perfect untuk testing & production

---

### 2. Dapatkan API Key (2 menit)

1. Setelah login, buka https://resend.com/api-keys
2. Klik tombol **Create API Key**
3. Berikan nama: `PLN Supplier Gathering` atau nama lain yang mudah dikenali
4. Pilih permission: **Full Access** (recommended)
5. Klik **Add**
6. **PENTING**: Copy API key yang muncul (format: `re_xxxxxxxxxxxxxxxx`)
   
   ⚠️ **API key hanya ditampilkan SEKALI!** Simpan di tempat aman.

7. Paste API key ke notepad/text editor sementara

---

### 3. Setup Domain (Opsional - 10-15 menit)

**Apakah perlu setup domain?**
- ✅ **YA** - Jika ingin email dari domain sendiri (contoh: noreply@plnsurabaya.com)
- ✅ **YA** - Untuk production (lebih professional & deliverability lebih baik)
- ❌ **TIDAK** - Jika hanya testing, bisa pakai `onboarding@resend.dev`

#### Setup Domain Step-by-Step:

1. **Buka Resend Dashboard → Domains**
   - https://resend.com/domains

2. **Add Domain**
   - Klik **Add Domain**
   - Masukkan domain: `yourdomain.com` (tanpa `www` atau subdomain)
   - Klik **Add**

3. **Tambahkan DNS Records**
   
   Resend akan menampilkan 3 DNS records yang harus ditambahkan:
   
   **a. SPF Record (TXT)**
   ```
   Type: TXT
   Name: @ atau yourdomain.com
   Value: v=spf1 include:_spf.resend.com ~all
   ```
   
   **b. DKIM Record (TXT)**
   ```
   Type: TXT
   Name: resend._domainkey
   Value: (copy dari Resend dashboard)
   ```
   
   **c. DMARC Record (TXT)** (opsional tapi recommended)
   ```
   Type: TXT
   Name: _dmarc
   Value: v=DMARC1; p=none
   ```

4. **Login ke Domain Provider**
   
   Tergantung provider Anda (Niagahoster, Cloudflare, Namecheap, dll):
   
   **Contoh: Cloudflare**
   - Login → pilih domain → DNS
   - Add Record → pilih Type: TXT
   - Masukkan Name dan Value sesuai di atas
   - Repeat untuk ketiga records
   
   **Contoh: cPanel/Niagahoster**
   - cPanel → Zone Editor
   - Add Record → TXT
   - Masukkan Name dan Value
   - Repeat untuk ketiga records

5. **Verifikasi di Resend**
   - Tunggu 5-15 menit (propagasi DNS)
   - Kembali ke Resend → Domains
   - Klik **Verify** pada domain Anda
   - Jika berhasil, status akan jadi ✅ **Verified**

6. **Test Kirim Email**
   - Resend Dashboard → Overview → Send Test Email
   - Gunakan format: `Nama Pengirim <noreply@yourdomain.com>`

---

### 4. Konfigurasi Code.gs (1 menit)

1. Buka file `Code.gs`
2. Cari baris 11-12 (section RESEND API CONFIGURATION)
3. Ganti nilai:

```javascript
// RESEND API CONFIGURATION
const RESEND_API_KEY = "re_your_actual_api_key_here"; // Paste API key dari step 2
const RESEND_FROM_EMAIL = "PLN Supplier Gathering <noreply@yourdomain.com>"; // Ganti domain
```

**Contoh Konfigurasi:**

**Dengan Domain Terverifikasi (PRODUCTION):**
```javascript
const RESEND_API_KEY = "re_9xWqP7mN5vK2jR8hT3fY6nL4sD1aB";
const RESEND_FROM_EMAIL = "PLN Supplier Gathering <noreply@plnsurabaya.com>";
```

**Tanpa Domain (TESTING ONLY):**
```javascript
const RESEND_API_KEY = "re_9xWqP7mN5vK2jR8hT3fY6nL4sD1aB";
const RESEND_FROM_EMAIL = "onboarding@resend.dev";
```

⚠️ **PENTING**: 
- Jangan commit API key ke GitHub!
- Jika menggunakan Git, simpan API key di environment variable atau file terpisah yang di-ignore

⚠️ **BATASAN `onboarding@resend.dev`**:
- **Hanya bisa mengirim email ke email yang terdaftar di akun Resend Anda sendiri**
- Contoh: Jika akun Resend Anda `sgplnnusantarapower@gmail.com`, maka email hanya bisa dikirim ke `sgplnnusantarapower@gmail.com`
- Untuk mengirim ke recipient lain (supplier), **WAJIB verifikasi domain sendiri**
- Testing dengan `onboarding@resend.dev` hanya untuk memastikan integrasi API berfungsi

---

### 5. Test Kirim Email (3 menit)

Sebelum deploy, test dulu apakah konfigurasi sudah benar:

1. **Buka Apps Script Editor**
2. **Copy-paste fungsi test ini** ke Code.gs (di paling bawah):

```javascript
/**
 * Test fungsi untuk verifikasi Resend API
 * JALANKAN MANUAL 1x untuk testing
 */
function testResendAPI() {
  try {
    // Test data
    const testData = {
      nama: "Test User",
      namaPerusahaan: "PT Test Company",
      email: "your-email@gmail.com", // GANTI dengan email Anda sendiri
      units: ["UP PAITON", "UP BRANTAS"],
    };
    
    Logger.log("🧪 Memulai test Resend API...");
    Logger.log("📧 Target email: " + testData.email);
    
    // Panggil fungsi sendQRCodeEmail
    const result = sendQRCodeEmail(testData);
    
    if (result) {
      Logger.log("✅ TEST BERHASIL!");
      Logger.log("📬 Email test berhasil dikirim ke: " + testData.email);
      Logger.log("🔍 Cek inbox/spam Anda untuk verifikasi");
      return "SUCCESS - Cek email Anda!";
    }
  } catch (error) {
    Logger.log("❌ TEST GAGAL!");
    Logger.log("Error: " + error.toString());
    
    // Troubleshooting hints
    if (error.toString().includes("RESEND_API_KEY")) {
      Logger.log("\n💡 SOLUSI: Pastikan RESEND_API_KEY sudah diisi di baris 11");
    } else if (error.toString().includes("401")) {
      Logger.log("\n💡 SOLUSI: API key tidak valid. Cek kembali di https://resend.com/api-keys");
    } else if (error.toString().includes("403")) {
      Logger.log("\n💡 SOLUSI: Domain tidak terverifikasi atau from email tidak valid");
    } else if (error.toString().includes("429")) {
      Logger.log("\n💡 SOLUSI: Rate limit tercapai. Tunggu beberapa saat lalu coba lagi");
    }
    
    return "FAILED - Lihat error di atas";
  }
}
```

3. **Edit email test**:
   - Ganti `your-email@gmail.com` dengan email Anda sendiri
   
4. **Jalankan fungsi**:
   - Pilih fungsi: `testResendAPI`
   - Klik ▶️ **Run**
   - Authorize jika diminta
   
5. **Cek hasil**:
   - Lihat **Execution log** (View → Logs atau Ctrl+Enter)
   - Jika berhasil: `✅ TEST BERHASIL!`
   - Jika gagal: Lihat error message dan ikuti petunjuk troubleshooting

6. **Cek email Anda**:
   - Buka inbox email yang Anda gunakan untuk test
   - Cari email dari "PLN Supplier Gathering"
   - Jika tidak ada di Inbox, cek folder **Spam/Junk**
   - Pastikan QR Code muncul dengan benar di email

**Expected Result:**
```
🧪 Memulai test Resend API...
📧 Target email: your-email@gmail.com
✅ Email berhasil dikirim ke: your-email@gmail.com (ID: 49a3999c-0ce1-4ea6-ab68-afcd6dc2e794)
✅ TEST BERHASIL!
📬 Email test berhasil dikirim ke: your-email@gmail.com
🔍 Cek inbox/spam Anda untuk verifikasi
```

---

### 6. Deploy & Setup Triggers (5 menit)

Jika test berhasil, lanjutkan ke deployment:

1. **Deploy Web App**:
   - Deploy → New deployment
   - Type: Web app
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy
   - Copy Web App URL

2. **Setup Triggers**:
   - Di Apps Script Editor
   - Pilih fungsi: `setupTriggers`
   - Klik ▶️ **Run**
   - Authorize jika diminta

3. **Verifikasi Triggers**:
   - Klik icon ⏰ **Triggers** di sidebar kiri
   - Pastikan ada 2 triggers aktif:
     - `processEmailQueue` (every 1 minute)
     - `generateRekapKehadiran` (every 1 minute)

4. **Test End-to-End**:
   - Buka Web App URL di browser
   - Isi form dan submit
   - Cek:
     - QR Code muncul instant di browser? ✅
     - Email masuk ke Sheet "Email Queue" dengan status `pending`? ✅
     - Tunggu 1 menit → Status berubah jadi `sent`? ✅
     - Email diterima? ✅

---

## 🔍 Troubleshooting

### Error: "RESEND_API_KEY belum dikonfigurasi"

**Penyebab**: API key belum diisi di Code.gs

**Solusi**:
1. Buka Code.gs baris 11
2. Ganti `re_xxxxxxxxxxxxxxxxxxxxxxxxxx` dengan API key asli
3. Save
4. Deploy ulang

---

### Error: "401 Unauthorized"

**Penyebab**: API key tidak valid atau salah

**Solusi**:
1. Buka https://resend.com/api-keys
2. Verifikasi API key masih aktif
3. Jika sudah lupa/hilang, buat API key baru
4. Update Code.gs dengan API key yang baru
5. Deploy ulang

---

### Error: "403 Forbidden" atau "Domain not verified"

**Penyebab**: Domain belum diverifikasi atau from email salah

**Solusi**:
1. Cek domain verification: https://resend.com/domains
2. Pastikan status domain: ✅ **Verified**
3. Jika belum verified, ikuti langkah setup domain di atas
4. Atau gunakan `onboarding@resend.dev` untuk testing (hanya bisa kirim ke email akun Resend Anda)

---

### Error: "You can only send testing emails to your own email address"

**Penyebab**: Menggunakan `onboarding@resend.dev` tapi mencoba kirim ke email lain

**Contoh Error:**
```
You can only send testing emails to your own email address (sgplnnusantarapower@gmail.com).
To send emails to other recipients, please verify a domain.
```

**Solusi**:
1. **Untuk Testing**: Ganti email test ke email akun Resend Anda
   ```javascript
   // Di fungsi testResendAPI(), ganti:
   email: "sgplnnusantarapower@gmail.com" // Email akun Resend Anda
   ```

2. **Untuk Production**: **WAJIB verifikasi domain**
   - Ikuti langkah "Setup Domain" di atas
   - Ganti `RESEND_FROM_EMAIL` di Code.gs:
   ```javascript
   const RESEND_FROM_EMAIL = "PLN Supplier Gathering <noreply@yourdomain.com>";
   ```
   - Deploy ulang

---

### Error: "429 Too Many Requests"

**Penyebab**: Rate limit tercapai (100 email/hari untuk free tier)

**Solusi**:
1. Tunggu hingga reset (daily limit reset setiap 24 jam)
2. Atau upgrade ke paid plan di https://resend.com/pricing
3. Email akan tetap masuk queue dan terkirim setelah limit reset

---

### Email masuk Spam

**Penyebab**: Domain belum setup dengan benar atau deliverability issue

**Solusi**:
1. Verifikasi SPF, DKIM, DMARC sudah benar
2. Gunakan domain terverifikasi (bukan onboarding@resend.dev)
3. Warming up: Kirim email bertahap (sedikit dulu, naikkan perlahan)
4. Minta penerima mark as "Not Spam" dan add to contacts

---

### QR Code tidak muncul di email

**Penyebab**: Email client blocking images atau error generate QR

**Solusi**:
1. Pastikan email client "Show Images" aktif
2. Test di berbagai email client (Gmail, Outlook, Yahoo, dll)
3. Cek execution log untuk error generate QR code
4. User tetap bisa download QR dari browser (instant display)

---

## 📊 Monitoring & Analytics

### Cek Email Queue Status

**Sheet "Email Queue":**
| Status | Arti |
|--------|------|
| `pending` | Menunggu dikirim |
| `sent` | ✅ Berhasil terkirim |
| `failed` | ❌ Gagal setelah 3x retry |

**Kolom "Error Message"**: Lihat detail error jika status `failed`

### Resend Dashboard

Monitor sending activity:
1. https://resend.com/overview
2. Lihat:
   - Total emails sent
   - Delivery rate
   - Bounce rate
   - Opens & clicks (jika tracking diaktifkan)

### Apps Script Execution Logs

Monitoring real-time:
1. Apps Script Editor → View → Executions
2. Filter: `processEmailQueue`
3. Lihat error logs dan status

---

## 📞 Support

**Resend Documentation:**
- API Reference: https://resend.com/docs/api-reference/introduction
- Guides: https://resend.com/docs
- Status: https://resend.com/status

**Questions?**
- Resend Discord: https://resend.com/discord
- Resend Support: support@resend.com

---

## ✅ Final Checklist

Sebelum production, pastikan semua ini sudah ✅:

- [ ] Resend account created & verified
- [ ] API key obtained & configured in Code.gs
- [ ] Domain verified (atau gunakan onboarding@resend.dev untuk test)
- [ ] Test email berhasil dikirim (fungsi `testResendAPI` passed)
- [ ] Web App deployed
- [ ] Triggers setup (`setupTriggers` dijalankan)
- [ ] 2 triggers aktif di Apps Script → Triggers (keduanya 1 menit)
- [ ] End-to-end test: Form submit → QR instant → Email masuk queue → Email terkirim
- [ ] Email received & QR code visible
- [ ] Sheet reordering executed (`reorderSheets`)

**Jika semua ✅, sistem siap production! 🎉**
