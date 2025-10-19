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

### 2. **PENTING: Setup Triggers (Jalankan 1x saja)**
Setelah deploy, jalankan fungsi ini **MANUAL** dari Apps Script Editor:

```javascript
setupTriggers()
```

Fungsi ini akan membuat 2 time-driven triggers:
- `processEmailQueue`: Jalan tiap **5 menit** (kirim email dari antrian)
- `generateRekapKehadiran`: Jalan tiap **5 menit** (update rekap otomatis)

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

### Cara Kerja:
1. **User submit form** → Email **TIDAK langsung terkirim** (instant response)
2. **QR Code langsung ditampilkan di browser** (primary)
3. Data email masuk ke **Sheet "Email Queue"** dengan status `pending` (backup)
4. Trigger `processEmailQueue` jalan tiap **5 menit**:
   - Ambil max 90 email dengan status `pending`
   - Kirim email + update status jadi `sent`
   - Retry 3x jika gagal, lalu status jadi `failed`

### Rate Limiting:
- **Max 90 email per run** (buffer 10 dari limit MailApp 100/hari)
- **Delay 1 detik** antar email untuk menghindari spike
- **Estimasi**: 90 email/5min = 18 email/menit = **~1080 email/jam**

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

- Trigger `generateRekapKehadiran` jalan tiap **5 menit**
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

**Q: Limit MailApp 100/hari tercapai?**
- Email queue akan pending sampai besok
- Atau upgrade ke Google Workspace (1500/hari)

---

## 📁 Sheet Structure

| Sheet Name | Fungsi |
|------------|--------|
| Data Kehadiran | Data submit step 1 (konfirmasi kehadiran) |
| Kuesioner UP Paiton | Data kuesioner unit Paiton |
| Kuesioner UP Brantas | Data kuesioner unit Brantas |
| Kuesioner UP Pacitan | Data kuesioner unit Pacitan |
| Rekap Kehadiran | Auto-generated summary (updated tiap 3 menit) |
| Buku Tamu | Data scan QR Code kehadiran |
| **Email Queue** | Antrian email (status: pending/sent/failed) |

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
- **Email capacity**: ~1000-1500/hari (dengan queue system)
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
- **Backup QR delivery**: Email (delay hingga 5 hari karena limit 100/hari)
- **Limit MailApp**: 100 email/hari (free Gmail)
- **Estimasi user**: 500 supplier → email butuh ~5 hari, tapi semua langsung dapat QR di browser ✅
- **Upgrade option**: Google Workspace ($6/bulan) → 1500 email/hari (opsional, tidak urgent)
