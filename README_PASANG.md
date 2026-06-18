# Sistem Admin Gym — Panduan Pasang dari Nol

Ada 2 bagian: **Frontend** (di GitHub Pages) dan **Backend** (di Google Apps Script).

---

## File & tempatnya

| File | Taruh di |
|------|----------|
| `index.html` | GitHub (repo Pages) |
| `style.css` | GitHub |
| `app.js` | GitHub |
| `config.js` | GitHub |
| `404.html` | GitHub |
| `Code.gs` | Google Apps Script (Ekstensi → Apps Script) |

> `Code.gs` JANGAN ditaruh di GitHub. Itu backend, jalan di Apps Script.

---

## LANGKAH 1 — Backend (Apps Script)

1. Buka Google Sheet kamu → menu **Ekstensi → Apps Script**.
2. Hapus semua isi file `Kode.gs`, **paste** isi `Code.gs` ini, lalu **Save** (Ctrl+S).
3. Pastikan baris paling atas `SPREADSHEET_ID` sudah berisi ID sheet kamu (cuma ID, bukan link).
4. Dropdown fungsi (atas) → pilih **`tesKoneksi`** → **Jalankan**.
   - Kalau muncul popup izin: Review permissions → pilih akun → Advanced → Go to ... (unsafe) → Allow.
   - Cek Log eksekusi: harus muncul nama file sheet kamu.
5. Dropdown → **`setupGymSheets`** → **Jalankan**. Ini bikin sheet & isi 200 slot kunci (Cowo 01-100 + Cewe 01-100).

### Deploy backend
6. **Terapkan (Deploy) → New deployment → Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy → **copy Web app URL** (berakhiran `/exec`).
7. Tes URL: buka di browser + `?action=cekversi`. Harus muncul JSON `{"ok":true,...,"version":"..."}`.

> Kalau nanti update Code.gs: **Deploy → Manage deployments → ikon pensil → Version: New version → Deploy.** (Bukan bikin deployment baru — biar URL tetap sama.)

---

## LANGKAH 2 — Frontend (GitHub)

1. Buka `config.js`, isi `SCRIPT_URL` dengan Web app URL dari langkah 6 di atas (yang `/exec`).
2. Upload `index.html`, `style.css`, `app.js`, `config.js`, `404.html` ke repo GitHub Pages kamu.
3. Tunggu 1-2 menit, lalu buka situs kamu. **Hard refresh (Ctrl+Shift+R)**.

---

## Cara pakai

- **Input Masuk**: isi Admin, Nama Pelanggan, Jenis Kunci, Nomor Kunci, Status = Masuk → Simpan.
  Di tabel Daftar Kunci, kunci itu jadi **Dipakai** + nama muncul + tombol **Keluar**.
- **Input Keluar**: klik tombol **Keluar** di tabel (atau isi manual Status = Keluar).

---

## Kalau ada masalah (fungsi penyelamat di Code.gs)

Jalankan dari dropdown editor Apps Script:

- **`perbaikiSemua`** — beresin semua sekaligus (buka merge + rapikan struktur + reset rekap bulanan). Pakai ini kalau bingung.
- `scanMerge` — cek sheet mana yang punya merge bermasalah (lihat Log).
- `bersihkanSemuaMerge` — buka semua merge di semua sheet.
- `perbaikiStrukturHarian` — rapikan kolom REKAP_HARIAN kalau geser.
- `resetRekapBulanan` — hapus & bangun ulang sheet bulanan.
- `cekVersi` — pastikan kode di editor versi terbaru.

---

## Catatan warna

- **Cewe = merah**, **Cowo = hitam** (di sheet & di web badge).
- Status: Dipakai = merah, Kosong = hijau.
