# indri

Web UI sederhana untuk jalankan analyzer JS:
- `ippd-hourly`
- `twamp-hourly`
- `twamp-daily`
- `s1-hourly`
- `s1-daily`
- `wpcsdm-step1`
- `wpcsdm-transform`

## Local run

```bash
cd /Users/ryanpratama/Desktop/calendar-sandbox/indri
npm install
npm run dev
```

Open: `http://localhost:3000`

Access code default: `indricantik`

## Struktur

- `server.js`: API + static web server
- `public/`: halaman login + upload UI
- `*.js analyzer`: logic analyzer di root folder ini
- `assets/`: file pendukung analyzer `wpcsdm` (opsional)

## Lifecycle file upload

- Semua file upload diproses di folder temporary (`os.tmpdir()`).
- Setelah output terbentuk, folder temporary langsung dihapus otomatis.
- File upload tidak disimpan permanen di project folder.

## Companion file untuk WPCSDM

Mode `wpcsdm-step1` dan `wpcsdm-transform` butuh file tambahan (`SFXL`, `sitelist`, `tagging`).

Jika file di-upload dari UI, file tersebut dipakai langsung.
Kalau tidak di-upload, server akan cari file dari urutan ini:

1. Env var path
2. Folder `/Users/ryanpratama/Desktop/calendar-sandbox/indri/assets/`
3. Folder parent project

### Env vars opsional

- `BOT_INDRI_CODE` (default `indricantik`)
- `BOT_INDRI_SFXL_PATH`
- `BOT_INDRI_SITELIST_PATH`
- `BOT_INDRI_TAGGING_PATH`

## Deploy Vercel

Deploy folder `indri` sebagai project Vercel.

Catatan: upload berbasis JSON base64, jadi ukuran request membesar sekitar 33%.
