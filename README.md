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

- Tanpa Supabase: file dikirim base64 ke API lalu diproses di folder temporary (`os.tmpdir()`).
- Dengan Supabase aktif: browser upload file langsung ke bucket, API hanya menerima path object.
- Setelah output terbentuk, folder temporary lokal selalu dihapus otomatis.
- Object upload di bucket dengan prefix `SUPABASE_UPLOAD_PREFIX` bisa dihapus otomatis setelah proses (`SUPABASE_DELETE_AFTER_PROCESS=true` default).

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
- `MAX_FILE_MB` (default `70`, dipakai untuk mode base64/fallback)

### Env vars Supabase (disarankan untuk Vercel)

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `SUPABASE_BUCKET`
- `SUPABASE_UPLOAD_PREFIX` (default `indri-uploads`)
- `SUPABASE_DELETE_AFTER_PROCESS` (default `true`)

## Deploy Vercel

Deploy folder `indri` sebagai project Vercel.

Jika env Supabase di-set, upload file tidak melewati body function Vercel (langsung ke bucket), jadi aman dari limit payload request ±4.5MB.
