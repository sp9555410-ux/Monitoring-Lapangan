# Monitoring Lapangan (PWA + Apps Script API)

## Setup sekali
1. Aktifkan Apps Script API: https://script.google.com/home/usersettings
2. `npm i` lalu `pip install -r requirements.txt`
3. `npx clasp login`
4. Salin `.clasp.json.example` jadi `.clasp.json`, isi scriptId (Project Settings > Script ID).
5. `npx clasp push -f` (menimpa Code.gs lama dengan versi baru; pastikan Index/GalianInput/Dashboard.html
   lama sudah Anda salin ke `src/` lebih dulu, lalu hapus dari proyek Apps Script bila masih ada).
6. Di editor Apps Script: Project Settings > Script Properties > tambah `API_TOKEN` = kode rahasia Anda.
   Jalankan fungsi `setup()` sekali bila sheet belum ada.
7. Deploy > New deployment > Web app (Execute as: Me, Access: Anyone). Salin URL `/exec`.
8. Salin `.env.example` jadi `.env`, isi `URL_EXEC`.
9. Salin Index.html, GalianInput.html, Dashboard.html ke `src/`.

## Sehari-hari
- `npm run dev`      : auto-build + preview http://localhost:5500
- `npm run push`     : kirim Code.gs ke Apps Script
- `clasp deploy -i DEPLOYMENT_ID -d "catatan"` : perbarui API tanpa mengganti URL /exec
- `npm run deploy:web` : build + deploy PWA ke Netlify (`npx netlify-cli login` sekali)

## Pasang di HP
Android Chrome: menu ⋮ > Install app. iPhone Safari: Share > Add to Home Screen.
# Monitoring-Lapangan
# Monitoring-Lapangan
