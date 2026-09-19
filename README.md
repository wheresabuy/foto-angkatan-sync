# Otomasi Pas Foto 3x4 Teknik Komputer 2025

Sistem otomatisasi sinkronisasi dan generator dokumen pas foto 3x4 (PDF & Word) untuk angkatan Teknik Komputer 2025.

Bekerja secara mandiri di cloud via **GitHub Actions** (24/7 tanpa laptop menyala) dan dapat dijalankan di lokal via **terminal/shortcut**.

---

## Fitur
* **Incremental Sync Google Drive**: Hanya mengunduh dan memproses foto baru secara paralel.
* **Auto 3x4 Cropping**: Penskalaan dan framing wajah proporsional rasio 3:4 (354x472 px @ 300 DPI) dengan panduan potong tipis.
* **Dual Output Documents**:
  1. `Foto_3x4_Teman_Angkatan_Siap_Gunting.pdf` & `.docx`: Hanya berisi foto yang sudah tersedia (tanpa kotak kosong), siap cetak & potong.
  2. `Foto_3x4_Teman_Angkatan_Lengkap.pdf` & `.docx`: Memuat 106 mahasiswa berurutan sesuai NRP dengan kotak garis putus-putus untuk yang belum upload.
* **Direct Drive Upload**: Otomatis mengunggah dokumen PDF terbaru langsung ke folder Google Drive angkatan agar bisa diakses langsung dari HP.
* **Desktop Notification**: Mengirim notifikasi pop-up Ubuntu/GNOME saat dijalankan di komputer lokal.

---

## Penggunaan Lokal
```bash
# Jalankan sinkronisasi sekali jalan
./run_sync.sh

# Jalankan pemantau background berkala
./run_sync.sh --watch

# Paksa buat ulang seluruh dokumen PDF & Word
./run_sync.sh --force
```

---

## Pengaturan GitHub Secrets (Untuk 24/7 Cloud Sync)
Di pengaturan repository GitHub (**Settings > Secrets and variables > Actions > New repository secret**), tambahkan 3 secrets berikut:

1. `GDRIVE_CLIENT_ID`: `44438659992-7kgjeitenc16ssihbtdjbgguch7ju55s.apps.googleusercontent.com`
2. `GDRIVE_CLIENT_SECRET`: `-gMLuQyDiI0XrQS_vx_mhuYF`
3. `GDRIVE_REFRESH_TOKEN`: (Lihat file lokal `.env.github` di laptop)
