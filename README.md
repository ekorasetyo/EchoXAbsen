# Scanlog Importer

Aplikasi desktop (Lazarus / Free Pascal) untuk mengimpor CSV scanlog mesin absensi,
dijalankan di Windows 7 atau lebih baru. Tanpa DLL tambahan (memakai `TBufDataset`).

## Fitur
- Import CSV dengan deteksi otomatis pemisah `,` `;` atau Tab
- Baris judul tambahan di atas header dilewati otomatis
- Cari berdasarkan PIN / Nama
- Hitung Jam Kerja per hari (scan terakhir - scan pertama)
- Tab Rekap per Pegawai: hari hadir, scan tak lengkap, total dan rata-rata jam
- Export hasil filter ke CSV (pemisah `;`)

## Format CSV
```
PIN;NIP;Nama;Jabatan;Departemen;Kantor;Tanggal;Scan 1;Scan 2;Scan 3
2;;PEGAWAI 01;;;;30-10-2026;10:58:27;19:04:02;
```
Tanggal `dd-mm-yyyy`, jam `hh:mm:ss`. Contoh data samaran ada di `contoh_scanlog.csv`.

## Build
1. Install [Lazarus](https://www.lazarus-ide.org/) (32-bit disarankan untuk Windows 7).
2. Buka `scanlog.lpi`, lalu tekan F9 atau Shift+F9.
3. Hasil: `scanlog.exe`.

Atau lewat command line: `lazbuild scanlog.lpi`

## Catatan
- Form dibuat lewat kode (tanpa `.lfm`).
- Jam kerja dihitung per hari, jadi shift lintas hari belum digabung.
- Jangan commit data absensi asli ke repo publik.

## Lisensi
MIT

## Unduh .exe
Setiap push ke `main` akan membangun `scanlog.exe` lewat GitHub Actions
(tab **Actions** > run terbaru > Artifacts). Untuk rilis: `git tag v1.0 && git push --tags`,
file akan muncul di tab **Releases**.

## Screenshot
Letakkan screenshot di `docs/screenshot.png`, lalu tambahkan `![Scanlog Importer](docs/screenshot.png)`.
