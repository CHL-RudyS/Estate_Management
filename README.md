# Estate Management — Mock Up

Mock up aplikasi **Estate Management** (Marchand Hype Station, PT. Griya Gardenia Indah) yang berdiri sendiri, dengan tampilan mengikuti FAT Program. Dari FAT hanya layar Login yang diambil.

Buka `index.html` di browser, lalu klik **Masuk**. Akun contoh (kata sandi bebas): `admin@marchand.id`, `billing@marchand.id`, `finance@marchand.id`, `teknik@marchand.id`, `lina@marchand.id` (non-aktif).

## Menu
Dashboard · Unit & Tenant · Tarif & Kontrak · Meteran Air · Invoice (SC & PL, Air, Internet, Profit Sharing, Sewa, Fee Sewa) · Penerimaan · Aging Piutang · Deposit Renovasi · Laporan IPL · Pengaturan (Umum, Pengguna)

## Sumber data
Data diambil dari Excel "01. Base Invoice" (Database Invoice, Invoice IPL SC/PL, Air, Internet, Profit Sharing, Sewa, Fee Sewa, Penerimaan & Pembayaran, Report IPL Marchand, Database BAST). Invoice ditampilkan untuk Mei–Okt 2026; aging dihitung per 6 Nov 2026.

> `index.html` berisi data tenant asli (nama, telepon, email). Simpan repo ini sebagai **private**.

## Membangun ulang
```bash
python3 tools/extract.py "<folder 01. Base Invoice>" data.json
python3 tools/build.py mockup/template.html data.json index.html assets/logo-chl-sm.png
```
`extract.py` membutuhkan `openpyxl`. File Excel dan `data.json` tidak disimpan di repo.
