# 📤 WhatsApp Blast - Sender

Script Python ini digunakan untuk mengirimkan pesan promosi dan gambar poster workshop secara otomatis ke banyak nomor WhatsApp menggunakan data dari file Excel. Sangat cocok digunakan untuk keperluan promosi seperti event, workshop, seminar, dan lainnya.

## 📌 Deskripsi

Dengan bantuan `pywhatkit`, script ini akan:
- Membaca data nama dan nomor HP dari file `Excel`
- Memformat nomor HP ke format internasional (contoh: +62xxx)
- Mengirimkan **gambar poster** dan **caption** secara otomatis ke masing-masing kontak
- Menambahkan jeda agar tidak dianggap spam oleh WhatsApp Web

## ⚙️ Fitur

- ✅ Support format nomor lokal dan internasional
- ✅ Kirim pesan personal dengan nama penerima
- ✅ Kirim gambar + caption
- ✅ Bekerja dengan WhatsApp Web
- ✅ Penanganan error saat pengiriman

## 📝 Struktur File

```
whatsapp-blast-uiux/
├── main.py          # Script utama
├── data2.xlsx       # File Excel berisi data nama & nomor HP
├── poster.jpg       # Gambar poster yang akan dikirim
└── README.md        # Penjelasan project
```

## 📂 Format File Excel (`data.xlsx`)

| No | Nama        | NomorHP      |
|----|-------------|--------------|
| 1  | Budi Setiawan | 081234567890 |
| 2  | Ani Lestari   | 628123456789 |
| ...| ...           | ...          |

Kolom:
- **No** → Nomor urut
- **Nama** → Nama penerima
- **NomorHP** → Nomor telepon yang akan diformat otomatis

## 🚀 Cara Instalasi & Jalankan

### 1. Clone Repository (atau download ZIP)

```bash
git clone https://github.com/mattrahendra/Pesan-WA-Otomatis.git
cd Pesan-WA-Otomatis
```

### 2. Install Library yang Dibutuhkan

Pastikan kamu sudah menginstall Python. Lalu jalankan:

```bash
pip install -r requirements.txt
```

Atau install manual:

```bash
pip install pandas pywhatkit openpyxl
```

### 3. Install Library Tambahan

Jangan lupa untuk install library di CMD

```bash
pip install pywin32
```


### 4. Jalankan Script

```bash
python main.py
```

> Pastikan:
> - WhatsApp Web aktif dan login di browser default (Chrome/Edge)
> - File `poster.jpg` tersedia dalam direktori
> - `data.xlsx` sudah diisi dengan benar

## ❗ Catatan Penting

- `pywhatkit` akan membuka WhatsApp Web di browser default.
- Pesan akan dijadwalkan otomatis. Jangan tutup browser saat proses berjalan.
- Tambahkan `time.sleep` untuk menghindari terlalu cepat mengirim ke banyak nomor.

## 💻 Dibuat Dengan

- Python
- pandas
- pywhatkit
- openpyxl

## 📬 Kontak

Jika kamu mengalami kendala atau ingin kontribusi, silakan hubungi atau buat issue di repo ini 🙌
