# 📤 WhatsApp Bulk Sender - UI/UX Workshop Poster Sender

Script Python ini digunakan untuk mengirimkan pesan promosi dan gambar poster workshop secara otomatis ke banyak nomor WhatsApp menggunakan data dari file Excel. Dilengkapi dengan antarmuka grafis (GUI) berbasis Tkinter, script ini sangat cocok untuk keperluan promosi seperti event, workshop, seminar, dan lainnya.

## 📌 Deskripsi

Dengan bantuan `pywhatkit` dan antarmuka GUI, script ini akan:
- Membaca data nama dan nomor HP dari file Excel.
- Memformat nomor HP ke format internasional (contoh: `+62xxx`).
- Menampilkan pratinjau data Excel dalam tabel sebelum pengiriman.
- Mengirimkan **gambar poster** dan **caption** yang dipersonalisasi ke masing-masing kontak.
- Menambahkan jeda acak untuk menghindari deteksi spam oleh WhatsApp Web.
- Menyediakan kontrol untuk memulai, menjeda, dan menghentikan pengiriman pesan.

## ⚙️ Fitur

- ✅ Antarmuka grafis yang ramah pengguna.
- ✅ Pratinjau data Excel dalam tabel yang dapat digulir.
- ✅ Dukungan format nomor lokal dan internasional.
- ✅ Pesan personal dengan variabel `{nama}`, `{nomor}`, dan `{no}`.
- ✅ Kirim gambar poster dengan caption kustom.
- ✅ Pengaturan jeda minimum dan maksimum antar pesan.
- ✅ Pilih rentang penerima untuk pengiriman.
- ✅ Penanganan error dengan log aktivitas.
- ✅ Progress bar untuk memantau proses pengiriman.

## 📝 Struktur File

```
Pesan-WA-Otomatis/
├── main.py  # Script utama dengan GUI
├── requirements.txt      # Daftar dependensi
├── README.md             # Penjelasan proyek
└── .gitignore            # File untuk mengabaikan file tertentu
```

## 📂 Format File Excel

| No | Nama        | NomorHP      |
|----|-------------|--------------|
| 1  | Budi Setiawan | 081234567890 |
| 2  | Ani Lestari   | 628123456789 |
| ...| ...           | ...          |

Kolom:
- **No**: Nomor urut atau ID penerima.
- **Nama**: Nama penerima untuk personalisasi pesan.
- **NomorHP**: Nomor telepon, akan diformat otomatis ke format internasional.

**Catatan**: File Excel tidak boleh memiliki header; data dimulai dari baris pertama.

## 🚀 Cara Instalasi & Jalankan

### 1. Clone Repository (atau download ZIP)

```bash
git clone https://github.com/yourusername/Pesan-WA-Otomatis.git
cd Pesan-WA-Otomatis
```

Ganti `yourusername` dengan nama pengguna GitHub Anda.

### 2. Install Library yang Dibutuhkan

Pastikan Anda memiliki Python 3.6+ terinstal. Instal dependensi dengan:

```bash
pip install -r requirements.txt
```

Atau instal manual:

```bash
pip install pandas pywhatkit openpyxl
```

Untuk Windows, instal juga `pywin32` jika diperlukan:

```bash
pip install pywin32
```

### 3. Siapkan File

- **File Excel**: Buat file Excel (misalnya, `data.xlsx`) dengan format di atas.
- **Poster**: Siapkan gambar poster (JPG, JPEG, atau PNG) di direktori proyek.
- **WhatsApp Web**: Pastikan Anda sudah login ke WhatsApp Web di browser default (Chrome/Edge).

### 4. Jalankan Script

```bash
python main.py
```

### 5. Gunakan GUI

- **Pilih File**: Klik "Browse" untuk memilih file Excel dan poster.
- **Pratinjau Data**: Klik "Preview Data" untuk melihat isi file Excel.
- **Sesuaikan Pengaturan**:
  - Atur jeda minimum dan maksimum (dalam detik).
  - Tentukan rentang penerima (misalnya, mulai dari 1 hingga 100).
  - Edit template pesan di area teks (gunakan variabel `{nama}`, `{nomor}`, `{no}`).
- **Kirim Pesan**: Klik "Start" untuk memulai, "Pause" untuk menjeda, atau "Stop" untuk menghentikan.
- **Pantau Proses**: Lihat progress bar dan log untuk status pengiriman.

## ❗ Catatan Penting

- **WhatsApp Web**: Script menggunakan `pywhatkit` yang membuka WhatsApp Web di browser default. Jangan tutup browser selama proses pengiriman.
- **Jeda**: Jeda acak (default 10-15 detik) membantu mencegah pemblokiran oleh WhatsApp. Sesuaikan dengan kebutuhan.
- **Format Nomor**: Nomor HP otomatis diformat ke `+62xxx`. Nomor tidak valid akan dilewati.
- **Kepatuhan**: Gunakan script secara bertanggung jawab dan patuhi syarat layanan WhatsApp serta peraturan lokal tentang pengiriman pesan massal.

## 💻 Dibuat Dengan

- Python
- Tkinter (untuk GUI)
- pandas (untuk membaca Excel)
- pywhatkit (untuk pengiriman WhatsApp)
- openpyxl (untuk dukungan Excel)

## 📬 Kontak

Jika Anda mengalami kendala atau ingin berkontribusi, silakan buat issue atau pull request di repository ini 🙌

## 📄 Lisensi

Proyek ini dilisensikan di bawah [MIT License](LICENSE). Lihat file `LICENSE` untuk detailnya.