"""
Modul ini mengirimkan pesan WhatsApp massal dengan poster gambar menggunakan pywhatkit.
"""

import os
import time
import pandas as pd
import pywhatkit as kit

# Baca file Excel
df = pd.read_excel("data2.xlsx", header=None)
df.columns = ["No", "Nama", "NomorHP"]

# Format nomor
def format_nomor(nomor):
    nomor_str = str(nomor)
    if nomor_str.startswith("0"):
        return "+62" + nomor_str[1:]
    elif nomor_str.startswith("8"):
        return "+62" + nomor_str
    elif nomor_str.startswith("62"):
        return "+" + nomor_str
    else:
        return nomor_str

df["NomorFormatted"] = df["NomorHP"].apply(format_nomor)

# Path untuk file gambar poster
POSTER_PATH = "poster.jpg"

# Pastikan file poster ada
if not os.path.exists(POSTER_PATH):
    print(f"Error: File poster '{POSTER_PATH}' tidak ditemukan.")
    exit()

# Kirim pesan ke semua nomor
for index, row in df.iterrows():
    nama = row["Nama"]
    nomor = row["NomorFormatted"]
    pesan = f"""📣 INFO WORKSHOP INATECHNO

Hi {nama}

🎨 Tertarik jadi UI/UX Designer? Tapi bingung mulai dari mana?
Tenang! Inatechno hadir dengan Workshop UI/UX Fundamental: Design Digital untuk Pemula 🔥

💡 Kamu akan belajar:
✅ Konsep dasar UI & UX
✅ Desain prinsip yang bisa langsung dipraktikkan
✅ Cara bikin wireframe & prototype pakai Figma
✅ Menyusun user flow
✅ Cara menilai kualitas desain digital

📅 Tanggal: 26 April 2025
📍 Hybrid Workshop - Kuota Terbatas!

🎟 Daftar sekarang: bit.ly/work-ina"""

    try:
        print(f"Mengirim pesan dan poster ke {nama} ({nomor})...")
        
        # Kirim gambar dengan caption - menggunakan parameter yang benar
        kit.sendwhats_image(
            receiver=nomor,
            img_path=POSTER_PATH,
            caption=pesan,
            wait_time=30,
            tab_close=True
        )
        
        print(f"Berhasil mengirim ke {nama}")
        time.sleep(10)  # jeda antar pesan, biar gak dianggap spam
    except Exception as e:
        print(f"Gagal kirim ke {nama} ({nomor}) - {e}")

print("Proses pengiriman pesan selesai!")
