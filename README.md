# afpi_test

📊 AFPI Test Project - Ticket Size & Statistical Analysis

Project ini adalah bagian dari assessment analisis data menggunakan Python, SQLite, dan Streamlit, untuk menghitung dan memvisualisasikan Ticket Size berdasarkan data transaksi (TRX) dan nominal pendanaan (NPP) per bulan.

📁 Struktur Script

File

Deskripsi

npp_test.py

Ekstrak dan proses data NPP dari file Excel ke database SQLite dan Excel. Validasi tren NPP per bulan.

trx_test.py

Ekstrak dan proses data TRX dari file Excel ke database SQLite dan Excel. Validasi tren TRX per bulan.

ticket_size.py

Menghitung Ticket Size bulanan dari data NPP dan TRX, simpan hasil ke SQLite dan export ke Excel.

chart_visualisasi.py

Membuat dashboard interaktif menggunakan Streamlit untuk visualisasi Ticket Size, TRX, dan NPP.



⚙️ Alur Kerja

Buat skrip untuk membaca semua tab sheet TRX & NPP tersebut ke dalam sebuah
database dengan format tabel & kolom
a. Nama Tabel : TRX
Field : id | bulan | trx
b. Nama Tabel : NPP
Field : id | bulan | npp

npp_test.py ➔ Membaca data NPP dari file Excel sheet 16.

trx_test.py ➔ Membaca data TRX dari file Excel sheet 18.

Rules Validasi
TRX bulan H > TRX bulan H-1
NPP bulan H >= NPP bulan H-1

Data dibersihkan, diproses, lalu disimpan ke database SQLite dan file Excel baru dan data trx dan npp ada validasi yang berbentuk txt file.

Hitung Ticket Size:

ticket_size.py ➔ Menggabungkan data NPP dan TRX untuk menghitung Ticket Size (Ticket Size = NPP / TRX) per bulan.

Dashboard Visualisasi:

chart_visualisasi.py ➔ Membuat dashboard visual dengan filter bulan dan 3 tab grafik (Ticket Size, TRX, NPP) menggunakan Streamlit.

🛠️ Tools & Library yang Digunakan

Python 3.10+

pandas

openpyxl

SQLAlchemy

SQLite3

Streamlit
