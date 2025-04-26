import pandas as pd
from sqlalchemy import create_engine, text

# === SETUP DATABASE SQLITE ===
db_path = "C:/Users/Stanford/Documents/test afpi/ticket_size.db"
engine = create_engine(f"sqlite:///{db_path}")

# === LOAD FILE NPP & TRX ===
npp_path = "C:/Users/Stanford/Documents/test afpi/test Hasil NPP STATISTIK LPBBTI Desember 2024.xlsx"
trx_path = "C:/Users/Stanford/Documents/test afpi/test Hasil TRX STATISTIK LPBBTI Desember 2024.xlsx"

# Karena file sudah rapi, pakai header=0
df_npp = pd.read_excel(npp_path, engine="openpyxl", header=0)
df_trx = pd.read_excel(trx_path, engine="openpyxl", header=0)

print("\nPreview Kolom df_npp:", df_npp.columns)
print("\nPreview Kolom df_trx:", df_trx.columns)

# Pastikan kolom jumlah numerik
df_npp["NPP"] = pd.to_numeric(df_npp["NPP"], errors="coerce").fillna(0)
df_trx["Trx"] = pd.to_numeric(df_trx["Trx"], errors="coerce").fillna(0)

# Rename kolom tanggal jadi seragam
df_npp = df_npp.rename(columns={"Tanggal": "Bulan"})
df_trx = df_trx.rename(columns={"Tanggal": "Bulan"})

# === SIMPAN DATA KE DATABASE ===
df_npp.to_sql("npp", engine, if_exists="replace", index=False)
df_trx.to_sql("trx", engine, if_exists="replace", index=False)
print("✅ Data NPP dan TRX berhasil disimpan ke SQLite.")

# === HITUNG TICKET SIZE ===
npp_total = df_npp.groupby("Bulan")["NPP"].sum()
trx_total = df_trx.groupby("Bulan")["Trx"].sum()

ticket_size = (npp_total / trx_total).round(6).reset_index()
ticket_size = ticket_size.rename(columns={0: "ticket_size"})

# Tambahkan id
ticket_size.insert(0, "id", range(1, len(ticket_size) + 1))

# === HAPUS VIEW / TABLE ticket_size KALAU ADA ===
with engine.connect() as conn:
    result = conn.execute(text("""
        SELECT type FROM sqlite_master 
        WHERE name='ticket_size'
    """)).fetchone()

    if result:
        if result[0] == 'view':
            conn.execute(text("DROP VIEW IF EXISTS ticket_size"))
            print("🗑 View 'ticket_size' dihapus.")
        elif result[0] == 'table':
            conn.execute(text("DROP TABLE IF EXISTS ticket_size"))
            print("🗑 Table 'ticket_size' dihapus.")

# Simpan hasil ke database
ticket_size.to_sql("ticket_size", engine, if_exists="replace", index=False)
print("✅ Tabel 'ticket_size' berhasil dibuat.")

# === EXPORT KE EXCEL ===
output_excel = "C:/Users/Stanford/Documents/test afpi/ticket_size.xlsx"
ticket_size.to_excel(output_excel, index=False, engine="openpyxl")
print(f"📄 Hasil ticket_size sudah diekspor ke: {output_excel}")

# === TAMPILKAN HASIL ===
print("\n=== Hasil Ticket Size ===")
print(ticket_size)
