import pandas as pd
import sqlite3

# === SETUP PATH ===
file_path = "C:/Users/Stanford/Documents/test afpi/STATISTIK LPBBTI Desember 2024.xlsx"
output_excel = "C:/Users/Stanford/Documents/test afpi/test Hasil NPP STATISTIK LPBBTI Desember 2024.xlsx"
output_txt = "C:/Users/Stanford/Documents/test afpi/NPP_Error_amount.txt"
db_path = "C:/Users/Stanford/Documents/test afpi/NPP.db"  # Path untuk database SQLite

# === LOAD DATA ===
df = pd.read_excel(file_path, sheet_name=22, header=None, engine="openpyxl")

# === EXTRACT BULAN ===
start_col = 2  # Mulai dari kolom C
bulan_list = []

for col in range(start_col, df.shape[1]):
    val = df.iloc[1, col]
    try:
        bulan = pd.to_datetime(val).strftime('%Y-%m')
        bulan_list.append((col, bulan))
    except:
        break  # Stop saat bukan bulan lagi

# === EXTRACT PROVINSI & BERSIHKAN ===
provinsi = df.iloc[3:43, 1].drop(index=[3, 10]).reset_index(drop=True)
provinsi = provinsi.str.replace(r"^\d+\.\s*", "", regex=True)

# === EXTRACT TRANSAKSI PER BULAN ===
df_list = []
for col, bln in bulan_list:
    npp = df.iloc[3:43, col].drop(index=[3, 10]).reset_index(drop=True)
    df_temp = pd.DataFrame({
        "Bulan": bln,
        "Provinsi": provinsi,
        "NPP": npp
    })
    df_list.append(df_temp)

# === GABUNGKAN HASIL ===
data_final = pd.concat(df_list, ignore_index=True)

# === KONVERSI TIPE DATA + BULATKAN 2 DESIMAL ===
data_final["NPP"] = pd.to_numeric(data_final["NPP"], errors="coerce").fillna(0).astype(float)
data_final["NPP"] = data_final["NPP"].round(2)  # <--- BULATKAN 2 DESIMAL

# === SIMPAN KE EXCEL ===
data_final.to_excel(output_excel, index=False, engine="openpyxl")
print(f"✅ File hasil disimpan ke: {output_excel}")

# === VALIDASI PENURUNAN ===
total_per_bulan = data_final.groupby("Bulan")["NPP"].sum().sort_index()

errors = []
prev = None
for bulan, total in total_per_bulan.items():
    if prev:
        prev_bulan, prev_total = prev
        if total < prev_total:
            errors.append(
                f"❌ Bulan {bulan}: NPP ({total:,.2f}) LEBIH KECIL dari bulan sebelumnya {prev_bulan}: ({prev_total:,.2f})"
            )
            print(f"❌ Bulan {bulan}: {total:,.2f} LEBIH KECIL dari bulan sebelumnya {prev_bulan}: {prev_total:,.2f}")
        elif total == prev_total:
            print(f"ℹ️ Bulan {bulan}: NPP ({total:,.2f}) SAMA dengan bulan sebelumnya {prev_bulan}: ({prev_total:,.2f})")
        else:
            print(f"✅ Bulan {bulan}: NPP ({total:,.2f}) NAIK dibanding bulan sebelumnya {prev_bulan}: ({prev_total:,.2f})")
    prev = (bulan, total)

# === SIMPAN ERROR (JIKA ADA) ===
if errors:
    with open(output_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(errors))
    print(f"\n❗ Ditemukan {len(errors)} penurunan. Cek file: {output_txt}")
else:
    print("\n✅ Semua NPP bulan valid (tidak ada penurunan).")

# === SIMPAN KE DATABASE SQLITE ===
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Buat tabel jika belum ada
cursor.execute("""
CREATE TABLE IF NOT EXISTS npp (
    Bulan TEXT,
    Provinsi TEXT,
    NPP REAL
)
""")

# Bersihkan data lama (opsional)
cursor.execute("DELETE FROM npp")

# Insert data
for _, row in data_final.iterrows():
    cursor.execute("""
    INSERT INTO npp (Bulan, Provinsi, NPP)
    VALUES (?, ?, ?)
    """, (row["Bulan"], row["Provinsi"], row["NPP"]))

# Commit dan tutup koneksi
conn.commit()
conn.close()

print(f"✅ Data berhasil dimasukkan ke database SQLite di: {db_path}")
