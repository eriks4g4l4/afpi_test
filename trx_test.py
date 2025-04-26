import pandas as pd
import sqlite3

# === SETUP PATH ===
file_path = "C:/Users/Stanford/Documents/test afpi/STATISTIK LPBBTI Desember 2024.xlsx"
output_excel = "C:/Users/Stanford/Documents/test afpi/test Hasil TRX STATISTIK LPBBTI Desember 2024.xlsx"
output_txt = "C:/Users/Stanford/Documents/test afpi/trx_Error_amount.txt"
db_path = "C:/Users/Stanford/Documents/test afpi/trx.db"  # path untuk database SQLite

# === LOAD DATA ===
df = pd.read_excel(file_path, sheet_name=20, header=None, engine="openpyxl")

# === EXTRACT bulan ===
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
    trx = df.iloc[3:43, col].drop(index=[3, 10]).reset_index(drop=True)
    df_temp = pd.DataFrame({
        "Bulan": bln,
        "Provinsi": provinsi,
        "Trx": trx
    })
    df_list.append(df_temp)

# === GABUNGKAN & SIMPAN HASIL ===
data_final = pd.concat(df_list, ignore_index=True)
data_final.to_excel(output_excel, index=False, engine="openpyxl")
print(f"✅ File berhasil disimpan ke: {output_excel}")

# === SIMPAN KE DATABASE SQLITE ===
conn = sqlite3.connect(db_path)
table_name = "TRX"

# Membuat tabel kalau belum ada
create_table_query = f"""
CREATE TABLE IF NOT EXISTS {table_name} (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    Bulan TEXT,
    Provinsi TEXT,
    Trx REAL
);
"""
conn.execute(create_table_query)
conn.commit()

# Insert data
data_final.to_sql(table_name, conn, if_exists="append", index=False)
conn.close()

print(f"✅ Data berhasil diinsert ke database: {db_path} (table: {table_name})")

# === VALIDASI PENURUNAN ===
data_final["Trx"] = pd.to_numeric(data_final["Trx"], errors="coerce").fillna(0)
total_per_bulan = data_final.groupby("Bulan")["Trx"].sum().sort_index()

errors = []
prev = None
for bulan, total in total_per_bulan.items():
    if prev:
        prev_bulan, prev_total = prev
        if total < prev_total:
            errors.append(
                f"⚠️ Trx bulan {bulan} ({total:.2f}) lebih kecil dari bulan sebelumnya {prev_bulan} ({prev_total:.2f})"
            )
    prev = (bulan, total)

# === SIMPAN ERROR (JIKA ADA) ===
if errors:
    with open(output_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(errors))
    print(f"❗ Ditemukan {len(errors)} penurunan. Cek file: {output_txt}")
else:
    print("✅ Tidak ada penurunan Trx yang mencurigakan.")
