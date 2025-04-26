import streamlit as st
import pandas as pd
from sqlalchemy import create_engine

# === SETUP CONNECTION DATABASE ===
db_path = "C:/Users/Stanford/Documents/test afpi/ticket_size.db"
engine = create_engine(f"sqlite:///{db_path}")

# === LOAD DATA TRX dan NPP ===
trx_path = "C:/Users/Stanford/Documents/test afpi/test Hasil TRX STATISTIK LPBBTI Desember 2024.xlsx"
npp_path = "C:/Users/Stanford/Documents/test afpi/test Hasil NPP STATISTIK LPBBTI Desember 2024.xlsx"

df_trx = pd.read_excel(trx_path, engine="openpyxl")
df_npp = pd.read_excel(npp_path, engine="openpyxl")

# Pastikan nama kolom
if "Tanggal" in df_trx.columns:
    df_trx = df_trx.rename(columns={"Tanggal": "Bulan", "Jumlah": "Trx"})
if "Tanggal" in df_npp.columns:
    df_npp = df_npp.rename(columns={"Tanggal": "Bulan", "Jumlah": "NPP"})

df_trx["Bulan"] = pd.to_datetime(df_trx["Bulan"])
df_npp["Bulan"] = pd.to_datetime(df_npp["Bulan"])

# === HITUNG TICKET SIZE ===
df_merge = pd.merge(df_npp, df_trx, on="Bulan", how="inner")
df_merge["ticket_size"] = df_merge["NPP"] / df_merge["Trx"]

# === STREAMLIT CONFIG ===
st.set_page_config(
    page_title="Dashboard Ticket Size LPBBTI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === DARK MODE TOGGLE ===
dark_mode = st.sidebar.toggle("🌙 Dark Mode", value=True)

if dark_mode:
    st.markdown(
        """
        <style>
            .main {
                background-color: #0E1117;
                color: white;
            }
        </style>
        """, unsafe_allow_html=True
    )

# === HEADER ===
st.title("📊 Dashboard Ticket Size LPBBTI")
st.markdown("---")

# === FILTER BULAN ===
st.sidebar.header("Filter Bulan")
bulan_min = df_merge["Bulan"].min().to_pydatetime()
bulan_max = df_merge["Bulan"].max().to_pydatetime()

selected_range = st.sidebar.slider(
    "Pilih rentang bulan:",
    min_value=bulan_min,
    max_value=bulan_max,
    value=(bulan_min, bulan_max),
    format="YYYY-MM"
)

# === FILTER DATA ===
mask = (df_merge["Bulan"] >= selected_range[0]) & (df_merge["Bulan"] <= selected_range[1])
filtered_data = df_merge.loc[mask]

# Pisahkan untuk masing-masing
filtered_ticket = filtered_data[["Bulan", "ticket_size"]]
filtered_trx = filtered_data[["Bulan", "Trx"]]
filtered_npp = filtered_data[["Bulan", "NPP"]]

# === VISUALISASI TABS ===
tab1, tab2, tab3 = st.tabs(["📈 Ticket Size", "💵 TRX", "💰 NPP"])

# === TAB 1: Ticket Size ===
with tab1:
    st.subheader("📈 Visualisasi Ticket Size")

    # --- Tampilkan Chart ---
    st.line_chart(
        filtered_ticket.set_index(filtered_ticket["Bulan"].dt.strftime('%Y-%m'))["ticket_size"]
    )

    # --- Tabel Ringkasan per Bulan ---
    st.markdown("### 📄 Tabel Data Ticket Size per Bulan")

    summary_ticket = (
        filtered_ticket
        .groupby(filtered_ticket["Bulan"].dt.strftime('%Y-%m'))
        .agg({"ticket_size": "mean"})  # Pakai mean kalau misal ada double data
        .reset_index()
    )

    summary_ticket = summary_ticket.rename(columns={"Bulan": "Bulan", "ticket_size": "Ticket Size"})
    summary_ticket["Ticket Size"] = summary_ticket["Ticket Size"].round(6)

    st.dataframe(summary_ticket, use_container_width=True)


# === TAB 2: TRX ===
with tab2:
    st.subheader("📈 Visualisasi Jumlah TRX")

    # Ambil data TRX detail
    display_trx = filtered_trx.copy()
    display_trx["Bulan"] = display_trx["Bulan"].dt.strftime('%Y-%m')

    # Chart Summary (total per bulan saja)
    trx_summary = (
        display_trx
        .groupby("Bulan")["Trx"]
        .sum()
        .reset_index()
    )

    st.line_chart(
        trx_summary.set_index("Bulan")["Trx"]
    )

    # Tabel TRX lengkap (Bulan + Provinsi + Trx)
    st.markdown("### 📄 Tabel Data TRX (Detail Provinsi)")
    st.dataframe(display_trx, use_container_width=True)

# === TAB 3: NPP ===
with tab3:
    st.subheader("📈 Visualisasi Jumlah NPP")

    # Ambil data NPP detail
    display_npp = filtered_npp.copy()
    display_npp["Bulan"] = display_npp["Bulan"].dt.strftime('%Y-%m')

    # Chart Summary (total per bulan saja)
    npp_summary = (
        display_npp
        .groupby("Bulan")["NPP"]
        .sum()
        .reset_index()
    )

    st.line_chart(
        npp_summary.set_index("Bulan")["NPP"]
    )

    # Tabel NPP lengkap (Bulan + Provinsi + NPP)
    st.markdown("### 📄 Tabel Data NPP (Detail Provinsi)")
    st.dataframe(display_npp, use_container_width=True)

