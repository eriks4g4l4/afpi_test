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
df_merge = pd.merge(df_npp, df_trx, on=["Bulan", "Provinsi"], how="inner")
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

# === FILTERS ===
st.sidebar.header("Filter Data")

# Filter Bulan
bulan_min = df_merge["Bulan"].min().to_pydatetime()
bulan_max = df_merge["Bulan"].max().to_pydatetime()

selected_range = st.sidebar.slider(
    "Pilih rentang bulan:",
    min_value=bulan_min,
    max_value=bulan_max,
    value=(bulan_min, bulan_max),
    format="YYYY-MM"
)

# Filter Provinsi
provinsi_options = ["Semua"] + sorted(df_merge["Provinsi"].unique().tolist())
selected_provinsi = st.sidebar.selectbox("Pilih Provinsi:", provinsi_options)

# === FILTER DATA ===
mask_bulan = (df_merge["Bulan"] >= selected_range[0]) & (df_merge["Bulan"] <= selected_range[1])

if selected_provinsi != "Semua":
    mask_provinsi = df_merge["Provinsi"] == selected_provinsi
else:
    mask_provinsi = pd.Series([True] * len(df_merge))

filtered_data = df_merge[mask_bulan & mask_provinsi]

# === TABS ===
tab1, tab2, tab3 = st.tabs(["📈 Ticket Size", "💵 TRX", "💰 NPP"])

# === TAB 1: Ticket Size ===
with tab1:
    st.subheader("📈 Visualisasi Ticket Size")

    # Pastikan kolom Bulan dalam datetime format
    filtered_data["Bulan"] = pd.to_datetime(filtered_data["Bulan"], errors="coerce")

    # Sum semua NPP dan TRX per bulan
    ticket_summary = (
        filtered_data
        .groupby(filtered_data["Bulan"].dt.to_period('M'))
        .agg({"NPP": "sum", "Trx": "sum"})
        .reset_index()
    )

    # Kembalikan kolom Bulan jadi format string YYYY-MM
    ticket_summary["Bulan"] = ticket_summary["Bulan"].dt.strftime('%Y-%m')

    # Hitung Ticket Size
    ticket_summary["Ticket_Size"] = (ticket_summary["NPP"] / ticket_summary["Trx"]).round(6)

    # === Visualisasi Line Chart ===
    if not ticket_summary.empty:
        st.line_chart(
            ticket_summary.set_index("Bulan")["Ticket_Size"]
        )
    else:
        st.warning("Data tidak tersedia untuk filter yang dipilih.")

# Tampilkan Tabel Summary dengan 6 angka decimal
st.markdown("### 📄 Tabel Data Ticket Size per Bulan")
st.dataframe(
    ticket_summary[["Bulan", "Ticket_Size"]],
    use_container_width=True,
    column_config={
        "Ticket_Size": st.column_config.NumberColumn(format="%.6f")
    }
)

# === TAB 2: TRX ===
with tab2:
    st.subheader("📈 Visualisasi Jumlah TRX")

    display_trx = filtered_data[["Bulan", "Provinsi", "Trx"]].copy()
    display_trx["Bulan"] = display_trx["Bulan"].dt.strftime('%Y-%m')

    summary_trx = (
        display_trx
        .groupby(["Bulan", "Provinsi"])["Trx"]
        .sum()
        .reset_index()
    )

    if not summary_trx.empty:
        st.line_chart(
            summary_trx.pivot(index="Bulan", columns="Provinsi", values="Trx")
        )
    else:
        st.warning("Data tidak tersedia untuk filter yang dipilih.")

    st.markdown("### 📄 Tabel Data TRX per Bulan dan Provinsi")
    st.dataframe(display_trx, use_container_width=True)

# === TAB 3: NPP ===
with tab3:
    st.subheader("📈 Visualisasi Jumlah NPP")

    display_npp = filtered_data[["Bulan", "Provinsi", "NPP"]].copy()
    display_npp["Bulan"] = display_npp["Bulan"].dt.strftime('%Y-%m')

    summary_npp = (
        display_npp
        .groupby(["Bulan", "Provinsi"])["NPP"]
        .sum()
        .reset_index()
    )

    if not summary_npp.empty:
        st.line_chart(
            summary_npp.pivot(index="Bulan", columns="Provinsi", values="NPP")
        )
    else:
        st.warning("Data tidak tersedia untuk filter yang dipilih.")

    st.markdown("### 📄 Tabel Data NPP per Bulan dan Provinsi")
    st.dataframe(display_npp, use_container_width=True)
