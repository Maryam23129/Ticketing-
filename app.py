import streamlit as st
import pandas as pd
from io import BytesIO
import re

def load_excel(file):
    return pd.read_excel(file)

def extract_total_summary(summary_df):
    summary_df["CETAK BOARDING PASS"] = pd.to_datetime(summary_df["CETAK BOARDING PASS"], errors='coerce')
    summary_df = summary_df[~summary_df["CETAK BOARDING PASS"].isna()]
    summary_df["TARIF"] = pd.to_numeric(summary_df["TARIF"], errors='coerce')
    return summary_df["TARIF"].sum()

def extract_total_invoice(invoice_df):
    filtered = invoice_df[invoice_df['STATUS'].str.lower() == 'dibayar']
    return filtered['HARGA'].sum()

def extract_total_rekening(rekening_df):
    rekening_df = rekening_df.iloc[12:, [1, 2, 5]].dropna()
    rekening_df.columns = ['Tanggal', 'Remark', 'Credit']
    rekening_df = rekening_df[rekening_df['Remark'].str.contains("DARI MIDI UTAMA INDONESIA", case=False, na=False)]
    rekening_df['Credit'] = rekening_df['Credit'].replace('[^0-9.]', '', regex=True).astype(float)
    rekening_df['TanggalKode'] = rekening_df['Remark'].str.extract(r'^(\S+)')[0].str[-4:]
    rekening_df['Bulan'] = rekening_df['TanggalKode'].str[:2]
    rekening_df['Tanggal'] = rekening_df['TanggalKode'].str[2:]
    rekening_df['Tanggal Transaksi'] = pd.to_datetime('2025' + rekening_df['Bulan'] + rekening_df['Tanggal'], format='%Y%m%d', errors='coerce')
    return rekening_df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
        workbook = writer.book
        worksheet = writer.sheets['Rekapitulasi']
        currency_format = workbook.add_format({'num_format': '"Rp" #,##0'})
        for col_num, column in enumerate(df.columns):
            fmt = currency_format if column in ['Nominal Tiket Terjual', 'Invoice', 'Uang Masuk', 'Selisih'] else None
            worksheet.set_column(col_num, col_num, 20, fmt)
        border_format = workbook.add_format({'border': 1})
        worksheet.conditional_format(0, 0, len(df), len(df.columns) - 1, {'type': 'no_blanks', 'format': border_format})
        worksheet.conditional_format(0, 0, len(df), len(df.columns) - 1, {'type': 'blanks', 'format': border_format})
    output.seek(0)
    return output

st.set_page_config(page_title="Dashboard Rekonsiliasi Pendapatan Ticketing", layout="wide")

st.markdown("""
<h1 style='text-align: center;'>📊 Dashboard Rekonsiliasi Pendapatan Ticketing 🚢💰</h1>
<p style='text-align: center; font-size: 18px;'>Aplikasi ini digunakan untuk membandingkan data tiket terjual, invoice, ringkasan tiket, dan pemasukan dari rekening koran guna memastikan kesesuaian pendapatan.</p>
""", unsafe_allow_html=True)

st.sidebar.title("Upload File")
if st.sidebar.button("🔄 Reset File Upload"):
    for key in ["main_upload", "extra_upload"]:
        if key in st.session_state:
            del st.session_state[key]
    st.experimental_rerun()

uploaded_files = st.sidebar.file_uploader("📁 Upload Semua File Sekaligus", type=["xlsx"], accept_multiple_files=True, key="main_upload")
if st.sidebar.button("➕ Tambah File Lagi"):
    st.sidebar.file_uploader("📁 Upload Tambahan", type=["xlsx"], accept_multiple_files=True, key="extra_upload")

uploaded_tiket_files = []
uploaded_invoice = uploaded_summary = uploaded_rekening = None
all_files = uploaded_files + st.session_state.get("extra_upload", []) if uploaded_files else st.session_state.get("extra_upload", [])

if all_files:
    st.write("📂 Semua file terdeteksi:", [file.name for file in all_files])

    for file in all_files:
        fname = file.name.lower()
        if "tiket" in fname:
            uploaded_tiket_files.append(file)
        elif "invoice" in fname:
            uploaded_invoice = file
        elif "summary" in fname:
            uploaded_summary = file
        elif "rekening" in fname or "acc_statement" in fname:
            uploaded_rekening = file

    st.write("✅ Tiket terdeteksi:", [file.name for file in uploaded_tiket_files])

if uploaded_tiket_files and uploaded_invoice and uploaded_summary and uploaded_rekening:
    st.success("Semua file berhasil diupload. Memproses rekapitulasi...")

    b2b_list = []
    for tiket_file in uploaded_tiket_files:
        df_tiket = load_excel(tiket_file)
        row = df_tiket[df_tiket.apply(lambda r: r.astype(str).str.contains("TOTAL JUMLAH \\(B2B\\)", regex=True).any(), axis=1)]
        if not row.empty:
            pendapatan = pd.to_numeric(row.iloc[0, 4], errors='coerce')
            pelabuhan = next((p.capitalize() for p in ["merak", "bakauheni", "ketapang", "gilimanuk", "ciwandan", "panjang"] if p in tiket_file.name.lower()), "Tidak diketahui")
            b2b_list.append({"Pelabuhan": pelabuhan, "Pendapatan": pendapatan})

    invoice_df = load_excel(uploaded_invoice)
    total_invoice_dibayar = extract_total_invoice(invoice_df)

    match = re.search(r'(\d{4}-\d{2}-\d{2})\s*s[\-_]d\s*(\d{4}-\d{2}-\d{2})', uploaded_invoice.name)
    if match:
        tanggal_awal_str, tanggal_akhir_str = match.groups()
        tanggal_transaksi = f"{pd.to_datetime(tanggal_awal_str).strftime('%d-%m-%Y')} s.d {pd.to_datetime(tanggal_akhir_str).strftime('%d-%m-%Y')}"
    else:
        tanggal_transaksi = "Tanggal tidak tersedia"

    

    summary_df = load_excel(uploaded_summary)
    _ = extract_total_summary(summary_df)

    rekening_df = load_excel(uploaded_rekening)
    rekening_detail_df = extract_total_rekening(rekening_df)
    total_rekening_midi = rekening_detail_df['Credit'].sum()

    pelabuhan_list = ["Merak", "Bakauheni", "Ketapang", "Gilimanuk", "Ciwandan", "Panjang"]
    df = pd.DataFrame({
        "No": list(range(1, len(pelabuhan_list) + 1)),
        "Tanggal Transaksi": [tanggal_transaksi] * len(pelabuhan_list),
        "Pelabuhan Asal": pelabuhan_list,
        "Nominal Tiket Terjual": [
            next((b['Pendapatan'] for b in b2b_list if b['Pelabuhan'].lower() == pel.lower()), 0)
            for pel in pelabuhan_list
        ],
        "Invoice": [total_invoice_dibayar] + [0] * (len(pelabuhan_list) - 1),
        "Uang Masuk": [total_rekening_midi] + [0] * (len(pelabuhan_list) - 1),
        "Selisih": [total_invoice_dibayar - total_rekening_midi] + [0] * (len(pelabuhan_list) - 1)
    })

    total_row = {
        "No": "",
        "Tanggal Transaksi": "",
        "Pelabuhan Asal": "TOTAL",
        "Nominal Tiket Terjual": df["Nominal Tiket Terjual"].sum(),
        "Invoice": df["Invoice"].sum(),
        "Uang Masuk": df["Uang Masuk"].sum(),
        "Selisih": df["Selisih"].sum()
    }
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    formatted_df = df.copy()
    for col in ["Nominal Tiket Terjual", "Invoice", "Uang Masuk", "Selisih"]:
        formatted_df[col] = formatted_df[col].apply(lambda x: f"Rp {x:,.0f}" if isinstance(x, (int, float)) and x != 0 else "")

    st.subheader("📄 Tabel Rekapitulasi Hasil Rekonsiliasi")
    st.dataframe(formatted_df, use_container_width=True)

    output_excel = to_excel(df)
    st.download_button(
        label="📥 Download Rekapitulasi",
        data=output_excel,
        file_name="rekapitulasi_rekonsiliasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Silakan upload semua file yang dibutuhkan untuk menampilkan tabel hasil rekonsiliasi.")
