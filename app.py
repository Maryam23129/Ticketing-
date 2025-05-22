import streamlit as st
import pandas as pd
from io import BytesIO
import re

def load_excel(file):
    return pd.read_excel(file)

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

def to_excel(df_pelabuhan, df_total):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_pelabuhan.to_excel(writer, index=False, sheet_name='Per Pelabuhan')
        df_total.to_excel(writer, index=False, sheet_name='Total Keseluruhan')
        
        for sheet_name in ['Per Pelabuhan', 'Total Keseluruhan']:
            worksheet = writer.sheets[sheet_name]
            workbook = writer.book
            currency_format = workbook.add_format({'num_format': '"Rp" #,##0'})
            border_format = workbook.add_format({'border': 1})
            bold = workbook.add_format({'bold': True})

            df_used = df_pelabuhan if sheet_name == 'Per Pelabuhan' else df_total
            for col_num, column in enumerate(df_used.columns):
                fmt = currency_format if column in ['Nominal Tiket Terjual', 'Invoice', 'Uang Masuk', 'Selisih'] else None
                worksheet.set_column(col_num, col_num, 22, fmt)
            worksheet.conditional_format(0, 0, len(df_used), len(df_used.columns) - 1, {'type': 'no_blanks', 'format': border_format})
            worksheet.conditional_format(0, 0, len(df_used), len(df_used.columns) - 1, {'type': 'blanks', 'format': border_format})
            worksheet.write(f'A{len(df_used)}', 'TOTAL', bold)

    output.seek(0)
    return output

st.set_page_config(page_title="Dashboard Rekonsiliasi Pendapatan Ticketing", layout="wide")

st.markdown("""
<h1 style='text-align: center;'>📊 Dashboard Rekonsiliasi Pendapatan Ticketing 🚢💰</h1>
<p style='text-align: center; font-size: 18px;'>Aplikasi ini digunakan untuk membandingkan data tiket terjual, invoice, dan pemasukan dari rekening koran guna memastikan kesesuaian pendapatan.</p>
""", unsafe_allow_html=True)

st.sidebar.title("Upload File")
uploaded_files = st.sidebar.file_uploader("📁 Upload Semua File Sekaligus", type=["xlsx"], accept_multiple_files=True)

uploaded_tiket_files = []
uploaded_invoice = uploaded_summary = uploaded_rekening = None

if uploaded_files:
    for file in uploaded_files:
        fname = file.name.lower()
        if "tiket" in fname:
            uploaded_tiket_files.append(file)
        elif "invoice" in fname:
            uploaded_invoice = file
        elif "summary" in fname:
            uploaded_summary = file
        elif "rekening" in fname or "acc_statement" in fname:
            uploaded_rekening = file

if uploaded_tiket_files and uploaded_invoice and uploaded_summary and uploaded_rekening:
    invoice_df = load_excel(uploaded_invoice)
    invoice_df['HARGA'] = pd.to_numeric(invoice_df['HARGA'], errors='coerce')
    filtered_invoice = invoice_df[invoice_df['STATUS'].str.lower() == 'dibayar']

    df_total = filtered_invoice[['TANGGAL INVOICE', 'HARGA']].copy()
    df_total = df_total.rename(columns={'TANGGAL INVOICE': 'Tanggal Transaksi', 'HARGA': 'Invoice'})
    df_total['Tanggal Transaksi'] = pd.to_datetime(df_total['Tanggal Transaksi'], errors='coerce')
    df_total = df_total.sort_values('Tanggal Transaksi')
    df_total['Tanggal Transaksi'] = df_total['Tanggal Transaksi'].dt.strftime('%d-%m-%y')
    df_total['Uang Masuk'] = ''
    df_total['Selisih'] = ''

    rekening_df = load_excel(uploaded_rekening)
    rekening_detail_df = extract_total_rekening(rekening_df)
    rekening_detail_df['Tanggal Transaksi'] = pd.to_datetime(rekening_detail_df['Tanggal Transaksi'], errors='coerce').dt.strftime('%d-%m-%y')

    df_total['Uang Masuk'] = df_total['Tanggal Transaksi'].map(
        lambda tgl: rekening_detail_df[rekening_detail_df['Tanggal Transaksi'] == tgl]['Credit'].sum()
    )
    df_total['Selisih'] = df_total['Invoice'] - df_total['Uang Masuk']

    df_total_total = pd.DataFrame({
        'Tanggal Transaksi': ['TOTAL'],
        'Invoice': [df_total['Invoice'].sum()],
        'Uang Masuk': [df_total['Uang Masuk'].sum()],
        'Selisih': [df_total['Selisih'].sum()]
    })
    df_total = pd.concat([df_total, df_total_total], ignore_index=True)

    pelabuhan_list = ["Merak", "Bakauheni", "Ketapang", "Gilimanuk", "Ciwandan", "Panjang"]
    df_pelabuhan = pd.DataFrame({
        "No": list(range(1, len(pelabuhan_list)+1)),
        "Tanggal Transaksi": ['-'] * len(pelabuhan_list),
        "Pelabuhan Asal": pelabuhan_list,
        "Nominal Tiket Terjual": [0] * len(pelabuhan_list),
        "Naik Turun Golongan": [''] * len(pelabuhan_list)
    })
    total_row = {
        "No": "",
        "Tanggal Transaksi": "",
        "Pelabuhan Asal": "TOTAL",
        "Nominal Tiket Terjual": df_pelabuhan["Nominal Tiket Terjual"].sum(),
        "Naik Turun Golongan": ""
    }
    df_pelabuhan = pd.concat([df_pelabuhan, pd.DataFrame([total_row])], ignore_index=True)

    st.success("✅ Rekonsiliasi selesai!")

    st.subheader("📄 Tabel Rekapitulasi Rekonsiliasi Per Pelabuhan")
    st.dataframe(df_pelabuhan, use_container_width=True)

    st.subheader("📄 Tabel Rekapitulasi Total Keseluruhan")
    st.dataframe(df_total, use_container_width=True)

    output_excel = to_excel(df_pelabuhan, df_total)
    st.download_button(
        label="📥 Download Rekapitulasi Excel",
        data=output_excel,
        file_name="rekapitulasi_rekonsiliasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Silakan upload semua file yang dibutuhkan.")
