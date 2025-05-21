import streamlit as st
import pandas as pd
from io import BytesIO

def load_excel(file):
    return pd.read_excel(file)

def rekonsiliasi(tiket_terjual, invoice, summary, rekening):
    # Contoh logika dasar rekonsiliasi: gabungkan dan bandingkan berdasarkan kolom 'Nomor Tiket' atau 'Nomor Invoice'
    result = pd.merge(tiket_terjual, invoice, on='Nomor Invoice', how='outer', suffixes=('_tiket', '_invoice'))
    result = pd.merge(result, summary, on='Nomor Invoice', how='outer')
    result = pd.merge(result, rekening, left_on='Nomor Invoice', right_on='Deskripsi', how='outer')
    result['Status Rekonsiliasi'] = result.apply(lambda row: 'Cocok' if row['Jumlah_tiket'] == row['Jumlah_invoice'] == row['Debit'] else 'Tidak Cocok', axis=1)
    return result

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Rekonsiliasi')
    output.seek(0)
    return output

st.set_page_config(page_title="Dashboard Rekonsiliasi Pendapatan Ticketing", layout="wide")

# Header dengan ikon
st.markdown("""
    <h1 style='text-align: center;'>📊 Dashboard Rekonsiliasi Pendapatan Ticketing 🚢💰</h1>
    <p style='text-align: center; font-size: 18px;'>Aplikasi ini digunakan untuk membandingkan data tiket terjual, invoice, ringkasan tiket, dan pemasukan dari rekening koran guna memastikan kesesuaian pendapatan.</p>
""", unsafe_allow_html=True)

st.sidebar.title("Upload File")

uploaded_tiket = st.sidebar.file_uploader("Upload Tiket Terjual", type=["xlsx"])
uploaded_invoice = st.sidebar.file_uploader("Upload Invoice", type=["xlsx"])
uploaded_summary = st.sidebar.file_uploader("Upload Ticket Summary", type=["xlsx"])
uploaded_rekening = st.sidebar.file_uploader("Upload Rekening Koran", type=["xlsx"])

if uploaded_tiket and uploaded_invoice and uploaded_summary and uploaded_rekening:
    st.success("Semua file berhasil diupload. Memproses rekonsiliasi...")
    tiket_df = load_excel(uploaded_tiket)
    invoice_df = load_excel(uploaded_invoice)
    summary_df = load_excel(uploaded_summary)
    rekening_df = load_excel(uploaded_rekening)

    hasil_rekonsiliasi = rekonsiliasi(tiket_df, invoice_df, summary_df, rekening_df)

    st.subheader("Hasil Rekonsiliasi")
    st.dataframe(hasil_rekonsiliasi)

    output_excel = to_excel(hasil_rekonsiliasi)
    st.download_button(
        label="Download Hasil Rekonsiliasi",
        data=output_excel,
        file_name="hasil_rekonsiliasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Silakan upload semua file yang dibutuhkan untuk melakukan rekonsiliasi.")
