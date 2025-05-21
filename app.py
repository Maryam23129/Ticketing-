import streamlit as st
import pandas as pd
from io import BytesIO

def load_excel(file):
    return pd.read_excel(file)

def extract_total_invoice(invoice_df):
    filtered = invoice_df[invoice_df['STATUS'].str.lower() == 'dibayar']
    return filtered['HARGA'].sum()

def extract_total_b2b(df):
    row = df[df.apply(lambda r: r.astype(str).str.contains("TOTAL JUMLAH \(B2B\)", regex=True).any(), axis=1)]
    if not row.empty:
        jumlah_tiket = pd.to_numeric(row.iloc[0, 3], errors='coerce')
        pendapatan = pd.to_numeric(row.iloc[0, 4], errors='coerce')
        return jumlah_tiket, pendapatan
    return None, None

def rekonsiliasi(tiket_terjual, invoice, summary, rekening, jumlah_b2b=None, pendapatan_b2b=None):
    result = pd.merge(tiket_terjual, invoice, on='Nomor Invoice', how='outer', suffixes=('_tiket', '_invoice'))
    result = pd.merge(result, summary, on='Nomor Invoice', how='outer')
    result = pd.merge(result, rekening, left_on='Nomor Invoice', right_on='Deskripsi', how='outer')

    if jumlah_b2b is not None:
        result['Validasi Jumlah Tiket'] = result['Jumlah_tiket'] == jumlah_b2b
    if pendapatan_b2b is not None:
        result['Validasi Pendapatan'] = result['Jumlah_invoice'] == pendapatan_b2b

    result['Status Rekonsiliasi'] = result.apply(
        lambda row: 'Cocok' if row.get('Validasi Jumlah Tiket', True) and row.get('Validasi Pendapatan', True) and row['Jumlah_tiket'] == row['Jumlah_invoice'] == row['Debit'] else 'Tidak Cocok', axis=1
    )

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

uploaded_tiket = st.sidebar.file_uploader("📁 Upload Tiket Terjual", type=["xlsx"])
uploaded_invoice = st.sidebar.file_uploader("📁 Upload Invoice", type=["xlsx"])
uploaded_summary = st.sidebar.file_uploader("📁 Upload Ticket Summary", type=["xlsx"])
uploaded_rekening = st.sidebar.file_uploader("📁 Upload Rekening Koran", type=["xlsx"])

if uploaded_tiket and uploaded_invoice and uploaded_summary and uploaded_rekening:
    st.success("Semua file berhasil diupload. Memproses rekonsiliasi...")
    tiket_df = load_excel(uploaded_tiket)
    jumlah_tiket_b2b, pendapatan_b2b = extract_total_b2b(tiket_df)
    st.write(f"📈 Jumlah Tiket B2B: {jumlah_tiket_b2b}")
    st.write(f"💵 Pendapatan B2B: Rp {pendapatan_b2b:,.0f}")
    invoice_df = load_excel(uploaded_invoice)
    total_invoice_dibayar = extract_total_invoice(invoi
    summary_df = load_excel(uploaded_summary)
    rekening_df = load_excel(uploaded_rekening)

    hasil_rekonsiliasi = rekonsiliasi(tiket_df, invoice_df, summary_df, rekening_df, jumlah_tiket_b2b, pendapatan_b2b)

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
