import streamlit as st
import pandas as pd
from io import BytesIO

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

def extract_total_b2b(df):
    row = df[df.apply(lambda r: r.astype(str).str.contains("TOTAL JUMLAH \(B2B\)", regex=True).any(), axis=1)]
    if not row.empty:
        jumlah_tiket = pd.to_numeric(row.iloc[0, 3], errors='coerce')
        pendapatan = pd.to_numeric(row.iloc[0, 4], errors='coerce')
        return jumlah_tiket, pendapatan
    return None, None

def extract_total_rekening(rekening_df):
    rekening_df = rekening_df.iloc[12:, [1, 2, 5]].dropna()
    rekening_df.columns = ['Tanggal', 'Remark', 'Credit']
    rekening_df = rekening_df[rekening_df['Remark'].str.contains("DARI MIDI UTAMA INDONESIA", case=False, na=False)]
    rekening_df['Credit'] = rekening_df['Credit'].replace('[^0-9.]', '', regex=True).astype(float)
    rekening_df['TanggalKode'] = rekening_df['Remark'].str.extract(r'^(\S+)')[0].str[-4:]
    rekening_df['Bulan'] = rekening_df['TanggalKode'].str[:2]
    rekening_df['Tanggal'] = rekening_df['TanggalKode'].str[2:]
    rekening_df['Tanggal Transaksi'] = pd.to_datetime('2025' + rekening_df['Bulan'] + rekening_df['Tanggal'], format='%Y%m%d', errors='coerce')
    return rekening_df, rekening_df['Credit'].sum()

def rekonsiliasi(tiket_terjual, invoice, summary, rekening, jumlah_b2b=None, pendapatan_b2b=None, total_invoice_dibayar=None, total_rekening_midi=None):
    result = pd.merge(tiket_terjual, invoice, on='Nomor Invoice', how='outer', suffixes=('_tiket', '_invoice'))
    result = pd.merge(result, summary, on='Nomor Invoice', how='outer')
    result = pd.merge(result, rekening, left_on='Nomor Invoice', right_on='Deskripsi', how='outer')

    if jumlah_b2b is not None:
        result['Validasi Jumlah Tiket'] = result['Jumlah_tiket'] == jumlah_b2b
    if pendapatan_b2b is not None:
        result['Validasi Pendapatan'] = result['Jumlah_invoice'] == pendapatan_b2b
    if total_invoice_dibayar is not None:
        result['Validasi Invoice Dibayar'] = result['Jumlah_invoice'] == total_invoice_dibayar
    if total_rekening_midi is not None:
        result['Validasi Rekening'] = result['Jumlah_invoice'] == total_rekening_midi

    result['Status Rekonsiliasi'] = result.apply(lambda row: 'Cocok' if all([
        row.get('Validasi Jumlah Tiket', True),
        row.get('Validasi Pendapatan', True),
        row.get('Validasi Invoice Dibayar', True),
        row.get('Validasi Rekening', True),
        row['Jumlah_tiket'] == row['Jumlah_invoice'] == row['Debit']
    ]) else 'Tidak Cocok', axis=1)

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
    total_invoice_dibayar = extract_total_invoice(invoice_df)
    st.write(f"🧾 Total Invoice Dibayar: Rp {total_invoice_dibayar:,.0f}")
    summary_df = load_excel(uploaded_summary)
    total_summary_tarif = extract_total_summary(summary_df)
    st.write(f"🧾 Total Tarif dari Ticket Summary: Rp {total_summary_tarif:,.0f}")
    rekening_df = load_excel(uploaded_rekening)
    rekening_detail_df, total_rekening_midi = extract_total_rekening(rekening_df)
    st.write(f"🏦 Total Kredit dari MIDI UTAMA INDONESIA: Rp {total_rekening_midi:,.0f}")

    hasil_rekonsiliasi = rekonsiliasi(tiket_df, invoice_df, summary_df, rekening_df, jumlah_tiket_b2b, pendapatan_b2b, total_invoice_dibayar, total_rekening_midi)

    # Buat tabel final sesuai template
    tanggal_transaksi = tiket_df[tiket_df.apply(lambda r: r.astype(str).str.contains("TOTAL JUMLAH \\(B2B\\)", regex=True).any(), axis=1)].iloc[0, 4]
    pelabuhan_list = ["Merak", "Bakauheni", "Ketapang", "Gilimanuk", "Ciwandan", "Panjang"]

    tabel_template = pd.DataFrame({
        "No": list(range(1, len(pelabuhan_list) + 1)),
        "Tanggal Transaksi": [tanggal_transaksi] * len(pelabuhan_list),
        "Pelabuhan Asal": pelabuhan_list,
        "Nominal Tiket Terjual": [pendapatan_b2b] + [0] * (len(pelabuhan_list) - 1),
        "Invoice": [total_invoice_dibayar] + [0] * (len(pelabuhan_list) - 1),
        "Uang Masuk": [total_rekening_midi] + [0] * (len(pelabuhan_list) - 1),
        "Selisih": [total_invoice_dibayar - total_rekening_midi] + [0] * (len(pelabuhan_list) - 1)
    })

    st.subheader("Rekapitulasi Hasil Rekonsiliasi")
    st.dataframe(tabel_template)

    # Export ke Excel
    def to_excel_summary(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
        output.seek(0)
        return output

    output_summary_excel = to_excel_summary(tabel_template)
    st.download_button(
        label="Download Tabel Rekapitulasi",
        data=output_summary_excel,
        file_name="rekapitulasi_rekonsiliasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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
