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

def rekonsiliasi(tiket_terjual, invoice, summary, rekening, jumlah_b2b=None, pendapatan_b2b=None, total_invoice_dibayar=None):
    result = pd.merge(tiket_terjual, invoice, on='Nomor Invoice', how='outer', suffixes=('_tiket', '_invoice'))
    result = pd.merge(result, summary, on='Nomor Invoice', how='outer')
    result = pd.merge(result, rekening, left_on='Nomor Invoice', right_on='Deskripsi', how='outer')

    if jumlah_b2b is not None:
        result['Validasi Jumlah Tiket'] = result['Jumlah_tiket'] == jumlah_b2b
    if pendapatan_b2b is not None:
        result['Validasi Pendapatan'] = result['Jumlah_invoice'] == pendapatan_b2b

        if total_invoice_dibayar is not None:
        result['Validasi Invoice Dibayar'] = result['Jumlah_invoice'] == total_invoice_dibayar
