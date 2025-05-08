import pandas as pd
# import tabula  # untuk PDF
from tabula import read_pdf
import os

# Jika datanya berupa PDF
def extract_from_pdf(pdf_path):
    # Ekstrak semua tabel dari PDF
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    return tables

# Jika datanya berupa Excel - HAPUS TANDA # DI SEMUA BARIS INI
def process_excel_reports(excel_paths):
    all_data = []
    
    for path in excel_paths:
        # Baca file Excel
        df = pd.read_excel(path)
        
        # Identifikasi baris header
        # Karena format laporan biasanya memiliki header dan footer
        header_row = df[df.iloc[:, 0] == 'Kode'].index[0]
        
        # Baca ulang dengan header yang benar
        df = pd.read_excel(path, header=header_row)
        
        # Tambahkan metadata dari laporan
        # Misalnya, ambil informasi divisi, periode dari bagian atas laporan
        meta_info = pd.read_excel(path, nrows=5)
        
        # Ekstrak periode dan informasi divisi
        periode = meta_info.iloc[1, 1] if 'Periode' in str(meta_info.iloc[1, 0]) else None
        divisi = meta_info.iloc[3, 1] if 'Divisi' in str(meta_info.iloc[3, 0]) else None
        
        # Tambahkan sebagai kolom baru
        df['Periode'] = periode
        df['Divisi'] = divisi
        
        # Bersihkan data
        df = df[~df['Kode'].isna()]  # Hapus baris tanpa kode produk
        
        all_data.append(df)
    
    # Gabungkan semua dataframe
    combined_data = pd.concat(all_data, ignore_index=True)
    
    return combined_data

if __name__ == "__main__":
    # Definisikan excel_files di sini
    excel_files = ["data/SalesByProductBrand-ID.pdf"]  # Sesuaikan dengan nama file Anda
    
    if "pdf" in excel_files[0].lower():
        # Gunakan fungsi extract_from_pdf jika file PDF
        tables = extract_from_pdf(excel_files[0])
        # Proses hasil ekstraksi
        print(f"Jumlah tabel yang diekstrak: {len(tables)}")
        print("Sample data dari tabel pertama:")
        if tables:
            print(tables[0].head())
    else:
        # Gunakan fungsi process_excel_reports jika file Excel
        result_df = process_excel_reports(excel_files)
        
        # Simpan hasil ke CSV
        result_df.to_csv("data/processed_sales_data.csv", index=False)
        
        # Tampilkan informasi hasil konversi
        print(f"Data berhasil dikonversi dengan {len(result_df)} baris")
        print("Sample data:")
        print(result_df.head())