import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# Konfigurasi tampilan halaman web
st.set_page_config(page_title="Konverter BCA ke Excel", page_icon="🏦", layout="wide")

# --- GERBANG KEAMANAN (KATA SANDI) ---
def check_password():
    def password_entered():
        # Memeriksa kata sandi dengan data di dalam Streamlit Secrets
        if st.session_state["password"] == st.secrets["app_password"]: 
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 Akses Terbatas")
        st.text_input("Masukkan kata sandi untuk menggunakan aplikasi ini:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 Akses Terbatas")
        st.text_input("Masukkan kata sandi untuk menggunakan aplikasi ini:", type="password", on_change=password_entered, key="password")
        st.error("❌ Kata sandi salah. Silakan coba lagi.")
        return False
    return True

# --- APLIKASI UTAMA (Hanya Terbuka Jika Sandi Benar) ---
if check_password():
    st.title("🏦 BCA e-Statement to Excel Converter")
    st.write("Unggah file PDF e-Statement BCA secara satuan atau massal. Sistem akan mengekstrak transaksi, memisahkan bulan ke dalam sheet, dan mengkalkulasi saldo secara otomatis.")

    # Fungsi bantuan
    def is_date(text):
        return bool(re.match(r'^\d{2}/\d{2}$', text.strip()))

    def is_mutasi_or_saldo(text):
        return bool(re.match(r'^[0-9,]+\.[0-9]{2}$', text.strip()))

    def is_cbg(text):
        return bool(re.match(r'^\d{4}$', text.strip()))

    def rapikan_nama_entitas(teks):
        teks = re.sub(r'\b([a-zA-Z0-9\s]+?)\s+(PT|UD|CV|Tbk)\b', r'\2 \1', teks, flags=re.IGNORECASE)
        return re.sub(r'\s+', ' ', teks).strip()

    map_bulan = {
        'JANUARI': 'Jan', 'FEBRUARI': 'Feb', 'MARET': 'Mar', 'APRIL': 'Apr',
        'MEI': 'Mei', 'JUNI': 'Jun', 'JULI': 'Jul', 'AGUSTUS': 'Agt',
        'SEPTEMBER': 'Sep', 'OKTOBER': 'Okt', 'NOVEMBER': 'Nov', 'DESEMBER': 'Des'
    }

    uploaded_files = st.file_uploader("Pilih file PDF e-Statement BCA", type="pdf", accept_multiple_files=True)

    if uploaded_files:
        if st.button("🚀 Proses Data Sekarang"):
            with st.spinner("Sedang membaca PDF dan mengkalkulasi saldo secara presisi"):
                sheet_data = {}
                global_norek = "UNKNOWN"
                global_year = "YYYY"

                for file in uploaded_files:
                    pdf_norek = "UNKNOWN"
                    pdf_month = "MMM"
                    pdf_year = "YYYY"
                    
                    with pdfplumber.open(file) as pdf:
                        for page in pdf.pages:
                            text = page.extract_text()
                            if text:
                                if pdf_norek == "UNKNOWN":
                                    norek_match = re.search(r'REKENING\s*[:\n\s]*(\d{10,})', text)
                                    if norek_match:
                                        pdf_norek = norek_match.group(1)
                                
                                if pdf_year == "YYYY":
                                    periode_match = re.search(r'PERIODE\s*[:\n\s]*([A-Z]+)\s+(\d{4})', text)
                                    if periode_match:
                                        bulan_full = periode_match.group(1)
                                        pdf_year = periode_match.group(2)
                                        pdf_month = map_bulan.get(bulan_full, 'MMM')
                                        
                            if pdf_norek != "UNKNOWN" and pdf_year != "YYYY":
                                break
                    
                    if global_norek == "UNKNOWN" and pdf_norek != "UNKNOWN":
                        global_norek = pdf_norek
                    if global_year == "YYYY" and pdf_year != "YYYY":
                        global_year = pdf_year
                        
                    parsed_data = []
                    
                    with pdfplumber.open(file) as pdf:
                        for page in pdf.pages:
                            words = page.extract_words()
                            if not words:
                                continue
                            
                            lines = []
                            current_line = []
                            current_top = words[0]['top']
                            
                            for w in words:
                                if abs(w['top'] - current_top) <= 3:
                                    current_line.append(w)
                                else:
                                    current_line.sort(key=lambda x: x['x0'])
                                    lines.append(current_line)
                                    current_line = [w]
                                    current_top = w['top']
                            if current_line:
                                current_line.sort(key=lambda x: x['x0'])
                                lines.append(current_line)
                                
                            x_ket, x_cbg, x_mut, x_sal = 60, 280, 350, 450
                            is_table = False
                            
                            for line in lines:
                                line_text = " ".join([w['text'] for w in line])
                                
                                if 'TANGGAL' in line_text and 'KETERANGAN' in line_text and 'MUTASI' in line_text:
                                    is_table = True
                                    for w in line:
                                        if w['text'] == 'KETERANGAN': x_ket = w['x0'] - 5
                                        elif w['text'] == 'CBG': x_cbg = w['x0'] - 5
                                        elif w['text'] == 'MUTASI': x_mut = w['x0'] - 5
                                        elif w['text'] == 'SALDO': x_sal = w['x0'] - 5
                                    continue
                                
                                if re.search(r'(SALDO AKHIR|CATATAN:|BERSAMBUNG|MUTASI CR|SALDO AWAL\s*:)', line_text, re.IGNORECASE):
                                    is_table = False
                                    
                                if not is_table or not line:
                                    continue

                                first_word = line[0]['text']
                                
                                if is_date(first_word):
                                    tgl = first_word
                                    ket_words = []
                                    cbg = mut = sal = ""
                                    
                                    for w in line[1:]:
                                        text = w['text']
                                        x = w['x0']
                                        
                                        if x >= x_sal and is_mutasi_or_saldo(text):
                                            sal = text
                                        elif x >= x_mut and is_mutasi_or_saldo(text):
                                            mut = text
                                        elif x >= x_mut and text == 'DB':
                                            mut += ' DB'
                                        elif x >= x_cbg and x < x_mut and is_cbg(text):
                                            cbg = text
                                        else:
                                            ket_words.append(text)
                                            
                                    parsed_data.append({
                                        'Tanggal': tgl,
                                        'Keterangan': " ".join(ket_words),
                                        'Cabang': cbg,
                                        'Mutasi': mut,
                                        'Saldo_Asli': sal
                                    })
                                    
                                else:
                                    extra_ket = " ".join([w['text'] for w in line])
                                    if parsed_data:
                                        parsed_data[-1]['Keterangan'] += " " + extra_ket

                    final_data = []
                    for row in parsed_data:
                        mutasi_str = row['Mutasi'].replace(',', '')
                        saldo_str = row['Saldo_Asli'].replace(',', '')
                        
                        tipe = ""
                        nominal = mutasi_str
                        if "DB" in mutasi_str:
                            tipe = "DB"
                            nominal = mutasi_str.replace("DB", "").strip()
                        elif mutasi_str:
                            tipe = "CR"
                            nominal = mutasi_str.replace("CR", "").strip()
                            
                        if not nominal:
                            nominal = 0.0
                        else:
                            nominal = float(nominal)
                            
                        final_data.append({
                            'Tanggal': row['Tanggal'],
                            'Keterangan': rapikan_nama_entitas(row['Keterangan']),
                            'Cabang': row['Cabang'], 
                            'Nominal': nominal,
                            'Tipe': tipe,
                            'Saldo_Asli': saldo_str
                        })

                    df = pd.DataFrame(final_data)

                    if not df.empty:
                        df['Tanggal'] = df['Tanggal'].replace(r'^\s*$', pd.NA, regex=True).ffill()
                        df['Tanggal'] = df['Tanggal'].apply(lambda x: f"{str(x).strip()}/{pdf_year}" if pd.notna(x) and len(str(x).strip()) == 5 else x)
                        
                        # Bagian yang diperbaiki (errors='coerce')
                        df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
                        
                        saldo_terhitung = []
                        current_saldo = 0.0

                        for idx, row in df.iterrows():
                            if 'SALDO AWAL' in str(row['Keterangan']).upper():
                                try:
                                    current_saldo = float(row['Saldo_Asli'])
                                except ValueError:
                                    current_saldo = 0.0
                            else:
                                if row['Tipe'] == 'CR':
                                    current_saldo += row['Nominal']
                                elif row['Tipe'] == 'DB':
                                    current_saldo -= row['Nominal']
                                    
                            saldo_terhitung.append(current_saldo)

                        df['Saldo'] = saldo_terhitung
                        df['Nominal'] = df['Nominal'].apply(lambda x: x if x > 0 else None)
                        df = df.drop(columns=['Saldo_Asli'])
                        
                        base_sheet_name = pdf_month
                        sheet_name = base_sheet_name
                        counter = 1
                        while sheet_name in sheet_data:
                            sheet_name = f"{base_sheet_name} ({counter})"
                            counter += 1
                            
                        sheet_data[sheet_name] = df

                if sheet_data:
                    output_filename = f"Hasil Convert BCA - Bulk_{global_norek}_{global_year}.xlsx"
                    
                    output = io.BytesIO()
                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                    
                    month_order = list(map_bulan.values())
                    def sort_key(sheet_name):
                        base_name = sheet_name.split(' ')[0]
                        if base_name in month_order:
                            return month_order.index(base_name)
                        return 99

                    sorted_sheets = sorted(sheet_data.keys(), key=sort_key)
                    
                    workbook  = writer.book
                    format_comma = workbook.add_format({'num_format': '#,##0.00'})
                    format_text = workbook.add_format({'num_format': '@'})

                    for sheet_name in sorted_sheets:
                        df = sheet_data[sheet_name]
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                        worksheet = writer.sheets[sheet_name]

                        worksheet.set_column('A:A', 12)  
                        worksheet.set_column('B:B', 50)  
                        worksheet.set_column('C:C', 10, format_text)
                        worksheet.set_column('D:D', 18, format_comma)  
                        worksheet.set_column('E:E', 6)   
                        worksheet.set_column('F:F', 20, format_comma)  
                        
                    writer.close()
                    processed_data = output.getvalue()
                    
                    st.success(f"✅ Data berhasil diekstrak menjadi {len(sorted_sheets)} sheet bulanan, suksme.")
                    
                    st.download_button(
                        label="📥 Unduh File Excel Sekarang",
                        data=processed_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Gagal mendeteksi tabel BCA. Pastikan file yang diunggah benar.")
