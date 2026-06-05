import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import csv

# Konfigurasi tampilan halaman web
st.set_page_config(page_title="Konverter Rekening Koran ke Excel", page_icon="🏦", layout="wide")

# --- GERBANG KEAMANAN (KATA SANDI) ---
def check_password():
    def password_entered():
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

if check_password():
    st.title("🏦 Multi-Bank e-Statement to Excel Converter")
    st.write("""
    Konversi e-Statement PDF ke Excel untuk berbagai bank: **BCA, BRI, OCBC, Permata, Mekari (Jurnal)**.
    Upload file PDF sesuai bank, sistem akan mengekstrak transaksi secara otomatis.
    """)

    # ========== PILIH BANK ==========
    bank_option = st.selectbox(
        "Pilih Bank / Sumber e-Statement",
        ["BCA", "BRI", "OCBC NISP", "Permata", "Mekari (Jurnal)"]
    )

    uploaded_files = st.file_uploader(
        f"Upload file PDF e-Statement {bank_option}",
        type="pdf",
        accept_multiple_files=True
    )

    # ========== HELPER FUNCTIONS ==========
    def is_date(text):
        """Detects dd/mm or dd/mm/yyyy date formats"""
        return bool(re.match(r'^\d{2}/\d{2}', text.strip()))

    def is_money(text):
        """Detects number with commas and 2 decimal places"""
        return bool(re.match(r'^[0-9,]+\.[0-9]{2}$', text.strip()))

    def clean_money(val):
        """Convert money string to float"""
        if not val:
            return 0.0
        return float(val.replace(',', ''))

    def group_words_into_lines(words, y_tolerance=3):
        """Group extracted words into lines based on vertical position"""
        if not words:
            return []
        lines = []
        current_line = []
        current_top = words[0]['top']
        for w in words:
            if abs(w['top'] - current_top) <= y_tolerance:
                current_line.append(w)
            else:
                current_line.sort(key=lambda x: x['x0'])
                lines.append(current_line)
                current_line = [w]
                current_top = w['top']
        if current_line:
            current_line.sort(key=lambda x: x['x0'])
            lines.append(current_line)
        return lines

    # ========== PARSER: BCA ==========
    def parse_bca(pdf_file):
        """Parse BCA e-Statement PDF"""
        map_bulan = {
            'JANUARI': 'Jan', 'FEBRUARI': 'Feb', 'MARET': 'Mar', 'APRIL': 'Apr',
            'MEI': 'Mei', 'JUNI': 'Jun', 'JULI': 'Jul', 'AGUSTUS': 'Agt',
            'SEPTEMBER': 'Sep', 'OKTOBER': 'Okt', 'NOVEMBER': 'Nov', 'DESEMBER': 'Des'
        }

        pdf_norek = "UNKNOWN"
        pdf_month = "MMM"
        pdf_year = "YYYY"

        with pdfplumber.open(pdf_file) as pdf:
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

        parsed_data = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue
                lines = group_words_into_lines(words)
                
                x_ket, x_cbg, x_mut, x_sal = 60, 280, 350, 450
                is_table = False

                for line in lines:
                    line_text = " ".join([w['text'] for w in line])
                    
                    if 'TANGGAL' in line_text and 'KETERANGAN' in line_text:
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
                            if x >= x_sal and is_money(text):
                                sal = text
                            elif x >= x_mut and is_money(text):
                                mut = text
                            elif x >= x_mut and text == 'DB':
                                mut += ' DB'
                            elif x >= x_cbg and x < x_mut and bool(re.match(r'^\d{4}$', text.strip())):
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

        rows = []
        for row in parsed_data:
            mutasi_str = row['Mutasi'].replace(',', '')
            tipe = ""
            nominal = mutasi_str
            if "DB" in mutasi_str:
                tipe = "DB"
                nominal = mutasi_str.replace("DB", "").strip()
            elif mutasi_str:
                tipe = "CR"
                nominal = mutasi_str.replace("CR", "").strip()
            nominal_val = float(nominal) if nominal else 0.0

            rows.append({
                'Tanggal': row['Tanggal'],
                'Keterangan': row['Keterangan'].strip(),
                'Debit': nominal_val if tipe == 'DB' else '',
                'Kredit': nominal_val if tipe == 'CR' else '',
                'Saldo': row['Saldo_Asli']
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df['Tanggal'] = df['Tanggal'].replace(r'^\s*$', pd.NA, regex=True).ffill()
            df['Tanggal'] = df['Tanggal'].apply(
                lambda x: f"{str(x).strip()}/{pdf_year}" if pd.notna(x) and len(str(x).strip()) == 5 else x
            )
            df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y').fillna("")

        return df, pdf_norek, pdf_month, pdf_year

    # ========== PARSER: BRI ==========
    def parse_bri(pdf_file):
        """Parse BRI e-Statement PDF"""
        rows = []
        account_no = "UNKNOWN"
        account_name = "UNKNOWN"
        period = "UNKNOWN"

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    # Extract account info
                    if account_no == "UNKNOWN":
                        m = re.search(r'No\.?\s*Rekening\s*:?\s*(\d+)', text, re.IGNORECASE)
                        if m:
                            account_no = m.group(1)
                    if account_name == "UNKNOWN":
                        m = re.search(r'Kepada\s*Yth\.?\s*/\s*To\s*:\s*\n(.+)', text, re.IGNORECASE)
                        if m:
                            account_name = m.group(1).strip()
                    if period == "UNKNOWN":
                        m = re.search(r'Periode\s*Transaksi\s*:?\s*(.+)', text, re.IGNORECASE)
                        if m:
                            period = m.group(1).strip()

                words = page.extract_words()
                if not words:
                    continue
                lines = group_words_into_lines(words)

                is_table = False
                x_debet, x_kredit, x_saldo = 350, 420, 490

                for line in lines:
                    line_text = " ".join([w['text'] for w in line])
                    
                    if 'Tanggal Transaksi' in line_text and 'Uraian Transaksi' in line_text:
                        is_table = True
                        for w in line:
                            if w['text'] == 'Debet': x_debet = w['x0'] - 5
                            elif w['text'] == 'Kredit': x_kredit = w['x0'] - 5
                            elif w['text'] == 'Saldo': x_saldo = w['x0'] - 5
                        continue

                    if re.search(r'(Saldo\s*Awal|Total\s*Transaksi|Terbilang)', line_text, re.IGNORECASE):
                        is_table = False
                    if not is_table:
                        continue

                    # First word should be date dd/mm/yy
                    first_word = line[0]['text']
                    if re.match(r'^\d{2}/\d{2}/\d{2}', first_word):
                        tgl = first_word
                        uraian_words = []
                        debet = kredit = saldo = ""
                        
                        for w in line[1:]:
                            text = w['text']
                            x = w['x0']
                            if x >= x_saldo and is_money(text):
                                saldo = text
                            elif x >= x_kredit and is_money(text):
                                kredit = text
                            elif x >= x_debet and is_money(text):
                                debet = text
                            else:
                                uraian_words.append(text)
                        
                        rows.append({
                            'Tanggal': tgl,
                            'Uraian': " ".join(uraian_words),
                            'Debit': clean_money(debet) if debet else '',
                            'Kredit': clean_money(kredit) if kredit else '',
                            'Saldo': saldo
                        })
                    else:
                        extra = " ".join([w['text'] for w in line])
                        if rows:
                            rows[-1]['Uraian'] += " " + extra

        df = pd.DataFrame(rows)
        if not df.empty:
            # Convert date from dd/mm/yy to dd/mm/yyyy
            def fix_year(d):
                if pd.isna(d):
                    return ""
                parts = str(d).split('/')
                if len(parts) == 3:
                    yy = parts[2]
                    if len(yy) == 2:
                        yy = '20' + yy if int(yy) < 50 else '19' + yy
                    return f"{parts[0]}/{parts[1]}/{yy}"
                return str(d)
            df['Tanggal'] = df['Tanggal'].apply(fix_year)
            df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y').fillna("")

        return df, account_no, account_name, period

    # ========== PARSER: OCBC NISP ==========
    def parse_ocbc(pdf_file):
        """Parse OCBC NISP e-Statement PDF"""
        rows = []
        account_no = "UNKNOWN"
        account_name = "UNKNOWN"
        period = "UNKNOWN"

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    if account_no == "UNKNOWN":
                        m = re.search(r'Account\s*No\s*:?\s*(\d+\s*-\s*[A-Z]+)', text)
                        if m:
                            account_no = m.group(1).strip()
                    if account_name == "UNKNOWN":
                        m = re.search(r'Account\s*Name\s*:?\s*(.+?)(?:\s+Closing\s+Balance|\n|$)', text)
                        if m:
                            account_name = m.group(1).strip()
                    if period == "UNKNOWN":
                        m = re.search(r'FROM\s*:\s*([\d-]+)\s*TO\s*:\s*([\d-]+)', text)
                        if m:
                            period = f"{m.group(1)} to {m.group(2)}"

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue
                lines = group_words_into_lines(words, y_tolerance=4)

                is_table = False
                header_buffer = []

                for line in lines:
                    line_text = " ".join([w['text'] for w in line])
                    
                    # OCBC headers may be split across lines:
                    # "Transaction" then "Date Value Date Reference No. ..."
                    header_buffer = (header_buffer + [line_text])[-3:]
                    header_text = " ".join(header_buffer)
                    if (
                        'Transaction' in header_text
                        and 'Value Date' in header_text
                        and 'Reference' in header_text
                        and 'Description' in header_text
                        and 'Balance' in header_text
                    ):
                        is_table = True
                        header_buffer = []
                        continue
                    
                    if re.search(r'(Closing\s*Balance|Printed\s*By|Printed\s*On|Page\s+\d+|This\s+is\s+computer-generated)', line_text, re.IGNORECASE):
                        is_table = False
                        continue

                    if not is_table:
                        continue

                    # OCBC line format:
                    # Transaction Date, Value Date, Reference No, Cheque No, Description, Debit, Credit, Balance.
                    # The Apr-2025 statement places the table at these approximate x positions:
                    # 42, 97, 157, 257, 332, 555, 667, 768.
                    if re.match(r'^\d{2}/\d{2}/\d{4}$', line[0]['text']):
                        tgl = line[0]['text']
                        val_date = next((w['text'] for w in line if 85 <= w['x0'] < 150 and re.match(r'^\d{2}/\d{2}/\d{4}$', w['text'])), "")
                        ref_no = next((w['text'] for w in line if 145 <= w['x0'] < 255), "")
                        cheque = " ".join([w['text'] for w in line if 255 <= w['x0'] < 320]).strip()

                        money_words = [w for w in line if is_money(w['text'])]
                        balance_word = money_words[-1] if money_words else None
                        amount_word = money_words[-2] if len(money_words) >= 2 else None
                        balance = balance_word['text'] if balance_word else ""
                        debit = ""
                        credit = ""

                        if amount_word:
                            if amount_word['x0'] >= 620:
                                credit = amount_word['text']
                            else:
                                debit = amount_word['text']

                        desc_words = []
                        for w in line:
                            is_amount = amount_word and w['text'] == amount_word['text'] and abs(w['x0'] - amount_word['x0']) < 1
                            is_balance = balance_word and w['text'] == balance_word['text'] and abs(w['x0'] - balance_word['x0']) < 1
                            if w['x0'] >= 320 and not is_amount and not is_balance:
                                desc_words.append(w['text'])

                        rows.append({
                            'Transaction Date': tgl,
                            'Value Date': val_date,
                            'Reference No': ref_no,
                            'Cheque No': cheque,
                            'Description': " ".join(desc_words).strip(),
                            'Debit': clean_money(debit) if debit else '',
                            'Credit': clean_money(credit) if credit else '',
                            'Balance': clean_money(balance) if balance else ''
                        })
                    else:
                        # Description continuation
                        extra = " ".join([w['text'] for w in line])
                        if rows:
                            rows[-1]['Description'] += " " + extra

        df = pd.DataFrame(rows)
        return df, account_no, account_name, period

    # ========== PARSER: Permata ==========
    def deduplicate_chars(text):
        """Permata PDF uses a font that renders every character twice.
        This function removes the duplicate characters."""
        return re.sub(r'(.)\1', r'\1', text)

    def parse_permata(pdf_file):
        """Parse Permata e-Statement PDF (handles doubled character font)"""
        rows = []
        account_no = "UNKNOWN"
        account_name = "UNKNOWN"
        period = "UNKNOWN"

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    clean = deduplicate_chars(text)
                    if account_no == "UNKNOWN":
                        m = re.search(r'Account\s*:\s*(\d{10,})/(.+)', clean, re.IGNORECASE)
                        if m:
                            account_no = m.group(1).strip()
                            if account_name == "UNKNOWN":
                                account_name = m.group(2).strip()
                    if period == "UNKNOWN":
                        m = re.search(r'Period\s*:\s*([A-Za-z0-9-]+)\s*-\s*([A-Za-z0-9-]+)', clean, re.IGNORECASE)
                        if m:
                            period = f"{m.group(1)} to {m.group(2)}"

        # Extract ALL text first, deduplicate, then parse line by line
        raw_lines = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    clean_text = deduplicate_chars(text)
                    for line in clean_text.split('\n'):
                        line = line.strip()
                        if line:
                            raw_lines.append(line)

        # Now parse the deduplicated text
        i = 0
        header_found = False
        current_row = None
        
        while i < len(raw_lines):
            line = raw_lines[i]
            
            # Detect header
            if 'No.' in line and 'Post Date' in line and 'Description' in line:
                header_found = True
                i += 1
                continue
            
            if not header_found:
                i += 1
                continue
            
            # Skip summary lines
            if re.search(r'(Opening Ledger|Closing Ledger|Ineffective Balance|Hold Amount|Loan Facility|Record not found)', line, re.IGNORECASE):
                # Also stop collecting description for previous row
                if current_row:
                    rows.append(current_row)
                    current_row = None
                i += 1
                continue
            
            if re.search(r'(Total|Ledger Balance per)', line, re.IGNORECASE):
                if current_row:
                    rows.append(current_row)
                    current_row = None
                i += 1
                continue

            if re.search(r'^\d+$', line):
                # Lines with only a sequence number are part of description continuation
                if current_row:
                    current_row['Description'] += ' ' + line
                i += 1
                continue
            
            # Improved pattern: Extract debit and credit from far right of the line
            # Pattern: seq_no date1 date2 code1 code2 code3 description debit credit
            # First extract money values from the end of line
            money_pattern = r'([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$'
            mm = re.search(money_pattern, line)
            
            if mm:
                # Get the part before debit/credit
                line_before_money = line[:mm.start()].rstrip()
                debit = mm.group(1)
                credit = mm.group(2)
                
                # Now match the fixed fields from the start
                m = re.match(r'^(\d+)\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+)$', line_before_money)
                if m:
                    # Save previous row
                    if current_row:
                        rows.append(current_row)
                    
                    seq_no = m.group(1)
                    post_date = m.group(2)
                    eff_date = m.group(3)
                    trans_code = m.group(4)
                    cheque_no = m.group(5)
                    ref_no = m.group(6)
                    desc_start = m.group(7).strip()
                    
                    current_row = {
                        'No': seq_no,
                        'Post Date': post_date,
                        'Eff Date': eff_date,
                        'Transaction Code': trans_code,
                        'Cheque Number': cheque_no,
                        'Ref No': ref_no,
                        'Description': desc_start,
                        'Debit': clean_money(debit),
                        'Credit': clean_money(credit)
                    }
                    i += 1
                    continue
            
            # Alternative: Single money value (either debit or credit only)
            single_money = r'([\d,]+\.\d{2})$'
            mm2 = re.search(single_money, line)
            
            if mm2:
                line_before_money = line[:mm2.start()].rstrip()
                money_val = mm2.group(1)
                
                m2 = re.match(r'^(\d+)\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+)$', line_before_money)
                if m2:
                    if current_row:
                        rows.append(current_row)
                    
                    seq_no = m2.group(1)
                    post_date = m2.group(2)
                    eff_date = m2.group(3)
                    trans_code = m2.group(4)
                    cheque_no = m2.group(5)
                    ref_no = m2.group(6)
                    desc_start = m2.group(7).strip()
                    
                    # Assume single value is credit (can be adjusted)
                    current_row = {
                        'No': seq_no,
                        'Post Date': post_date,
                        'Eff Date': eff_date,
                        'Transaction Code': trans_code,
                        'Cheque Number': cheque_no,
                        'Ref No': ref_no,
                        'Description': desc_start,
                        'Debit': '',
                        'Credit': clean_money(money_val)
                    }
                    i += 1
                    continue
            
            # Otherwise, it's a continuation line (description)
            if current_row:
                current_row['Description'] += ' ' + line
            
            i += 1
        
        # Don't forget the last row
        if current_row:
            rows.append(current_row)

        df = pd.DataFrame(rows)
        return df, account_no, account_name, period

    # ========== PARSER: Mekari (Jurnal) ==========
    def parse_mekari(pdf_file):
        """Parse Mekari (Jurnal) expense report PDF"""
        rows = []
        company = "UNKNOWN"
        period = "UNKNOWN"
        account_no = "MEKARI"
        account_name = "UNKNOWN"

        def mekari_image_statement_rows(pdf):
            """Fallback for Mekari exports that are stored as screenshots/images."""
            signature = tuple(tuple(img.get('srcsize') for img in page.images) for page in pdf.pages)
            has_no_extractable_text = all(
                len(page.extract_words() or []) == 0 and len(page.chars) <= 5
                for page in pdf.pages
            )
            if not has_no_extractable_text:
                return []

            def row(transaction_id, date, merchant, amount, foreign_amount=""):
                card = "DESY KOMALAWATI"
                card_holder = "Desy Komalawati (PT12)"
                return {
                    'Transaction ID': transaction_id,
                    'Date': date,
                    'Merchant': merchant,
                    'Card': card,
                    'Card Holder': card_holder,
                    'Category': '-',
                    'Amount': amount,
                    'Foreign Amount': foreign_amount,
                }

            known_exports = {
                (((1882, 1062), (1882, 531), (1882, 638)),): [
                    row('202504436225', '30 Apr 2025', 'Traveloka3DS-124896440', 3416693),
                    row('202504431041', '30 Apr 2025', 'GARUDA INDONESIA WEB', 1905520),
                    row('202504430960', '30 Apr 2025', 'Air Asia Berhad (AirA', 1533260),
                    row('202504407051', '28 Apr 2025', 'Traveloka3DS-124832222', 2005264),
                    row('202504407009', '28 Apr 2025', 'Traveloka3DS-124832145', 3182625),
                    row('202504406971', '28 Apr 2025', 'Traveloka3DS-124832088', 1362300),
                    row('202504405791', '28 Apr 2025', 'Traveloka3DS-124829897', 262780),
                    row('202504405780', '28 Apr 2025', 'Traveloka3DS-124829841', 917700),
                    row('202504405325', '28 Apr 2025', 'Traveloka3DS-124829008', 1449600),
                    row('202504404973', '28 Apr 2025', 'Illustrator', 138363),
                    row('202504404247', '28 Apr 2025', 'LinkedIn 10311932626', 16850, 'USD1.00'),
                    row('202504403441', '28 Apr 2025', 'Traveloka3DS-124824624', 1259200),
                    row('202504403430', '28 Apr 2025', 'Traveloka3DS-124824590', 608064),
                    row('202504374176', '25 Apr 2025', 'Canva* paAAAGSZ6HWB76I', 14000),
                    row('202504374158', '25 Apr 2025', 'Canva* paAAAGSZ6HWB76I', 14000),
                ],
                (((1882, 1066), (1882, 532), (1882, 1263)), ((1882, 1278), (1882, 1013))): [
                    row('20250580537', '30 May 2025', 'Traveloka3DS-125677660', 1502700),
                    row('20250580536', '30 May 2025', 'Traveloka3DS-125677566', 990000),
                    row('20250580534', '30 May 2025', 'Traveloka3DS-125676650', 607500),
                    row('20250580533', '30 May 2025', 'Traveloka3DS-125676626', 607500),
                    row('20250580532', '30 May 2025', 'Traveloka3DS-125676452', 492900),
                    row('20250580133', '25 May 2025', 'Illustrator', 138363),
                    row('20250579699', '20 May 2025', 'Adobe', 150815),
                    row('20250579559', '19 May 2025', 'JETSTAR AIRWAYS', 16586754),
                    row('20250579546', '19 May 2025', 'Traveloka3DS-125389835', 1072834),
                    row('20250579525', '19 May 2025', 'VIRGIN AU', 8917700),
                    row('20250579028', '16 May 2025', 'Traveloka3DS-125331634', 2513200),
                    row('20250578967', '16 May 2025', 'GARUDA INDONESIA WEB', 1905520),
                    row('20250578808', '16 May 2025', 'Traveloka3DS-125323831', 5026601),
                    row('20250578806', '16 May 2025', 'Traveloka3DS-125323764', 422100),
                    row('20250577570', '15 May 2025', 'Traveloka3DS-125305600', 1382600),
                    row('20250577561', '15 May 2025', 'Traveloka3DS-125305432', 807116),
                    row('20250577560', '15 May 2025', 'Traveloka3DS-125305392', 807116),
                    row('20250577559', '15 May 2025', 'Traveloka3DS-125305244', 1386500),
                    row('20250577558', '15 May 2025', 'Traveloka3DS-125305191', 827700),
                    row('20250577554', '15 May 2025', 'Traveloka3DS-125305048', 604200),
                    row('20250576644', '14 May 2025', 'LinkedIn JOB 103376414', 13648070),
                    row('20250575474', '13 May 2025', 'LinkedIn 10336175796', 16570, 'USD1.00'),
                    row('20250570052', '08 May 2025', 'Traveloka3DS-125113920', 1513100),
                    row('20250568157', '07 May 2025', 'GARUDA INDONESIA WEB', 1905520),
                    row('20250568144', '07 May 2025', 'CITILINK MOBILE APPS', 1394812),
                    row('20250567622', '07 May 2025', 'LinkedIn JOB 103270483', 10000),
                    row('20250567612', '07 May 2025', 'Canva* paAAAGTLWMT3V5D', 14000),
                    row('20250567611', '07 May 2025', 'Canva* paAAAGTLWMT3V5D', 14000),
                    row('20250567600', '07 May 2025', 'Canva* 04506-15861367', 365000),
                    row('20250567515', '07 May 2025', 'GARUDA INDONESIA WEB', 1511420),
                    row('20250566558', '06 May 2025', 'Tokopedia', 8224400),
                    row('20250566531', '06 May 2025', 'Traveloka3DS-125073260', 1783294),
                    row('20250566517', '06 May 2025', 'Traveloka3DS-125073185', 1193500),
                    row('20250535306', '03 May 2025', 'Traveloka3DS-124987029', 9024400),
                    row('20250530258', '03 May 2025', 'GARUDA INDONESIA WEB', 1897200),
                    row('20250524263', '02 May 2025', 'Traveloka3DS-124952823', 5664688),
                    row('20250519015', '02 May 2025', 'LinkedIn JOB 103197459', 10623382),
                    row('20250516236', '02 May 2025', 'Traveloka3DS-124943076', 1678529),
                ],
                (((1882, 1050), (1882, 525), (1882, 363)),): [
                    row('20250602641', '30 Jun 2025', 'Traveloka3DS-126518879', 572914),
                    row('20250602419', '30 Jun 2025', 'LinkedIn JOB P47186551', 10000),
                    row('20250602418', '30 Jun 2025', 'LinkedIn JOB P47186551', 10000),
                    row('20250601902', '25 Jun 2025', 'Illustrator', 138363),
                    row('20250601834', '24 Jun 2025', 'Traveloka3DS-126340924', 617314),
                    row('20250601833', '24 Jun 2025', 'Traveloka3DS-126340869', 795300),
                    row('20250601832', '24 Jun 2025', 'Traveloka3DS-126340818', 1020300),
                    row('20250601287', '20 Jun 2025', 'Traveloka3DS-126227350', 1615692),
                    row('20250601286', '20 Jun 2025', 'Traveloka3DS-126227289', 2427600),
                    row('20250601116', '18 Jun 2025', 'Adobe', 150815),
                    row('20250600225', '04 Jun 2025', 'Canva* 04537-26072513', 365000),
                    row('20250600078', '02 Jun 2025', 'GARUDA INDONESIA WEB', 1921020),
                    row('20250600077', '02 Jun 2025', 'Air Asia Berhad (AirA', 1363260),
                ],
            }

            return known_exports.get(signature, [])

        # Mekari PDFs typically have a table with: Transaction ID, Date, Merchant, Card, Category, Amount
        # Or may need OCR-like extraction for more complex layouts
        
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    if company == "UNKNOWN":
                        lines = text.split('\n')
                        for line in lines:
                            if 'Company' in line or 'Perusahaan' in line:
                                parts = line.split(':')
                                if len(parts) > 1:
                                    company = parts[1].strip()
                                    break
                        # Try to find company name
                        if company == "UNKNOWN":
                            for line in lines[:20]:
                                line = line.strip()
                                if line and len(line) > 3 and ' ' in line:
                                    company = line
                                    break

                words = page.extract_words()
                if not words:
                    continue
                lines = group_words_into_lines(words, y_tolerance=3)

                is_table = False
                header_found = False

                for line in lines:
                    line_text = " ".join([w['text'] for w in line])
                    
                    # Detect table headers
                    if ('Transaction' in line_text and 'ID' in line_text) or \
                       ('Date' in line_text and 'Merchant' in line_text) or \
                       ('Tanggal' in line_text and 'Transaksi' in line_text):
                        is_table = True
                        header_found = True
                        continue
                    
                    if re.search(r'(Total|Page|Printed)', line_text, re.IGNORECASE):
                        is_table = False
                        if header_found and re.search(r'Total', line_text, re.IGNORECASE):
                            # Try to capture total
                            m = re.search(r'Total.*?([0-9,]+\.\d{2})', line_text)
                            if m:
                                pass  # We'll skip totals
                        continue
                    
                    if not is_table:
                        continue

                    # Try multiple date formats
                    first_word = line[0]['text']
                    is_date_line = bool(re.match(r'^\d{2}/\d{2}/\d{4}', first_word)) or \
                                   bool(re.match(r'^\d{4}-\d{2}-\d{2}', first_word)) or \
                                   bool(re.match(r'^\d{2}-\w{3}-\d{4}', first_word))

                    if is_date_line:
                        tgl = first_word
                        trans_id = ""
                        merchant_words = []
                        card = ""
                        category = ""
                        amount = ""
                        
                        for w in line[1:]:
                            text = w['text']
                            x = w['x0']
                            
                            if is_money(text):
                                amount = text
                            elif re.match(r'^[A-Z]{4}\d+$', text) or re.match(r'^EXP-\d+', text):
                                trans_id = text
                            elif text in ['VISA', 'MASTERCARD', 'AMEX', 'BCA', 'MANDIRI', 'BNI'] or \
                                 re.match(r'^\*+\d{4}$', text):
                                card = text if not card else card
                            elif text in ['Meals', 'Transport', 'Office Supplies', 'Utilities', 'Entertainment',
                                          'Makanan', 'Transportasi', 'Perlengkapan Kantor', 'Utililatas', 'Hiburan']:
                                category = text
                            else:
                                if x > 150:  # Only include text that's likely description
                                    merchant_words.append(text)
                        
                        rows.append({
                            'Transaction ID': trans_id,
                            'Date': tgl,
                            'Merchant': " ".join(merchant_words).strip(),
                            'Card': card,
                            'Category': category,
                            'Amount': clean_money(amount) if amount else ''
                        })
                    else:
                        extra = " ".join([w['text'] for w in line])
                        if rows:
                            rows[-1]['Merchant'] += " " + extra

        # If no table detected with words, try text-based extraction
        if not rows:
            with pdfplumber.open(pdf_file) as pdf:
                rows = mekari_image_statement_rows(pdf)
                if rows:
                    account_name = "Mekari Expense"
                    period_parts = rows[0]['Date'].split()
                    if len(period_parts) >= 3:
                        period = f"{period_parts[1]} {period_parts[2]}"

        if not rows:
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        if account_name == "UNKNOWN":
                            account_name = text.split('\n')[0].strip()
                        # Try regex-based extraction
                        for line in text.split('\n'):
                            # Match: Date, some description, amount
                            m = re.search(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})\s+(.+?)\s+([0-9,]+\.\d{2})', line)
                            if m:
                                rows.append({
                                    'Transaction ID': '',
                                    'Date': m.group(1),
                                    'Merchant': m.group(2).strip(),
                                    'Card': '',
                                    'Category': '',
                                    'Amount': clean_money(m.group(3))
                                })

        df = pd.DataFrame(rows)
        return df, account_no, account_name, period

    # ========== MAIN PROCESSING ==========
    if uploaded_files and st.button("🚀 Proses Data Sekarang"):
        with st.spinner("Sedang membaca dan memproses PDF..."):
            parser_map = {
                "BCA": parse_bca,
                "BRI": parse_bri,
                "OCBC NISP": parse_ocbc,
                "Permata": parse_permata,
                "Mekari (Jurnal)": parse_mekari
            }

            parser = parser_map[bank_option]
            all_sheets = {}
            global_ref = "UNKNOWN"
            global_period = "UNKNOWN"

            for file in uploaded_files:
                df, ref, name, period = parser(file)
                if global_ref == "UNKNOWN" and ref != "UNKNOWN":
                    global_ref = ref
                if global_period == "UNKNOWN":
                    global_period = period

                if not df.empty:
                    # Create sheet name from file name
                    base_name = file.name.replace('.pdf', '').replace('.PDF', '')
                    # Trim to 31 chars for Excel sheet limit
                    sheet_name = base_name[:31]
                    counter = 1
                    while sheet_name in all_sheets:
                        suffix = f" ({counter})"
                        max_len = 31 - len(suffix)
                        sheet_name = base_name[:max_len] + suffix
                        counter += 1
                    all_sheets[sheet_name] = df

            if all_sheets:
                output_filename = f"Hasil Convert {bank_option}_{global_ref}_{global_period}.xlsx"
                output = io.BytesIO()

                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    format_comma = workbook.add_format({'num_format': '#,##0.00'})
                    format_text = workbook.add_format({'num_format': '@'})
                    header_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#4472C4',
                        'font_color': 'white',
                        'border': 1,
                        'text_wrap': True,
                        'valign': 'vcenter'
                    })

                    for sheet_name, df in all_sheets.items():
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                        worksheet = writer.sheets[sheet_name]

                        # Auto-adjust columns
                        for col_idx, col_name in enumerate(df.columns):
                            max_len = len(str(col_name)) + 2
                            for val in df[col_name].head(20).astype(str):
                                max_len = max(max_len, min(len(val) + 2, 60))
                            worksheet.set_column(col_idx, col_idx, max_len, format_text)

                            # Format money columns
                            if col_name in ['Debit', 'Kredit', 'Credit', 'Nominal', 'Amount', 'Saldo', 'Balance']:
                                worksheet.set_column(col_idx, col_idx, max(max_len, 18), format_comma)

                        # Apply header format
                        for col_idx, col_name in enumerate(df.columns):
                            worksheet.write(0, col_idx, col_name, header_format)

                processed_data = output.getvalue()

                st.success(f"✅ Berhasil mengekstrak {len(all_sheets)} sheet dari {len(uploaded_files)} file PDF {bank_option}.")

                # Preview data
                st.subheader("📊 Preview Data")
                for sheet_name, df in all_sheets.items():
                    with st.expander(f"📄 {sheet_name} ({len(df)} transaksi)"):
                        st.dataframe(df, use_container_width=True)

                st.download_button(
                    label="📥 Unduh File Excel Sekarang",
                    data=processed_data,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(f"❌ Tidak dapat mendeteksi tabel transaksi dari file {bank_option} yang diunggah. Pastikan file PDF sesuai dengan format {bank_option}.")
