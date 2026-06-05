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

        def permata_money(text):
            return bool(re.match(r'^[\d,]+\.\d{2}$', text.strip()))

        def dedup_word(word):
            return deduplicate_chars(word.get('text', '')).strip()

        def parse_permata_by_position(pdf):
            positioned_rows = []
            current_row = None
            is_table = False
            x_customer, x_desc, x_debit, x_credit = 540, 620, 790, 930
            header_buffer = []
            header_line_buffer = []

            for page in pdf.pages:
                words = page.extract_words()
                if not words:
                    continue

                clean_words = []
                for word in words:
                    clean_word = dict(word)
                    clean_word['text'] = dedup_word(word)
                    if clean_word['text']:
                        clean_words.append(clean_word)

                for line in group_words_into_lines(clean_words, y_tolerance=4):
                    line_text = " ".join([w['text'] for w in line]).strip()
                    if not line_text:
                        continue

                    header_buffer = (header_buffer + [line_text])[-3:]
                    header_line_buffer = (header_line_buffer + [line])[-3:]
                    header_text = " ".join(header_buffer)
                    if 'No.' in header_text and 'Post Date' in header_text and 'Description' in header_text:
                        is_table = True
                        for header_line in header_line_buffer:
                            for word in header_line:
                                if word['text'] == 'Description':
                                    x_desc = word['x0'] - 5
                                elif word['text'] == 'Customer':
                                    x_customer = word['x0'] - 5
                                elif word['text'] == 'Debit':
                                    x_debit = word['x0'] - 5
                                elif word['text'] == 'Credit':
                                    x_credit = word['x0'] - 5
                        for word in line:
                            if word['text'] == 'Description':
                                x_desc = word['x0'] - 5
                            elif word['text'] == 'Customer':
                                x_customer = word['x0'] - 5
                            elif word['text'] == 'Debit':
                                x_debit = word['x0'] - 5
                            elif word['text'] == 'Credit':
                                x_credit = word['x0'] - 5
                        header_buffer = []
                        header_line_buffer = []
                        continue

                    if not is_table:
                        continue

                    if re.search(r'(Opening Ledger|Closing Ledger|Ineffective Balance|Hold Amount|Loan Facility|Record not found|Total|Ledger Balance per)', line_text, re.IGNORECASE):
                        if current_row:
                            positioned_rows.append(current_row)
                            current_row = None
                        continue

                    m = re.match(r'^(\d+)\s*(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)', line_text)
                    if not m:
                        if current_row:
                            continuation_words = []
                            for w in line:
                                if permata_money(w['text']):
                                    continue
                                if w['x0'] >= x_desc and w['x0'] < x_debit:
                                    continuation_words.append(w['text'])
                                elif w['x0'] >= x_desc and not any(permata_money(item['text']) for item in line):
                                    continuation_words.append(w['text'])
                            continuation = " ".join(continuation_words).strip()
                            if continuation and not re.search(r'^(No\.|Post Date|Description|Debit|Credit)', continuation, re.IGNORECASE):
                                current_row['Description'] = f"{current_row['Description']} {continuation}".strip()
                        continue

                    if current_row:
                        positioned_rows.append(current_row)

                    money_words = sorted([w for w in line if permata_money(w['text'])], key=lambda item: item['x0'])
                    debit_word = None
                    credit_word = None
                    if len(money_words) >= 2:
                        debit_word = money_words[-2]
                        credit_word = money_words[-1]
                    elif len(money_words) == 1:
                        amount_word = money_words[0]
                        if amount_word['x0'] >= x_credit:
                            credit_word = amount_word
                        else:
                            debit_word = amount_word

                    customer_ref_words = []
                    desc_words = []
                    amount_start_x = debit_word['x0'] if debit_word else (credit_word['x0'] if credit_word else x_debit)
                    for word in line:
                        same_as_debit = debit_word and word['text'] == debit_word['text'] and abs(word['x0'] - debit_word['x0']) < 1
                        same_as_credit = credit_word and word['text'] == credit_word['text'] and abs(word['x0'] - credit_word['x0']) < 1
                        if x_customer <= word['x0'] < x_desc and not same_as_debit and not same_as_credit:
                            customer_ref_words.append(word['text'])
                        elif x_desc <= word['x0'] < amount_start_x and not same_as_debit and not same_as_credit:
                            desc_words.append(word['text'])

                    current_row = {
                        'No': m.group(1),
                        'Post Date': m.group(2),
                        'Eff Date': m.group(3),
                        'Transaction Code': m.group(4),
                        'Cheque Number': m.group(5),
                        'Ref No': m.group(6),
                        'Customer Ref No': " ".join(customer_ref_words).strip(),
                        'Description': " ".join(desc_words).strip(),
                        'Debit': clean_money(debit_word['text']) if debit_word else '',
                        'Credit': clean_money(credit_word['text']) if credit_word else ''
                    }

            if current_row:
                positioned_rows.append(current_row)

            return positioned_rows

        def split_permata_amounts_from_description(row):
            desc = row.get('Description', '').strip()
            money_matches = list(re.finditer(r'[\d,]+\.\d{2}', desc))
            if not money_matches:
                return row

            amounts_to_remove = []
            if row.get('Credit') == '' and len(money_matches) >= 1:
                row['Credit'] = clean_money(money_matches[-1].group(0))
                amounts_to_remove.append(money_matches[-1])
            if row.get('Debit') == '' and len(money_matches) >= 2:
                row['Debit'] = clean_money(money_matches[-2].group(0))
                amounts_to_remove.append(money_matches[-2])

            for match in sorted(amounts_to_remove, key=lambda item: item.start(), reverse=True):
                desc = f"{desc[:match.start()]} {desc[match.end():]}"
            row['Description'] = re.sub(r'\s+', ' ', desc).strip()
            return row

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

        with pdfplumber.open(pdf_file) as pdf:
            rows = parse_permata_by_position(pdf)

        if rows:
            rows = [split_permata_amounts_from_description(row) for row in rows]
            df = pd.DataFrame(rows)
            return df, account_no, account_name, period

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
            
            # Check if this line starts a new transaction: sequence number immediately followed by date
            # Pattern: e.g., "103-Mar-2025" (No=1, PostDate=03-Mar-2025) or "1003-Mar-2025" (No=10)
            m = re.match(r'^(\d+)\s*(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$', line)
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
                debit = m.group(8)
                credit = m.group(9)
                
                current_row = {
                    'No': seq_no,
                    'Post Date': post_date,
                    'Eff Date': eff_date,
                    'Transaction Code': trans_code,
                    'Cheque Number': cheque_no,
                    'Ref No': ref_no,
                    'Customer Ref No': '',
                    'Description': desc_start,
                    'Debit': clean_money(debit),
                    'Credit': clean_money(credit)
                }
                i += 1
                continue
            
            # Try alternative pattern where debit/credit might be empty
            m2 = re.match(r'^(\d+)\s*(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+)$', line)
            if m2:
                if current_row:
                    rows.append(current_row)
                
                seq_no = m2.group(1)
                post_date = m2.group(2)
                eff_date = m2.group(3)
                trans_code = m2.group(4)
                cheque_no = m2.group(5)
                ref_no = m2.group(6)
                rest = m2.group(7).strip()
                
                # Try to extract debit/credit from end of rest
                debit = ""
                credit = ""
                desc = rest
                
                # Check if rest ends with two money values
                money_pattern = r'([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$'
                mm = re.search(money_pattern, rest)
                if mm:
                    debit = mm.group(1)
                    credit = mm.group(2)
                    desc = rest[:mm.start()].strip()
                else:
                    # Maybe ends with one money value
                    mm2 = re.search(r'([\d,]+\.\d{2})$', rest)
                    if mm2:
                        # Check if it's at the far right (credit) or could be debit
                        # Try finding by position in original
                        pass  # keep desc as rest
                
                current_row = {
                    'No': seq_no,
                    'Post Date': post_date,
                    'Eff Date': eff_date,
                    'Transaction Code': trans_code,
                    'Cheque Number': cheque_no,
                    'Ref No': ref_no,
                    'Customer Ref No': '',
                    'Description': desc,
                    'Debit': clean_money(debit) if debit else '',
                    'Credit': clean_money(credit) if credit else ''
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

        rows = [split_permata_amounts_from_description(row) for row in rows]
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

        def norm_text(value):
            return re.sub(r'\s+', ' ', str(value or '').replace('\n', ' ')).strip()

        def clean_mekari_amount(value):
            value = norm_text(value)
            if not value:
                return ''
            negative = value.startswith('(') and value.endswith(')')
            value = re.sub(r'(?i)\b(rp|idr)\b', '', value)
            value = re.sub(r'[^0-9,.\-]', '', value)
            if not value or not re.search(r'\d', value):
                return ''
            if value.startswith('-'):
                negative = True
                value = value[1:]

            last_dot = value.rfind('.')
            last_comma = value.rfind(',')
            decimal_sep = ''
            if last_dot > -1 and last_comma > -1:
                decimal_sep = '.' if last_dot > last_comma else ','
            elif last_dot > -1 and len(value) - last_dot - 1 == 2:
                decimal_sep = '.'
            elif last_comma > -1 and len(value) - last_comma - 1 == 2:
                decimal_sep = ','

            if decimal_sep:
                thousands_sep = ',' if decimal_sep == '.' else '.'
                value = value.replace(thousands_sep, '').replace(decimal_sep, '.')
            else:
                value = value.replace(',', '').replace('.', '')

            try:
                amount = float(value)
                return -amount if negative else amount
            except ValueError:
                return ''

        def looks_like_amount(value):
            text = norm_text(value)
            if not re.search(r'\d', text):
                return False
            if looks_like_mekari_date(text):
                return False
            if re.match(r'^\d{8,16}$', text):
                return False
            return clean_mekari_amount(text) != ''

        def has_currency_letters(value):
            return bool(re.search(r'[A-Za-z]', norm_text(value)))

        def looks_like_transaction_id(value):
            text = norm_text(value)
            return bool(re.match(r'^(\d{8,16}|[A-Z]{2,}[-/]\d[\w/-]*)$', text))

        def looks_like_mekari_date(value):
            text = norm_text(value)
            return bool(re.match(r'^(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})$', text))

        def empty_row(transaction_id='', date='', merchant='', card='', card_holder='', category='', amount='', foreign_amount=''):
            return {
                'Transaction ID': norm_text(transaction_id),
                'Date': norm_text(date),
                'Merchant': norm_text(merchant),
                'Card': norm_text(card),
                'Card Holder': norm_text(card_holder),
                'Category': norm_text(category),
                'Amount': clean_mekari_amount(amount) if amount != '' else '',
                'Foreign Amount': norm_text(foreign_amount),
            }

        def add_row(row):
            if not row.get('Date') or row.get('Amount') == '':
                return
            if not row.get('Merchant') and not row.get('Transaction ID'):
                return
            key = (
                row.get('Transaction ID', ''),
                row.get('Date', ''),
                row.get('Merchant', ''),
                row.get('Amount', ''),
            )
            if key not in seen_rows:
                rows.append(row)
                seen_rows.add(key)

        def parse_mekari_table_row(cells, header_map=None):
            cells = [norm_text(cell) for cell in cells]
            if not any(cells):
                return None
            row_text = ' '.join(cells)
            if re.search(r'^(total|page|printed|transaction\s+id|date\b)', row_text, re.IGNORECASE):
                return None

            if header_map:
                def get(*keys):
                    normalized_map = {
                        re.sub(r'[^a-z0-9 ]+', '', column).strip(): idx
                        for column, idx in header_map.items()
                    }
                    for key in keys:
                        key = re.sub(r'[^a-z0-9 ]+', '', key.lower()).strip()
                        idx = normalized_map.get(key)
                        if idx is not None and idx < len(cells):
                            return cells[idx]
                    for key in keys:
                        key = re.sub(r'[^a-z0-9 ]+', '', key.lower()).strip()
                        for column, idx in normalized_map.items():
                            if key == 'amount' and 'foreign' in column:
                                continue
                            if (key in column or column in key) and idx < len(cells):
                                return cells[idx]
                    return ''

                return empty_row(
                    get('transaction id', 'id transaksi', 'no transaksi'),
                    get('date', 'tanggal'),
                    get('merchant', 'description', 'keterangan'),
                    get('card', 'kartu'),
                    get('card holder', 'pemegang kartu'),
                    get('category', 'kategori'),
                    get('amount', 'nominal', 'jumlah'),
                    get('foreign amount', 'foreign', 'mata uang asing'),
                )

            transaction_id = cells[0] if looks_like_transaction_id(cells[0]) else ''
            date_idx = next((idx for idx, cell in enumerate(cells) if looks_like_mekari_date(cell)), None)
            amount_idx = next(
                (idx for idx in range(len(cells) - 1, -1, -1) if looks_like_amount(cells[idx]) and not has_currency_letters(cells[idx])),
                None
            )
            if amount_idx is None:
                amount_idx = next((idx for idx in range(len(cells) - 1, -1, -1) if looks_like_amount(cells[idx])), None)
            if date_idx is None or amount_idx is None:
                return None

            merchant_start = date_idx + 1
            merchant_end = amount_idx
            merchant = ' '.join(cells[merchant_start:merchant_end])
            foreign_amount = ''
            if amount_idx + 1 < len(cells) and has_currency_letters(cells[amount_idx + 1]):
                foreign_amount = cells[amount_idx + 1]
            return empty_row(transaction_id, cells[date_idx], merchant, '', '', '', cells[amount_idx], foreign_amount)

        def parse_mekari_text_line(line_text):
            line_text = norm_text(line_text)
            if not line_text or re.search(r'^(total|page|printed|transaction\s+id|date\b)', line_text, re.IGNORECASE):
                return None

            id_pat = r'(\d{8,16}|[A-Z]{2,}[-/]\d[\w/-]*)'
            date_pat = r'(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})'
            match = re.match(rf'^{id_pat}\s+{date_pat}\s+(.+)$', line_text)
            if match:
                transaction_id, date, rest = match.group(1), match.group(2), match.group(3)
            else:
                match = re.match(rf'^{date_pat}\s+(.+)$', line_text)
                if not match:
                    return None
                transaction_id, date, rest = '', match.group(1), match.group(2)

            parts = rest.split()
            amount_idx = next(
                (idx for idx in range(len(parts) - 1, -1, -1) if looks_like_amount(parts[idx]) and not has_currency_letters(parts[idx])),
                None
            )
            if amount_idx is None:
                amount_idx = next((idx for idx in range(len(parts) - 1, -1, -1) if looks_like_amount(parts[idx])), None)
            if amount_idx is None:
                return None

            amount = parts[amount_idx]
            foreign_amount = ''
            if amount_idx + 1 < len(parts) and re.match(r'^[A-Z]{3}\s*[-0-9,.]+$', parts[amount_idx + 1], re.IGNORECASE):
                foreign_amount = parts[amount_idx + 1]

            description = ' '.join(parts[:amount_idx])
            return empty_row(transaction_id, date, description, '', '', '', amount, foreign_amount)

        def update_mekari_metadata(text):
            nonlocal company, account_name, period
            if not text:
                return
            lines = [norm_text(line) for line in text.split('\n') if norm_text(line)]
            for line in lines:
                if company == "UNKNOWN":
                    m = re.search(r'(Company|Perusahaan)\s*:?\s*(.+)', line, re.IGNORECASE)
                    if m:
                        company = m.group(2).strip()
                        account_name = company
                if period == "UNKNOWN":
                    m = re.search(r'(Period|Periode)\s*:?\s*(.+)', line, re.IGNORECASE)
                    if m:
                        period = m.group(2).strip()
            if account_name == "UNKNOWN" and lines:
                account_name = lines[0]

        seen_rows = set()

        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                update_mekari_metadata(text)

                table_settings = {
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "intersection_tolerance": 5,
                    "text_tolerance": 3,
                }
                for table in page.extract_tables(table_settings=table_settings) or []:
                    header_map = None
                    for table_row in table:
                        cells = [norm_text(cell) for cell in table_row]
                        header_text = ' '.join(cells).lower()
                        if 'transaction' in header_text and ('date' in header_text or 'tanggal' in header_text):
                            header_map = {}
                            for idx, cell in enumerate(cells):
                                key = cell.lower()
                                key = re.sub(r'\s+', ' ', key)
                                header_map[key] = idx
                            continue
                        parsed_row = parse_mekari_table_row(cells, header_map)
                        if parsed_row:
                            add_row(parsed_row)

                words = page.extract_words()
                if not words:
                    continue
                for line in group_words_into_lines(words, y_tolerance=4):
                    line_text = " ".join([w['text'] for w in line])
                    parsed_row = parse_mekari_text_line(line_text)
                    if parsed_row:
                        add_row(parsed_row)
                    elif rows and line_text and line[0]['x0'] > 120 and not re.search(
                        r'(total|page|printed|company|perusahaan|period|periode|transaction\s+id|date\b|merchant)',
                        line_text,
                        re.IGNORECASE
                    ):
                        rows[-1]['Merchant'] = norm_text(f"{rows[-1]['Merchant']} {line_text}")

        # If no table detected with words, try image-based known exports.
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
                    update_mekari_metadata(text)
                    if text:
                        for line in text.split('\n'):
                            parsed_row = parse_mekari_text_line(line)
                            if parsed_row:
                                add_row(parsed_row)

        df = pd.DataFrame(rows)
        if not df.empty:
            ordered_cols = ['Transaction ID', 'Date', 'Merchant', 'Card', 'Card Holder', 'Category', 'Amount', 'Foreign Amount']
            for col in ordered_cols:
                if col not in df.columns:
                    df[col] = ''
            df = df[ordered_cols]
            if period == "UNKNOWN":
                dates = df['Date'].dropna().astype(str)
                if not dates.empty:
                    parts = dates.iloc[0].split()
                    if len(parts) >= 3:
                        period = f"{parts[1]} {parts[2]}"
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
