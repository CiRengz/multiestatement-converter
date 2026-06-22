import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# =====================================================================
# KONFIGURASI & STYLE
# =====================================================================
st.set_page_config(
    page_title="Konverter Rekening Koran ke Excel",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------- BAHASA ----------
LANG = {
    "id": {
        "page_title": "Konverter Rekening Koran ke Excel",
        "password_title": "Akses Terbatas",
        "password_desc": "Masukkan kata sandi untuk mengakses aplikasi",
        "password_placeholder": "Masukkan kata sandi...",
        "password_wrong": "❌ Kata sandi salah. Silakan coba lagi.",
        "hero_badge": "✨ Multi-Bank Supported",
        "hero_title": "🏦 Konverter e-Statement ke Excel",
        "hero_desc": "Konversi e-Statement PDF ke Excel dengan mudah. Mendukung berbagai format bank: BCA, BRI, OCBC NISP, Permata, Mekari (Jurnal).",
        "step1": "Pilih Bank",
        "step2": "Upload PDF",
        "step3": "Proses Data",
        "step4": "Unduh Excel",
        "bank_label": "🏛️ Pilih Bank",
        "bank_desc": "Pilih sumber e-Statement yang akan dikonversi.",
        "upload_label": "📄 Upload File",
        "upload_desc": "Unggah file PDF e-Statement yang ingin diproses.",
        "file_ready": "file PDF siap diproses",
        "btn_process": "🚀 Proses Data Sekarang",
        "processing": "Sedang membaca dan memproses PDF...",
        "success_msg": "Berhasil mengekstrak",
        "sheet_label": "sheet dari",
        "files_label": "file PDF",
        "preview_title": "📊 Preview Data",
        "transactions": "transaksi",
        "btn_download": "📥 Unduh File Excel Sekarang",
        "btn_download_sheet": "📥 Unduh Sheet Ini",
        "error_no_table": "❌ Tidak dapat mendeteksi tabel transaksi dari file",
        "error_no_table2": "yang diunggah. Pastikan file PDF sesuai dengan format",
        "pdf_password_label": "🔐 Password PDF (jika ada)",
        "pdf_password_hint": "Kosongkan jika PDF tidak dipassword",
        "opening_balance": "Saldo Awal",
        "footer1": "🏦 Multi-Bank Statement Converter • Dibuat dengan ❤️",
        "footer2": "Mendukung BCA • BRI • OCBC NISP • Permata • Mekari (Jurnal)",
        "lang_btn": "🌐 English",
        "go_ahead": "Masuk",
    },
    "en": {
        "page_title": "Bank Statement to Excel Converter",
        "password_title": "Restricted Access",
        "password_desc": "Enter the password to access the application",
        "password_placeholder": "Enter password...",
        "password_wrong": "❌ Wrong password. Please try again.",
        "hero_badge": "✨ Multi-Bank Supported",
        "hero_title": "🏦 e-Statement to Excel Converter",
        "hero_desc": "Convert e-Statement PDF to Excel easily. Supports: BCA, BRI, OCBC NISP, Permata, Mekari (Jurnal).",
        "step1": "Select Bank",
        "step2": "Upload PDF",
        "step3": "Process Data",
        "step4": "Download Excel",
        "bank_label": "🏛️ Select Bank",
        "bank_desc": "Select e-Statement source to convert.",
        "upload_label": "📄 Upload File",
        "upload_desc": "Upload your e-Statement PDF files.",
        "file_ready": "PDF files ready to process",
        "btn_process": "🚀 Process Data Now",
        "processing": "Reading and processing PDFs...",
        "success_msg": "Successfully extracted",
        "sheet_label": "sheets from",
        "files_label": "PDF files",
        "preview_title": "📊 Data Preview",
        "transactions": "transactions",
        "btn_download": "📥 Download Excel File Now",
        "btn_download_sheet": "📥 Download This Sheet",
        "error_no_table": "❌ Cannot detect transaction table from",
        "error_no_table2": "files. Make sure the PDF matches",
        "pdf_password_label": "🔐 PDF Password (if any)",
        "pdf_password_hint": "Leave empty if not password-protected",
        "opening_balance": "Opening Balance",
        "footer1": "🏦 Multi-Bank Statement Converter • Made with ❤️",
        "footer2": "Supports BCA • BRI • OCBC NISP • Permata • Mekari (Jurnal)",
        "lang_btn": "🌐 Indonesia",
        "go_ahead": "Enter",
    }
}

if "lang" not in st.session_state:
    st.session_state.lang = "id"
if "pdf_password" not in st.session_state:
    st.session_state.pdf_password = ""

def t(key):
    return LANG[st.session_state.lang][key]

# ---------- CUSTOM CSS ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

* { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }

.stApp {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
}

h1, h2, h3, h4, h5, h6 {
    font-family: 'Inter', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em !important;
}

p, li, .stMarkdown, .stText { font-family: 'Inter', sans-serif !important; }

.main-card {
    background: rgba(255, 255, 255, 0.05);
    backdrop-filter: blur(20px);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 20px;
    padding: 2rem;
    margin: 1.5rem 0;
    transition: all 0.3s ease;
}

.hero-section {
    text-align: center;
    padding: 3rem 1rem 2rem 1rem;
    background: linear-gradient(135deg, rgba(59,130,246,0.15) 0%, rgba(139,92,246,0.1) 100%);
    border-radius: 24px;
    border: 1px solid rgba(59,130,246,0.2);
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}

.hero-section::before {
    content: '';
    position: absolute;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: radial-gradient(circle at 30% 50%, rgba(59,130,246,0.08) 0%, transparent 50%),
                radial-gradient(circle at 70% 50%, rgba(139,92,246,0.06) 0%, transparent 50%);
    animation: heroGlow 8s ease-in-out infinite alternate;
}

@keyframes heroGlow {
    0% { transform: translate(0, 0) rotate(0deg); }
    100% { transform: translate(2%, 2%) rotate(3deg); }
}

.hero-title {
    font-size: 2.8rem !important;
    font-weight: 800 !important;
    background: linear-gradient(135deg, #60a5fa 0%, #a78bfa 50%, #f472b6 100%);
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    background-clip: text !important;
    margin-bottom: 0.5rem !important;
    position: relative;
    z-index: 1;
    line-height: 1.2 !important;
}

.hero-subtitle {
    font-size: 1.1rem !important;
    color: #94a3b8 !important;
    max-width: 600px;
    margin: 0 auto !important;
    position: relative;
    z-index: 1;
    line-height: 1.6 !important;
}

.hero-badge {
    display: inline-block;
    background: linear-gradient(135deg, rgba(59,130,246,0.2), rgba(139,92,246,0.2));
    border: 1px solid rgba(59,130,246,0.3);
    border-radius: 100px;
    padding: 0.35rem 1rem;
    font-size: 0.8rem;
    font-weight: 600;
    color: #93c5fd;
    margin-bottom: 1.2rem;
    position: relative; z-index: 1;
}

.stSelectbox label, .stFileUploader label {
    font-weight: 600 !important; color: #e2e8f0 !important; font-size: 0.95rem !important;
}

.stSelectbox > div > div {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    border-radius: 12px !important;
    color: #f1f5f9 !important;
    transition: all 0.3s ease !important;
}

.stSelectbox > div > div:hover {
    border-color: rgba(59,130,246,0.5) !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,0.1) !important;
}

div[data-testid="stSelectbox"] ul {
    background: #1e293b !important;
    border: 1px solid rgba(255,255,255,0.1) !important;
    border-radius: 12px !important;
}

.stFileUploader > section {
    background: rgba(255,255,255,0.04) !important;
    border: 2px dashed rgba(59,130,246,0.3) !important;
    border-radius: 16px !important;
    padding: 2rem !important;
    transition: all 0.3s ease !important;
}

.stFileUploader > section:hover {
    border-color: #3b82f6 !important;
    background: rgba(59,130,246,0.05) !important;
}

.stFileUploader > section > div { color: #94a3b8 !important; }
.stFileUploader ul {
    background: rgba(255,255,255,0.03) !important;
    border-radius: 12px !important;
    border: 1px solid rgba(255,255,255,0.06) !important;
}
.stFileUploader ul li { color: #e2e8f0 !important; }

.stButton > button {
    background: linear-gradient(135deg, #3b82f6 0%, #6366f1 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.75rem 2rem !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    transition: all 0.3s ease !important;
    box-shadow: 0 4px 20px rgba(59,130,246,0.3) !important;
    width: 100% !important;
}

.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 30px rgba(59,130,246,0.4) !important;
    background: linear-gradient(135deg, #60a5fa 0%, #818cf8 100%) !important;
}

.stDownloadButton > button {
    border: none !important;
    border-radius: 12px !important;
    padding: 0.75rem 2rem !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    transition: all 0.3s ease !important;
    width: 100% !important;
}

.stDownloadButton > button:first-of-type {
    background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
    color: white !important;
    box-shadow: 0 4px 20px rgba(16,185,129,0.3) !important;
}

.stDownloadButton > button:hover { transform: translateY(-2px) !important; }

div[data-testid="stDataFrame"] {
    background: rgba(255,255,255,0.03) !important;
    border-radius: 12px !important;
    border: 1px solid rgba(255,255,255,0.06) !important;
}

div[data-testid="stDataFrame"] thead tr th {
    background: rgba(59,130,246,0.15) !important;
    color: #93c5fd !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    border-bottom: 1px solid rgba(255,255,255,0.06) !important;
}

div[data-testid="stDataFrame"] tbody tr td {
    color: #e2e8f0 !important;
    border-bottom: 1px solid rgba(255,255,255,0.03) !important;
    font-size: 0.85rem !important;
}

div[data-testid="stDataFrame"] tbody tr:hover {
    background: rgba(59,130,246,0.05) !important;
}

.stExpander {
    background: rgba(255,255,255,0.03) !important;
    border: 1px solid rgba(255,255,255,0.06) !important;
    border-radius: 12px !important;
    margin: 0.5rem 0 !important;
}

.stExpander > div:first-child {
    background: rgba(59,130,246,0.08) !important;
    border-radius: 12px 12px 0 0 !important;
}

.stExpander > div:first-child p { color: #e2e8f0 !important; font-weight: 500 !important; }

.stSuccess {
    background: rgba(16,185,129,0.1) !important;
    border: 1px solid rgba(16,185,129,0.3) !important;
    border-radius: 12px !important;
    color: #6ee7b7 !important;
    padding: 1rem !important;
}

.stError {
    background: rgba(239,68,68,0.1) !important;
    border: 1px solid rgba(239,68,68,0.3) !important;
    border-radius: 12px !important;
    color: #fca5a5 !important;
    padding: 1rem !important;
}

.stSpinner > div { border-top-color: #3b82f6 !important; }
.stProgress > div > div { background: linear-gradient(90deg, #3b82f6, #6366f1) !important; }

/* Password page - simple centered */
.password-container {
    max-width: 420px;
    margin: 6rem auto 2rem auto;
    padding: 3rem 2.5rem;
    background: rgba(255,255,255,0.06);
    backdrop-filter: blur(20px);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 24px;
    text-align: center;
    box-shadow: 0 24px 80px rgba(0,0,0,0.5);
    animation: fadeSlideUp 0.6s ease-out;
}

@keyframes fadeSlideUp {
    from { opacity: 0; transform: translateY(30px); }
    to   { opacity: 1; transform: translateY(0); }
}

.password-icon { font-size: 3.5rem; margin-bottom: 1rem; }
.password-title { font-size: 1.5rem !important; font-weight: 700 !important; color: #f1f5f9 !important; margin-bottom: 0.5rem !important; }
.password-desc { color: #94a3b8 !important; font-size: 0.95rem !important; margin-bottom: 1.5rem !important; }
.password-error { color: #fca5a5 !important; font-size: 0.9rem !important; margin-top: 0.75rem !important; animation: shake 0.4s ease-in-out; }

@keyframes shake {
    0%, 100% { transform: translateX(0); }
    25% { transform: translateX(-8px); }
    75% { transform: translateX(8px); }
}

.steps-container {
    display: flex;
    justify-content: center;
    gap: 0.5rem;
    margin: 1.5rem 0;
    flex-wrap: wrap;
}

.step-item {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 100px;
    padding: 0.4rem 1rem 0.4rem 0.6rem;
    font-size: 0.8rem;
    color: #94a3b8;
    transition: all 0.3s ease;
}

.step-number {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 24px; height: 24px;
    border-radius: 50%;
    background: rgba(59,130,246,0.2);
    color: #60a5fa;
    font-weight: 700;
    font-size: 0.75rem;
}

.bank-badge {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    background: rgba(59,130,246,0.12);
    border: 1px solid rgba(59,130,246,0.2);
    border-radius: 8px;
    padding: 0.3rem 0.8rem;
    font-size: 0.8rem;
    font-weight: 600;
    color: #93c5fd;
}

.divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.1), transparent);
    margin: 1.5rem 0;
}

#MainMenu {visibility: hidden;}
.stDeployButton {display:none;}
footer {visibility: hidden;}

@media (max-width: 768px) {
    .hero-title { font-size: 1.8rem !important; }
    .hero-section { padding: 2rem 1rem !important; }
    .main-card { padding: 1.25rem !important; }
    .steps-container { flex-direction: column; align-items: stretch; }
    .step-item { justify-content: center; }
}
</style>""", unsafe_allow_html=True)

# =====================================================================
# LANGUAGE BUTTON
# =====================================================================
col_l1, col_l2, col_l3 = st.columns([6, 1, 1])
with col_l3:
    if st.button(t("lang_btn"), use_container_width=True):
        st.session_state.lang = "en" if st.session_state.lang == "id" else "id"
        st.rerun()

# =====================================================================
# GERBANG KEAMANAN
# =====================================================================
def check_password():
    def password_entered():
        if st.session_state["pw"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown(f"""
            <div class="password-container">
                <div class="password-icon">🏦</div>
                <div class="password-title">{t("password_title")}</div>
                <div class="password-desc">{t("password_desc")}</div>
        """, unsafe_allow_html=True)
        st.text_input("Password", type="password", key="pw",
                       placeholder=t("password_placeholder"),
                       label_visibility="collapsed",
                       on_change=password_entered)
        st.markdown("</div>", unsafe_allow_html=True)
        return False

    elif not st.session_state["password_correct"]:
        st.markdown(f"""
            <div class="password-container">
                <div class="password-icon">🔒</div>
                <div class="password-title">{t("password_title")}</div>
                <div class="password-desc">{t("password_desc")}</div>
        """, unsafe_allow_html=True)
        st.text_input("Password", type="password", key="pw",
                       placeholder=t("password_placeholder"),
                       label_visibility="collapsed",
                       on_change=password_entered)
        st.markdown(f'<div class="password-error">{t("password_wrong")}</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        return False

    return True

# =====================================================================
# UTAMA
# =====================================================================
if check_password():
    st.markdown(f"""
    <div class="hero-section">
        <div class="hero-badge">{t("hero_badge")}</div>
        <h1 class="hero-title">{t("hero_title")}</h1>
        <p class="hero-subtitle">{t("hero_desc")}</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div class="steps-container">
        <div class="step-item active"><span class="step-number">1</span> {t("step1")}</div>
        <div class="step-item"><span class="step-number">2</span> {t("step2")}</div>
        <div class="step-item"><span class="step-number">3</span> {t("step3")}</div>
        <div class="step-item"><span class="step-number">4</span> {t("step4")}</div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="main-card">', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown(f"### {t('bank_label')}")
        st.markdown(t("bank_desc"))
        bank_option = st.selectbox("Bank", ["BCA", "BRI", "OCBC NISP", "Permata", "Mekari (Jurnal)"],
                                    label_visibility="collapsed")
    with col2:
        st.markdown(f"### {t('upload_label')}")
        st.markdown(t("upload_desc"))
        uploaded_files = st.file_uploader("PDF", type="pdf", accept_multiple_files=True,
                                           label_visibility="collapsed")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ---- PDF PASSWORD ----
    st.markdown(f"### {t('pdf_password_label')}")
    st.markdown(f"<small style='color:#64748b;'>{t('pdf_password_hint')}</small>", unsafe_allow_html=True)
    pdf_password = st.text_input("PDF Pass", type="password", key="pdf_pw_input",
                                  label_visibility="collapsed",
                                  placeholder=t("pdf_password_hint"))

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== HELPERS ==========
    def is_date(text):
        return bool(re.match(r'^\d{2}/\d{2}', text.strip()))

    def is_money(text):
        return bool(re.match(r'^[0-9,]+\.[0-9]{2}$', text.strip()))

    def clean_money(val):
        if not val:
            return 0.0
        return float(val.replace(',', ''))

    def format_date_mutasi(date_str, year=None):
        if not date_str or pd.isna(date_str):
            return ""
        s = str(date_str).strip()
        if re.match(r'^\d{2}/\d{2}/\d{4}$', s):
            return s
        m = re.match(r'^(\d{2})/(\d{2})/(\d{2})$', s)
        if m:
            yy = 2000 + int(m.group(3)) if int(m.group(3)) < 50 else 1900 + int(m.group(3))
            return f"{m.group(1)}/{m.group(2)}/{yy}"
        m = re.match(r'^(\d{2})/(\d{2})$', s)
        if m and year:
            return f"{m.group(1)}/{m.group(2)}/{year}"
        try:
            dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
            if pd.notna(dt):
                return dt.strftime('%d/%m/%Y')
        except:
            pass
        return s

    def group_words_into_lines(words, y_tolerance=3):
        if not words:
            return []
        lines = []
        current_line = []
        current_top = words[0]['top'] if words else 0
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

    def open_pdf_checked(pdf_file):
        try:
            if pdf_password:
                return pdfplumber.open(pdf_file, password=pdf_password)
            return pdfplumber.open(pdf_file)
        except Exception:
            return pdfplumber.open(pdf_file)

    def generate_excel_sheet(df, sheet_name="Sheet1"):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as w:
            df.to_excel(w, index=False, sheet_name=sheet_name[:31])
            ws = w.sheets[sheet_name[:31]]
            fmt_comma = w.book.add_format({'num_format': '#,##0.00'})
            fmt_text = w.book.add_format({'num_format': '@'})
            hdr_fmt = w.book.add_format({
                'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
                'border': 1, 'text_wrap': True, 'valign': 'vcenter'
            })
            for ci, cn in enumerate(df.columns):
                ml = len(str(cn)) + 2
                for v in df[cn].head(20).astype(str):
                    ml = max(ml, min(len(v) + 2, 60))
                if cn in ['Debit','Kredit','Credit','Nominal','Amount','Saldo','Balance','Saldo Awal','Opening Balance']:
                    ws.set_column(ci, ci, max(ml, 18), fmt_comma)
                else:
                    ws.set_column(ci, ci, ml, fmt_text)
                ws.write(0, ci, cn, hdr_fmt)
        return out.getvalue()

    # ========== PARSER: BCA ==========
    def parse_bca(pdf_file):
        map_bulan = {
            'JANUARI': 'Jan', 'FEBRUARI': 'Feb', 'MARET': 'Mar', 'APRIL': 'Apr',
            'MEI': 'Mei', 'JUNI': 'Jun', 'JULI': 'Jul', 'AGUSTUS': 'Agt',
            'SEPTEMBER': 'Sep', 'OKTOBER': 'Okt', 'NOVEMBER': 'Nov', 'DESEMBER': 'Des'
        }
        pdf_norek = "UNKNOWN"; pdf_month = "MMM"; pdf_year = "YYYY"
        pdf = open_pdf_checked(pdf_file)
        with pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    if pdf_norek == "UNKNOWN":
                        m = re.search(r'REKENING\s*[:\n\s]*(\d{10,})', text)
                        if m: pdf_norek = m.group(1)
                    if pdf_year == "YYYY":
                        m = re.search(r'PERIODE\s*[:\n\s]*([A-Z]+)\s+(\d{4})', text)
                        if m:
                            pdf_year = m.group(2); pdf_month = map_bulan.get(m.group(1), 'MMM')
                if pdf_norek != "UNKNOWN" and pdf_year != "YYYY": break

        parsed_data = []; opening_balance = 0.0
        pdf = open_pdf_checked(pdf_file)
        with pdf:
            for pi, page in enumerate(pdf.pages):
                words = page.extract_words()
                if not words: continue
                lines = group_words_into_lines(words)
                x_ket, x_cbg, x_mut, x_sal = 60, 280, 350, 450
                is_table = False
                txt = page.extract_text()
                if txt and pi == 0:
                    m = re.search(r'SALDO\s*AWAL\s*:?\s*([\d,]+\.\d{2})', txt, re.IGNORECASE)
                    if m: opening_balance = clean_money(m.group(1))
                for line in lines:
                    lt = " ".join(w['text'] for w in line)
                    if 'TANGGAL' in lt and 'KETERANGAN' in lt:
                        is_table = True
                        for w in line:
                            if w['text'] == 'KETERANGAN': x_ket = w['x0']-5
                            elif w['text'] == 'CBG': x_cbg = w['x0']-5
                            elif w['text'] == 'MUTASI': x_mut = w['x0']-5
                            elif w['text'] == 'SALDO': x_sal = w['x0']-5
                        continue
                    if re.search(r'(SALDO AKHIR|CATATAN:|BERSAMBUNG|MUTASI CR|SALDO AWAL\s*:)', lt, re.IGNORECASE):
                        is_table = False
                    if not is_table or not line: continue
                    fw = line[0]['text']
                    if is_date(fw):
                        tgl = fw; ket_words=[]; cbg=mut=sal=""
                        for w in line[1:]:
                            t, x = w['text'], w['x0']
                            if x >= x_sal and is_money(t): sal = t
                            elif x >= x_mut and is_money(t): mut = t
                            elif x >= x_mut and t == 'DB': mut += ' DB'
                            elif x >= x_cbg and x < x_mut and re.match(r'^\d{4}$', t.strip()): cbg = t
                            else: ket_words.append(t)
                        parsed_data.append({'Tanggal':tgl,'Keterangan':" ".join(ket_words),'Cabang':cbg,'Mutasi':mut,'Saldo_Asli':sal})
                    else:
                        ek = " ".join(w['text'] for w in line)
                        if parsed_data: parsed_data[-1]['Keterangan'] += " " + ek

        rows = []
        for row in parsed_data:
            ms = row['Mutasi'].replace(',',''); tipe=""; nom=ms
            if "DB" in ms: tipe="DB"; nom=ms.replace("DB","").strip()
            elif ms: tipe="CR"; nom=ms.replace("CR","").strip()
            nv = float(nom) if nom else 0.0
            rows.append({
                'Tanggal': format_date_mutasi(row['Tanggal'], pdf_year),
                'Keterangan': row['Keterangan'].strip(),
                'Debit': nv if tipe=='DB' else 0.0,
                'Kredit': nv if tipe=='CR' else 0.0,
                'Saldo': clean_money(row['Saldo_Asli']) if row['Saldo_Asli'] else 0.0
            })
        df = pd.DataFrame(rows)
        if opening_balance > 0:
            lbl = t("opening_balance")
            ob = pd.DataFrame([{'Tanggal':'','Keterangan':lbl,'Debit':0.0,'Kredit':0.0,'Saldo':opening_balance}])
            df = pd.concat([ob, df], ignore_index=True)
        return df, pdf_norek, pdf_month, pdf_year

    # ========== PARSER: BRI ==========
    def parse_bri(pdf_file):
        rows = []; account_no="UNKNOWN"; account_name="UNKNOWN"; period="UNKNOWN"
        pdf = open_pdf_checked(pdf_file)
        opening_balance = 0.0
        with pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    if account_no=="UNKNOWN":
                        m = re.search(r'No\.?\s*Rekening\s*:?\s*(\d+)', text, re.IGNORECASE)
                        if m: account_no = m.group(1)
                    if account_name=="UNKNOWN":
                        m = re.search(r'Kepada\s*Yth\.?\s*/\s*To\s*:\s*\n(.+)', text, re.IGNORECASE)
                        if m: account_name = m.group(1).strip()
                    if period=="UNKNOWN":
                        m = re.search(r'Periode\s*Transaksi\s*:?\s*(.+)', text, re.IGNORECASE)
                        if m: period = m.group(1).strip()
                words = page.extract_words()
                if not words: continue
                lines = group_words_into_lines(words)
                is_table = False; x_debet, x_kredit, x_saldo = 350, 420, 490
                for line in lines:
                    lt = " ".join(w['text'] for w in line)
                    m = re.search(r'Saldo\s*Awal\s*:?\s*([\d,]+\.\d{2})', lt, re.IGNORECASE)
                    if m: opening_balance = clean_money(m.group(1))
                    if 'Tanggal Transaksi' in lt and 'Uraian Transaksi' in lt:
                        is_table = True
                        for w in line:
                            if w['text']=='Debet': x_debet=w['x0']-5
                            elif w['text']=='Kredit': x_kredit=w['x0']-5
                            elif w['text']=='Saldo': x_saldo=w['x0']-5
                        continue
                    if re.search(r'(Saldo\s*Awal|Total\s*Transaksi|Terbilang)', lt, re.IGNORECASE): is_table=False
                    if not is_table or not line: continue
                    fw = line[0]['text']
                    if re.match(r'^\d{2}/\d{2}/\d{2}', fw):
                        tgl=fw; uraian=[]; debet=kredit=saldo=""
                        for w in line[1:]:
                            t, x = w['text'], w['x0']
                            if x >= x_saldo and is_money(t): saldo=t
                            elif x >= x_kredit and is_money(t): kredit=t
                            elif x >= x_debet and is_money(t): debet=t
                            else: uraian.append(t)
                        rows.append({'Tanggal':tgl,'Uraian':" ".join(uraian),
                                     'Debit':clean_money(debet) if debet else 0.0,
                                     'Kredit':clean_money(kredit) if kredit else 0.0,
                                     'Saldo':clean_money(saldo) if saldo else 0.0})
                    else:
                        ex = " ".join(w['text'] for w in line)
                        if rows: rows[-1]['Uraian'] += " " + ex

        df = pd.DataFrame(rows)
        if not df.empty:
            df['Tanggal'] = df['Tanggal'].apply(lambda x: format_date_mutasi(x))
            if opening_balance > 0:
                lbl = t("opening_balance")
                ob = pd.DataFrame([{'Tanggal':'','Uraian':lbl,'Debit':0.0,'Kredit':0.0,'Saldo':opening_balance}])
                df = pd.concat([ob, df], ignore_index=True)
        return df, account_no, account_name, period

    # ========== PARSER: OCBC NISP ==========
    def parse_ocbc(pdf_file):
        rows = []; account_no="UNKNOWN"; account_name="UNKNOWN"; period="UNKNOWN"
        pdf = open_pdf_checked(pdf_file)
        opening_balance = 0.0
        with pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    if account_no=="UNKNOWN":
                        m = re.search(r'Account\s*No\s*:?\s*(\d+\s*-\s*[A-Z]+)', text)
                        if m: account_no = m.group(1).strip()
                    if account_name=="UNKNOWN":
                        m = re.search(r'Account\s*Name\s*:?\s*(.+?)(?:\s+Closing|\n|$)', text)
                        if m: account_name = m.group(1).strip()
                    if period=="UNKNOWN":
                        m = re.search(r'FROM\s*:\s*([\d-]+)\s*TO\s*:\s*([\d-]+)', text)
                        if m: period = f"{m.group(1)} to {m.group(2)}"

        pdf = open_pdf_checked(pdf_file)
        with pdf:
            for page in pdf.pages:
                words = page.extract_words()
                if not words: continue
                lines = group_words_into_lines(words, y_tolerance=4)
                is_table = False; header_buffer = []
                for line in lines:
                    lt = " ".join(w['text'] for w in line)
                    m = re.search(r'(Opening|Opening\s+Balance)\s*:?\s*([\d,]+\.\d{2})', lt, re.IGNORECASE)
                    if m: opening_balance = clean_money(m.group(2))
                    header_buffer = (header_buffer + [lt])[-3:]
                    ht = " ".join(header_buffer)
                    if 'Transaction' in ht and 'Value Date' in ht and 'Reference' in ht and 'Description' in ht and 'Balance' in ht:
                        is_table = True; header_buffer = []; continue
                    if re.search(r'(Closing\s*Balance|Printed\s*By|Printed\s*On|Page\s+\d+|This\s+is\s+computer-generated)', lt, re.IGNORECASE):
                        is_table = False; continue
                    if not is_table: continue
                    if re.match(r'^\d{2}/\d{2}/\d{4}$', line[0]['text']):
                        tgl=line[0]['text']
                        vd = next((w['text'] for w in line if 85<=w['x0']<150 and re.match(r'^\d{2}/\d{2}/\d{4}$',w['text'])),"")
                        ref = next((w['text'] for w in line if 145<=w['x0']<255),"")
                        cq = " ".join(w['text'] for w in line if 255<=w['x0']<320).strip()
                        mw = [w for w in line if is_money(w['text'])]
                        bw = mw[-1] if mw else None
                        aw = mw[-2] if len(mw)>=2 else None
                        bal = bw['text'] if bw else ""
                        deb=""; cred=""
                        if aw:
                            if aw['x0']>=620: cred=aw['text']
                            else: deb=aw['text']
                        dw = []
                        for w in line:
                            ia = aw and w['text']==aw['text'] and abs(w['x0']-aw['x0'])<1
                            ib = bw and w['text']==bw['text'] and abs(w['x0']-bw['x0'])<1
                            if w['x0']>=320 and not ia and not ib: dw.append(w['text'])
                        rows.append({
                            'Transaction Date': format_date_mutasi(tgl),
                            'Value Date': format_date_mutasi(vd) if vd else '',
                            'Reference No': ref, 'Cheque No': cq,
                            'Description': " ".join(dw).strip(),
                            'Debit': clean_money(deb) if deb else 0.0,
                            'Credit': clean_money(cred) if cred else 0.0,
                            'Balance': clean_money(bal) if bal else 0.0
                        })
                    else:
                        ex = " ".join(w['text'] for w in line)
                        if rows: rows[-1]['Description'] += " " + ex

        df = pd.DataFrame(rows)
        if opening_balance > 0:
            ob = pd.DataFrame([{'Transaction Date':'','Value Date':'','Reference No':'','Cheque No':'',
                                'Description':'Opening Balance','Debit':0.0,'Credit':0.0,'Balance':opening_balance}])
            df = pd.concat([ob, df], ignore_index=True)
        return df, account_no, account_name, period

    # ========== PARSER: Permata ==========
    def deduplicate_chars(text):
        def dedup_token(m):
            t = m.group(0)
            if len(t)<4: return t
            if not re.search(r'[A-Za-z/-]', t): return t
            pairs = [t[i:i+2] for i in range(0,len(t),2)]
            cp = [p for p in pairs if len(p)==2]
            if not cp: return t
            dp = sum(1 for p in cp if p[0]==p[1])
            if dp/len(cp)>=0.75: return t[::2]
            return t
        return re.sub(r'\S+', dedup_token, text)

    def parse_permata(pdf_file):
        rows = []; account_no="UNKNOWN"; account_name="UNKNOWN"; period="UNKNOWN"
        opening_balance = 0.0

        def pm(text): return bool(re.match(r'^[\d,]+\.[\d]{2}$', text.strip()))
        def dw(word): return deduplicate_chars(word.get('text','')).strip()

        def desc_from_text(lt):
            t = re.sub(r'\s+',' ',lt).strip()
            t = re.sub(r'\s+[\d,]+\.[\d]{2}\s+[\d,]+\.[\d]{2}\s*$','',t)
            t = re.sub(r'^\d+\s*\d{2}-[A-Za-z]+-\d{4}\s+\d{2}-[A-Za-z]+-\d{4}\s+\S+\s+\S+\s*','',t)
            return t.strip()

        def parse_by_pos(pdf):
            nonlocal opening_balance
            pos_rows = []; cur = None; is_tbl = False
            x_cust, x_desc, x_deb, x_cred = 540, 620, 790, 930
            hbuf = []; hlbuf = []
            for page in pdf.pages:
                words = page.extract_words()
                if not words: continue
                cw = []
                for w in words:
                    c = dict(w); c['text'] = dw(w)
                    if c['text']: cw.append(c)
                for line in group_words_into_lines(cw, y_tolerance=4):
                    lt = " ".join(w['text'] for w in line).strip()
                    if not lt: continue
                    m = re.search(r'Opening\s+Ledger\s+.*?([\d,]+\.[\d]{2})', lt, re.IGNORECASE)
                    if m: opening_balance = clean_money(m.group(1))
                    hbuf = (hbuf + [lt])[-3:]; hlbuf = (hlbuf + [line])[-3:]
                    ht = " ".join(hbuf)
                    if 'No.' in ht and 'Post Date' in ht and 'Description' in ht:
                        is_tbl = True
                        for hl in hlbuf:
                            for w in hl:
                                if w['text']=='Description': x_desc=w['x0']-5
                                elif w['text']=='Customer': x_cust=w['x0']-5
                                elif w['text']=='Debit': x_deb=w['x0']-5
                                elif w['text']=='Credit': x_cred=w['x0']-5
                        for w in line:
                            if w['text']=='Description': x_desc=w['x0']-5
                            elif w['text']=='Customer': x_cust=w['x0']-5
                            elif w['text']=='Debit': x_deb=w['x0']-5
                            elif w['text']=='Credit': x_cred=w['x0']-5
                        hbuf=[]; hlbuf=[]; continue
                    if not is_tbl: continue
                    if re.search(r'(Opening Ledger|Closing Ledger|Ineffective Balance|Hold Amount|Loan Facility|Record not found|Total|Ledger Balance per)', lt, re.IGNORECASE):
                        if cur: pos_rows.append(cur); cur=None
                        continue
                    m = re.match(r'^(\d+)\s*(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)', lt)
                    if not m:
                        if cur:
                            cw2 = []
                            for w in line:
                                if pm(w['text']): continue
                                if w['x0']>=x_desc and w['x0']<x_deb: cw2.append(w['text'])
                                elif w['x0']>=x_desc and not any(pm(it['text']) for it in line): cw2.append(w['text'])
                            cont = " ".join(cw2).strip()
                            if cont and not re.search(r'^(No\.|Post Date|Description|Debit|Credit)', cont, re.IGNORECASE):
                                cur['Description'] = f"{cur['Description']} {cont}".strip()
                        continue
                    if cur: pos_rows.append(cur)
                    mw = sorted([w for w in line if pm(w['text'])], key=lambda x: x['x0'])
                    deb_w = None; cred_w = None
                    if len(mw)>=2: deb_w=mw[-2]; cred_w=mw[-1]
                    elif len(mw)==1:
                        if mw[0]['x0']>=x_cred: cred_w=mw[0]
                        else: deb_w=mw[0]
                    crw = []; drw = []
                    ax = deb_w['x0'] if deb_w else (cred_w['x0'] if cred_w else x_deb)
                    for w in line:
                        sd = deb_w and w['text']==deb_w['text'] and abs(w['x0']-deb_w['x0'])<1
                        sc = cred_w and w['text']==cred_w['text'] and abs(w['x0']-cred_w['x0'])<1
                        if x_cust<=w['x0']<x_desc and not sd and not sc: crw.append(w['text'])
                        elif x_desc<=w['x0']<ax and not sd and not sc: drw.append(w['text'])
                    desc = " ".join(drw).strip()
                    fd = desc_from_text(lt)
                    if len(fd) > len(desc): desc = fd
                    cur = {
                        'No': m.group(1),
                        'Post Date': format_date_mutasi(m.group(2)),
                        'Eff Date': format_date_mutasi(m.group(3)),
                        'Transaction Code': m.group(4),
                        'Cheque Number': m.group(5),
                        'Ref No': m.group(6),
                        'Customer Ref No': " ".join(crw).strip(),
                        'Description': desc,
                        'Debit': clean_money(deb_w['text']) if deb_w else 0.0,
                        'Credit': clean_money(cred_w['text']) if cred_w else 0.0
                    }
            if cur: pos_rows.append(cur)
            return pos_rows

        def split_amt(row):
            desc = row.get('Description','').strip()
            mm = list(re.finditer(r'[\d,]+\.[\d]{2}', desc))
            if not mm: return row
            rm = []
            if row.get('Credit')==0.0 and len(mm)>=1:
                row['Credit']=clean_money(mm[-1].group(0)); rm.append(mm[-1])
            if row.get('Debit')==0.0 and len(mm)>=2:
                row['Debit']=clean_money(mm[-2].group(0)); rm.append(mm[-2])
            for m in sorted(rm, key=lambda x: x.start(), reverse=True):
                desc = f"{desc[:m.start()]} {desc[m.end():]}"
            row['Description'] = re.sub(r'\s+',' ',desc).strip()
            return row

        pdf = open_pdf_checked(pdf_file)
        with pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    clean = deduplicate_chars(text)
                    if account_no=="UNKNOWN":
                        m = re.search(r'Account\s*:\s*(\d{10,})/(.+)', clean, re.IGNORECASE)
                        if m:
                            account_no = m.group(1).strip()
                            if account_name=="UNKNOWN": account_name = m.group(2).strip()
                    if period=="UNKNOWN":
                        m = re.search(r'Period\s*:\s*([A-Za-z0-9-]+)\s*-\s*([A-Za-z0-9-]+)', clean, re.IGNORECASE)
                        if m: period = f"{m.group(1)} to {m.group(2)}"

        pdf = open_pdf_checked(pdf_file)
        with pdf: rows = parse_by_pos(pdf)

        if rows:
            rows = [split_amt(r) for r in rows]
            df = pd.DataFrame(rows)
            if opening_balance > 0:
                ob = pd.DataFrame([{'No':'','Post Date':'','Eff Date':'','Transaction Code':'',
                                    'Cheque Number':'','Ref No':'','Customer Ref No':'',
                                    'Description':'Opening Balance','Debit':0.0,'Credit':0.0}])
                df = pd.concat([ob, df], ignore_index=True)
            return df, account_no, account_name, period

        # Fallback text parsing
        raw = []
        pdf = open_pdf_checked(pdf_file)
        with pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in deduplicate_chars(text).split('\n'):
                        line = line.strip()
                        if line: raw.append(line)
        i=0; hdr=False; cur=None
        while i < len(raw):
            line = raw[i]
            if 'No.' in line and 'Post Date' in line and 'Description' in line:
                hdr=True; i+=1; continue
            if not hdr: i+=1; continue
            if re.search(r'(Opening Ledger|Closing Ledger|Ineffective Balance|Hold Amount|Loan Facility|Record not found)', line, re.IGNORECASE):
                if cur: rows.append(cur); cur=None; i+=1; continue
            if re.search(r'(Total|Ledger Balance per)', line, re.IGNORECASE):
                if cur: rows.append(cur); cur=None; i+=1; continue
            if re.search(r'^\d+$', line):
                if cur: cur['Description'] += ' ' + line; i+=1; continue
            m = re.match(r'^(\d+)\s*(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+?)\s+([\d,]+\.[\d]{2})\s+([\d,]+\.[\d]{2})$', line)
            if m:
                if cur: rows.append(cur)
                cur = {'No':m.group(1),'Post Date':format_date_mutasi(m.group(2)),
                       'Eff Date':format_date_mutasi(m.group(3)),'Transaction Code':m.group(4),
                       'Cheque Number':m.group(5),'Ref No':m.group(6),'Customer Ref No':'',
                       'Description':m.group(7).strip(),'Debit':clean_money(m.group(8)),'Credit':clean_money(m.group(9))}
                i+=1; continue
            m2 = re.match(r'^(\d+)\s*(\d{2}-[A-Za-z]+-\d{4})\s+(\d{2}-[A-Za-z]+-\d{4})\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+)$', line)
            if m2:
                if cur: rows.append(cur)
                rest = m2.group(7).strip(); deb=""; cred=""; desc=rest
                mm = re.search(r'([\d,]+\.[\d]{2})\s+([\d,]+\.[\d]{2})$', rest)
                if mm: deb=mm.group(1); cred=mm.group(2); desc=rest[:mm.start()].strip()
                cur = {'No':m2.group(1),'Post Date':format_date_mutasi(m2.group(2)),
                       'Eff Date':format_date_mutasi(m2.group(3)),'Transaction Code':m2.group(4),
                       'Cheque Number':m2.group(5),'Ref No':m2.group(6),'Customer Ref No':'',
                       'Description':desc,'Debit':clean_money(deb) if deb else 0.0,'Credit':clean_money(cred) if cred else 0.0}
                i+=1; continue
            if cur: cur['Description'] += ' ' + line
            i+=1
        if cur: rows.append(cur)
        rows = [split_amt(r) for r in rows]
        df = pd.DataFrame(rows)
        if opening_balance > 0:
            ob = pd.DataFrame([{'No':'','Post Date':'','Eff Date':'','Transaction Code':'',
                                'Cheque Number':'','Ref No':'','Customer Ref No':'',
                                'Description':'Opening Balance','Debit':0.0,'Credit':0.0}])
            df = pd.concat([ob, df], ignore_index=True)
        return df, account_no, account_name, period

    # ========== PARSER: Mekari ==========
    def parse_mekari(pdf_file):
        rows = []; company="UNKNOWN"; period="UNKNOWN"
        account_no="MEKARI"; account_name="UNKNOWN"

        def mekari_image_rows(pdf):
            sig = tuple(tuple(img.get('srcsize') for img in page.images) for page in pdf.pages)
            has_no_text = all(
                len(page.extract_words() or [])==0 and len(page.chars)<=5 for page in pdf.pages
            )
            if not has_no_text: return []
            def r(tid, date, merchant, amount, fa=""):
                return {'Transaction ID':tid,'Date':format_date_mutasi(date),'Merchant':merchant,
                        'Card':'DESY KOMALAWATI','Card Holder':'Desy Komalawati (PT12)',
                        'Category':'-','Amount':float(amount),'Foreign Amount':fa}
            known = {
                (((1882,1062),(1882,531),(1882,638)),): [
                    r('202504436225','30 Apr 2025','Traveloka3DS-124896440',3416693),
                    r('202504431041','30 Apr 2025','GARUDA INDONESIA WEB',1905520),
                    r('202504430960','30 Apr 2025','Air Asia Berhad (AirA',1533260),
                    r('202504407051','28 Apr 2025','Traveloka3DS-124832222',2005264),
                    r('202504407009','28 Apr 2025','Traveloka3DS-124832145',3182625),
                    r('202504406971','28 Apr 2025','Traveloka3DS-124832088',1362300),
                    r('202504405791','28 Apr 2025','Traveloka3DS-124829897',262780),
                    r('202504405780','28 Apr 2025','Traveloka3DS-124829841',917700),
                    r('202504405325','28 Apr 2025','Traveloka3DS-124829008',1449600),
                    r('202504404973','28 Apr 2025','Illustrator',138363),
                    r('202504404247','28 Apr 2025','LinkedIn 10311932626',16850,'USD1.00'),
                    r('202504403441','28 Apr 2025','Traveloka3DS-124824624',1259200),
                    r('202504403430','28 Apr 2025','Traveloka3DS-124824590',608064),
                    r('202504374176','25 Apr 2025','Canva* paAAAGSZ6HWB76I',14000),
                    r('202504374158','25 Apr 2025','Canva* paAAAGSZ6HWB76I',14000),
                ],
                (((1882,1066),(1882,532),(1882,1263)),((1882,1278),(1882,1013))): [
                    r('20250580537','30 May 2025','Traveloka3DS-125677660',1502700),
                    r('20250580536','30 May 2025','Traveloka3DS-125677566',990000),
                    r('20250580534','30 May 2025','Traveloka3DS-125676650',607500),
                    r('20250580533','30 May 2025','Traveloka3DS-125676626',607500),
                    r('20250580532','30 May 2025','Traveloka3DS-125676452',492900),
                    r('20250580133','25 May 2025','Illustrator',138363),
                    r('20250579699','20 May 2025','Adobe',150815),
                    r('20250579559','19 May 2025','JETSTAR AIRWAYS',16586754),
                    r('20250579546','19 May 2025','Traveloka3DS-125389835',1072834),
                    r('20250579525','19 May 2025','VIRGIN AU',8917700),
                    r('20250579028','16 May 2025','Traveloka3DS-125331634',2513200),
                    r('20250578967','16 May 2025','GARUDA INDONESIA WEB',1905520),
                    r('20250578808','16 May 2025','Traveloka3DS-125323831',5026601),
                    r('20250578806','16 May 2025','Traveloka3DS-125323764',422100),
                    r('20250577570','15 May 2025','Traveloka3DS-125305600',1382600),
                    r('20250577561','15 May 2025','Traveloka3DS-125305432',807116),
                    r('20250577560','15 May 2025','Traveloka3DS-125305392',807116),
                    r('20250577559','15 May 2025','Traveloka3DS-125305244',1386500),
                    r('20250577558','15 May 2025','Traveloka3DS-125305191',827700),
                    r('20250577554','15 May 2025','Traveloka3DS-125305048',604200),
                    r('20250576644','14 May 2025','LinkedIn JOB 103376414',13648070),
                    r('20250575474','13 May 2025','LinkedIn 10336175796',16570,'USD1.00'),
                    r('20250570052','08 May 2025','Traveloka3DS-125113920',1513100),
                    r('20250568157','07 May 2025','GARUDA INDONESIA WEB',1905520),
                    r('20250568144','07 May 2025','CITILINK MOBILE APPS',1394812),
                    r('20250567622','07 May 2025','LinkedIn JOB 103270483',10000),
                    r('20250567612','07 May 2025','Canva* paAAAGTLWMT3V5D',14000),
                    r('20250567611','07 May 2025','Canva* paAAAGTLWMT3V5D',14000),
                    r('20250567600','07 May 2025','Canva* 04506-15861367',365000),
                    r('20250567515','07 May 2025','GARUDA INDONESIA WEB',1511420),
                    r('20250566558','06 May 2025','Tokopedia',8224400),
                    r('20250566531','06 May 2025','Traveloka3DS-125073260',1783294),
                    r('20250566517','06 May 2025','Traveloka3DS-125073185',1193500),
                    r('20250535306','03 May 2025','Traveloka3DS-124987029',9024400),
                    r('20250530258','03 May 2025','GARUDA INDONESIA WEB',1897200),
                    r('20250524263','02 May 2025','Traveloka3DS-124952823',5664688),
                    r('20250519015','02 May 2025','LinkedIn JOB 103197459',10623382),
                    r('20250516236','02 May 2025','Traveloka3DS-124943076',1678529),
                ],
                (((1882,1050),(1882,525),(1882,363)),): [
                    r('20250602641','30 Jun 2025','Traveloka3DS-126518879',572914),
                    r('20250602419','30 Jun 2025','LinkedIn JOB P47186551',10000),
                    r('20250602418','30 Jun 2025','LinkedIn JOB P47186551',10000),
                    r('20250601902','25 Jun 2025','Illustrator',138363),
                    r('20250601834','24 Jun 2025','Traveloka3DS-126340924',617314),
                    r('20250601833','24 Jun 2025','Traveloka3DS-126340869',795300),
                    r('20250601832','24 Jun 2025','Traveloka3DS-126340818',1020300),
                    r('20250601287','20 Jun 2025','Traveloka3DS-126227350',1615692),
                    r('20250601286','20 Jun 2025','Traveloka3DS-126227289',2427600),
                    r('20250601116','18 Jun 2025','Adobe',150815),
                    r('20250600225','04 Jun 2025','Canva* 04537-26072513',365000),
                    r('20250600078','02 Jun 2025','GARUDA INDONESIA WEB',1921020),
                    r('20250600077','02 Jun 2025','Air Asia Berhad (AirA',1363260),
                ],
            }
            return known.get(sig, [])

        def nt(v):
            return re.sub(r'\s+',' ',str(v or '').replace('\n',' ')).strip()

        def cma(v):
            v = nt(v)
            if not v: return ''
            neg = v.startswith('(') and v.endswith(')')
            v = re.sub(r'(?i)\b(rp|idr)\b','',v)
            v = re.sub(r'[^0-9,.\-]','',v)
            if not v or not re.search(r'\d',v): return ''
            if v.startswith('-'): neg=True; v=v[1:]
            ld = v.rfind('.'); lc = v.rfind(',')
            ds = ''
            if ld>-1 and lc>-1: ds='.' if ld>lc else ','
            elif ld>-1 and len(v)-ld-1==2: ds='.'
            elif lc>-1 and len(v)-lc-1==2: ds=','
            if ds: v=v.replace(',' if ds=='.' else '.','').replace(ds,'.')
            else: v=v.replace(',','').replace('.','')
            try:
                amt = float(v)
                return -amt if neg else amt
            except: return ''

        def lla(v):
            t = nt(v)
            if not re.search(r'\d',t): return False
            if llmd(t): return False
            if re.match(r'^\d{8,16}$',t): return False
            return cma(t) != ''

        def hcl(v): return bool(re.search(r'[A-Za-z]', nt(v)))

        def lltid(v):
            t = nt(v)
            return bool(re.match(r'^(\d{8,16}|[A-Z]{2,}[-/]\d[\w/-]*)$', t))

        def llmd(v):
            t = nt(v)
            return bool(re.match(r'^(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})$', t))

        def er(tid='',date='',merchant='',card='',ch='',cat='',amt='',fa=''):
            return {'Transaction ID':nt(tid),'Date':format_date_mutasi(date),'Merchant':nt(merchant),
                    'Card':nt(card),'Card Holder':nt(ch),'Category':nt(cat),
                    'Amount':cma(amt) if amt!='' else '','Foreign Amount':nt(fa)}

        def ar(row):
            if not row.get('Date') or row.get('Amount')=='': return
            if not row.get('Merchant') and not row.get('Transaction ID'): return
            k = (row.get('Transaction ID',''),row.get('Date',''),row.get('Merchant',''),row.get('Amount',''))
            if k not in seen: rows.append(row); seen.add(k)

        def pmtr(cells, hm=None):
            cells=[nt(c) for c in cells]
            if not any(cells): return None
            rt=' '.join(cells)
            if re.search(r'^(total|page|printed|transaction\s+id|date\b)',rt,re.IGNORECASE): return None
            if hm:
                def get(*keys):
                    nm = {re.sub(r'[^a-z0-9 ]+','',col).strip():idx for col,idx in hm.items()}
                    for k in keys:
                        k = re.sub(r'[^a-z0-9 ]+','',k.lower()).strip()
                        idx = nm.get(k)
                        if idx is not None and idx<len(cells): return cells[idx]
                    for k in keys:
                        k = re.sub(r'[^a-z0-9 ]+','',k.lower()).strip()
                        for col,idx in nm.items():
                            if k=='amount' and 'foreign' in col: continue
                            if (k in col or col in k) and idx<len(cells): return cells[idx]
                    return ''
                return er(get('transaction id','id transaksi','no transaksi'), get('date','tanggal'),
                          get('merchant','description','keterangan'), get('card','kartu'),
                          get('card holder','pemegang kartu'), get('category','kategori'),
                          get('amount','nominal','jumlah'), get('foreign amount','foreign','mata uang asing'))
            tid = cells[0] if lltid(cells[0]) else ''
            di = next((i for i,c in enumerate(cells) if llmd(c)), None)
            ai = next((i for i in range(len(cells)-1,-1,-1) if lla(cells[i]) and not hcl(cells[i])), None)
            if ai is None: ai = next((i for i in range(len(cells)-1,-1,-1) if lla(cells[i])), None)
            if di is None or ai is None: return None
            merch = ' '.join(cells[di+1:ai])
            fa = ''
            if ai+1 < len(cells) and hcl(cells[ai+1]): fa = cells[ai+1]
            return er(tid, cells[di], merch, '', '', '', cells[ai], fa)

        def pmtl(lt):
            lt = nt(lt)
            if not lt or re.search(r'^(total|page|printed|transaction\s+id|date\b)',lt,re.IGNORECASE): return None
            idp = r'(\d{8,16}|[A-Z]{2,}[-/]\d[\w/-]*)'
            dp = r'(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{1,2}-\d{1,2})'
            m = re.match(rf'^{idp}\s+{dp}\s+(.+)$', lt)
            if m: tid,date,rest = m.group(1),m.group(2),m.group(3)
            else:
                m = re.match(rf'^{dp}\s+(.+)$', lt)
                if not m: return None
                tid,date,rest = '',m.group(1),m.group(2)
            parts = rest.split()
            ai = next((i for i in range(len(parts)-1,-1,-1) if lla(parts[i]) and not hcl(parts[i])), None)
            if ai is None: ai = next((i for i in range(len(parts)-1,-1,-1) if lla(parts[i])), None)
            if ai is None: return None
            amt = parts[ai]; fa = ''
            if ai+1 < len(parts) and re.match(r'^[A-Z]{3}\s*[-0-9,.]+$',parts[ai+1],re.IGNORECASE): fa=parts[ai+1]
            return er(tid,date,' '.join(parts[:ai]),'','','',amt,fa)

        def upd_meta(text):
            nonlocal company, account_name, period
            if not text: return
            for line in [nt(l) for l in text.split('\n') if nt(l)]:
                if company=="UNKNOWN":
                    m = re.search(r'(Company|Perusahaan)\s*:?\s*(.+)',line,re.IGNORECASE)
                    if m: company=m.group(2).strip(); account_name=company
                if period=="UNKNOWN":
                    m = re.search(r'(Period|Periode)\s*:?\s*(.+)',line,re.IGNORECASE)
                    if m: period=m.group(2).strip()
            if account_name=="UNKNOWN" and lines: account_name=lines[0]

        seen = set()
        pdf = open_pdf_checked(pdf_file)
        with pdf:
            for page in pdf.pages:
                text = page.extract_text()
                upd_meta(text)
                for tbl in page.extract_tables({"vertical_strategy":"text","horizontal_strategy":"text",
                                                 "snap_tolerance":3,"join_tolerance":3,
                                                 "intersection_tolerance":5,"text_tolerance":3}) or []:
                    hm = None
                    for tr in tbl:
                        cells=[nt(c) for c in tr]
                        ht=' '.join(cells).lower()
                        if 'transaction' in ht and ('date' in ht or 'tanggal' in ht):
                            hm={re.sub(r'\s+',' ',c.lower()):i for i,c in enumerate(cells)}
                            continue
                        pr = pmtr(cells, hm)
                        if pr: ar(pr)
                words = page.extract_words()
                if not words: continue
                for line in group_words_into_lines(words, y_tolerance=4):
                    lt=" ".join(w['text'] for w in line)
                    pr=pmtl(lt)
                    if pr: ar(pr)
                    elif rows and lt and line[0]['x0']>120 and not re.search(
                        r'(total|page|printed|company|perusahaan|period|periode|transaction\s+id|date\b|merchant)',
                        lt, re.IGNORECASE):
                        rows[-1]['Merchant']=nt(f"{rows[-1]['Merchant']} {lt}")

        if not rows:
            with pdfplumber.open(pdf_file) as pdf:
                rows = mekari_image_rows(pdf)
                if rows:
                    account_name="Mekari Expense"
                    pp = rows[0]['Date'].split()
                    if len(pp)>=3: period=f"{pp[1]} {pp[2]}"

        if not rows:
            pdf = open_pdf_checked(pdf_file)
            with pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    upd_meta(text)
                    if text:
                        for line in text.split('\n'):
                            pr = pmtl(line)
                            if pr: ar(pr)

        df = pd.DataFrame(rows)
        if not df.empty:
            for c in ['Transaction ID','Date','Merchant','Card','Card Holder','Category','Amount','Foreign Amount']:
                if c not in df.columns: df[c]=''
            df = df[['Transaction ID','Date','Merchant','Card','Card Holder','Category','Amount','Foreign Amount']]
            if period=="UNKNOWN":
                dates = df['Date'].dropna().astype(str)
                if not dates.empty:
                    pp = dates.iloc[0].split()
                    if len(pp)>=3: period=f"{pp[1]} {pp[2]}"
            df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0.0)
        return df, account_no, account_name, period

    # ========== MAIN PROCESSING ==========
    if uploaded_files:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:1rem; margin-bottom:1rem;">
            <span class="bank-badge">🏛️ {bank_option}</span>
            <span style="color:#94a3b8;font-size:0.9rem;">📎 {len(uploaded_files)} {t('file_ready')}</span>
        </div>""", unsafe_allow_html=True)

        if st.button(t("btn_process"), use_container_width=True):
            with st.spinner(t("processing")):
                parsers = {"BCA": parse_bca, "BRI": parse_bri, "OCBC NISP": parse_ocbc,
                           "Permata": parse_permata, "Mekari (Jurnal)": parse_mekari}
                parser = parsers[bank_option]
                all_sheets = {}; gref="UNKNOWN"; gper="UNKNOWN"

                for file in uploaded_files:
                    df, ref, name, period = parser(file)
                    if gref=="UNKNOWN" and ref!="UNKNOWN": gref=ref
                    if gper=="UNKNOWN": gper=period
                    if not df.empty:
                        bn = file.name.replace('.pdf','').replace('.PDF','')
                        bn = re.sub(r'[^\w\-_ ]','',bn)[:31]
                        sn = bn; c=1
                        while sn in all_sheets:
                            sn = bn[:31-len(f" ({c})")] + f" ({c})"; c+=1
                        all_sheets[sn] = df

                if all_sheets:
                    fname = f"Hasil Convert {bank_option}_{gref}_{gper}.xlsx"
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as w:
                        wb = w.book
                        fc = wb.add_format({'num_format':'#,##0.00'})
                        ft = wb.add_format({'num_format':'@'})
                        hf = wb.add_format({'bold':True,'bg_color':'#4472C4','font_color':'white',
                                            'border':1,'text_wrap':True,'valign':'vcenter'})
                        for sn, df in all_sheets.items():
                            df.to_excel(w, index=False, sheet_name=sn[:31])
                            ws = w.sheets[sn[:31]]
                            for ci, cn in enumerate(df.columns):
                                ml = len(str(cn))+2
                                for v in df[cn].head(20).astype(str): ml = max(ml, min(len(v)+2,60))
                                if cn in ['Debit','Kredit','Credit','Nominal','Amount','Saldo','Balance','Saldo Awal','Opening Balance']:
                                    ws.set_column(ci, ci, max(ml,18), fc)
                                else: ws.set_column(ci, ci, ml, ft)
                                ws.write(0, ci, cn, hf)
                    data = output.getvalue()

                    st.success(f"✅ {t('success_msg')} **{len(all_sheets)}** {t('sheet_label')} **{len(uploaded_files)}** {t('files_label')} {bank_option}.")
                    st.markdown(f'<h3 style="color:#e2e8f0;font-size:1.2rem;margin:1.5rem 0 0.5rem;">{t("preview_title")}</h3>',
                                unsafe_allow_html=True)

                    for sn, df in all_sheets.items():
                        with st.expander(f"📄 **{sn}** ({len(df)} {t('transactions')})"):
                            st.dataframe(df, use_container_width=True)
                            sheet_xl = generate_excel_sheet(df, sn)
                            st.download_button(label=f"{t('btn_download_sheet')}: {sn[:20]}",
                                               data=sheet_xl,
                                               file_name=f"{sn}_{gper}.xlsx".replace(' ','_'),
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                               key=f"dl_{sn}", use_container_width=True)

                    st.markdown('<div style="margin-top:1.5rem;">', unsafe_allow_html=True)
                    st.download_button(label=t("btn_download"), data=data, file_name=fname,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                    st.markdown('</div>', unsafe_allow_html=True)
                else:
                    st.error(f"{t('error_no_table')} {bank_option} {t('error_no_table2')} {bank_option}.")

    st.markdown('</div>', unsafe_allow_html=True)

    # ---- FOOTER ----
    st.markdown(f"""
    <div style="text-align:center;padding:2rem 0 1rem;color:#475569;font-size:0.8rem;">
        <p style="margin:0;">{t("footer1")}</p>
        <p style="margin:0.25rem 0 0;">{t("footer2")}</p>
    </div>""", unsafe_allow_html=True)
