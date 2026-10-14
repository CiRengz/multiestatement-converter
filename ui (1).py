"""
ui.py - Tema futuristik & komponen tampilan untuk aplikasi e-Statement Converter.
Letakkan di folder yang sama dengan app.py, lalu `import ui`.
Tidak mengubah logika parser sama sekali.
"""
import html
import pandas as pd
import streamlit as st

BANKS = ["BCA", "BRI", "BNI", "Mandiri Kopra", "OCBC NISP", "Permata", "Mekari (Jurnal)"]


def _h(s: str) -> str:
    """Ratakan HTML jadi satu baris (menghindari markdown menganggap indentasi sebagai code block)."""
    return " ".join(line.strip() for line in s.strip().splitlines())


CSS = """
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@500;700;900&family=Space+Grotesk:wght@400;500;600&family=JetBrains+Mono:wght@500;700&display=swap');

:root{
  --cyan:#00e5ff; --violet:#7c3aed; --pink:#ff3df5; --bg:#05070f;
  --panel:rgba(12,18,36,.72); --line:rgba(0,229,255,.22); --text:#e6edf7; --muted:#8b9ab8;
}

/* ---------- Latar: grid animasi + aurora ---------- */
.stApp{
  font-family:'Space Grotesk',sans-serif; color:var(--text);
  background:
    linear-gradient(rgba(0,229,255,.045) 1px,transparent 1px) 0 0/48px 48px,
    linear-gradient(90deg,rgba(0,229,255,.045) 1px,transparent 1px) 0 0/48px 48px,
    radial-gradient(1100px 600px at 8% -10%,rgba(124,58,237,.32),transparent 60%),
    radial-gradient(900px 520px at 100% 0%,rgba(0,229,255,.20),transparent 55%),
    radial-gradient(800px 500px at 50% 120%,rgba(255,61,245,.14),transparent 60%),
    var(--bg);
  background-attachment:fixed;
  animation:gridmove 8s linear infinite;
}
@keyframes gridmove{to{background-position:0 48px,48px 0,0 0,0 0,0 0,0 0}}

[data-testid="stHeader"]{background:transparent}
footer{visibility:hidden}
.block-container{max-width:1080px;padding-top:2.2rem;padding-bottom:4rem}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-thumb{background:linear-gradient(var(--cyan),var(--violet));border-radius:8px}
::-webkit-scrollbar-track{background:transparent}

/* ---------- Hero ---------- */
.hero{text-align:center;padding:1.2rem 0 .6rem}
.badge{display:inline-flex;align-items:center;gap:8px;padding:6px 14px;border-radius:999px;
  border:1px solid var(--line);background:rgba(0,229,255,.07);
  font:600 .72rem 'JetBrains Mono',monospace;letter-spacing:.18em;color:var(--cyan)}
.dot{width:8px;height:8px;border-radius:50%;background:#2cff9a;box-shadow:0 0 0 0 rgba(44,255,154,.7);animation:pulse 1.8s infinite}
@keyframes pulse{70%{box-shadow:0 0 0 10px rgba(44,255,154,0)}100%{box-shadow:0 0 0 0 rgba(44,255,154,0)}}
.hero-title{font-family:'Orbitron',sans-serif;font-weight:900;margin:.9rem 0 .4rem;
  font-size:clamp(1.9rem,5vw,3.4rem);letter-spacing:.04em;line-height:1.1;color:#fff}
.hero-title span{background:linear-gradient(90deg,var(--cyan),var(--violet) 55%,var(--pink));
  -webkit-background-clip:text;background-clip:text;color:transparent;
  filter:drop-shadow(0 0 18px rgba(0,229,255,.35))}
.hero-sub{max-width:640px;margin:0 auto;color:var(--muted);font-size:1.02rem;line-height:1.6}
.chips{display:flex;flex-wrap:wrap;justify-content:center;gap:8px;margin-top:1.2rem}
.chip{padding:6px 14px;border-radius:10px;font:500 .8rem 'JetBrains Mono',monospace;color:var(--text);
  background:rgba(124,58,237,.12);border:1px solid rgba(124,58,237,.35);transition:.25s}
.chip:hover{transform:translateY(-2px);border-color:var(--cyan);box-shadow:0 0 16px rgba(0,229,255,.35)}
.hero-line{height:2px;margin:1.6rem auto 0;max-width:520px;border-radius:2px;
  background:linear-gradient(90deg,transparent,var(--cyan),var(--violet),transparent);
  background-size:200% 100%;animation:sweep 3s linear infinite}
@keyframes sweep{to{background-position:-200% 0}}

/* ---------- Judul langkah ---------- */
.step{display:flex;align-items:center;gap:14px;margin:2.3rem 0 .7rem}
.step-num{width:40px;height:40px;flex:none;border-radius:11px;display:grid;place-items:center;
  font:900 1rem 'Orbitron',sans-serif;color:#05070f;
  background:linear-gradient(135deg,var(--cyan),var(--violet));box-shadow:0 0 20px rgba(0,229,255,.45)}
.step-title{font:700 1.15rem 'Orbitron',sans-serif;letter-spacing:.05em;color:#fff;margin:0}
.step-sub{font-size:.88rem;color:var(--muted);margin:2px 0 0}

/* ---------- Widget ---------- */
[data-baseweb="select"]>div,[data-baseweb="input"]{
  background:var(--panel)!important;border:1px solid var(--line)!important;border-radius:12px!important;
  backdrop-filter:blur(10px);transition:.25s}
[data-baseweb="select"]>div:hover,[data-baseweb="input"]:focus-within{
  border-color:var(--cyan)!important;box-shadow:0 0 0 3px rgba(0,229,255,.12),0 0 22px rgba(0,229,255,.18)}
[data-testid="stTextInput"] input{color:var(--text)!important;letter-spacing:.2em}

[data-testid="stFileUploaderDropzone"],[data-testid="stFileUploader"] section{
  border:1.6px dashed rgba(0,229,255,.5)!important;border-radius:18px!important;
  background:linear-gradient(160deg,rgba(0,229,255,.06),rgba(124,58,237,.06))!important;
  padding:1.6rem!important;transition:.3s}
[data-testid="stFileUploaderDropzone"]:hover,[data-testid="stFileUploader"] section:hover{
  border-color:var(--cyan)!important;box-shadow:0 0 34px rgba(0,229,255,.22) inset,0 0 26px rgba(0,229,255,.18)}
[data-testid="stFileUploaderFile"]{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:2px 8px}

.stButton>button,[data-testid="stDownloadButton"]>button,[data-testid="stBaseButton-secondary"],[data-testid="stBaseButton-primary"]{
  width:100%;border:0!important;border-radius:14px!important;padding:.85rem 1.2rem!important;
  background:linear-gradient(120deg,var(--cyan),var(--violet) 60%,var(--pink))!important;
  background-size:180% 100%!important;color:#05070f!important;
  font:800 .95rem 'Orbitron',sans-serif!important;letter-spacing:.12em;text-transform:uppercase;
  box-shadow:0 8px 30px rgba(124,58,237,.4),0 0 0 1px rgba(255,255,255,.08) inset;transition:.3s}
.stButton>button:hover,[data-testid="stDownloadButton"]>button:hover{
  background-position:100% 0!important;transform:translateY(-2px);
  box-shadow:0 12px 40px rgba(0,229,255,.45),0 0 0 1px rgba(255,255,255,.2) inset}
.stButton>button p,[data-testid="stDownloadButton"]>button p{color:#05070f!important;font-weight:800!important}

[data-testid="stAlert"]{background:var(--panel)!important;border:1px solid var(--line)!important;
  border-radius:14px!important;backdrop-filter:blur(10px)}
[data-testid="stExpander"]{border:1px solid var(--line)!important;border-radius:16px!important;
  background:var(--panel);backdrop-filter:blur(10px);overflow:hidden;margin-bottom:.6rem}
[data-testid="stExpander"] summary:hover{color:var(--cyan)}
[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:12px;overflow:hidden}
[data-testid="stSpinner"]{color:var(--cyan)}

/* ---------- Kartu statistik ---------- */
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px;margin:.8rem 0 1.4rem}
.stat{position:relative;overflow:hidden;padding:18px 20px;border-radius:16px;
  background:linear-gradient(160deg,rgba(22,32,64,.8),rgba(10,15,32,.8));border:1px solid var(--line);
  transition:.3s}
.stat:hover{transform:translateY(-3px);border-color:var(--cyan);box-shadow:0 10px 34px rgba(0,229,255,.18)}
.stat::before{content:"";position:absolute;left:0;top:0;height:2px;width:100%;
  background:linear-gradient(90deg,var(--cyan),var(--violet),var(--pink))}
.stat-label{font:600 .7rem 'JetBrains Mono',monospace;letter-spacing:.16em;text-transform:uppercase;color:var(--muted)}
.stat-value{font:700 1.45rem 'JetBrains Mono',monospace;color:#fff;margin:6px 0 2px;word-break:break-word}
.stat-sub{font-size:.78rem;color:var(--cyan)}

/* ---------- Empty state & login ---------- */
.empty{padding:1.6rem;text-align:center;border-radius:16px;color:var(--muted);
  border:1px dashed rgba(139,154,184,.35);background:rgba(12,18,36,.4)}
.empty b{color:var(--cyan)}
.lock{width:74px;height:74px;margin:0 auto 10px;border-radius:20px;display:grid;place-items:center;font-size:2rem;
  background:linear-gradient(135deg,rgba(0,229,255,.18),rgba(124,58,237,.25));border:1px solid var(--line);
  box-shadow:0 0 40px rgba(0,229,255,.3);animation:float 3.5s ease-in-out infinite}
@keyframes float{50%{transform:translateY(-6px)}}
.foot{text-align:center;margin-top:3rem;font:500 .72rem 'JetBrains Mono',monospace;letter-spacing:.2em;color:#5d6b88}
"""


def inject_css():
    st.markdown("<style>" + _h(CSS) + "</style>", unsafe_allow_html=True)


def hero():
    chips = "".join(f'<span class="chip">{html.escape(b)}</span>' for b in BANKS)
    st.markdown(_h(f"""
    <div class="hero">
      <div class="badge"><span class="dot"></span>SYSTEM ONLINE</div>
      <h1 class="hero-title">e-STATEMENT <span>CONVERTER</span></h1>
      <p class="hero-sub">Ubah PDF rekening koran multi-bank menjadi Excel yang rapi dalam hitungan detik.
      Pilih bank, unggah file, dan biarkan sistem mengekstrak transaksinya secara otomatis.</p>
      <div class="chips">{chips}</div>
      <div class="hero-line"></div>
    </div>"""), unsafe_allow_html=True)


def login_header():
    st.markdown(_h("""
    <div class="hero" style="padding-top:3rem">
      <div class="lock">🔒</div>
      <h1 class="hero-title" style="font-size:clamp(1.5rem,4vw,2.4rem)">AKSES <span>TERBATAS</span></h1>
      <p class="hero-sub">Masukkan kata sandi untuk membuka e-Statement Converter.</p>
      <div class="hero-line" style="max-width:300px"></div>
    </div>"""), unsafe_allow_html=True)


def step_header(num: int, title: str, sub: str = ""):
    st.markdown(_h(f"""
    <div class="step"><div class="step-num">{num}</div>
      <div><p class="step-title">{html.escape(title)}</p><p class="step-sub">{html.escape(sub)}</p></div>
    </div>"""), unsafe_allow_html=True)


def empty_state(msg: str):
    st.markdown(f'<div class="empty">{msg}</div>', unsafe_allow_html=True)


def _idr(x: float) -> str:
    return "Rp " + f"{x:,.0f}".replace(",", ".")


def summary_cards(all_sheets: dict, n_files: int):
    """Kartu ringkasan hasil ekstraksi (otomatis menyesuaikan kolom tiap bank)."""
    dfs = list(all_sheets.values())
    tx = sum(len(d) for d in dfs)

    def tot(names):
        s, found = 0.0, False
        for d in dfs:
            for n in names:
                if n in d.columns:
                    s += pd.to_numeric(d[n], errors="coerce").fillna(0).sum()
                    found = True
        return s, found

    deb, fd = tot(["Debit"])
    kre, fk = tot(["Kredit", "Credit"])
    amt, fa = tot(["Amount"])

    items = [
        ("File diproses", str(n_files), "dokumen PDF"),
        ("Total transaksi", f"{tx:,}".replace(",", "."), "baris data"),
    ]
    if fd:
        items.append(("Total debit", _idr(deb), "dana keluar"))
    if fk:
        items.append(("Total kredit", _idr(kre), "dana masuk"))
    if fa and not (fd or fk):
        items.append(("Total nominal", _idr(amt), "seluruh transaksi"))

    cards = "".join(
        f'<div class="stat"><div class="stat-label">{html.escape(l)}</div>'
        f'<div class="stat-value">{html.escape(v)}</div><div class="stat-sub">{html.escape(s)}</div></div>'
        for l, v, s in items
    )
    st.markdown(f'<div class="stats">{cards}</div>', unsafe_allow_html=True)


def footer():
    st.markdown('<div class="foot">E-STATEMENT CONVERTER · v2.0</div>', unsafe_allow_html=True)
