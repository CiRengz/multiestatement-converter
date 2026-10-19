"""
ui.py - Sistem tema + komponen tampilan untuk e-Statement Converter.
5 tema bisa dipilih lewat dropdown (ui.theme_picker()). Tidak mengubah logika parser.
"""
import html
import pandas as pd
import streamlit as st

BANKS = ["BCA", "BRI", "BNI", "Mandiri Kopra", "OCBC NISP", "Permata", "Mekari (Jurnal)"]
DEFAULT_THEME = "neon"


def _h(s: str) -> str:
    """Ratakan HTML jadi satu baris (supaya markdown tidak menganggap indentasi sebagai code block)."""
    return " ".join(line.strip() for line in s.strip().splitlines())


# ======================================================================
#  DEFINISI TEMA  (variabel CSS + CSS tambahan khusus tema)
# ======================================================================
DEFAULTS = {
    "bw": "1px", "shadow-card": "none", "btn-border": "0", "btn-transform": "uppercase",
    "btn-tracking": ".12em", "btn-weight": "800", "head-tracking": ".05em", "head-transform": "none",
    "title-filter": "none", "df-filter": "none", "df-border": "1px solid var(--line)", "blur": "10px",
    "bg-anim": "none", "focus": "0 0 0 3px rgba(0,229,255,.12)", "num-border": "0", "num-shadow": "none",
    "chip-hover": "none", "glow": "none", "empty-border": "rgba(139,154,184,.4)", "stat-bar-h": "2px",
    "ok": "#2cff9a", "badge-radius": "999px",
}

THEMES = {
    # ---------------------------------------------------------------- 1
    "neon": {
        "label": "⚡ Neon Cyber",
        "fonts": "family=Orbitron:wght@500;700;900&family=Space+Grotesk:wght@400;500;600&family=JetBrains+Mono:wght@500;700",
        "vars": {
            "bg": "linear-gradient(rgba(0,229,255,.045) 1px,transparent 1px) 0 0/48px 48px,"
                  "linear-gradient(90deg,rgba(0,229,255,.045) 1px,transparent 1px) 0 0/48px 48px,"
                  "radial-gradient(1100px 600px at 8% -10%,rgba(124,58,237,.32),transparent 60%),"
                  "radial-gradient(900px 520px at 100% 0%,rgba(0,229,255,.20),transparent 55%),"
                  "radial-gradient(800px 500px at 50% 120%,rgba(255,61,245,.14),transparent 60%),#05070f",
            "bg-anim": "gridmove 8s linear infinite",
            "panel": "rgba(12,18,36,.72)", "solid": "#0c1224", "input-bg": "rgba(12,18,36,.72)",
            "line": "rgba(0,229,255,.22)", "line-strong": "#00e5ff",
            "text": "#e6edf7", "title": "#ffffff", "muted": "#8b9ab8",
            "a1": "#00e5ff", "a2": "#7c3aed", "a3": "#ff3df5",
            "grad": "linear-gradient(120deg,#00e5ff,#7c3aed 60%,#ff3df5)",
            "grad-text": "linear-gradient(90deg,#00e5ff,#7c3aed 55%,#ff3df5)",
            "title-filter": "drop-shadow(0 0 18px rgba(0,229,255,.35))",
            "btn-text": "#05070f", "btn-shadow": "0 8px 30px rgba(124,58,237,.4)",
            "btn-shadow-hover": "0 12px 40px rgba(0,229,255,.45)",
            "radius": "16px", "radius-sm": "12px",
            "font-head": "'Orbitron',sans-serif", "font-body": "'Space Grotesk'", "font-mono": "'JetBrains Mono',monospace",
            "badge-bg": "rgba(0,229,255,.07)", "badge-text": "#00e5ff",
            "chip-bg": "rgba(124,58,237,.12)", "chip-border": "rgba(124,58,237,.35)",
            "chip-hover": "0 0 16px rgba(0,229,255,.35)",
            "stat-bg": "linear-gradient(160deg,rgba(22,32,64,.8),rgba(10,15,32,.8))",
            "drop-bg": "linear-gradient(160deg,rgba(0,229,255,.06),rgba(124,58,237,.06))",
            "drop-border": "1.6px dashed rgba(0,229,255,.5)",
            "num-shadow": "0 0 20px rgba(0,229,255,.45)", "glow": "0 0 40px rgba(0,229,255,.3)",
        },
        "extra": "",
    },
    # ---------------------------------------------------------------- 2
    "aurora": {
        "label": "🌌 Aurora Glass",
        "fonts": "family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700",
        "vars": {
            "bg": "radial-gradient(circle at 20% 20%,rgba(99,102,241,.55),transparent 45%) 0% 0%/140% 140% no-repeat,"
                  "radial-gradient(circle at 80% 15%,rgba(236,72,153,.40),transparent 45%) 100% 0%/140% 140% no-repeat,"
                  "radial-gradient(circle at 50% 90%,rgba(20,184,166,.40),transparent 50%) 50% 100%/140% 140% no-repeat,#070816",
            "bg-anim": "drift 18s ease-in-out infinite",
            "panel": "rgba(255,255,255,.07)", "solid": "#171934", "input-bg": "rgba(255,255,255,.08)",
            "line": "rgba(255,255,255,.16)", "line-strong": "rgba(255,255,255,.5)",
            "text": "#f1f3ff", "title": "#ffffff", "muted": "#a9aedb",
            "a1": "#818cf8", "a2": "#ec4899", "a3": "#2dd4bf",
            "grad": "linear-gradient(120deg,#6366f1,#ec4899 60%,#f59e0b)",
            "grad-text": "linear-gradient(90deg,#a5b4fc,#f0abfc 45%,#5eead4,#a5b4fc)",
            "btn-text": "#ffffff", "btn-shadow": "0 10px 34px rgba(236,72,153,.35)",
            "btn-shadow-hover": "0 14px 42px rgba(99,102,241,.55)",
            "radius": "24px", "radius-sm": "16px", "blur": "22px",
            "shadow-card": "0 8px 32px rgba(0,0,0,.35)",
            "font-head": "'Plus Jakarta Sans',sans-serif", "font-body": "'Plus Jakarta Sans'", "font-mono": "'JetBrains Mono',monospace",
            "head-tracking": "-.01em", "btn-tracking": ".04em", "btn-transform": "none", "btn-weight": "700",
            "badge-bg": "rgba(255,255,255,.08)", "badge-text": "#c7d2fe",
            "chip-bg": "rgba(255,255,255,.07)", "chip-border": "rgba(255,255,255,.18)",
            "chip-hover": "0 0 22px rgba(167,139,250,.5)",
            "stat-bg": "rgba(255,255,255,.07)",
            "drop-bg": "rgba(255,255,255,.05)", "drop-border": "1.6px dashed rgba(255,255,255,.35)",
            "focus": "0 0 0 3px rgba(129,140,248,.25)", "glow": "0 0 50px rgba(129,140,248,.5)",
            "empty-border": "rgba(255,255,255,.25)",
        },
        "extra": ".hero-title span{background-size:200% 100%;animation:sweep 6s linear infinite}"
                 ".hero .hero-title{font-weight:800}",
    },
    # ---------------------------------------------------------------- 3
    "terminal": {
        "label": "💻 Terminal Hacker",
        "fonts": "family=JetBrains+Mono:wght@400;500;700;800",
        "vars": {
            "bg": "repeating-linear-gradient(0deg,rgba(0,255,156,.04) 0 1px,transparent 1px 3px),"
                  "radial-gradient(900px 500px at 50% -10%,rgba(0,255,156,.14),transparent 60%),#020805",
            "panel": "rgba(0,22,11,.78)", "solid": "#04140b", "input-bg": "rgba(0,22,11,.78)",
            "line": "rgba(0,255,156,.35)", "line-strong": "#00ff9c",
            "text": "#c8ffe0", "title": "#e8fff2", "muted": "#5fae85",
            "a1": "#00ff9c", "a2": "#7dff6b", "a3": "#ffd400",
            "grad": "linear-gradient(120deg,#00ff9c,#7dff6b)",
            "grad-text": "linear-gradient(90deg,#00ff9c,#b6ff6b)",
            "title-filter": "drop-shadow(0 0 12px rgba(0,255,156,.5))",
            "btn-text": "#02130a", "btn-shadow": "0 0 18px rgba(0,255,156,.35)",
            "btn-shadow-hover": "0 0 34px rgba(0,255,156,.7)",
            "radius": "4px", "radius-sm": "3px", "blur": "4px",
            "font-head": "'JetBrains Mono',monospace", "font-body": "'JetBrains Mono'", "font-mono": "'JetBrains Mono',monospace",
            "head-transform": "uppercase", "head-tracking": ".06em", "btn-tracking": ".08em",
            "badge-bg": "rgba(0,255,156,.08)", "badge-text": "#00ff9c", "badge-radius": "3px",
            "chip-bg": "rgba(0,255,156,.06)", "chip-border": "rgba(0,255,156,.35)",
            "chip-hover": "0 0 14px rgba(0,255,156,.5)",
            "stat-bg": "rgba(0,22,11,.8)",
            "drop-bg": "rgba(0,255,156,.04)", "drop-border": "1.5px dashed rgba(0,255,156,.55)",
            "focus": "0 0 0 2px rgba(0,255,156,.25)", "glow": "0 0 36px rgba(0,255,156,.4)",
            "num-shadow": "0 0 14px rgba(0,255,156,.6)", "ok": "#00ff9c",
            "empty-border": "rgba(0,255,156,.35)",
        },
        "extra": ".hero .hero-title::after{content:'_';color:var(--a1);-webkit-text-fill-color:var(--a1);"
                 "animation:blink 1s steps(1) infinite;margin-left:6px}"
                 ".step-title::before{content:'$ ';color:var(--a1)}"
                 ".stat-value{text-shadow:0 0 12px rgba(0,255,156,.55)}"
                 ".hero .hero-sub::before{content:'// ';color:var(--a1)}",
    },
    # ---------------------------------------------------------------- 4
    "brutal": {
        "label": "🟨 Neo-Brutalism",
        "fonts": "family=Archivo+Black&family=Space+Grotesk:wght@400;500;600;700&family=JetBrains+Mono:wght@500;700",
        "vars": {
            "light": "1",
            "bg": "radial-gradient(rgba(0,0,0,.14) 1.3px,transparent 1.3px) 0 0/22px 22px,#fff4d6",
            "panel": "#ffffff", "solid": "#ffffff", "input-bg": "#ffffff",
            "line": "#111111", "line-strong": "#111111",
            "text": "#111111", "title": "#111111", "muted": "#444444",
            "a1": "#ffd60a", "a2": "#ff6b9d", "a3": "#4cc9f0",
            "grad": "linear-gradient(#ffd60a,#ffd60a)",
            "grad-text": "linear-gradient(#ff6b9d,#ff6b9d)",
            "btn-text": "#111111", "btn-shadow": "5px 5px 0 #111", "btn-shadow-hover": "7px 7px 0 #111",
            "btn-border": "3px solid #111", "btn-tracking": ".04em", "btn-weight": "700",
            "radius": "12px", "radius-sm": "10px", "blur": "0px", "bw": "3px",
            "shadow-card": "5px 5px 0 #111", "focus": "7px 7px 0 #111",
            "font-head": "'Archivo Black',sans-serif", "font-body": "'Space Grotesk'", "font-mono": "'JetBrains Mono',monospace",
            "head-tracking": ".01em", "head-transform": "uppercase",
            "badge-bg": "#4cc9f0", "badge-text": "#111111", "ok": "#16a34a",
            "chip-bg": "#ffffff", "chip-border": "#111111", "chip-hover": "5px 5px 0 #111",
            "stat-bg": "#ffffff", "stat-bar-h": "10px",
            "drop-bg": "#fffbe8", "drop-border": "3px dashed #111",
            "num-border": "3px solid #111", "num-shadow": "3px 3px 0 #111",
            "df-filter": "invert(1) hue-rotate(180deg)", "df-border": "0 solid transparent",
            "empty-border": "#111", "glow": "6px 6px 0 #111",
        },
        "extra": ".hero .hero-title span{-webkit-background-clip:border-box;background-clip:border-box;color:#111;"
                 "padding:0 .22em;border:3px solid #111;box-shadow:6px 6px 0 #111;display:inline-block;transform:rotate(-1.5deg)}"
                 ".chip{box-shadow:3px 3px 0 #111;font-weight:700}"
                 ".chip:nth-child(3n+1){background:#ffd60a}.chip:nth-child(3n+2){background:#ff6b9d}.chip:nth-child(3n){background:#4cc9f0}"
                 ".badge{box-shadow:3px 3px 0 #111;font-weight:700}"
                 ".hero-line{height:5px;background:#111;animation:none}"
                 ".stat:hover{transform:translate(-2px,-2px);box-shadow:8px 8px 0 #111;border-color:#111}"
                 ".stButton>button:hover,[data-testid='stDownloadButton']>button:hover{transform:translate(-2px,-2px)}"
                 ".stButton>button:active,[data-testid='stDownloadButton']>button:active{transform:translate(3px,3px);box-shadow:0 0 0 #111!important}"
                 ".lock{border:3px solid #111;background:#ffd60a;animation:none;font-size:2rem}"
                 ".stat::before{background:linear-gradient(90deg,#ffd60a 0 33%,#ff6b9d 33% 66%,#4cc9f0 66%)}",
    },
    # ---------------------------------------------------------------- 5
    "minimal": {
        "label": "☁️ Clean Light",
        "fonts": "family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700",
        "vars": {
            "light": "1",
            "bg": "radial-gradient(900px 480px at 50% -10%,rgba(99,102,241,.13),transparent 60%),"
                  "radial-gradient(700px 400px at 100% 10%,rgba(6,182,212,.08),transparent 60%),#f6f7fb",
            "panel": "#ffffff", "solid": "#ffffff", "input-bg": "#ffffff",
            "line": "#e3e6f0", "line-strong": "#6366f1",
            "text": "#1a1d2e", "title": "#12141f", "muted": "#6b7185",
            "a1": "#4f46e5", "a2": "#8b5cf6", "a3": "#06b6d4",
            "grad": "linear-gradient(135deg,#4f46e5,#7c3aed)",
            "grad-text": "linear-gradient(90deg,#4f46e5,#9333ea 60%,#06b6d4)",
            "btn-text": "#ffffff", "btn-shadow": "0 6px 18px rgba(79,70,229,.30)",
            "btn-shadow-hover": "0 10px 26px rgba(79,70,229,.42)",
            "radius": "16px", "radius-sm": "12px", "blur": "0px",
            "shadow-card": "0 1px 2px rgba(16,24,40,.06),0 8px 24px rgba(16,24,40,.06)",
            "focus": "0 0 0 3px rgba(99,102,241,.22)",
            "font-head": "'Inter',sans-serif", "font-body": "'Inter'", "font-mono": "'JetBrains Mono',monospace",
            "head-tracking": "-.01em", "btn-tracking": ".02em", "btn-transform": "none", "btn-weight": "600",
            "badge-bg": "rgba(79,70,229,.09)", "badge-text": "#4f46e5", "ok": "#16a34a",
            "chip-bg": "#ffffff", "chip-border": "#e3e6f0", "chip-hover": "0 4px 14px rgba(79,70,229,.18)",
            "stat-bg": "#ffffff",
            "drop-bg": "#fbfbff", "drop-border": "1.6px dashed #b9bdf5",
            "df-filter": "invert(1) hue-rotate(180deg)", "df-border": "0 solid transparent",
            "empty-border": "#cfd3e6", "glow": "0 10px 40px rgba(79,70,229,.25)",
            "num-shadow": "0 6px 14px rgba(79,70,229,.35)",
        },
        "extra": ".hero .hero-title{font-weight:800}",
    },
}


# ======================================================================
#  CSS DASAR (memakai variabel tema)
# ======================================================================
BASE_CSS = """
@keyframes gridmove{to{background-position:0 48px,48px 0,0 0,0 0,0 0,0 0}}
@keyframes drift{0%,100%{background-position:0% 0%,100% 0%,50% 100%,0 0}50%{background-position:30% 40%,70% 30%,60% 70%,0 0}}
@keyframes pulse{70%{box-shadow:0 0 0 10px rgba(44,255,154,0)}100%{box-shadow:0 0 0 0 rgba(44,255,154,0)}}
@keyframes sweep{to{background-position:-200% 0}}
@keyframes float{50%{transform:translateY(-6px)}}
@keyframes blink{50%{opacity:0}}

.stApp{font-family:var(--font-body),sans-serif;color:var(--text);background:var(--bg);background-attachment:fixed;animation:var(--bg-anim)}
[data-testid="stHeader"]{background:transparent}
[data-testid="stHeader"] *,[data-testid="stToolbar"] *{color:var(--text)!important}
[data-testid="stElementContainer"]:has(style){display:none}
footer{visibility:hidden}
.block-container{max-width:1080px;padding-top:2.2rem;padding-bottom:4rem}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-thumb{background:var(--grad);border-radius:8px}
::-webkit-scrollbar-track{background:transparent}

/* ---------- Hero ---------- */
.hero{display:flex;flex-direction:column;align-items:center;text-align:center;width:100%;padding:1.2rem 0 .6rem}
.badge{display:inline-flex;align-items:center;gap:8px;padding:6px 14px;border-radius:var(--badge-radius);
  border:var(--bw) solid var(--line);background:var(--badge-bg);
  font:600 .72rem var(--font-mono);letter-spacing:.18em;color:var(--badge-text)}
.dot{width:8px;height:8px;border-radius:50%;background:var(--ok);box-shadow:0 0 0 0 rgba(44,255,154,.7);animation:pulse 1.8s infinite}
.hero .hero-title{width:100%;text-align:center;font-family:var(--font-head);font-weight:900;margin:.9rem 0 .4rem;padding:0;
  font-size:clamp(1.9rem,5vw,3.4rem);letter-spacing:var(--head-tracking);text-transform:var(--head-transform);line-height:1.15;color:var(--title)}
.hero-title span{background:var(--grad-text);-webkit-background-clip:text;background-clip:text;color:transparent;filter:var(--title-filter)}
.hero .hero-sub{max-width:640px;width:100%;margin:0!important;text-align:center;color:var(--muted);font-size:1.02rem;line-height:1.6}
.chips{display:flex;flex-wrap:wrap;justify-content:center;gap:10px;margin-top:1.3rem}
.chip{padding:6px 14px;border-radius:calc(var(--radius-sm) - 2px);font:500 .8rem var(--font-mono);color:var(--text);
  background:var(--chip-bg);border:var(--bw) solid var(--chip-border);transition:.25s}
.chip:hover{transform:translateY(-2px);border-color:var(--line-strong);box-shadow:var(--chip-hover)}
.hero-line{height:2px;margin:1.6rem 0 0;width:100%;max-width:520px;border-radius:2px;
  background:linear-gradient(90deg,transparent,var(--a1),var(--a2),transparent);background-size:200% 100%;animation:sweep 3s linear infinite}

/* ---------- Judul langkah ---------- */
.step{display:flex;align-items:center;gap:14px;margin:2.3rem 0 .7rem}
.step-num{width:40px;height:40px;flex:none;border-radius:calc(var(--radius-sm) - 1px);display:grid;place-items:center;
  font:900 1rem var(--font-head);color:var(--btn-text);background:var(--grad);border:var(--num-border);box-shadow:var(--num-shadow)}
.step-title{font:700 1.15rem var(--font-head);letter-spacing:var(--head-tracking);text-transform:var(--head-transform);color:var(--title);margin:0}
.step-sub{font-size:.88rem;color:var(--muted);margin:2px 0 0}

/* ---------- Widget ---------- */
[data-baseweb="select"]>div,[data-baseweb="input"]{
  background:var(--input-bg)!important;border:var(--bw) solid var(--line)!important;border-radius:var(--radius-sm)!important;
  box-shadow:var(--shadow-card);backdrop-filter:blur(var(--blur));transition:.25s}
[data-baseweb="select"]>div:hover,[data-baseweb="input"]:focus-within{border-color:var(--line-strong)!important;box-shadow:var(--focus)}
[data-baseweb="select"] div,[data-baseweb="select"] span,.stApp input{color:var(--text)!important}
[data-baseweb="select"] svg,[data-baseweb="input"] svg{color:var(--muted)!important;fill:var(--muted)!important}
[data-testid="stTextInput"] input{letter-spacing:.2em}
[data-testid="stWidgetLabel"] p,[data-testid="stWidgetLabel"] label{color:var(--text)!important}
[data-baseweb="popover"]>div,[data-baseweb="popover"] ul{background:var(--solid)!important;border-radius:var(--radius-sm)!important}
[data-baseweb="popover"] li,[data-baseweb="popover"] li *{color:var(--text)!important}
[data-baseweb="popover"] li:hover,[data-baseweb="popover"] li[aria-selected="true"]{background:var(--chip-bg)!important}

[data-testid="stFileUploaderDropzone"],[data-testid="stFileUploader"] section{
  border:var(--drop-border)!important;border-radius:var(--radius)!important;background:var(--drop-bg)!important;
  padding:1.6rem!important;transition:.3s;box-shadow:var(--shadow-card)}
[data-testid="stFileUploaderDropzone"]:hover,[data-testid="stFileUploader"] section:hover{border-color:var(--line-strong)!important;box-shadow:var(--glow)}
[data-testid="stFileUploaderDropzone"] span,[data-testid="stFileUploaderDropzone"] small,[data-testid="stFileUploaderDropzone"] div{color:var(--muted)}
[data-testid="stFileUploaderFile"]{background:var(--panel);border:var(--bw) solid var(--line);border-radius:10px;padding:2px 8px}
[data-testid="stFileUploaderFile"] *{color:var(--text)!important}

.stButton>button,[data-testid="stDownloadButton"]>button,[data-testid="stBaseButton-secondary"],[data-testid="stBaseButton-primary"]{
  width:100%;border:var(--btn-border)!important;border-radius:var(--radius-sm)!important;padding:.85rem 1.2rem!important;
  background:var(--grad)!important;background-size:180% 100%!important;color:var(--btn-text)!important;
  font:var(--btn-weight) .95rem var(--font-head)!important;letter-spacing:var(--btn-tracking);text-transform:var(--btn-transform);
  box-shadow:var(--btn-shadow);transition:.25s}
.stButton>button:hover,[data-testid="stDownloadButton"]>button:hover{background-position:100% 0!important;transform:translateY(-2px);box-shadow:var(--btn-shadow-hover)}
.stButton>button p,[data-testid="stDownloadButton"]>button p,[data-testid="stBaseButton-secondary"] p{color:var(--btn-text)!important;font-weight:var(--btn-weight)!important}

[data-testid="stAlert"]{background:var(--panel)!important;border:var(--bw) solid var(--line)!important;
  border-radius:var(--radius-sm)!important;backdrop-filter:blur(var(--blur));box-shadow:var(--shadow-card)}
[data-testid="stAlert"] *{color:var(--text)!important}
[data-testid="stExpander"]{border:var(--bw) solid var(--line)!important;border-radius:var(--radius)!important;
  background:var(--panel);backdrop-filter:blur(var(--blur));overflow:hidden;margin-bottom:.8rem;box-shadow:var(--shadow-card)}
[data-testid="stExpander"] summary,[data-testid="stExpander"] summary *{color:var(--text)!important}
[data-testid="stExpander"] summary:hover,[data-testid="stExpander"] summary:hover *{color:var(--a1)!important}
[data-testid="stDataFrame"]{border:var(--df-border);border-radius:12px;overflow:hidden;filter:var(--df-filter)}
[data-testid="stSpinner"],[data-testid="stSpinner"] *{color:var(--text)!important}

/* ---------- Kartu statistik ---------- */
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:16px;margin:.8rem 0 1.4rem}
.stat{position:relative;overflow:hidden;padding:20px 20px 18px;border-radius:var(--radius);background:var(--stat-bg);
  border:var(--bw) solid var(--line);box-shadow:var(--shadow-card);backdrop-filter:blur(var(--blur));transition:.3s}
.stat:hover{transform:translateY(-3px);border-color:var(--line-strong);box-shadow:var(--glow)}
.stat::before{content:"";position:absolute;left:0;top:0;height:var(--stat-bar-h);width:100%;background:var(--grad)}
.stat-label{font:600 .7rem var(--font-mono);letter-spacing:.16em;text-transform:uppercase;color:var(--muted)}
.stat-value{font:700 1.45rem var(--font-mono);color:var(--title);margin:6px 0 2px;word-break:break-word}
.stat-sub{font-size:.78rem;color:var(--badge-text);font-weight:600}

/* ---------- Empty state, login, footer ---------- */
.empty{padding:1.6rem;text-align:center;border-radius:var(--radius);color:var(--muted);
  border:var(--bw) dashed var(--empty-border);background:var(--panel)}
.empty b{color:var(--badge-text)}
.lock{width:74px;height:74px;margin:0 auto 10px;border-radius:20px;display:grid;place-items:center;font-size:2rem;
  background:var(--chip-bg);border:var(--bw) solid var(--line);box-shadow:var(--glow);animation:float 3.5s ease-in-out infinite}
.foot{text-align:center;margin-top:3rem;font:500 .72rem var(--font-mono);letter-spacing:.2em;color:var(--muted)}
"""


# ======================================================================
#  FUNGSI TEMA
# ======================================================================
def current_theme() -> str:
    key = st.session_state.get("ui_theme", DEFAULT_THEME)
    return key if key in THEMES else DEFAULT_THEME


def build_css(theme_key: str) -> str:
    t = THEMES[theme_key]
    v = {**DEFAULTS, **t["vars"]}
    v.pop("light", None)
    root = ":root{" + "".join(f"--{k}:{val};" for k, val in v.items()) + "}"
    imp = f"@import url('https://fonts.googleapis.com/css2?{t['fonts']}&display=swap');"
    return imp + root + BASE_CSS + t["extra"]


def inject_css():
    """Suntikkan CSS sesuai tema yang dipilih (panggil di awal app.py, sebelum elemen lain)."""
    st.session_state.setdefault("ui_theme", DEFAULT_THEME)
    st.markdown("<style>" + _h(build_css(current_theme())) + "</style>", unsafe_allow_html=True)


def theme_picker():
    """Dropdown pemilih tema di pojok kanan atas."""
    st.session_state.setdefault("ui_theme", DEFAULT_THEME)
    _, right = st.columns([3, 1.4])
    with right:
        st.selectbox(
            "🎨 Tema tampilan",
            list(THEMES.keys()),
            format_func=lambda k: THEMES[k]["label"],
            key="ui_theme",
        )


# ======================================================================
#  KOMPONEN
# ======================================================================
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
    <div class="hero" style="padding-top:2rem">
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
