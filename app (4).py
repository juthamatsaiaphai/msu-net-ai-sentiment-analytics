import html
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================
# ตั้งค่าหน้าเว็บ
# ==========================

st.set_page_config(
    page_title="MSU Net / Wi-Fi Satisfaction Dashboard",
    page_icon="📶",
    layout="wide",
)

EXCEL_NAME = "สรุปผล.xlsx"
BASE_DIR = Path(__file__).parent

# ชื่อคอลัมน์ในไฟล์ (ไม่สนตัวพิมพ์/ช่องว่าง/_ ) -> ชื่อมาตรฐานในโค้ด
ALIASES = {
    "responseid": "response_id",
    "sentiment": "sentiment",
    "keyissues": "key_issues",
    "issuelocation": "issue_location",
    "positivefeedback": "positive_feedback",
    "aisummary": "summary",
    "summary": "summary",
    "aisuggestions": "suggestions",
    "suggestions": "suggestions",
    "processingstatus": "status",
    "processedat": "processed_at",
    "errormessage": "error",
    "retrycount": "retry",
    "score": "score",
}
REQUIRED = ["sentiment", "key_issues", "summary", "suggestions"]

NAVY, NAVY_DARK, YELLOW, CYAN = "#0b3a73", "#062a57", "#f7c61f", "#0ea5e9"

SENT_ORDER = ["Positive", "Neutral", "Negative"]
SENT_COLORS = {"Positive": "#22c55e", "Neutral": "#f59e0b", "Negative": "#ef4444"}
ACTION_COLORS = {
    "🔴 ด่วน": "#ef4444",
    "🟡 ติดตาม": "#f59e0b",
    "🟢 เฝ้าระวัง": "#22c55e",
}

# ข้อความในไฟล์คั่นหลายรายการด้วย " | "
SPLIT_RE = r"\s*\|\s*"

# กฎจัดกลุ่มข้อความ (จับคำสำคัญ ไม่ใช่การวิเคราะห์ด้วย AI)
ISSUE_THEMES = {
    "สัญญาณหลุด / ไม่เสถียร": ["หลุด", "ไม่เสถียร", "ผันผวน", "สวิง", "มาๆ หายๆ", "ไม่สม่ำเสมอ", "ขัดข้อง"],
    "อินเทอร์เน็ตช้า": ["ช้า", "ความเร็ว", "หน่วง", "โหลด"],
    "พื้นที่ไม่ครอบคลุม / จุดอับ": ["ครอบคลุม", "จุดอับ", "เข้าไม่ถึง", "ไปไม่ถึง", "ไม่มีบริการ", "สัญญาณอ่อน", "อ่อนมาก"],
    "ผู้ใช้หนาแน่น / ช่วงพีค": ["จำนวนมาก", "พร้อมกัน", "หลายคน", "พักเที่ยง", "ช่วงค่ำ", "แออัด", "ใกล้สอบ", "ช่วงเช้า"],
    "ล็อกอิน / Captive Portal": ["ล็อกอิน", "captive", "ยืนยันตัวตน", "ประวัติการเชื่อมต่อ", "ip address"],
    "Roaming / สลับอาคาร": ["roaming", "สลับ", "เปลี่ยนอาคาร", "ย้ายห้อง", "access point"],
    "อัปโหลด / แบนด์วิดท์": ["อัปโหลด", "แบนด์วิดท์", "ดาวน์โหลด"],
}

LOCATION_GROUPS = {
    "อาคารเรียน / ห้องเรียน": ["อาคารเรียน", "ห้องเรียน", "ตึกเรียน", "ตึก rn", "อาคาร rn", "ใต้ตึก", "เลคเชอร์"],
    "โรงอาหาร": ["โรงอาหาร"],
    "สนามกีฬา / ตลาดน้อย": ["สนามกีฬา", "ตลาดน้อย"],
    "หอพัก": ["หอพัก"],
    "อาคาร/ห้องปฏิบัติการ": ["ปฏิบัติการ"],
    "สำนักวิทยบริการ": ["สำนักวิทย"],
    "ลานกิจกรรม": ["ลานกิจกรรม"],
    "ไม่ระบุ": ["ไม่ระบุ"],
}

POSITIVE_THEMES = {
    "เชื่อมต่อ / ใช้งานง่าย": ["ง่าย", "สะดวก", "ไม่ซับซ้อน"],
    "ความเร็วดี": ["เร็ว", "ความเร็ว"],
    "สัญญาณแรง": ["แรง"],
    "ครอบคลุม / มีให้ใช้": ["ครอบคลุม", "กระจาย", "หลายพื้นที่", "มี wi-fi", "มีเครือข่าย"],
    "ใช้ฟรี / ประหยัด": ["ฟรี", "ประหยัด"],
    "เสถียร / ปลอดภัย": ["เสถียร", "ปลอดภัย"],
    "มีคู่มือ / แนะนำ": ["คู่มือ", "แนะนำ"],
}

SUGGESTION_THEMES = {
    "ปรับปรุงความเสถียร": ["เสถียร", "หลุด"],
    "เพิ่มความเร็ว / แบนด์วิดท์": ["ความเร็ว", "แบนด์วิดท์", "ช้า", "อัปโหลด"],
    "ขยายพื้นที่ / เพิ่มจุดบริการ": ["ครอบคลุม", "เพิ่มจุด", "ติดตั้ง", "ขยาย", "access point", "พื้นที่"],
    "รองรับผู้ใช้พร้อมกัน": ["รองรับ", "ผู้ใช้งานพร้อมกัน", "ผู้ใช้จำนวน"],
    "ปรับระบบล็อกอิน": ["ล็อกอิน", "ยืนยันตัวตน", "portal", "ประวัติ", "บันทึก"],
    "แจ้งเตือน / สื่อสาร": ["แจ้งเตือน", "ประชาสัมพันธ์"],
}

OTHER = "อื่น ๆ"

# (id ของส่วน, ชื่อบนเมนู, หัวข้อเล็กสีฟ้า)
SECTIONS = [
    ("overview", "ภาพรวม & สรุปผู้ตอบ", "OVERVIEW"),
    ("sentiment", "Sentiment Analysis", "SENTIMENT"),
    ("issues", "Issue Intensity", "ISSUES"),
    ("spatial", "Spatial & Location", "LOCATION"),
    ("priority", "Priority Matrix & Action", "PRIORITY"),
    ("insights", "Insights & Recommendations", "INSIGHTS"),
    ("explorer", "Data Explorer", "DATA EXPLORER"),
]
EYEBROW = {sid: eb for sid, _, eb in SECTIONS}
TITLE = {sid: label for sid, label, _ in SECTIONS}

# ==========================
# ฟังก์ชันประมวลผลข้อมูล
# ==========================


def norm_col(name):
    return re.sub(r"[\s_]+", "", str(name)).lower()


def standardize(d):
    d = d.copy()
    d.columns = [ALIASES.get(norm_col(c), str(c).strip()) for c in d.columns]
    return d


def split_items(value):
    if pd.isna(value):
        return []
    return [p.strip() for p in re.split(SPLIT_RE, str(value)) if p.strip()]


def tag(text, rules):
    """คืนรายชื่อกลุ่มทั้งหมดที่ข้อความตรงกับคำสำคัญ (หนึ่งข้อความอยู่ได้หลายกลุ่ม)"""
    t = str(text).lower()
    hits = [name for name, kws in rules.items() if any(k in t for k in kws)]
    return hits or [OTHER]


def tag_cell(value, rules):
    """จัดกลุ่มทั้งช่อง (รวมทุกข้อความในช่อง) ไม่ซ้ำกัน"""
    if pd.isna(value) or not str(value).strip():
        return []
    seen = []
    for item in split_items(value):
        for h in tag(item, rules):
            if h not in seen:
                seen.append(h)
    return seen


def prepare(d):
    d = d.dropna(subset=["sentiment"]).copy()
    d["sentiment"] = d["sentiment"].astype(str).str.strip().str.title()
    if "score" in d.columns:
        d["score"] = pd.to_numeric(d["score"], errors="coerce")
    d["n_issues"] = d["key_issues"].apply(lambda v: len(split_items(v)))
    d["issue_themes"] = d["key_issues"].apply(lambda v: tag_cell(v, ISSUE_THEMES))
    if "issue_location" in d.columns:
        d["loc_groups"] = d["issue_location"].apply(lambda v: tag_cell(v, LOCATION_GROUPS))
    else:
        d["loc_groups"] = [[] for _ in range(len(d))]
    d["pos_themes"] = (
        d["positive_feedback"].apply(lambda v: tag_cell(v, POSITIVE_THEMES))
        if "positive_feedback" in d.columns
        else [[] for _ in range(len(d))]
    )
    return d.reset_index(drop=True)


def theme_table(d, col, overall_neg):
    """ตารางต่อกลุ่ม : จำนวนผู้ตอบที่กล่าวถึง, % Negative, และแนวทาง"""
    ex = d[["sentiment", col]].explode(col).dropna(subset=[col])
    if ex.empty:
        return pd.DataFrame(columns=["Theme", "Count", "Share", "NegPct", "Action"])
    g = (
        ex.groupby(col)
        .agg(Count=(col, "size"), NegPct=("sentiment", lambda s: round((s == "Negative").mean() * 100, 1)))
        .reset_index()
        .rename(columns={col: "Theme"})
    )
    g["Share"] = (g["Count"] / len(d) * 100).round(1)
    med = g["Count"].median()

    def action(r):
        hi, neg = r["Count"] >= med, r["NegPct"] >= overall_neg
        if hi and neg:
            return "🔴 ด่วน"
        return "🟡 ติดตาม" if (hi or neg) else "🟢 เฝ้าระวัง"

    g["Action"] = g.apply(action, axis=1)
    return g.sort_values("Count", ascending=False).reset_index(drop=True)


def top_theme(tbl, skip=(OTHER, "ไม่ระบุ")):
    t = tbl[~tbl["Theme"].isin(skip)] if not tbl.empty else tbl
    return None if t.empty else t.iloc[0]


# ==========================
# หน้าตา (CSS / HTML)
# ==========================

CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700&display=swap');
html{scroll-behavior:smooth;}
header[data-testid="stHeader"],#MainMenu,footer{display:none !important;}
.stApp{background:#f4f7fc;}
.stApp,.stApp p,.stApp label,.stApp li,.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp h5,.stApp input,.stApp textarea,.stApp button,.stApp [data-testid="stMetricValue"],.stApp [data-testid="stMetricLabel"],.stApp [data-testid="stCaptionContainer"],.stApp summary p{font-family:'Prompt','Noto Sans Thai',sans-serif;color:#1f2a44;}
.block-container{padding-top:7.2rem;max-width:1240px;padding-bottom:0;}

.topnav{position:fixed;top:0;left:0;right:0;z-index:999;height:76px;display:flex;align-items:center;justify-content:space-between;gap:24px;padding:0 4vw;background:#fff;border-bottom:4px solid #f7c61f;overflow-x:auto;white-space:nowrap;}
.brand{display:flex;align-items:center;gap:12px;flex-shrink:0;}
.brand-dot{width:34px;height:34px;border-radius:50%;background:radial-gradient(circle at 35% 35%,#2f7fd6,#0b3a73);display:flex;align-items:center;justify-content:center;color:#fff;font-size:17px;}
.brand-t{font-weight:700;font-size:1.05rem;color:#0b3a73;line-height:1.1;}
.brand-s{font-size:.72rem;color:#6b7a99;}
.links{display:flex;gap:26px;align-items:center;}
.links a{font-size:.88rem;font-weight:500;color:#1f2a44 !important;text-decoration:none !important;padding:6px 0;transition:color .15s ease;}
.links a:hover{color:#1596c8 !important;}

.hero{position:relative;overflow:hidden;border-radius:34px;background:linear-gradient(135deg,#0d4a90 0%,#062a57 100%);padding:64px 48px 140px;text-align:center;color:#fff;}
.hero:before{content:"";position:absolute;right:-90px;top:-110px;width:380px;height:380px;border-radius:50%;background:rgba(255,255,255,.06);}
.hero:after{content:"";position:absolute;left:-70px;bottom:-130px;width:300px;height:300px;border-radius:50%;background:rgba(247,198,31,.10);}
.hero-eyebrow{position:relative;z-index:1;font-size:.8rem;font-weight:600;letter-spacing:.2em;color:#f7c61f;}
.hero-title{position:relative;z-index:1;margin:14px 0 12px;font-size:2.4rem;font-weight:700;line-height:1.3;color:#fff;}
.hero-sub{position:relative;z-index:1;font-size:1rem;color:#c7d6ee;margin-bottom:28px;}
.hero-btns{position:relative;z-index:1;display:flex;gap:14px;justify-content:center;flex-wrap:wrap;}
.btn-y,.btn-o{display:inline-block;padding:13px 30px;border-radius:999px;font-weight:600;font-size:.95rem;text-decoration:none !important;transition:transform .15s ease;}
.btn-y{background:#f7c61f;color:#0b2a55 !important;}
.btn-o{border:2px solid rgba(255,255,255,.75);color:#fff !important;}
.btn-y:hover,.btn-o:hover{transform:translateY(-2px);}

.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:22px;margin:-92px 28px 8px;position:relative;z-index:2;}
.kpi{background:#fff;border-radius:26px;padding:28px 18px 24px;text-align:center;box-shadow:0 14px 40px rgba(10,40,90,.14);}
.kpi-ic{width:64px;height:64px;border-radius:18px;margin:0 auto 14px;display:flex;align-items:center;justify-content:center;font-size:28px;}
.kpi-t{font-weight:600;font-size:1.02rem;color:#1f2a44;}
.kpi-n{font-size:2.6rem;font-weight:700;color:#0b3a73;line-height:1.15;margin:4px 0;}
.kpi-d{font-size:.84rem;color:#6b7a99;}
@media (max-width:900px){.kpi-grid{grid-template-columns:repeat(2,1fr);margin:-92px 4px 8px;}.hero{padding:44px 22px 130px;}.hero-title{font-size:1.7rem;}}

.anchor{position:relative;top:-100px;visibility:hidden;height:0;}
.sec-head{text-align:center;margin:44px 0 26px;}
.eyebrow{font-size:.78rem;font-weight:600;letter-spacing:.18em;color:#1d7fd1;}
.sec-title{font-size:1.9rem;font-weight:700;color:#1f2a44;margin-top:4px;}
.sec-bar{width:54px;height:4px;border-radius:4px;background:#1d7fd1;margin:12px auto 10px;}
.sec-desc{color:#6b7a99;font-size:.95rem;max-width:760px;margin:0 auto;}

.callout{background:#fff;border-radius:18px;padding:18px 22px;border-left:5px solid #1d7fd1;box-shadow:0 8px 26px rgba(10,40,90,.07);height:100%;}
.callout-t{font-weight:600;color:#0b3a73;margin-bottom:4px;}
.callout-b{color:#4a5877;font-size:.93rem;line-height:1.6;}

[data-testid="stPlotlyChart"]{background:#fff;border-radius:22px;padding:10px 10px 0;box-shadow:0 8px 26px rgba(10,40,90,.07);}
[data-testid="stDataFrame"]{border-radius:18px;overflow:hidden;box-shadow:0 8px 26px rgba(10,40,90,.07);}
[data-testid="stMetric"]{background:#fff;border-radius:20px;padding:18px 16px;box-shadow:0 8px 26px rgba(10,40,90,.07);align-items:center;text-align:center;}
[data-testid="stMetricLabel"]{justify-content:center;}
[data-testid="stMetricValue"]{color:#0b3a73;font-weight:700;}
[data-testid="stExpander"]{background:#fff;border:none;border-radius:16px;box-shadow:0 6px 20px rgba(10,40,90,.06);}
[data-testid="stExpander"] details{border:none;border-radius:16px;}
div[role="radiogroup"]{gap:10px;}
div[role="radiogroup"] label{background:#fff;border:1px solid #dbe4f3;border-radius:999px;padding:6px 20px;cursor:pointer;}
div[role="radiogroup"] label > div:first-child{display:none;}
div[role="radiogroup"] label:has(input:checked){background:#0ea5e9;border-color:#0ea5e9;}
div[role="radiogroup"] label:has(input:checked) p{color:#fff !important;font-weight:600;}
div[data-baseweb="input"],div[data-baseweb="select"] > div{border-radius:999px !important;background:#fff;}
[data-testid="stDownloadButton"] button{background:#f7c61f;color:#0b2a55;border:none;border-radius:999px;font-weight:600;padding:8px 24px;}
[data-testid="stDownloadButton"] button:hover{background:#e8b810;color:#0b2a55;}
[data-testid="stAlert"]{border-radius:16px;}

.footer{margin:56px -4vw 0;background:#062a57;color:#c7d6ee;text-align:center;padding:26px 16px;font-size:.88rem;border-top:4px solid #f7c61f;}
.footer b{color:#fff;font-weight:600;}
</style>"""


def nav_html():
    links = "".join(
        f'<a href="#{sid}" target="_self">{html.escape(label)}</a>' for sid, label, _ in SECTIONS
    )
    return (
        '<div class="topnav">'
        '<div class="brand"><div class="brand-dot">📶</div>'
        '<div><div class="brand-t">Wi-Fi Insight</div>'
        '<div class="brand-s">วิเคราะห์ความคิดเห็นด้วย AI</div></div></div>'
        f'<div class="links">{links}</div></div>'
    )


def hero_html(total, counts, updated):
    def kpi(icon, bg, title, n, desc):
        return (
            f'<div class="kpi"><div class="kpi-ic" style="background:{bg}">{icon}</div>'
            f'<div class="kpi-t">{title}</div><div class="kpi-n">{n}</div>'
            f'<div class="kpi-d">{desc}</div></div>'
        )

    pct = lambda n: f"{n / total:.0%} ของผู้ตอบ"
    upd = f" · ประมวลผลล่าสุด {updated}" if updated else ""
    return (
        '<div class="hero">'
        '<div class="hero-eyebrow">STUDENT SATISFACTION ANALYTICS</div>'
        '<div class="hero-title">ระบบวิเคราะห์ความคิดเห็นการใช้บริการ<br>MSU Net / Wi-Fi ด้วย AI</div>'
        f'<div class="hero-sub">วิเคราะห์ความคิดเห็นของผู้ใช้ {total:,} ราย ด้วยระบบ AI + n8n{upd}</div>'
        '<div class="hero-btns">'
        '<a class="btn-y" href="#priority" target="_self">ดูประเด็นเร่งด่วน</a>'
        '<a class="btn-o" href="#insights" target="_self">ข้อเสนอแนะจาก AI</a>'
        '</div></div>'
        '<div class="kpi-grid">'
        + kpi("👥", "#e3f0ff", "ผู้ตอบทั้งหมด", f"{total:,}", "รายการที่ AI วิเคราะห์แล้ว")
        + kpi("😞", "#fde8e8", "Negative", int(counts["Negative"]), pct(int(counts["Negative"])))
        + kpi("😐", "#fff4cc", "Neutral", int(counts["Neutral"]), pct(int(counts["Neutral"])))
        + kpi("😊", "#e3f7ec", "Positive", int(counts["Positive"]), pct(int(counts["Positive"])))
        + "</div>"
    )


def section(sid, desc=""):
    st.markdown(
        f'<div id="{sid}" class="anchor"></div>'
        '<div class="sec-head">'
        f'<div class="eyebrow">{EYEBROW[sid]}</div>'
        f'<div class="sec-title">{html.escape(TITLE[sid])}</div>'
        '<div class="sec-bar"></div>'
        f'<div class="sec-desc">{html.escape(desc)}</div></div>',
        unsafe_allow_html=True,
    )


def callout(title, body, accent="#1d7fd1"):
    return (
        f'<div class="callout" style="border-left-color:{accent}">'
        f'<div class="callout-t">{html.escape(title)}</div>'
        f'<div class="callout-b">{html.escape(body)}</div></div>'
    )


def style(fig, height=400):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=56, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend_title_text="",
        font=dict(family="Prompt, Noto Sans Thai, sans-serif", color="#1f2a44"),
        title=dict(font=dict(size=16, color=NAVY)),
    )
    return fig


st.markdown(CSS, unsafe_allow_html=True)
st.markdown(nav_html(), unsafe_allow_html=True)


@st.cache_data(show_spinner="กำลังโหลดข้อมูล...")
def load_data(source, stamp):
    sheets = pd.read_excel(source, sheet_name=None)
    for name, d in sheets.items():
        d = standardize(d)
        if all(c in d.columns for c in REQUIRED):
            return d, name, None
    info = {n: [str(c) for c in standardize(d).columns] for n, d in sheets.items()}
    return None, None, info


# ---- หาไฟล์ข้อมูล ----
path = BASE_DIR / EXCEL_NAME
if not path.exists():
    others = sorted(BASE_DIR.glob("*.xlsx"))
    path = others[0] if others else None

if path is not None:
    source, stamp = path, path.stat().st_mtime
else:
    st.title("📶 ระบบวิเคราะห์ความคิดเห็นการใช้บริการ MSU Net / Wi-Fi")
    st.warning(f"ไม่พบไฟล์ {EXCEL_NAME} ในโฟลเดอร์เดียวกับ app.py — อัปโหลดไฟล์ได้ที่นี่")
    up = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"])
    if up is None:
        st.stop()
    source, stamp = up, up.name + str(up.size)

raw, sheet_name, info = load_data(source, stamp)

if raw is None:
    st.title("📶 ระบบวิเคราะห์ความคิดเห็นการใช้บริการ MSU Net / Wi-Fi")
    st.error("ไม่พบชีทที่มีคอลัมน์ครบ : " + ", ".join(REQUIRED))
    st.write("ชีทและคอลัมน์ที่พบในไฟล์")
    st.json(info)
    st.stop()

df = prepare(raw)
if df.empty:
    st.warning("ไม่มีข้อมูลในชีท")
    st.stop()

total = len(df)
counts = df["sentiment"].value_counts().reindex(SENT_ORDER, fill_value=0)
positive, neutral, negative = (int(counts[s]) for s in SENT_ORDER)
overall_neg = round(negative / total * 100, 1)

issue_tbl = theme_table(df, "issue_themes", overall_neg)
loc_tbl = theme_table(df, "loc_groups", overall_neg)

updated = ""
if "processed_at" in df.columns:
    _t = pd.to_datetime(df["processed_at"], errors="coerce", utc=True).dt.tz_convert("Asia/Bangkok")
    if _t.notna().any():
        updated = _t.max().strftime("%d/%m/%Y %H:%M")

st.markdown(hero_html(total, counts, updated), unsafe_allow_html=True)

# ==========================
# 1) ภาพรวม & สรุปผู้ตอบ
# ==========================

section("overview", f"ข้อมูลจากชีท “{sheet_name}” จำนวน {total:,} รายการ")

m1, m2, m3, m4 = st.columns(4)
m1.metric("🚨 ประเด็นปัญหารวม", int(df["n_issues"].sum()))
m2.metric("📌 ประเด็นเฉลี่ยต่อผู้ตอบ", round(df["n_issues"].mean(), 2))
m3.metric("😞 สัดส่วน Negative", f"{overall_neg}%")
m4.metric("📍 พื้นที่ที่ถูกกล่าวถึง", int((loc_tbl["Theme"] != "ไม่ระบุ").sum()) if not loc_tbl.empty else 0)

if "score" in df.columns and df["score"].notna().any():
    sc = df["score"].value_counts().sort_index().reset_index()
    sc.columns = ["Score", "Count"]
    fig = px.bar(sc, x="Score", y="Count", text="Count", title="Satisfaction Score",
                 color_discrete_sequence=[NAVY])
    fig.update_xaxes(type="category")
    st.plotly_chart(style(fig, 340), use_container_width=True)

# คุณภาพการประมวลผลของระบบ (แยกจากความถูกต้องของผลวิเคราะห์ AI)
if "status" in df.columns:
    st.markdown("##### ⚙️ สถานะการประมวลผลของระบบอัตโนมัติ")
    ok = int((df["status"].astype(str).str.lower() == "processed").sum())
    err = int(df["error"].notna().sum()) if "error" in df.columns else 0
    retry = int(pd.to_numeric(df["retry"], errors="coerce").fillna(0).sum()) if "retry" in df.columns else 0
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("ประมวลผลสำเร็จ", f"{ok}/{total}")
    p2.metric("รายการที่มีข้อผิดพลาด", err)
    p3.metric("จำนวนการลองใหม่รวม", retry)
    p4.metric("ประมวลผลล่าสุด", updated or "-")
    st.caption("ตัวเลขนี้บอกว่าระบบทำงานครบหรือไม่ ไม่ได้บอกว่าผลวิเคราะห์ของ AI ถูกต้อง")

# ==========================
# 2) Sentiment Analysis
# ==========================

section("sentiment", "สัดส่วนความรู้สึกต่อบริการ")

sent_df = pd.DataFrame({"Sentiment": counts.index, "Count": counts.values})
sent_nz = sent_df[sent_df["Count"] > 0]
left, right = st.columns(2)

with left:
    fig = px.pie(sent_nz, values="Count", names="Sentiment", hole=0.5,
                 title="Overall Sentiment", color="Sentiment", color_discrete_map=SENT_COLORS)
    st.plotly_chart(style(fig), use_container_width=True)

with right:
    avg_issue = df.groupby("sentiment")["n_issues"].mean().round(2).reset_index()
    avg_issue.columns = ["Sentiment", "AvgIssues"]
    fig = px.bar(avg_issue, x="Sentiment", y="AvgIssues", text="AvgIssues", color="Sentiment",
                 color_discrete_map=SENT_COLORS, title="จำนวนประเด็นปัญหาเฉลี่ยต่อผู้ตอบ แยกตาม Sentiment")
    st.plotly_chart(style(fig), use_container_width=True)

if positive == 0:
    st.caption("หมายเหตุ : ในข้อมูลชุดนี้ไม่มีความคิดเห็นที่ AI จัดเป็น Positive")

# ==========================
# 3) Issue Intensity
# ==========================

section("issues", "ประเด็นปัญหาที่ผู้ตอบกล่าวถึงมากที่สุด (นับจำนวนผู้ตอบ หนึ่งคนอยู่ได้หลายประเด็น)")

if issue_tbl.empty:
    st.info("ไม่พบข้อมูลประเด็นปัญหา")
else:
    i1, i2 = st.columns([3, 2])
    with i1:
        fig = px.bar(issue_tbl.sort_values("Count"), x="Count", y="Theme", orientation="h",
                     text="Count", color="Count", color_continuous_scale="Blues",
                     title="จำนวนผู้ตอบที่กล่าวถึงแต่ละประเด็น")
        fig.update_coloraxes(showscale=False)
        fig.update_yaxes(title=None)
        st.plotly_chart(style(fig, 420), use_container_width=True)
    with i2:
        sent_issue = (
            df[["sentiment", "issue_themes"]].explode("issue_themes").dropna()
            .groupby(["issue_themes", "sentiment"]).size().reset_index(name="Count")
        )
        fig = px.bar(sent_issue, x="Count", y="issue_themes", color="sentiment", orientation="h",
                     color_discrete_map=SENT_COLORS, title="ประเด็น × Sentiment")
        fig.update_yaxes(title=None)
        st.plotly_chart(style(fig, 420), use_container_width=True)

with st.expander("ดูตัวอย่างข้อความตามประเด็น"):
    pick = st.selectbox("เลือกประเด็น", issue_tbl["Theme"].tolist() if not issue_tbl.empty else [])
    if pick:
        for _, r in df[df["issue_themes"].apply(lambda x: pick in x)].iterrows():
            st.markdown(f"**{r.get('response_id', '')}** · {r['sentiment']}")
            for it in split_items(r["key_issues"]):
                if pick in tag(it, ISSUE_THEMES):
                    st.markdown(f"- {it}")
    st.caption("การจัดกลุ่มประเด็นใช้คำสำคัญ ข้อความหนึ่งอาจอยู่ได้หลายกลุ่ม ข้อความที่ไม่ตรงกลุ่มใดจะอยู่ใน “อื่น ๆ”")

# ==========================
# 4) Spatial & Location
# ==========================

section("spatial", "จุดที่พบปัญหา จัดกลุ่มจากข้อความในคอลัมน์ Issue Location")

if loc_tbl.empty:
    st.info("ไฟล์ข้อมูลนี้ไม่มีคอลัมน์ Issue Location จึงยังแสดงผลตามพื้นที่ไม่ได้ — ใช้ไฟล์ สรุปผล.xlsx เวอร์ชันล่าสุดที่มีคอลัมน์นี้")
else:
    l1, l2 = st.columns(2)
    with l1:
        sent_loc = (
            df[["sentiment", "loc_groups"]].explode("loc_groups").dropna()
            .groupby(["loc_groups", "sentiment"]).size().reset_index(name="Count")
        )
        fig = px.bar(sent_loc, x="Count", y="loc_groups", color="sentiment", orientation="h",
                     color_discrete_map=SENT_COLORS, title="จำนวนผู้ตอบที่ระบุแต่ละพื้นที่")
        fig.update_yaxes(title=None, categoryorder="total ascending")
        st.plotly_chart(style(fig, 440), use_container_width=True)

    with l2:
        cross = (
            df[["issue_themes", "loc_groups"]]
            .explode("issue_themes").explode("loc_groups")
            .dropna().reset_index(drop=True)
        )
        cross = cross[cross["loc_groups"] != "ไม่ระบุ"]
        if cross.empty:
            st.info("ไม่มีข้อมูลเพียงพอสำหรับตารางความร้อน")
        else:
            heat = pd.crosstab(cross["issue_themes"], cross["loc_groups"])
            fig = px.imshow(heat, text_auto=True, aspect="auto", color_continuous_scale="Blues",
                            title="ประเด็นปัญหา × พื้นที่ (จำนวนคู่ที่พบ)")
            fig.update_xaxes(title=None, tickangle=-35)
            fig.update_yaxes(title=None)
            fig.update_coloraxes(showscale=False)
            st.plotly_chart(style(fig, 440), use_container_width=True)

    unspec = int(df["loc_groups"].apply(lambda x: "ไม่ระบุ" in x).sum())
    if unspec:
        st.caption(f"มี {unspec} รายการที่ไม่ระบุสถานที่ (ไม่นับในตารางความร้อน)")

# ==========================
# 5) Priority Matrix & Action
# ==========================

section("priority", "จัดลำดับจากจำนวนผู้ตอบที่กล่าวถึง × สัดส่วนที่เป็น Negative (ขวาบน = เร่งด่วน)")

if issue_tbl.empty:
    st.info("ไม่พบข้อมูลสำหรับจัดลำดับ")
else:
    fig = px.scatter(
        issue_tbl, x="Count", y="NegPct", size="Count", color="Action", text="Theme",
        hover_name="Theme", color_discrete_map=ACTION_COLORS,
        labels={"Count": "จำนวนผู้ตอบที่กล่าวถึง", "NegPct": "% Negative ของผู้ที่กล่าวถึง"},
        title="Priority Matrix : ประเด็นปัญหา",
    )
    fig.update_traces(textposition="top center")
    fig.add_vline(x=issue_tbl["Count"].median(), line_dash="dash", opacity=0.4)
    fig.add_hline(y=overall_neg, line_dash="dash", opacity=0.4,
                  annotation_text=f"Negative รวม {overall_neg}%")
    fig.update_yaxes(range=[0, 108])
    st.plotly_chart(style(fig, 500), use_container_width=True)

    show = issue_tbl.rename(columns={
        "Theme": "ประเด็น", "Count": "จำนวนผู้ตอบ", "Share": "% ของผู้ตอบทั้งหมด",
        "NegPct": "% Negative", "Action": "แนวทาง",
    })
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.caption(
        f"เกณฑ์ : ด่วน = จำนวนผู้ตอบ ≥ ค่ากลาง และ %Negative ≥ {overall_neg}% (ค่ารวมทั้งชุด) · "
        "ติดตาม = ผ่านเกณฑ์ข้อใดข้อหนึ่ง · เฝ้าระวัง = ไม่ผ่านทั้งสองข้อ"
    )

# ==========================
# 6) Insights & Recommendations
# ==========================

section("insights", "สรุป ข้อเสนอแนะ และข้อดีที่ผู้ตอบกล่าวถึง")

sugg = [s for v in df["suggestions"] for s in split_items(v)]
st_tbl = pd.DataFrame(columns=["Theme", "Count"])
if sugg:
    rows = [(th, s) for s in sugg for th in tag(s, SUGGESTION_THEMES)]
    st_tbl = pd.DataFrame(rows, columns=["Theme", "Text"]).groupby("Theme").size().reset_index(name="Count")
    st_tbl = st_tbl.sort_values("Count", ascending=False)

ti, tl = top_theme(issue_tbl), top_theme(loc_tbl)
ts = None if st_tbl.empty else (st_tbl[st_tbl["Theme"] != OTHER].iloc[0] if (st_tbl["Theme"] != OTHER).any() else None)

ca, cb, cc = st.columns(3)
with ca:
    if ti is not None:
        st.markdown(callout("ประเด็นที่พบมากที่สุด", f"{ti['Theme']} — {int(ti['Count'])} จาก {total} คน ({ti['Share']}%)", "#1d7fd1"),
                    unsafe_allow_html=True)
with cb:
    if tl is not None:
        st.markdown(callout("พื้นที่ที่ถูกกล่าวถึงมากที่สุด", f"{tl['Theme']} — {int(tl['Count'])} จาก {total} คน ({tl['Share']}%)", "#14b8a6"),
                    unsafe_allow_html=True)
with cc:
    if ts is not None:
        st.markdown(callout("ข้อเสนอแนะที่ AI เสนอบ่อยที่สุด", f"{ts['Theme']} — {int(ts['Count'])} ข้อความ", "#f7c61f"),
                    unsafe_allow_html=True)

st.write("")
i1, i2 = st.columns(2)

with i1:
    if sugg:
        fig = px.bar(st_tbl.sort_values("Count"), x="Count", y="Theme", orientation="h", text="Count",
                     color_discrete_sequence=["#10b981"], title="กลุ่มข้อเสนอแนะจาก AI (จำนวนข้อความ)")
        fig.update_yaxes(title=None)
        st.plotly_chart(style(fig, 360), use_container_width=True)
        with st.expander(f"ดูข้อเสนอแนะทั้งหมด ({len(set(sugg))} รายการไม่ซ้ำ)"):
            for t in sorted(set(sugg)):
                st.markdown(f"- {t}")
    else:
        st.info("ไม่มีข้อเสนอแนะ")

with i2:
    pos_tbl = theme_table(df, "pos_themes", overall_neg) if "positive_feedback" in df.columns else pd.DataFrame()
    if pos_tbl.empty:
        st.info("ไม่มีข้อมูลข้อดี")
    else:
        fig = px.bar(pos_tbl.sort_values("Count"), x="Count", y="Theme", orientation="h", text="Count",
                     color_discrete_sequence=[NAVY], title="ข้อดีที่ผู้ตอบกล่าวถึง (จำนวนผู้ตอบ)")
        fig.update_yaxes(title=None)
        st.plotly_chart(style(fig, 360), use_container_width=True)

st.markdown("##### 💡 AI Summary รายผู้ตอบ")
sel_s = st.radio("แสดงเฉพาะ", ["ทั้งหมด"] + [s for s in SENT_ORDER if counts[s] > 0],
                 horizontal=True, label_visibility="collapsed")
view_s = df if sel_s == "ทั้งหมด" else df[df["sentiment"] == sel_s]
for _, r in view_s.head(30).iterrows():
    with st.expander(f"{r.get('response_id', '')} · {r['sentiment']}"):
        st.write(r["summary"])

# ==========================
# 7) Data Explorer
# ==========================

section("explorer", "ค้นหาและกรองข้อมูลดิบ")

f1, f2, f3 = st.columns([2, 3, 3])
with f1:
    sel = st.multiselect("Sentiment", SENT_ORDER, default=[s for s in SENT_ORDER if counts[s] > 0])
with f2:
    loc_opts = sorted({g for lst in df["loc_groups"] for g in lst})
    sel_loc = st.multiselect("พื้นที่", loc_opts)
with f3:
    kw = st.text_input("ค้นหาข้อความ", placeholder="🔍 พิมพ์คำที่ต้องการค้นหา...")

view = df[df["sentiment"].isin(sel)]
if sel_loc:
    view = view[view["loc_groups"].apply(lambda x: any(g in x for g in sel_loc))]
if kw:
    text_cols = [c for c in raw.columns if raw[c].dtype == object]
    mask = view[text_cols].astype(str).apply(
        lambda col: col.str.contains(kw, case=False, na=False, regex=False)
    ).any(axis=1)
    view = view[mask]

out = view.drop(columns=["n_issues", "issue_themes", "loc_groups", "pos_themes"], errors="ignore")
st.caption(f"แสดง {len(out):,} จาก {total:,} รายการ")
st.dataframe(out, use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ ดาวน์โหลด CSV",
    out.to_csv(index=False).encode("utf-8-sig"),
    file_name="student_feedback.csv",
    mime="text/csv",
)

st.markdown(
    '<div class="footer"><b>ระบบวิเคราะห์ความคิดเห็นการใช้บริการ MSU Net / Wi-Fi ด้วย AI</b><br>'
    "Developed by Juthamat Saiaphai | AI + n8n Dashboard</div>",
    unsafe_allow_html=True,
)
