# -*- coding: utf-8 -*-
"""
اپلیکیشن گزارش تردد روزانه - پاکسازی، تشخیص تردد ناقص و تردد مکرر
اجرا:  streamlit run app.py
پیش‌نیاز:  pip install streamlit pandas openpyxl
"""
import re
import io
import datetime as dt
import pandas as pd
import streamlit as st

# ------------------------------------------------------------------
# ابزارهای نرمال‌سازی (ارقام فارسی/عربی، زمان، تاریخ، کد پرسنلی)
# ------------------------------------------------------------------
FA = "۰۱۲۳۴۵۶۷۸۹"
AR = "٠١٢٣٤٥٦٧٨٩"

def norm(v):
    """تبدیل هر مقدار به رشته با ارقام انگلیسی و حذف نیم‌فاصله"""
    if v is None:
        return ""
    s = str(v)
    for i, d in enumerate("0123456789"):
        s = s.replace(FA[i], d).replace(AR[i], d)
    s = s.replace("\u200c", " ").replace("\u066b", ":")
    return s.strip()

def parse_time(v):
    """تبدیل مقدار سلول به obj.time یا None"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, dt.datetime):
        return v.time()
    if isinstance(v, dt.time):
        return v
    m = re.search(r"(\d{1,2}):(\d{1,2})", norm(v))
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if h < 24 and mi < 60:
            return dt.time(h, mi)
    return None

def fmt_time(t):
    return t.strftime("%H:%M") if t is not None else ""

def parse_date(v):
    """استخراج تاریخ به شکل استاندارد YYYY/MM/DD"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y/%m/%d")
    s = norm(v).replace("-", "/").replace(".", "/")
    m = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})", s)
    if m:
        return f"{m.group(1)}/{int(m.group(2)):02d}/{int(m.group(3)):02d}"
    return s or None

def parse_code(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    m = re.search(r"\d{3,}", norm(v).replace(" ", ""))
    return m.group(0) if m else None

# ------------------------------------------------------------------
# ۱) تشخیص ساختار فایل (ردیف هدر + ادغام «اطلاعات تردد»)
# ------------------------------------------------------------------
def find_structure(raw):
    hdr = None
    for i in range(min(25, len(raw))):
        joined = " ".join(norm(x) for x in raw.iloc[i].tolist())
        if "خانوادگی" in joined and "کد" in joined and "تاریخ" in joined:
            hdr = i
            break
    if hdr is None:
        return None, None
    upper = [norm(x) for x in raw.iloc[hdr - 1].tolist()] if hdr > 0 else []
    names = []
    for j in range(raw.shape[1]):
        sub = norm(raw.iat[hdr, j]) if j < raw.shape[1] else ""
        up = upper[j] if j < len(upper) else ""
        if sub:                       # زیرستون واقعی (ورود گیت، خروج گیت، ...)
            names.append(sub)
        elif up and "اطلاعات" not in up:   # مقدار ادغام‌شده ردیف بالاتر (شروع/پایان شیفت)
            names.append(up)
        else:
            names.append("")
    return hdr, names

def map_columns(names):
    c = {"code": None, "name": None, "date": None, "day": None,
         "ss": None, "es": None, "in": None, "out": None, "status": None,
         "gates": []}
    for j, nm in enumerate(names):
        n = nm.strip()
        if not n:
            continue
        if "فرع" in n:
            continue
        elif "شروع" in n:
            c["ss"] = j
        elif "پایان" in n:
            c["es"] = j
        elif "ورود" in n:
            c["in"] = j
        elif "خروج" in n:
            c["out"] = j
        elif "وضعیت" in n:
            c["status"] = j
        elif "گیت" in n:
            c["gates"].append(j)
        elif "کد" in n:
            c["code"] = j
        elif "خانوادگی" in n or n == "نام":
            c["name"] = j
        elif "تاریخ" in n:
            c["date"] = j
        elif n == "روز":
            c["day"] = j
    # اتصال ستون‌های «گیت» به نزدیک‌ترین ستون ورود/خروج
    c["gate_in"], c["gate_out"] = None, None
    gates = sorted(c["gates"])
    if len(gates) == 1:
        g = gates[0]
        if c["in"] is not None and (c["out"] is None or abs(g - c["in"]) <= abs(g - c["out"])):
            c["gate_in"] = g
        else:
            c["gate_out"] = g
    elif len(gates) >= 2:
        dists = sorted(gates, key=lambda g: abs(g - (c["in"] if c["in"] is not None else 0)))
        c["gate_in"] = dists[0]
        rest = [g for g in gates if g != c["gate_in"]]
        c["gate_out"] = min(rest, key=lambda g: abs(g - (c["out"] if c["out"] is not None else 0))) if rest and c["out"] is not None else (rest[0] if rest else None)
    return c

# ------------------------------------------------------------------
# ۲) استخراج رکوردهای تردد با مدیریت سلول‌های ادغام‌شده (forward-fill)
# ------------------------------------------------------------------
def extract_records(raw, hdr, c):
    recs = []
    ctx = {"code": None, "name": None, "date": None, "day": None, "ss": None, "es": None}

    def val(r, j):
        if j is None or j >= raw.shape[1]:
            return None
        return raw.iat[r, j]

    for r in range(hdr + 1, len(raw)):
        name = norm(val(r, c["name"])) or None
        if name and ("شرکت" in name or "گزارش" in name or "کد" in name):
            continue  # ردیف‌های عنوان تکراری
        code = parse_code(val(r, c["code"]))
        if code:
            ctx["code"] = code
        if name:
            ctx["name"] = name
        d = parse_date(val(r, c["date"]))
        if d:
            ctx["date"] = d
        day = norm(val(r, c["day"])) or None
        if day:
            ctx["day"] = day
        ss = parse_time(val(r, c["ss"]))
        if ss:
            ctx["ss"] = ss
        es = parse_time(val(r, c["es"]))
        if es:
            ctx["es"] = es
        if not ctx["code"] or not ctx["date"]:
            continue

        t_in, t_out = parse_time(val(r, c["in"])), parse_time(val(r, c["out"]))
        if t_in is None and t_out is None:
            continue
        gin = norm(val(r, c["gate_in"])) if c["gate_in"] is not None else ""
        gout = norm(val(r, c["gate_out"])) if c["gate_out"] is not None else ""
        status = norm(val(r, c["status"])) if c["status"] is not None else ""
        base = {"کد پرسنلی": ctx["code"], "نام و نام خانوادگی": ctx["name"] or "",
                "تاریخ": ctx["date"], "روز": ctx["day"] or "",
                "شروع شیفت": fmt_time(ctx["ss"]), "پایان شیفت": fmt_time(ctx["es"]),
                "وضعیت": status}
        if t_in:
            recs.append({**base, "kind": "ورود", "time": t_in, "gate": gin})
        if t_out:
            recs.append({**base, "kind": "خروج", "time": t_out, "gate": gout})
    return recs

# ------------------------------------------------------------------
# ۳) طبقه‌بندی: تردد مکرر / تردد ناقص / تضاد
# ------------------------------------------------------------------
def classify_single(rec):
    """تشخیص نوع نقص برای رکوردهای تکی با مقایسه با شیفت"""
    t, kind, ss, es = rec["time"], rec["kind"], rec["ss"], rec["es"]
    m = lambda x: x.hour * 60 + x.minute
    if ss is None and es is None:
        return "خروج ثبت نشده است" if kind == "ورود" else "ورود ثبت نشده است"
    night = (ss is not None and es is not None and m(es) < m(ss))  # شیفت شبانه
    if kind == "ورود":
        if night:
            legitimate = m(t) >= m(ss) or m(t) <= m(es)
        else:
            legitimate = es is None or m(t) < m(es)
        return "خروج ثبت نشده است" if legitimate else "ورود ثبت نشده است (زمان ثبت‌شده در واقع خروج است)"
    else:  # خروج
        if night:
            legitimate = m(t) >= m(ss) or m(t) <= m(es)
        else:
            legitimate = ss is None or m(t) > m(ss)
        return "ورود ثبت نشده است" if legitimate else "خروج ثبت نشده است (زمان ثبت‌شده در واقع ورود است)"

def classify(recs):
    incomplete, multiple = [], []
    groups = {}
    for rec in recs:
        groups.setdefault((rec["کد پرسنلی"], rec["نام و نام خانوادگی"], rec["تاریخ"]), []).append(rec)

    for (code, name, date), rs in groups.items():
        rs = sorted(rs, key=lambda x: x["time"])
        ins = [r for r in rs if r["kind"] == "ورود"]
        outs = [r for r in rs if r["kind"] == "خروج"]
        base = {"کد پرسنلی": code, "نام و نام خانوادگی": name, "تاریخ": date,
                "روز": rs[0]["روز"], "شروع شیفت": rs[0]["شروع شیفت"], "پایان شیفت": rs[0]["پایان شیفت"]}

        # --- تردد مکرر: بیش از یک ورود یا بیش از یک خروج در یک تاریخ ---
        if len(ins) > 1 or len(outs) > 1:
            for r in rs:
                multiple.append({**base, "نوع": r["kind"], "زمان": fmt_time(r["time"]),
                                 "گیت": r["gate"], "وضعیت": r["وضعیت"],
                                 "تعداد ورود": len(ins), "تعداد خروج": len(outs)})
        # --- تضاد: ورود و خروج با زمان کاملاً برابر ---
        elif len(ins) == 1 and len(outs) == 1 and ins[0]["time"] == outs[0]["time"]:
            for r in rs:
                incomplete.append({**base, "نوع ثبت شده": r["kind"], "زمان ثبت شده": fmt_time(r["time"]),
                                   "گیت": r["gate"], "وضعیت": r["وضعیت"] or "تضاد",
                                   "شرح نقص": "تضاد: زمان ورود و خروج یکسان ثبت شده است"})
        # --- تردد ناقص: فقط یک رکورد ---
        elif len(rs) == 1:
            r = rs[0]
            incomplete.append({**base, "نوع ثبت شده": r["kind"], "زمان ثبت شده": fmt_time(r["time"]),
                               "گیت": r["gate"], "وضعیت": r["وضعیت"],
                               "شرح نقص": classify_single(r)})
        # در غیر این صورت (یک ورود + یک خروج معمولی) تردد کامل است
    return incomplete, multiple

# ------------------------------------------------------------------
# ۴) رابط کاربری Streamlit
# ------------------------------------------------------------------
st.set_page_config(page_title="گزارش تردد روزانه", page_icon="📋", layout="wide")
st.markdown("""<style>
.main, .stApp { direction: rtl; text-align: right; }
div[data-testid="stDataFrame"] { direction: ltr; }
h1, h2, h3 { text-align: right; }
</style>""", unsafe_allow_html=True)

st.title("📋 سامانه بررسی تردد روزانه پرسنل")
st.caption("پاکسازی اطلاعات تردد، تشخیص ورود/خروج ناقص و ترددهای مکرر بر اساس شیفت")

uploaded = st.file_uploader("فایل اکسل گزارش تردد را آپلود کنید", type=["xlsx", "xls"])

if uploaded:
    try:
        raw = pd.read_excel(uploaded, header=None, engine="openpyxl", sheet_name=0)
    except Exception as e:
        st.error(f"خطا در خواندن فایل: {e}")
        st.stop()

    hdr, names = find_structure(raw)
    if hdr is None:
        st.error("ردیف هدر (شامل «کد پرسنلی» و «نام و نام خانوادگی») پیدا نشد. ساختار فایل را بررسی کنید.")
        st.stop()

    cols = map_columns(names)
    missing = [k for k in ["code", "name", "date", "in", "out"] if cols[k] is None]
    if missing:
        st.error(f"ستون‌های ضروری شناسایی نشدند: {missing} — هدرهای کشف‌شده: {names}")
        st.stop()

    with st.spinner("در حال پاکسازی و تحلیل داده‌ها..."):
        recs = extract_records(raw, hdr, cols)
        incomplete, multiple = classify(recs)

        df_inc = pd.DataFrame(incomplete)
        df_mul = pd.DataFrame(multiple)
        df_clean = pd.DataFrame([{**r, "زمان": fmt_time(r["time"]), "نوع": r["kind"],
                                  "گیت": r["gate"]} for r in recs])
        df_clean = df_clean[["کد پرسنلی", "نام و نام خانوادگی", "تاریخ", "روز",
                             "شروع شیفت", "پایان شیفت", "نوع", "زمان", "گیت", "وضعیت"]]

    st.success(f"✅ پردازش انجام شد — {df_clean['کد پرسنلی'].nunique()} پرسنل، "
               f"{len(df_clean)} رکورد تردد، {len(df_inc)} مورد ناقص، {len(df_mul)} رکورد مکرر")

    c1, c2, c3 = st.columns(3)
    c1.metric("👥 تعداد پرسنل", df_clean["کد پرسنلی"].nunique())
    c2.metric("⚠️ تردد ناقص", len(df_inc))
    c3.metric("🔁 تردد مکرر (افراد)", df_mul[["کد پرسنلی", "تاریخ"]].drop_duplicates().shape[0] if len(df_mul) else 0)

    # --- جستجو ---
    st.subheader("🔎 جستجوی پرسنل")
    q = st.text_input("کد پرسنلی یا نام را وارد کنید (خروجی‌های زیر فیلتر می‌شوند):").strip()
    def flt(df):
        if not q or df.empty:
            return df
        mask = (df["کد پرسنلی"].astype(str).str.contains(q, na=False) |
                df["نام و نام خانوادگی"].astype(str).str.contains(q, na=False))
        return df[mask]

    tab1, tab2, tab3 = st.tabs(["⚠️ تردد ناقص", "🔁 تردد مکرر", "🧹 داده پاکسازی‌شده"])

    with tab1:
        st.dataframe(flt(df_inc), use_container_width=True, hide_index=True)
        if len(df_inc):
            st.download_button("⬇️ دانلود اکسل تردد ناقص",
                               to_excel_bytes({"تردد ناقص": df_inc}),
                               "tardad_naghes.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        st.dataframe(flt(df_mul), use_container_width=True, hide_index=True)
        if len(df_mul):
            st.download_button("⬇️ دانلود اکسل تردد مکرر",
                               to_excel_bytes({"تردد مکرر": df_mul}),
                               "tardad_mokarrar.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab3:
        st.dataframe(flt(df_clean), use_container_width=True, hide_index=True)

    # --- فایل ترکیبی ---
    if len(df_inc) or len(df_mul):
        sheets = {}
        if len(df_inc): sheets["تردد ناقص"] = df_inc
        if len(df_mul): sheets["تردد مکرر"] = df_mul
        st.download_button("⬇️ دانلود فایل کامل (هر دو گزارش)",
                           to_excel_bytes(sheets),
                           "gozaresh_tardad.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("لطفاً فایل اکسل را آپلود کنید.") 
