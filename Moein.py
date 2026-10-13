# -*- coding: utf-8 -*-
import re
import io
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ============================================================
#                     تنظیمات صفحه و ظاهر (RTL)
# ============================================================
st.set_page_config(page_title="پاکسازی غیبت‌ها (تردد ناقص)", page_icon="🧹", layout="wide")

st.markdown("""
<style>
    .main, .main * { direction: rtl; text-align: right; font-family: Vazirmatn, Tahoma, sans-serif; }
    .stDataFrame, .stDataFrame * { direction: ltr; }
    div[data-testid="stMetric"] { background:#f5f7fa; border-radius:10px; padding:10px; }
</style>
""", unsafe_allow_html=True)

st.title("🧹 پاکسازی فایل غیبت بر اساس ترددهای ناقص")
st.caption("ابتدا دو فایل اکسل را در سایدبار آپلود کنید، سپس تنظیمات هر بخش را بررسی و تحلیل را اجرا کنید.")

# ============================================================
#                       توابع کمکی
# ============================================================
FA_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
AR_DIGITS = "٠١٢٣٤٥٦٧٨٩"
DIGIT_MAP = {ord(f): str(i) for i, f in enumerate(FA_DIGITS)}
DIGIT_MAP.update({ord(a): str(i) for i, a in enumerate(AR_DIGITS)})

def norm(x) -> str:
    """نرمال‌سازی متن: تبدیل ارقام فارسی/عربی به انگلیسی + حذف نویسه‌های مخفی"""
    if x is None:
        return ""
    s = str(x).translate(DIGIT_MAP)
    return s.replace("\u200c", "").replace("\u200f", "").replace("\u200e", "").strip()

def extract_code(val):
    """استخراج اولین رشته‌ی عددی ۳ رقم به بالا (کد پرسنلی) از یک مقدار"""
    s = norm(val).replace(",", "").replace("٫", "")
    m = re.search(r"(\d{3,})", s)
    return m.group(1) if m else None

def extract_dates(val):
    """استخراج تاریخ‌ها با الگوی 1405/07/08 (با ارقام نرمال‌شده)"""
    return re.findall(r"\d{4}/\d{1,2}/\d{1,2}", norm(val))

# تگ‌های خط جدید: <br> ، <br/> ، </br> ، <\br> و نسخه‌های با فاصله
BR_RE = re.compile(r"<\s*[/\\]?\s*br\s*[/\\]?\s*>", re.I)

def split_br(val):
    return [p.strip() for p in BR_RE.split(norm(val)) if p.strip()]

def find_col(cols, keywords, exclude=None):
    """پیدا کردن اولین ستونی که شامل کلیدواژه باشد"""
    exclude = exclude or []
    for c in cols:
        cc = norm(c)
        if any(k in cc for k in keywords) and not any(e in cc for e in exclude):
            return c
    return None

def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    """تخت‌کردن ستون‌های مولتی‌ایندکس به یک سطح خوانا"""
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for tup in df.columns:
            parts = [str(p).strip() for p in tup
                     if p is not None and str(p).strip()
                     and not str(p).startswith("Unnamed")]
            new_cols.append(" - ".join(parts) if parts else "ستون بدون نام")
        df.columns = new_cols
    else:
        df.columns = [norm(c) if norm(c) else f"ستون {i+1}" for i, c in enumerate(df.columns)]
    # حذف نام‌های تکراری
    seen, out = {}, []
    for c in df.columns:
        if c in seen:
            seen[c] += 1
            out.append(f"{c} ({seen[c]})")
        else:
            seen[c] = 0
            out.append(c)
    df.columns = out
    return df

YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# ============================================================
#                پردازش فایل «تردد ناقص» (مرحله ۱)
# ============================================================
def process_tarefi(df: pd.DataFrame, title_col: str, msg_col: str) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        title = norm(r.get(title_col))
        msg   = norm(r.get(msg_col))
        code  = extract_code(title)
        # نام = بخش قبل از خط تیره (با حذف عبارت «کد پرسنلی» در صورت چسبیدگی)
        name = re.split(r"\s*[-–—]\s*", title)[0] if title else ""
        name = re.sub(r"کد\s*پرسنلی.*", "", name).strip()

        segments = split_br(msg)
        dates = []
        if segments:
            for seg in segments:               # هر بخش بعد از <br> یک ردیف جدا
                d = extract_dates(seg)
                if d:
                    dates.append(d[0])
        else:                                  # پیغام بدون تگ <br>
            dates = extract_dates(msg)[:1]

        if not dates:
            rows.append({"عنوان": name, "کد پرسنلی": code, "تاریخ": None})
        else:
            for d in dates:
                rows.append({"عنوان": name, "کد پرسنلی": code, "تاریخ": d})
    return pd.DataFrame(rows)

# ============================================================
#        کار با openpyxl برای حفظ ساختار مولتی‌ایندکس غیبت
# ============================================================
def detect_ws_columns(ws):
    """تشخیص خودکار ردیف هدر، ستون کد پرسنلی، ستون تاریخ و شروع داده‌ها"""
    header_row, code_col = None, None
    for r in range(1, min(15, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and "کد پرسنلی" in norm(v):
                header_row, code_col = r, c
                break
        if header_row:
            break
    if header_row is None:
        return None

    date_col = None
    for r in range(header_row, min(header_row + 2, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = norm(ws.cell(r, c).value)
            if "پایان غیبت" in v or "تاریخ" in v:
                date_col = c
                break
        if date_col:
            break

    data_start = header_row + 1
    # اگر ردیف بعد از هدر، هدر دوم (سطح زیرستون‌ها) باشد، یک ردیف پایین‌تر برو
    if not extract_code(ws.cell(data_start, code_col).value) and \
       not extract_dates(ws.cell(data_start, date_col).value if date_col else None) and \
       extract_code(ws.cell(data_start + 1, code_col).value):
        data_start += 1

    return {"header_row": header_row, "code_col": code_col,
            "date_col": date_col, "data_start": data_start}

def iterate_ghyabat_rows(ws, info):
    """پیمایش ردیف‌های داده غیبت با forward-fill روی کد پرسنلی (سلول‌های ادغام‌شده)"""
    last_code = None
    for r in range(info["data_start"], ws.max_row + 1):
        c = extract_code(ws.cell(r, info["code_col"]).value)
        if c:
            last_code = c
        code = c or last_code
        dv = extract_dates(ws.cell(r, info["date_col"]).value) if info["date_col"] else []
        yield r, code, (dv[0] if dv else None)

def build_highlighted(file_bytes, sheet_name, pairs):
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    info = detect_ws_columns(ws)
    if not info or not info["date_col"]:
        return None, None, 0
    n = 0
    for r, code, d in iterate_ghyabat_rows(ws, info):
        if code and d and (code, d) in pairs:
            n += 1
            for cidx in range(1, ws.max_column + 1):
                ws.cell(r, cidx).fill = YELLOW
    return wb, info, n

def build_cleaned(file_bytes, sheet_name, pairs):
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name]
    info = detect_ws_columns(ws)
    if not info or not info["date_col"]:
        return None, None, 0
    to_delete = [r for r, code, d in iterate_ghyabat_rows(ws, info)
                 if code and d and (code, d) in pairs]
    for r in reversed(to_delete):          # حذف از پایین به بالا
        ws.delete_rows(r)
    return wb, info, len(to_delete)

def wb_to_bytes(wb):
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()

# ============================================================
#                        سایدبار: آپلود
# ============================================================
with st.sidebar:
    st.header("📂 آپلود فایل‌ها")
    ghyabat_file = st.file_uploader("فایل غیبت (مولتی‌ایندکس)", type=["xlsx", "xls"])
    tarefi_file  = st.file_uploader("فایل لیست تردد ناقص", type=["xlsx", "xls"])

if not ghyabat_file or not tarefi_file:
    st.info("⬅️ لطفاً هر دو فایل اکسل را آپلود کنید تا تحلیل آغاز شود.")
    st.stop()

# ============================================================
#                 بخش ۱: پردازش فایل تردد ناقص
# ============================================================
st.header("۱️⃣ پردازش فایل تردد ناقص")

tf = pd.ExcelFile(tarefi_file)
tarefi_sheet = st.selectbox("شیت فایل تردد ناقص", tf.sheet_names,
                            index=len(tf.sheet_names) - 1)
tarefi_header_rows = st.number_input("تعداد ردیف‌های هدر فایل تردد", 0, 3, 0)

t_df = pd.read_excel(tf, sheet_name=tarefi_sheet, header=int(tarefi_header_rows))
t_df = flatten_columns(t_df)
st.write(f"**{len(t_df)}** ردیف و **{len(t_df.columns)}** ستون شناسایی شد.")

t_cols = list(t_df.columns)
title_col = st.selectbox("ستون «عنوان»", t_cols,
                         index=t_cols.index(find_col(t_cols, ["عنوان"]) or t_cols[0]))
msg_col = st.selectbox("ستون «پیغام»", t_cols,
                       index=t_cols.index(find_col(t_cols, ["پیغام", "پیام"]) or t_cols[-1]))

if st.button("🚀 پردازش فایل تردد ناقص", type="primary"):
    st.session_state["t_expanded"] = process_tarefi(t_df, title_col, msg_col)

if "t_expanded" in st.session_state:
    t_exp = st.session_state["t_expanded"]
    t_exp = t_exp.dropna(subset=["کد پرسنلی"])
    st.success(f"✅ پس از تفکیک عنوان و بازکردن تگ‌های <br>، تعداد ردیف‌ها به {len(t_exp)} رسید.")
    with st.expander("👁 مشاهده جدول تفکیک‌شده (عنوان / کد پرسنلی / تاریخ)"):
        st.dataframe(t_exp, use_container_width=True)

    n_dates = t_exp["تاریخ"].notna().sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("کد پرسنلی یکتا", t_exp["کد پرسنلی"].nunique())
    c2.metric("تاریخ استخراج‌شده", int(n_dates))
    c3.metric("تاریخ یکتا", t_exp["تاریخ"].dropna().nunique())

    # نقشه‌ی کد پرسنلی → نام برای استفاده در خروجی نهایی
    name_map = (t_exp.dropna(subset=["کد پرسنلی"])
                     .drop_duplicates("کد پرسنلی")
                     .set_index("کد پرسنلی")["عنوان"]
                     .to_dict())

    # جفت‌های (کد پرسنلی، تاریخ) برای مقایسه
    pairs = set(zip(t_exp["کد پرسنلی"], t_exp["تاریخ"]))
    pairs = {(c, d) for c, d in pairs if c and d}
else:
    st.warning("برای ادامه، دکمه‌ی پردازش فایل تردد را بزنید.")
    st.stop()

# ============================================================
#                 بخش ۲: پردازش فایل غیبت (مولتی‌ایندکس)
# ============================================================
st.header("۲️⃣ پردازش فایل غیبت و مقایسه")

gf = pd.ExcelFile(ghyabat_file)
ghyabat_sheet = st.selectbox("شیت فایل غیبت", gf.sheet_names)
g_header_start = st.number_input("ردیف شروع هدر (اگر بالای جدول، عنوان گزارش دارید)",
                                 0, 10, 0)
g_header_count = st.selectbox("تعداد سطح هدر", [1, 2], index=1)

header_arg = (list(range(int(g_header_start), int(g_header_start) + g_header_count))
              if g_header_count > 1 else int(g_header_start))
g_df = pd.read_excel(gf, sheet_name=ghyabat_sheet, header=header_arg)
g_df = flatten_columns(g_df)
st.write(f"**{len(g_df)}** ردیف و **{len(g_df.columns)}** ستون از فایل غیبت خوانده شد.")

g_cols = list(g_df.columns)
code_col_def = find_col(g_cols, ["کد پرسنلی"]) or g_cols[0]
date_col_def = find_col(g_cols, ["پایان غیبت"], exclude=["شرح"]) or \
               find_col(g_cols, ["تاریخ"]) or g_cols[-1]
roozane_def  = find_col(g_cols, ["غیبت روزانه", "روزانه"]) or g_cols[-1]

code_col   = st.selectbox("ستون «کارمند کد پرسنلی»", g_cols, index=g_cols.index(code_col_def))
date_col_g = st.selectbox("ستون «تاریخ» (برای مقایسه با تاریخ تردد)", g_cols,
                          index=g_cols.index(date_col_def))
roozane_col = st.selectbox("ستون «غیبت روزانه» (نوع غیبت)", g_cols,
                           index=g_cols.index(roozane_def))

# آماده‌سازی داده‌های غیبت در pandas
g_df["_code"] = g_df[code_col].map(extract_code).ffill()
g_df["_date"] = g_df[date_col_g].map(lambda v: (extract_dates(v) or [None])[0])
g_df["_matched"] = g_df.apply(
    lambda r: (r["_code"], r["_date"]) in pairs if r["_code"] and r["_date"] else False, axis=1)

with st.expander("👁 پیش‌نمایش فایل غیبت (با وضعیت تطبیق)"):
    st.dataframe(g_df.assign(تطبیق=g_df["_matched"]).drop(columns=["_code", "_date", "_matched"]),
                 use_container_width=True)

file_bytes = ghyabat_file.getvalue()

st.subheader("⚙️ اجرای تحلیل و تولید خروجی‌ها")
if st.button("🧮 اجرای تحلیل کامل", type="primary"):

    # ---------- خروجی ۱: غیبت با هایلایت ----------
    wb_h, info_h, n_high = build_highlighted(file_bytes, ghyabat_sheet, pairs)
    # ---------- خروجی ۲: غیبت پاک‌شده ----------
    wb_c, info_c, n_del = build_cleaned(file_bytes, ghyabat_sheet, pairs)

    if wb_h is None or wb_c is None:
        st.error("تشخیص خودکار ستون‌ها در فایل غیبت ناموفق بود؛ نام ستون‌ها را دستی بررسی کنید.")
        st.stop()

    # ---------- خروجی ۳: تعداد غیبت هر نفر (از فایل پاک‌شده) ----------
    cleaned = g_df[~g_df["_matched"]]
    records = cleaned[cleaned["_date"].notna()]          # فقط ردیف‌های واقعی غیبت
    summary = (records.groupby("_code").size()
               .reset_index(name="تعداد غیبت")
               .rename(columns={"_code": "کد پرسنلی"}))
    summary["نام و نام خانوادگی"] = summary["کد پرسنلی"].map(name_map).fillna("نامشخص")
    summary = summary[["کد پرسنلی", "نام و نام خانوادگی", "تعداد غیبت"]] \
        .sort_values("تعداد غیبت", ascending=False).reset_index(drop=True)

    # ---------- نمایش نتایج ----------
    st.subheader("📊 خلاصه نتایج")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("ردیف‌های هایلایت‌شده", n_high)
    m2.metric("ردیف‌های حذف‌شده", n_del)
    m3.metric("ردیف باقی‌مانده (غیبت خالص)", int(g_df["_matched"].shape[0] - g_df["_matched"].sum()))
    m4.metric("افراد دارای غیبت", len(summary))

    st.subheader("🏆 تعداد غیبت به تفکیک کارمند (نزولی)")
    st.dataframe(summary, use_container_width=True, hide_index=True)
    if len(summary):
        chart_df = summary.set_index("نام و نام خانوادگی")["تعداد غیبت"].head(25)
        st.bar_chart(chart_df)

    # ---------- دکمه‌های دانلود ----------
    st.subheader("⬇️ دانلود خروجی‌ها")
    d1, d2, d3 = st.columns(3)

    d1.download_button("📄 غیبت با هایلایت (xlsx)",
                       wb_to_bytes(wb_h),
                       file_name="غیبت_هایلایت_شده.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    d2.download_button("🧹 غیبت پاک‌شده (xlsx)",
                       wb_to_bytes(wb_c),
                       file_name="غیبت_خالص_بدون_تردد_ناقص.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # خروجی خلاصه: شیت اول = خلاصه، شیت دوم = تفکیک نوع غیبت
    sum_buf = io.BytesIO()
    with pd.ExcelWriter(sum_buf, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="تعداد غیبت", index=False)
        if roozane_col in records.columns:
            pd.crosstab(records["_code"], records[roozane_col]) \
              .rename(index=name_map).rename_axis("نام و نام خانوادگی") \
              .to_excel(w, sheet_name="تفکیک نوع غیبت")
    d3.download_button("📋 جدول تعداد غیبت‌ها (xlsx)",
                       sum_buf.getvalue(),
                       file_name="خلاصه_تعداد_غیبت.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")
    with st.expander("🔎 جزئیات فنی تشخیص ستون‌ها در فایل غیبت"):
        st.json({"ردیف هدر": info_h["header_row"],
                 "ستون کد پرسنلی": info_h["code_col"],
                 "ستون تاریخ": info_h["date_col"],
                 "شروع داده‌ها": info_h["data_start"]})
else:
    st.info("برای تولید هر سه خروجی، دکمه‌ی «اجرای تحلیل کامل» را بزنید.")
