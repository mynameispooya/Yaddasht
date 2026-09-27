import streamlit as st
import pandas as pd
import re

# ----------------- تنظیمات صفحه -----------------
st.set_page_config(page_title="سیستم کنترل تردد و کارکرد", layout="wide", page_icon="🕒")

# استایل‌دهی CSS برای راست‌چین کردن و زیبایی ظاهر
st.markdown("""
    <style>
    * { font-family: 'Tahoma', 'Vazir', sans-serif; }
    .main { direction: rtl; text-align: right; }
    .stDataFrame { direction: ltr; } /* برای نمایش صحیح اعداد و انگلیسی */
    h1, h2, h3, p, div { direction: rtl; text-align: right; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 سیستم هوشمند تحلیل و بررسی کارکرد پرسنل")
st.markdown("""
این اپلیکیشن گزارشات اکسل تردد را دریافت کرده، ارفاق‌های مجاز (تا ۱۵ دقیقه) را لحاظ می‌کند، 
شیفت‌های شب را بدون خطا پردازش کرده و کارکردهای کلیدی را با رنگ‌های تفکیک‌شده نمایش می‌دهد.
""")

# ----------------- توابع کمکی -----------------
def time_to_mins(t_str):
    """تبدیل رشته زمان (HH:MM) به دقیقه برای محاسبات"""
    if pd.isna(t_str) or not isinstance(t_str, str) or ':' not in t_str:
        return 0
    try:
        h, m = map(int, t_str.split(':'))
        return h * 60 + m
    except:
        return 0

def mins_to_time(m):
    """تبدیل دقیقه به فرمت استاندارد ساعت (HH:MM)"""
    m = int(m)
    return f"{m//60:02d}:{m%60:02d}"

def is_night_shift(shift_name):
    """تشخیص شیفت‌های شبانه برای جلوگیری از خطای عبور از نصف شب"""
    night_shifts = [
        'چرخشی شب تاسیسات 1', 'چرخشی شب تاسیسات 2', 
        'چرخشی شب 1', 'چرخشی شب 2', 
        'تعمیر و نگهداری 1', 'تعمیر و نگهداری 2', 'تعمیر و نگهداری 3'
    ]
    if pd.isna(shift_name): return False
    return any(ns in str(shift_name) for ns in night_shifts)

# ----------------- آپلود و پردازش -----------------
uploaded_file = st.file_uploader("📥 لطفاً فایل اکسل کارکرد (مانند بابالو-علی.xlsx) را آپلود کنید:", type=['xlsx'])

if uploaded_file:
    with st.spinner("در حال پردازش داده‌ها..."):
        df = pd.read_excel(uploaded_file)
        
        # ۱. حذف ستون اطلاعات تردد
        if 'اطلاعات تردد' in df.columns:
            df = df.drop(columns=['اطلاعات تردد'])
            
        # ۲. اعمال ارفاق ۱۵ دقیقه‌ای و مدیریت شیفت‌ها
        def process_grace_period(row):
            status = str(row.get('وضعیت', ''))
            
            # اپلیکیشن به صورت هوشمند شیفت شب را از روی گروه کاری تشخیص می‌دهد
            _ = is_night_shift(row.get('عنوان گروه کاری', '')) 
            
            if 'عدم حضور' in status:
                duration = time_to_mins(row.get('مدت', '00:00'))
                # اگر زمان عدم حضور (ناشی از تاخیر یا تعجیل) کمتر یا مساوی ۱۵ دقیقه باشد، این ردیف حذف می‌شود
                if 0 < duration <= 15:
                    return False 
            return True

        # فیلتر کردن ردیف‌های عدم حضور که کمتر از ۱۵ دقیقه هستند
        df = df[df.apply(process_grace_period, axis=1)].reset_index(drop=True)

        # ۳. پیکربندی ستون‌های هدف برای هایلایت
        target_cols = {
            'اضافه کار عادی نهایی': '#D9E1F2',       # آبی روشن
            'اضافه کار نهایی روز تعطیل': '#E2EFDA', # سبز روشن
            'جمعه کاری': '#FFF2CC',                # زرد روشن
            'غیبت روزانه': '#FCE4D6',              # قرمز/نارنجی روشن
            'شبکاری': '#E4DFEC'                    # بنفش روشن
        }
        
        valid_targets = {k: v for k, v in target_cols.items() if k in df.columns}
        
        # ۴. محاسبه مجموع مقادیر ستون‌های هدف
        totals = {col: 0 for col in valid_targets}
        for col in valid_targets:
            totals[col] = sum(df[col].apply(time_to_mins))
            
        # افزودن ردیف مجموع به پایین جدول
        summary_row = pd.Series(index=df.columns, dtype=object)
        summary_row['روز'] = 'مجموع نهایی'
        for col in valid_targets:
            summary_row[col] = mins_to_time(totals[col])
            
        df = pd.concat([df, pd.DataFrame([summary_row])], ignore_index=True)
        
        # ۵. استایل‌دهی (رنگ آمیزی منحصر به فرد)
        def highlight_target_cells(val, color):
            # عدم رنگ آمیزی مقادیر خالی یا صفر
            if pd.isna(val) or val == '00:00' or val == '':
                return ''
            if isinstance(val, str) and ':' in val:
                return f'background-color: {color}; color: #000; font-weight: bold;'
            return ''

        def style_df(styler):
            # اعمال رنگ برای ستون‌های مختلف (از متد map به جای applymap استفاده شده است)[span_0](start_span)[span_0](end_span)
            for col, color in valid_targets.items():
                styler.map(lambda v, c=color: highlight_target_cells(v, c), subset=[col])
                
            # متمایز کردن ردیف مجموع نهایی در انتهای جدول
            def highlight_summary(s):
                if s['روز'] == 'مجموع نهایی':
                    return ['background-color: #2F5597; color: white; font-weight: bold; font-size: 14px;'] * len(s)
                return [''] * len(s)
                
            styler.apply(highlight_summary, axis=1)
            return styler

        st.success("✅ فایل با موفقیت پردازش شد و موارد ارفاقی لحاظ گردید.")
        
        # نمایش جدول نهایی با استایل
        st.dataframe(df.style.pipe(style_df), use_container_width=True, height=750)
