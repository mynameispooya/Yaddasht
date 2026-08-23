import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# تنظیمات صفحه و راست‌چین کردن
st.set_page_config(page_title="محاسبه اضافه کاری", layout="wide", initial_sidebar_state="collapsed")

# CSS برای رابط کاربری مینیمال و راست‌چین
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;700&display=swap');
        body, .stApp { font-family: 'Vazirmatn', sans-serif; direction: rtl; }
        .main-title { text-align: center; color: #1E3A8A; margin-bottom: 20px; }
        .metric-box { background-color: #F0F9FF; padding: 20px; border-radius: 10px; border-right: 5px solid #0EA5E9; text-align: center; }
        .metric-value { font-size: 2em; font-weight: bold; color: #0EA5E9; }
        .metric-label { font-size: 1em; color: #475569; }
        .stDataFrame { width: 100%; }
    </style>
""", unsafe_allow_html=True)

# تابع تبدیل اعداد فارسی به انگلیسی
def fa_to_en(text):
    if pd.isna(text) or text == '*':
        return np.nan
    text = str(text)
    persian = '۰۱۲۳۴۵۶۷۸۹'
    english = '0123456789'
    translation_table = str.maketrans(persian, english)
    return text.translate(translation_table)

st.markdown('<h1 class="main-title">📊 سیستم استخراج و محاسبه اضافه‌کاری</h1>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 فایل اکسل تردد پرسنل را آپلود کنید", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # خواندن فایل اکسل با در نظر گرفتن هدرهای سطر 4 و 5 (ایندکس 4 و 5 در پایتون)
        df = pd.read_excel(uploaded_file, header=[4, 5])
        
        # ترکیب هدرهای چندلایه (MultiIndex) به یک هدر واحد و قابل خواندن
        df.columns = ['_'.join(str(c).strip() for c in col if str(c) != 'nan' and 'Unnamed' not in str(c)).strip() for col in df.columns.values]
        
        # پیدا کردن نام ستون‌های مورد نیاز بر اساس کلمات کلیدی
        entry_col = [c for c in df.columns if 'ورود' in c and 'گیت' in c]
        exit_col = [c for c in df.columns if 'خروج' in c and 'گیت' in c]
        code_col = [c for c in df.columns if 'کدپرسنلی' in c]
        name_col = [c for c in df.columns if 'نام' in c and 'خانوادگی' in c]

        if not entry_col or not exit_col or not code_col:
            st.error("ستون‌های مورد نیاز (ورود، خروج، کدپرسنلی) در فایل یافت نشدند. لطفا ساختار فایل را بررسی کنید.")
        else:
            # ساخت دیتافریم نهایی
            cols_to_keep = [code_col[0]] + ([name_col[0]] if name_col else []) + [entry_col[0], exit_col[0]]
            clean_df = df[cols_to_keep].copy()
            
            # نام‌گذاری مجدد ستون‌ها
            new_cols = ['کدپرسنلی', 'نام و نام خانوادگی', 'ورود', 'خروج'] if name_col else ['کدپرسنلی', 'ورود', 'خروج']
            clean_df.columns = new_cols
            
            # حذف ردیف‌های خالی و ردیف‌های "مجموع"
            clean_df = clean_df.dropna(subset=['کدپرسنلی'])
            clean_df = clean_df[~clean_df['کدپرسنلی'].astype(str).str.contains('مجموع', na=False)]
            
            # تبدیل اعداد فارسی به انگلیسی
            clean_df['کدپرسنلی'] = clean_df['کدپرسنلی'].apply(fa_to_en)
            clean_df['ورود'] = clean_df['ورود'].apply(fa_to_en)
            clean_df['خروج'] = clean_df['خروج'].apply(fa_to_en)
            
            # تبدیل زمان‌ها به فرمت datetime
            clean_df['ورود'] = pd.to_datetime(clean_df['ورود'], format='%H:%M', errors='coerce')
            clean_df['خروج'] = pd.to_datetime(clean_df['خروج'], format='%H:%M', errors='coerce')
            
            # حذف ردیف‌هایی که زمان ورود یا خروج ندارند
            clean_df = clean_df.dropna(subset=['ورود', 'خروج'])
            
            # محاسبه تفاضل زمان (با در نظر گرفتن شیفت‌های شبانه که خروج در روز بعد است)
            clean_df['تفاضل زمانی'] = np.where(
                clean_df['خروج'] < clean_df['ورود'],
                (clean_df['خروج'] + pd.Timedelta(days=1)) - clean_df['ورود'],
                clean_df['خروج'] - clean_df['ورود']
            )
            
            # قالب‌بندی تفاضل به صورت ساعت:دقیقه
            clean_df['اضافه کاری روز تعطیل'] = clean_df['تفاضل زمانی'].apply(
                lambda x: f"{int(x.total_seconds() // 3600):02d}:{int((x.total_seconds() % 3600) // 60):02d}"
            )
            
            # انتخاب ستون‌های نهایی برای نمایش
            final_cols = ['کدپرسنلی', 'نام و نام خانوادگی', 'ورود', 'خروج', 'اضافه کاری روز تعطیل'] if name_col else ['کدپرسنلی', 'ورود', 'خروج', 'اضافه کاری روز تعطیل']
            final_df = clean_df[final_cols].copy()
            
            # تبدیل مجدد زمان‌ها به فرمت رشته‌ای برای نمایش تمیزتر
            final_df['ورود'] = final_df['ورود'].dt.strftime('%H:%M')
            final_df['خروج'] = final_df['خروج'].dt.strftime('%H:%M')
            
            # محاسبه تعداد کدهای پرسنلی یکتا
            unique_count = final_df['کدپرسنلی'].nunique()

            # نمایش بخش اطلاعات کلی
            st.markdown("""
                <div class="metric-box">
                    <div class="metric-label">تعداد کل کدهای پرسنلی یکتا</div>
                    <div class="metric-value">{}</div>
                </div>
            """.format(unique_count), unsafe_allow_html=True)
            
            st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)
            
            # نمایش جدول داده‌ها
            st.markdown("### 📋 جدول استخراج شده")
            st.dataframe(final_df, use_container_width=True, height=500)
            
            # امکان دانلود فایل اکسل
            @st.cache_data
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='گزارش اضافه کاری')
                return output.getvalue()

            excel_data = convert_df_to_excel(final_df)
            
            st.download_button(
                label="⬇️ دانلود فایل اکسل خروجی",
                data=excel_data,
                file_name="overtime_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"خطا در پردازش فایل: {str(e)}")
        st.info("لطفا مطمئن شوید که فایل آپلود شده دقیقا مطابق ساختار سیستم تردد است.")
