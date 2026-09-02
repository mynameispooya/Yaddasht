# 3

import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="بررسی و تجمیع گزارش اضافه کاری", layout="wide")

# تابع تبدیل اعداد فارسی و عربی به انگلیسی
def convert_persian_to_english_numbers(text):
    if pd.isna(text): 
        return text
    persian_numbers = "۰۱۲۳۴۵۶۷۸۹"
    arabic_numbers  = "٠١٢٣٤٥٦٧٨٩"
    english_numbers = "0123456789"
    
    trans_table = str.maketrans(persian_numbers + arabic_numbers, english_numbers * 2)
    return str(text).translate(trans_table)

# تابع هوشمند برای یافتن سطر هدر (نام ستون‌ها) در فایل‌های دارای متادیتا
def load_excel_with_dynamic_header(file, required_columns):
    # خواندن ۳۰ سطر اول برای پیدا کردن هدر
    df_temp = pd.read_excel(file, header=None, nrows=30)
    header_idx = 0
    for idx, row in df_temp.iterrows():
        # تبدیل تمام مقادیر سطر به رشته
        row_values = [str(val).strip() for val in row.values if pd.notnull(val)]
        # اگر سطر فعلی شامل ستون‌های مورد نیاز ما بود، آن را هدر در نظر می‌گیریم
        if all(any(req in val for val in row_values) for req in required_columns):
            header_idx = idx
            break
            
    # بازنشانی نشانگر فایل و خواندن فایل با هدر صحیح
    file.seek(0)
    df = pd.read_excel(file, header=header_idx)
    # حذف فاصله‌های اضافی از نام ستون‌ها
    df.columns = df.columns.astype(str).str.strip()
    return df

# تابع هوشمند برای تبدیل زمان (HH:MM یا عدد) به دقیقه جهت محاسبات دقیق
def parse_time_to_minutes(time_val):
    if pd.isna(time_val): return 0
    time_val = convert_persian_to_english_numbers(str(time_val)).strip()
    if ':' in time_val:
        try:
            h, m = time_val.split(':')
            return int(h) * 60 + int(m)
        except:
            return 0
    else:
        try:
            return float(time_val) * 60  
        except:
            return 0

def minutes_to_time(minutes):
    if pd.isna(minutes): return "00:00"
    minutes = int(minutes)
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"

def highlight_diff(row):
    color = 'background-color: #ffcccc' if row['مغایرت دارد'] == 'بله' else 'background-color: #ccffcc'
    return [color] * len(row)

st.title("📊 سیستم تجمیع و بررسی مغایرت اضافه کاری")
st.write("لطفا فایل‌های اکسل خود را به ترتیب زیر آپلود کنید:")

st.header("مرحله ۱: آپلود گزارش جامع کارکرد (فایل اول)")
file1 = st.file_uploader("فایل اول را انتخاب کنید", type=['xlsx', 'xls'], key='file1')

st.header("مرحله ۲: آپلود گزارش تجمیعی اضافه کاری (فایل دوم)")
file2 = st.file_uploader("فایل دوم را انتخاب کنید", type=['xlsx', 'xls'], key='file2')

if file1 and file2:
    try:
        # خواندن هوشمند فایل‌ها با تشخیص خودکار سطر هدر
        cols1_required = ['عنوان واحد سازمانی', 'نام', 'نام خانوادگی', 'کارکرد اضافه کاری']
        cols2_required = ['کد پرسنلی', 'نام و نام خانوادگی', 'مجموع اضافه کار کلی']
        
        df1 = load_excel_with_dynamic_header(file1, ['عنوان واحد سازمانی', 'کارکرد اضافه کاری'])
        df2 = load_excel_with_dynamic_header(file2, ['کد پرسنلی', 'مجموع اضافه کار کلی'])

        if not all(col in df1.columns for col in cols1_required):
            st.error(f"❌ خطا: ستون‌های لازم در فایل اول پیدا نشدند: {cols1_required}")
            st.stop()
            
        if not all(col in df2.columns for col in cols2_required):
            st.error(f"❌ خطا: ستون‌های لازم در فایل دوم پیدا نشدند: {cols2_required}")
            st.stop()

        df1_sub = df1[cols1_required].copy()
        df2_sub = df2[cols2_required].copy()

        # همسان‌سازی نوع داده: تبدیل اعداد فارسی به انگلیسی + حذف فاصله‌ها + حذف صفرهای قبل از کد
        df1_sub['کد_پرسنلی_استاندارد'] = df1_sub['عنوان واحد سازمانی'].apply(convert_persian_to_english_numbers).astype(str).str.strip().str.lstrip('0')
        df2_sub['کد_پرسنلی_استاندارد'] = df2_sub['کد پرسنلی'].apply(convert_persian_to_english_numbers).astype(str).str.strip().str.lstrip('0')

        # ادغام بر اساس ستون استاندارد شده
        merged_df = pd.merge(df1_sub, df2_sub, on='کد_پرسنلی_استاندارد', how='inner')

        if merged_df.empty:
            st.warning("⚠️ هیچ رکورد مشترکی یافت نشد. لطفاً ساختار داده‌ها را بررسی کنید.")
            st.stop()

        merged_df['دقیقه_اضافه_فایل1'] = merged_df['کارکرد اضافه کاری'].apply(parse_time_to_minutes)
        merged_df['دقیقه_اضافه_فایل2'] = merged_df['مجموع اضافه کار کلی'].apply(parse_time_to_minutes)
        merged_df['مجموع نهایی اضافه کاری (دقیقه)'] = merged_df['دقیقه_اضافه_فایل1'] + merged_df['دقیقه_اضافه_فایل2']
        merged_df['مجموع نهایی اضافه کاری'] = merged_df['مجموع نهایی اضافه کاری (دقیقه)'].apply(minutes_to_time)

        final_df = pd.DataFrame({
            'کد پرسنلی': merged_df['کد پرسنلی'],
            'نام': merged_df['نام'],
            'نام و نام خانوادگی': merged_df['نام و نام خانوادگی'], 
            'مجموع نهایی اضافه کاری': merged_df['مجموع نهایی اضافه کاری'],
            'کد_پرسنلی_استاندارد': merged_df['کد_پرسنلی_استاندارد'] # برای مچ شدن با فایل سوم
        })

        st.success(f"✅ ادغام با موفقیت انجام شد! ({len(final_df)} رکورد مشترک یافت شد)")
        
        # نمایش بدون ستون استاندارد
        st.dataframe(final_df.drop(columns=['کد_پرسنلی_استاندارد']))

        st.header("مرحله ۳: بررسی و مقایسه با فایل جدید (فایل سوم)")
        file3 = st.file_uploader("فایل سوم را جهت مقایسه آپلود کنید", type=['xlsx', 'xls'], key='file3')

        if file3:
            df3 = load_excel_with_dynamic_header(file3, ['عنوان واحد سازمانی', 'کارکرد اضافه کاری'])

            if 'عنوان واحد سازمانی' not in df3.columns or 'کارکرد اضافه کاری' not in df3.columns:
                st.error("❌ فایل سوم باید دارای ستون‌های 'عنوان واحد سازمانی' و 'کارکرد اضافه کاری' باشد.")
            else:
                df3_sub = df3[['عنوان واحد سازمانی', 'کارکرد اضافه کاری']].copy()
                # استانداردسازی کد پرسنلی در فایل سوم
                df3_sub['کد_پرسنلی_استاندارد'] = df3_sub['عنوان واحد سازمانی'].apply(convert_persian_to_english_numbers).astype(str).str.strip().str.lstrip('0')
                df3_sub['کارکرد فایل جدید (دقیقه)'] = df3_sub['کارکرد اضافه کاری'].apply(parse_time_to_minutes)
                df3_sub.rename(columns={'کارکرد اضافه کاری': 'کارکرد در فایل جدید (سوم)'}, inplace=True)

                # ادغام فایل نهایی با فایل سوم
                comparison_df = pd.merge(final_df, df3_sub, on='کد_پرسنلی_استاندارد', how='left')
                
                comparison_df['مجموع محاسبه شده ما (دقیقه)'] = merged_df['مجموع نهایی اضافه کاری (دقیقه)']
                comparison_df['مغایرت دارد'] = comparison_df.apply(
                    lambda row: 'خیر' if abs(row['مجموع محاسبه شده ما (دقیقه)'] - row['کارکرد فایل جدید (دقیقه)']) < 2 else 'بله', axis=1
                )

                display_cols = ['کد پرسنلی', 'نام', 'نام و نام خانوادگی', 'مجموع نهایی اضافه کاری', 'کارکرد در فایل جدید (سوم)', 'مغایرت دارد']
                final_comparison = comparison_df[display_cols]

                st.subheader("🔍 نتیجه بررسی مغایرت‌ها (ردیف‌های قرمز دارای مغایرت هستند):")
                styled_df = final_comparison.style.apply(highlight_diff, axis=1)
                st.dataframe(styled_df, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_comparison.to_excel(writer, index=False, sheet_name='گزارش مغایرت')
                
                st.download_button(
                    label="📥 دانلود فایل نهایی گزارش (Excel)",
                    data=output.getvalue(),
                    file_name="Final_Overtime_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"خطایی در پردازش فایل‌ها رخ داده است: {e}")
 
