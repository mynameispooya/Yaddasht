import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="بررسی و تجمیع گزارش اضافه کاری", layout="wide")

# تابع هوشمند برای تبدیل زمان (HH:MM یا عدد) به دقیقه جهت محاسبات دقیق
def parse_time_to_minutes(time_val):
    if pd.isna(time_val): return 0
    time_val = str(time_val).strip()
    if ':' in time_val:
        try:
            h, m = time_val.split(':')
            return int(h) * 60 + int(m)
        except:
            return 0
    else:
        try:
            return float(time_val) * 60  # اگر ساعت به صورت اعشاری بود
        except:
            return 0

# تابع برای برگرداندن دقیقه به فرمت HH:MM
def minutes_to_time(minutes):
    if pd.isna(minutes): return "00:00"
    minutes = int(minutes)
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"

# تابع هایلایت کردن مغایرت‌ها
def highlight_diff(row):
    color = 'background-color: #ffcccc' if row['مغایرت دارد'] == 'بله' else 'background-color: #ccffcc'
    return [color] * len(row)

st.title("📊 سیستم تجمیع و بررسی مغایرت اضافه کاری")
st.write("لطفا فایل‌های اکسل خود را به ترتیب زیر آپلود کنید:")

# مرحله 1: آپلود فایل اول
st.header("مرحله ۱: آپلود گزارش جامع کارکرد (فایل اول)")
file1 = st.file_uploader("فایل اول را انتخاب کنید (مثلاً: گزارش جامع کارکرد مرداد 1405 (1).xlsx)", type=['xlsx', 'xls'], key='file1')

# مرحله 2: آپلود فایل دوم
st.header("مرحله ۲: آپلود گزارش تجمیعی اضافه کاری (فایل دوم)")
file2 = st.file_uploader("فایل دوم را انتخاب کنید (مثلاً: گزارش_تجمیعی_اضافه_کاری (1).xlsx)", type=['xlsx', 'xls'], key='file2')

if file1 and file2:
    try:
        # خواندن فایل ها
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # استخراج ستون‌های مورد نیاز و پاکسازی فضاهای خالی در نام ستون‌ها
        df1.columns = df1.columns.str.strip()
        df2.columns = df2.columns.str.strip()

        # بررسی وجود ستون‌ها در فایل اول
        cols1 = ['عنوان واحد سازمانی', 'نام', 'نام خانوادگی', 'کارکرد اضافه کاری']
        if not all(col in df1.columns for col in cols1):
            st.error(f"❌ خطا: فایل اول باید حتما شامل این ستون‌ها باشد: {cols1}")
            st.stop()
            
        # بررسی وجود ستون‌ها در فایل دوم
        cols2 = ['کد پرسنلی', 'نام و نام خانوادگی', 'مجموع اضافه کار کلی']
        if not all(col in df2.columns for col in cols2):
            st.error(f"❌ خطا: فایل دوم باید حتما شامل این ستون‌ها باشد: {cols2}")
            st.stop()

        df1_sub = df1[cols1].copy()
        df2_sub = df2[cols2].copy()

        # همسان‌سازی نوع داده برای مقایسه دقیق
        df1_sub['عنوان واحد سازمانی'] = df1_sub['عنوان واحد سازمانی'].astype(str).str.strip()
        df2_sub['کد پرسنلی'] = df2_sub['کد پرسنلی'].astype(str).str.strip()

        # ادغام دو جدول بر اساس عنوان واحد سازمانی (فایل 1) و کد پرسنلی (فایل 2)
        merged_df = pd.merge(df1_sub, df2_sub, left_on='عنوان واحد سازمانی', right_on='کد پرسنلی', how='inner')

        if merged_df.empty:
            st.warning("⚠️ هیچ رکورد مشترکی بین 'عنوان واحد سازمانی' در فایل اول و 'کد پرسنلی' در فایل دوم یافت نشد.")
            st.stop()

        # محاسبه مجموع اضافه کاری
        merged_df['دقیقه_اضافه_فایل1'] = merged_df['کارکرد اضافه کاری'].apply(parse_time_to_minutes)
        merged_df['دقیقه_اضافه_فایل2'] = merged_df['مجموع اضافه کار کلی'].apply(parse_time_to_minutes)
        merged_df['مجموع نهایی اضافه کاری (دقیقه)'] = merged_df['دقیقه_اضافه_فایل1'] + merged_df['دقیقه_اضافه_فایل2']
        merged_df['مجموع نهایی اضافه کاری'] = merged_df['مجموع نهایی اضافه کاری (دقیقه)'].apply(minutes_to_time)

        # ساخت فایل نهایی طبق خواسته شما (درج کامل نام و نام خانوادگی از فایل دوم)
        final_df = pd.DataFrame({
            'کد پرسنلی': merged_df['کد پرسنلی'],
            'نام': merged_df['نام'],
            'نام و نام خانوادگی': merged_df['نام و نام خانوادگی'],  # کل محتوای ستون از فایل دوم
            'مجموع نهایی اضافه کاری': merged_df['مجموع نهایی اضافه کاری']
        })

        st.success("✅ ادغام و محاسبات با موفقیت انجام شد!")
        st.dataframe(final_df)

        # مرحله 3: آپلود فایل سوم برای مقایسه
        st.header("مرحله ۳: بررسی و مقایسه با فایل جدید (فایل سوم)")
        file3 = st.file_uploader("فایل سوم را جهت مقایسه آپلود کنید (مثلاً: گزارش جامع کارکرد مرداد جدید.xlsx)", type=['xlsx', 'xls'], key='file3')

        if file3:
            df3 = pd.read_excel(file3)
            df3.columns = df3.columns.str.strip()

            if 'عنوان واحد سازمانی' not in df3.columns or 'کارکرد اضافه کاری' not in df3.columns:
                st.error("❌ فایل سوم باید دارای ستون‌های 'عنوان واحد سازمانی' (به عنوان کد پرسنلی) و 'کارکرد اضافه کاری' باشد.")
            else:
                df3_sub = df3[['عنوان واحد سازمانی', 'کارکرد اضافه کاری']].copy()
                df3_sub['عنوان واحد سازمانی'] = df3_sub['عنوان واحد سازمانی'].astype(str).str.strip()
                df3_sub['کارکرد فایل جدید (دقیقه)'] = df3_sub['کارکرد اضافه کاری'].apply(parse_time_to_minutes)
                df3_sub.rename(columns={'کارکرد اضافه کاری': 'کارکرد در فایل جدید (سوم)'}, inplace=True)

                # ادغام برای مقایسه
                comparison_df = pd.merge(final_df, df3_sub, left_on='کد پرسنلی', right_on='عنوان واحد سازمانی', how='left')
                
                # بررسی مغایرت
                comparison_df['مجموع محاسبه شده ما (دقیقه)'] = merged_df['مجموع نهایی اضافه کاری (دقیقه)']
                comparison_df['مغایرت دارد'] = comparison_df.apply(
                    lambda row: 'خیر' if abs(row['مجموع محاسبه شده ما (دقیقه)'] - row['کارکرد فایل جدید (دقیقه)']) < 2 else 'بله', axis=1
                ) # اختلاف 2 دقیقه ای را برای اطمینان از خطای گرد کردن نادیده میگیریم

                # تمیز کردن جدول برای نمایش
                display_cols = ['کد پرسنلی', 'نام', 'نام و نام خانوادگی', 'مجموع نهایی اضافه کاری', 'کارکرد در فایل جدید (سوم)', 'مغایرت دارد']
                final_comparison = comparison_df[display_cols]

                st.subheader("🔍 نتیجه بررسی مغایرت‌ها (ردیف‌های قرمز دارای مغایرت هستند):")
                # اعمال استایل هایلایت
                styled_df = final_comparison.style.apply(highlight_diff, axis=1)
                st.dataframe(styled_df, use_container_width=True)

                # امکان دانلود فایل نهایی مقایسه شده
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

