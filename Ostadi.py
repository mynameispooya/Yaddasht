import streamlit as st
import pandas as pd
import io

# تنظیمات اولیه صفحه
st.set_page_config(page_title="تطبیق کد ملی پرسنل", layout="wide", page_icon="📊")

st.title("اپلیکیشن افزودن کد ملی به لیست پرداخت")
st.write("این برنامه بر اساس شماره پرسنلی مشترک، کد ملی را از فایل اول استخراج کرده و به فایل دوم اضافه می‌کند.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("فایل اول: کدملی پرسنل")
    file1 = st.file_uploader("فایل حاوی کدملی (مانند کدملی پرسنل امیر.xlsx) را آپلود کنید", type=['xlsx', 'xls'])

with col2:
    st.subheader("فایل دوم: لیست نهایی پرداخت")
    file2 = st.file_uploader("فایل مقصد (مانند لیست نهایی پرداخت.xlsx) را آپلود کنید", type=['xlsx', 'xls'])

if file1 and file2:
    try:
        # خواندن فایل‌های اکسل
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        
        # بررسی وجود ستون‌های کلیدی در هر دو فایل
        if 'کد پرسنلی' not in df1.columns or 'کد ملی' not in df1.columns:
            st.error("فایل اول باید حتماً دارای ستون‌های 'کد پرسنلی' و 'کد ملی' باشد.")
        elif 'شماره پرسنلی' not in df2.columns:
            st.error("فایل دوم باید حتماً دارای ستون 'شماره پرسنلی' باشد.")
        else:
            # یکسان‌سازی فرمت داده‌های پرسنلی (تبدیل به رشته و حذف فاصله‌های اضافی و صفرهای اعشاری)
            df1['کد پرسنلی'] = df1['کد پرسنلی'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df2['شماره پرسنلی'] = df2['شماره پرسنلی'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # ایجاد یک دیکشنری نگاشت (Mapping) از فایل اول: {کد پرسنلی: کد ملی}
            # با استفاده از این روش، ترتیب فایل دوم به هیچ وجه تغییر نمی‌کند
            mapping_dict = df1.set_index('کد پرسنلی')['کد ملی'].to_dict()
            
            # اضافه کردن ستون کد ملی به فایل دوم
            df2['کد ملی'] = df2['شماره پرسنلی'].map(mapping_dict)
            
            st.success("✅ پردازش با موفقیت انجام شد! کدهای ملی به فایل دوم اضافه شدند.")
            
            # نمایش پیش‌نمایش داده‌های خروجی
            st.write("پیش‌نمایش ۵ سطر اول فایل خروجی:")
            st.dataframe(df2.head())
            
            # آماده‌سازی فایل خروجی برای دانلود
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df2.to_excel(writer, index=False, sheet_name='Result')
            processed_data = output.getvalue()
            
            # دکمه دانلود
            st.download_button(
                label="📥 دانلود فایل اکسل نهایی",
                data=processed_data,
                file_name="لیست_نهایی_همراه_با_کدملی.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"خطایی در پردازش فایل‌ها رخ داد: {e}")
