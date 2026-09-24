import streamlit as st
import pandas as pd
import re
from io import BytesIO

# تنظیمات اولیه صفحه استریم‌لیت
st.set_page_config(page_title="پردازش تردد ناقص", layout="centered")

st.title("ابزار پردازش لیست تردد ناقص (خروجی اکسل)")
st.write("فایل اکسل خود را آپلود کنید. پیام‌های هر شخص تجمیع شده و به صورت یک فایل اکسل جدید و مرتب قابل دانلود خواهد بود.")

# آپلود فایل
uploaded_file = st.file_uploader("انتخاب فایل اکسل (لیست-تردد-ناقص.xlsx)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # خواندن فایل اکسل
        df = pd.read_excel(uploaded_file)
        
        # بررسی وجود ستون‌های مورد نیاز
        if 'عنوان' not in df.columns or 'پیغام' not in df.columns:
            st.error("فایل آپلود شده باید حتماً دارای ستون‌های 'عنوان' و 'پیغام' باشد.")
        else:
            # حذف ردیف‌هایی که ستون عنوان در آن‌ها خالی است
            df = df.dropna(subset=['عنوان'])
            
            processed_data = []
            
            # گروه‌بندی داده‌ها بر اساس ستون 'عنوان'
            grouped = df.groupby('عنوان', sort=False)
            
            for title, group in grouped:
                person_messages = []
                
                # پیمایش تمام پیام‌های مربوط به این شخص در تاریخ‌های مختلف
                for msg in group['پیغام']:
                    if pd.isna(msg):
                        continue
                    
                    # شکستن پیام‌ها بر اساس تگ </br> یا <br>
                    parts = re.split(r'</br>|<br\s*/?>', str(msg), flags=re.IGNORECASE)
                    
                    for part in parts:
                        clean_part = part.strip()
                        if clean_part:
                            person_messages.append(clean_part)
                            
                # چسباندن تمام پیام‌های شخص با کاراکتر خط جدید (\n)
                # تا در یک سلول اکسل زیر هم قرار بگیرند
                final_message = "\n".join(person_messages)
                
                # اضافه کردن به لیست داده‌های نهایی
                processed_data.append({
                    'عنوان': str(title).strip(),
                    'پیغام': final_message
                })
            
            # تبدیل لیست پردازش شده به یک دیتافریم جدید
            out_df = pd.DataFrame(processed_data)
            
            st.success("پردازش فایل با موفقیت انجام شد!")
            
            # نمایش پیش‌نمایش جدول در اپلیکیشن
            st.write("**پیش‌نمایش داده‌های خروجی:**")
            st.dataframe(out_df, use_container_width=True)
            
            # ایجاد فایل اکسل در حافظه موقت (RAM) برای دانلود
            output = BytesIO()
            # استفاده از xlsxwriter برای تنظیمات ظاهر اکسل (مانند Wrap Text)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                out_df.to_excel(writer, index=False, sheet_name='تردد ناقص')
                
                workbook = writer.book
                worksheet = writer.sheets['تردد ناقص']
                
                # فعال کردن قابلیت شکستن متن (Wrap Text) برای خوانایی پیام‌های چند خطی در اکسل
                wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
                
                # تنظیم عرض ستون‌ها
                worksheet.set_column('A:A', 25, wrap_format) # ستون عنوان
                worksheet.set_column('B:B', 90, wrap_format) # ستون پیغام
                
                # اعمال استایل به هدرها
                for col_num, value in enumerate(out_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)

            processed_excel = output.getvalue()
            
            st.write("---")
            # قرار دادن دکمه دانلود فایل اکسل
            st.download_button(
                label="📥 دانلود فایل اکسل نهایی",
                data=processed_excel,
                file_name="Processed_Attendance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"خطایی در حین خواندن یا پردازش فایل رخ داد: {e}")
