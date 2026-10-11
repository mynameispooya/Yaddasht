import streamlit as st
import openpyxl
from copy import copy
import io

# تنظیمات اولیه صفحه
st.set_page_config(page_title="انتقال هایلایت اکسل", layout="centered")
st.title("انتقال هایلایت‌ها بین دو فایل اکسل")
st.markdown("فایل اول (دارای هایلایت) و فایل دوم (بدون هایلایت) را آپلود کنید. این برنامه رنگ پس‌زمینه سلول‌های فایل اول را دقیقاً به همان مکان در فایل دوم منتقل می‌کند.")

# ایجاد دو کادر برای آپلود فایل‌ها
col1, col2 = st.columns(2)

with col1:
    file_colored = st.file_uploader("۱. فایل مبدأ (رنگ‌دار) را آپلود کنید", type=['xlsx'])

with col2:
    file_plain = st.file_uploader("۲. فایل مقصد (بدون رنگ) را آپلود کنید", type=['xlsx'])

if file_colored and file_plain:
    st.divider()
    if st.button("انتقال رنگ‌ها و ایجاد فایل خروجی", type="primary", use_container_width=True):
        try:
            with st.spinner("در حال پردازش و انتقال رنگ‌ها..."):
                # بارگذاری هر دو ورک‌بوک با استفاده از openpyxl
                # پارامتر data_only=False باعث می‌شود فرمول‌ها و استایل‌ها حفظ شوند
                wb_colored = openpyxl.load_workbook(file_colored, data_only=False)
                wb_plain = openpyxl.load_workbook(file_plain, data_only=False)
                
                # فرض می‌کنیم کار روی شیت اول (فعال) انجام می‌شود
                ws_colored = wb_colored.active
                ws_plain = wb_plain.active
                
                # پیمایش همزمان ردیف‌ها و سلول‌های هر دو شیت
                for row_col, row_pln in zip(ws_colored.iter_rows(), ws_plain.iter_rows()):
                    for cell_col, cell_pln in zip(row_col, row_pln):
                        # اگر سلول در فایل اول دارای رنگ پس‌زمینه (fill) باشد
                        if cell_col.fill and cell_col.fill.fill_type:
                            # کپی کردن استایل رنگ به سلول متناظر در فایل دوم
                            cell_pln.fill = copy(cell_col.fill)
                
                # ذخیره فایل مقصدِ تغییر یافته در مموری (BytesIO)
                output = io.BytesIO()
                wb_plain.save(output)
                output.seek(0)
                
            st.success("✅ انتقال رنگ‌ها با موفقیت انجام شد!")
            
            # دکمه دانلود فایل نهایی
            st.download_button(
                label="📥 دانلود فایل اکسل نهایی (همراه با هایلایت‌ها)",
                data=output,
                file_name="Highlighted_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"خطایی رخ داد: {e}")
