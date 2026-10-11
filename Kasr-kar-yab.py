import streamlit as st
import pandas as pd
import io

# تابع نرمال‌سازی کدهای پرسنلی (تبدیل اعداد فارسی به انگلیسی و حذف صفرهای قبل از عدد)
def normalize_code(text):
    if pd.isna(text):
        return ""
    text = str(text).strip()
    persian_to_english = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')
    arabic_to_english = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
    text = text.translate(persian_to_english).translate(arabic_to_english)
    # حذف صفرهای پیشرو (مثلا 000100 با 100 یکی در نظر گرفته می‌شود)
    text = text.lstrip('0')
    return text

# تنظیمات اولیه صفحه
st.set_page_config(page_title="استخراج کسر کار پرسنل", layout="wide")
st.title("جستجو و استخراج کسر کار ماهانه و چرخشی")

st.markdown("فایل **گزارش کسر کار** را آپلود کنید و کدهای پرسنلی مد نظر را وارد نمایید.")

# ایجاد دو ستون برای چیدمان بهتر
col1, col2 = st.columns([1, 2])

with col1:
    uploaded_file = st.file_uploader("آپلود فایل اکسل (xlsx, xls)", type=['xlsx', 'xls'])
    personnel_codes_input = st.text_area("کدهای پرسنلی را وارد کنید (هر کد در یک خط):", height=250)
    search_btn = st.button("جستجو و تهیه خروجی", use_container_width=True)

with col2:
    if search_btn:
        if uploaded_file is not None and personnel_codes_input.strip():
            try:
                # خواندن ۲۰ خط اول برای پیدا کردن هوشمند ردیف هدر (سرتیترها)
                temp_df = pd.read_excel(uploaded_file, nrows=20)
                
                header_row_idx = None
                for idx, row in temp_df.iterrows():
                    # اگر کلمه کسرکار چرخشی در ردیف بود، آن ردیف هدر است
                    row_str = row.astype(str).str.replace(' ', '')
                    if row_str.str.contains('کسرکارچرخشی').any():
                        header_row_idx = idx
                        break
                
                if header_row_idx is not None:
                    # برگرداندن پوینتر فایل به ابتدا و خواندن دیتای اصلی از ردیف پیدا شده
                    uploaded_file.seek(0)
                    df = pd.read_excel(uploaded_file, header=header_row_idx + 1)
                    
                    # حذف ستون‌های خالی (Unnamed)
                    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                    
                    # پیدا کردن نام دقیق ستون کد پرسنلی (ممکن است شامل کاراکترهای اضافی مانند \n باشد)
                    emp_col = [col for col in df.columns if 'کد پرسنلی' in col][0]
                    
                    # اعمال نرمال‌سازی روی کدهای پرسنلی فایل اکسل
                    df['کد_پرسنلی_جستجو'] = df[emp_col].apply(normalize_code)
                    
                    # پردازش کدهای ورودی کاربر
                    input_codes = [normalize_code(code) for code in personnel_codes_input.split('\n') if code.strip()]
                    
                    # فیلتر کردن دیتافریم بر اساس کدهای وارد شده
                    filtered_df = df[df['کد_پرسنلی_جستجو'].isin(input_codes)]
                    
                    if not filtered_df.empty:
                        # پیدا کردن دقیق ستون‌های مورد نیاز برای نمایش
                        cols_to_show = []
                        target_cols = [emp_col, 'نام', 'نام خانوادگی', 'عنوان ماه', 'تاریخ', 'کسرکار ماهانه', 'کسرکار چرخشی']
                        
                        for target in target_cols:
                            # جستجوی ستون بدون حساسیت به فاصله
                            matches = [c for c in df.columns if target.replace(' ', '') in c.replace(' ', '')]
                            if matches:
                                cols_to_show.append(matches[0])
                        
                        # آماده‌سازی دیتافریم نهایی (بدون ستون جستجوی کمکی)
                        final_df = filtered_df[cols_to_show]
                        
                        st.success(f"✅ تعداد {len(final_df)} رکورد یافت شد.")
                        st.dataframe(final_df, use_container_width=True)
                        
                        # ساخت فایل اکسل خروجی در مموری
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            final_df.to_excel(writer, index=False, sheet_name='نتیجه کسر کار')
                        
                        st.download_button(
                            label="📥 دانلود فایل اکسل خروجی",
                            data=output.getvalue(),
                            file_name="Kasre_Kar_Result.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.warning("هیچ رکوردی با کدهای پرسنلی وارد شده در فایل یافت نشد.")
                else:
                    st.error("ساختار فایل اکسل معتبر نیست. ستون 'کسرکار چرخشی' یافت نشد.")
                    
            except Exception as e:
                st.error(f"خطایی در پردازش فایل رخ داد: {e}")
        else:
            st.warning("لطفاً فایل اکسل را آپلود کرده و حداقل یک کد پرسنلی وارد کنید.")
