import streamlit as st
import pandas as pd
import re

# تنظیمات صفحه اپلیکیشن
st.set_page_config(page_title="تحلیل مرخصی استعلاجی", page_icon="💊", layout="wide")

def clean_and_load_excel(uploaded_file):
    """
    این تابع فایل اکسل را می‌خواند و مشکل ستون 'اطلاعات تردد' و هدرهای چندگانه را حل می‌کند.
    """
    # خواندن کل فایل بدون در نظر گرفتن هدر برای پیدا کردن ردیف اصلی تیترها
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    header_idx = -1
    # جستجو برای پیدا کردن ردیفی که شامل 'وضعیت' و 'تاریخ' است (ردیف اصلی تیترها)
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if "وضعیت" in row_str and "تاریخ" in row_str and "کدپرسنلی" in row_str:
            header_idx = idx
            break
            
    if header_idx == -1:
        st.error(f"ساختار فایل {uploaded_file.name} معتبر نیست. ردیف تیترها پیدا نشد.")
        return None
        
    # تنظیم ردیف پیدا شده به عنوان هدر و حذف ردیف‌های اضافی بالای آن (مانند 'اطلاعات تردد')
    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = df_raw.iloc[header_idx].values
    
    # حذف ردیف‌هایی که کد پرسنلی یا تاریخ ندارند (ردیف‌های خالی انتهای فایل)
    df = df.dropna(subset=['کدپرسنلی', 'تاریخ'])
    
    # مرتب‌سازی نام ستون‌ها برای جلوگیری از وجود فاصله (Space) اضافه
    df.columns = [str(col).strip() for col in df.columns]
    
    return df

def is_pure_sick_leave(statuses):
    """
    این تابع بررسی می‌کند که آیا در وضعیت‌های ثبت شده برای یک روز، 
    فقط و فقط "مرخصی استعلاجی" وجود دارد یا با چیز دیگری ترکیب شده است.
    """
    # فیلتر کردن مقادیر خالی
    valid_statuses = [str(s) for s in statuses if pd.notna(s) and str(s).strip() not in ('', 'nan')]
    if not valid_statuses:
        return False
        
    combined_status = " ".join(valid_statuses)
    
    # اگر اصلا کلمه استعلاجی در آن روز نیست، پس فالس برگردان
    if "استعلاجی" not in combined_status:
        return False
        
    # حذف کلمات مربوط به مرخصی استعلاجی برای بررسی اینکه آیا کلمه دیگری باقی می‌ماند یا خیر
    cleaned = combined_status.replace("مرخصی", "").replace("استعلاجی", "")
    
    # حذف تمام کاراکترهای غیرحروفی (مثل فاصله‌ها، خط تیره، اسلش، اعداد و...)
    cleaned = re.sub(r'\W+', '', cleaned)
    cleaned = re.sub(r'\d+', '', cleaned)
    
    # اگر بعد از حذف 'مرخصی' و 'استعلاجی'، کلمه دیگری (مثل 'استحقاقی'، 'حضور' و...) 
    # در متن باقی مانده باشد، یعنی خالص نیست و نباید شمرده شود.
    if len(cleaned.strip()) > 0:
        return False
        
    return True

# رابط کاربری Streamlit
st.title("📊 سیستم تحلیل و شمارش مرخصی‌های استعلاجی")
st.markdown("""
این برنامه فایل‌های اکسل حضور و غیاب (با ساختار مشابه `تعداد-استعلاجی.xlsx`) را دریافت کرده، 
هدرهای آن را اصلاح می‌کند و به صورت هوشمند مرخصی‌های استعلاجی خالص را استخراج می‌نماید.
روزهایی که استعلاجی با موارد دیگر (مثل استحقاقی یا حضور) ترکیب شده باشد، شمرده نخواهند شد.
""")

st.divider()

# آپلود یک یا چند فایل
uploaded_files = st.file_uploader("فایل‌های اکسل خود را اینجا آپلود کنید", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 پردازش فایل‌ها"):
        for file in uploaded_files:
            st.subheader(f"📄 نتایج پردازش فایل: {file.name}")
            
            df = clean_and_load_excel(file)
            
            if df is not None:
                results = []
                # گروه‌بندی بر اساس کدپرسنلی و نام
                grouped = df.groupby(['کدپرسنلی', 'نام و نام خانوادگی'])
                
                for (emp_code, emp_name), group in grouped:
                    valid_sick_leaves = []
                    
                    # گروه‌بندی داخلی بر اساس تاریخ برای بررسی ترکیب وضعیت‌ها در یک روز مشخص
                    date_grouped = group.groupby(['تاریخ', 'روز'])
                    
                    for (date, day), d_group in date_grouped:
                        statuses = d_group['وضعیت'].tolist()
                        
                        # بررسی شرط حیاتی: آیا این روز خالصاً استعلاجی است؟
                        if is_pure_sick_leave(statuses):
                            valid_sick_leaves.append({
                                'تاریخ': date,
                                'روز': day
                            })
                            
                    # اگر شخص مرخصی استعلاجی معتبر داشت، اطلاعاتش را ذخیره کن
                    if valid_sick_leaves:
                        # مرتب‌سازی تاریخ‌ها تا اولین و آخرین روز مشخص شود
                        valid_sick_leaves = sorted(valid_sick_leaves, key=lambda x: x['تاریخ'])
                        start_date = valid_sick_leaves[0]
                        end_date = valid_sick_leaves[-1]
                        
                        results.append({
                            'کد پرسنلی': int(float(emp_code)), # تبدیل به عدد صحیح برای زیبایی
                            'نام و نام خانوادگی': emp_name,
                            'شروع استعلاجی': f"{start_date['روز']} {start_date['تاریخ']}",
                            'پایان استعلاجی': f"{end_date['روز']} {end_date['تاریخ']}",
                            'تعداد روزهای استعلاجی': len(valid_sick_leaves)
                        })
                
                # نمایش خروجی به شکل کارتی و زیبا
                if results:
                    for res in results:
                        with st.container():
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown(f"**کد پرسنلی:** `{res['کد پرسنلی']}`")
                                st.markdown(f"**نام و نام خانوادگی:** {res['نام و نام خانوادگی']}")
                                st.markdown(f"**تعداد روز محاسبه شده:** {res['تعداد روزهای استعلاجی']} روز")
                            with col2:
                                st.info(f"**شروع استعلاجی:** {res['شروع استعلاجی']}")
                                st.error(f"**پایان استعلاجی:** {res['پایان استعلاجی']}")
                            st.divider()
                else:
                    st.warning("هیچ مرخصی استعلاجی معتبری در این فایل یافت نشد.")
