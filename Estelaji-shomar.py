#2
import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="تحلیل مرخصی استعلاجی", page_icon="💊", layout="wide")

def clean_and_load_excel(uploaded_file):
    """
    این تابع به صورت داینامیک فایل اکسل را بررسی کرده و ستون‌ها را بر اساس نامشان پیدا می‌کند.
    این روش مشکل فایل‌های دارای هدرهای ترکیبی (Multi-Index) را به طور کامل حل می‌کند.
    """
    # خواندن فایل بدون فرض هیچ هدری
    df = pd.read_excel(uploaded_file, header=None)
    
    col_map = {}
    status_row_idx = -1
    
    # جستجو در تمام ردیف‌ها برای پیدا کردن شماره ستون (Index) هر تیتر
    for idx, row in df.iterrows():
        row_vals = row.dropna().astype(str).values
        # پیدا کردن ردیفی که شامل کلمات کلیدی تاریخ و پرسنلی است
        if any('کدپرسنلی' in val.replace(' ', '') for val in row_vals):
            for col_idx, val in enumerate(row.values):
                if pd.notna(val):
                    v_str = str(val).strip().replace(' ', '')
                    if 'کدپرسنلی' in v_str:
                        col_map['کدپرسنلی'] = col_idx
                    elif 'تاریخ' in v_str and 'تاریخ' not in col_map: 
                        col_map['تاریخ'] = col_idx
                    elif 'نامونامخانوادگی' in v_str:
                        col_map['نام و نام خانوادگی'] = col_idx
                    elif 'روز' == v_str: 
                        col_map['روز'] = col_idx
                        
        # پیدا کردن ردیف وضعیت که ممکن است در سطر پایین‌تر (به خاطر ادغام سلول‌ها) باشد
        if any('وضعیت' in val for val in row_vals):
            for col_idx, val in enumerate(row.values):
                if pd.notna(val) and 'وضعیت' in str(val).strip():
                    col_map['وضعیت'] = col_idx
                    status_row_idx = idx
                    break
    
    # بررسی اینکه آیا همه ستون‌های حیاتی پیدا شدند؟
    required_keys = ['وضعیت', 'روز', 'تاریخ', 'نام و نام خانوادگی', 'کدپرسنلی']
    if not all(k in col_map for k in required_keys) or status_row_idx == -1:
        missing = [k for k in required_keys if k not in col_map]
        st.error(f"ساختار فایل '{uploaded_file.name}' معتبر نیست. ستون‌های زیر پیدا نشدند: {', '.join(missing)}")
        return None
        
    # برش داده‌ها از یک ردیف بعد از کلمه 'وضعیت'
    data = df.iloc[status_row_idx + 1:].copy()
    
    # فقط ستون‌های شناسایی شده استخراج می‌شوند
    data = data[list(col_map.values())]
    data.columns = list(col_map.keys())
    
    # پر کردن سلول‌های ادغام شده (Merged Cells) که در پانداس به شکل خالی می‌افتند
    data['تاریخ'] = data['تاریخ'].ffill()
    data['کدپرسنلی'] = data['کدپرسنلی'].ffill()
    data['نام و نام خانوادگی'] = data['نام و نام خانوادگی'].ffill()
    data['روز'] = data['روز'].ffill()
    
    # حذف ردیف‌هایی که وضعیت در آن‌ها خالی است
    data = data.dropna(subset=['وضعیت'])
    
    return data

def is_pure_sick_leave(statuses):
    """
    بررسی می‌کند که آیا در یک روز خاص فقط 'مرخصی استعلاجی' ثبت شده است 
    یا اینکه با موارد دیگر (استحقاقی، حضور و ...) ترکیب شده است.
    """
    valid_statuses = [str(s) for s in statuses if pd.notna(s) and str(s).strip() not in ('', 'nan')]
    if not valid_statuses:
        return False
        
    combined_status = " ".join(valid_statuses)
    
    if "استعلاجی" not in combined_status:
        return False
        
    # کلمات مرخصی و استعلاجی حذف می‌شوند
    cleaned = combined_status.replace("مرخصی", "").replace("استعلاجی", "")
    cleaned = re.sub(r'\W+', '', cleaned)
    cleaned = re.sub(r'\d+', '', cleaned)
    
    # اگر بعد از حذف استعلاجی، کلمه‌ای مثل حضور یا استحقاقی بماند، False برمی‌گردد
    if len(cleaned.strip()) > 0:
        return False
        
    return True

st.title("📊 سیستم تحلیل و شمارش مرخصی‌های استعلاجی")
st.markdown("""
این برنامه به صورت هوشمند ستون‌ها را حتی در فایل‌های **Multi-Index (هدرهای تو در تو و سلول‌های ادغام شده)** تشخیص می‌دهد.
روزهایی که استعلاجی با موارد دیگر ترکیب شده باشد (مثل استحقاقی یا حضور)، به عنوان استعلاجی شمرده **نخواهند شد**.
""")

st.divider()

uploaded_files = st.file_uploader("فایل‌های اکسل خود را اینجا آپلود کنید", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 پردازش فایل‌ها"):
        for file in uploaded_files:
            st.subheader(f"📄 نتایج پردازش فایل: {file.name}")
            
            df = clean_and_load_excel(file)
            
            if df is not None:
                results = []
                grouped = df.groupby(['کدپرسنلی', 'نام و نام خانوادگی'])
                
                for (emp_code, emp_name), group in grouped:
                    valid_sick_leaves = []
                    date_grouped = group.groupby(['تاریخ', 'روز'])
                    
                    for (date, day), d_group in date_grouped:
                        statuses = d_group['وضعیت'].tolist()
                        
                        # بررسی شرط خلوص مرخصی استعلاجی
                        if is_pure_sick_leave(statuses):
                            valid_sick_leaves.append({
                                'تاریخ': date,
                                'روز': day
                            })
                            
                    if valid_sick_leaves:
                        valid_sick_leaves = sorted(valid_sick_leaves, key=lambda x: x['تاریخ'])
                        start_date = valid_sick_leaves[0]
                        end_date = valid_sick_leaves[-1]
                        
                        try:
                            # حذف اعشار از کد پرسنلی اگر به صورت float خوانده شده
                            clean_emp_code = int(float(emp_code))
                        except ValueError:
                            clean_emp_code = emp_code
                            
                        results.append({
                            'کد پرسنلی': clean_emp_code,
                            'نام و نام خانوادگی': emp_name,
                            'شروع استعلاجی': f"{start_date['روز']} {start_date['تاریخ']}",
                            'پایان استعلاجی': f"{end_date['روز']} {end_date['تاریخ']}",
                            'تعداد روزهای استعلاجی': len(valid_sick_leaves)
                        })
                
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
