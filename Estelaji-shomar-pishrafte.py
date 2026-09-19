import streamlit as st
import pandas as pd
import re
import base64

# تنظیمات صفحه اپلیکیشن
st.set_page_config(page_title="تحلیل مرخصی استعلاجی", page_icon="💊", layout="wide")

def shamsi_to_days(shamsi_date_str):
    """
    تبدیل تاریخ شمسی به تعداد روز جهت محاسبه فاصله بین تاریخ‌ها
    برای تشخیص دقیق گپ‌ها و بخش‌بندی کردن مرخصی‌ها
    """
    try:
        # تبدیل اعداد فارسی به انگلیسی
        persian_to_eng = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')
        clean_date = str(shamsi_date_str).translate(persian_to_eng).strip()
        parts = clean_date.split('/')
        
        if len(parts) != 3: return 0
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        
        # محاسبه تقریبی روزها از یک مبدا زمانی
        y_diff = y - 1300
        days = y_diff * 365 + (y_diff // 4)
        days_in_months = [0, 31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
        days += sum(days_in_months[:m]) + d
        
        return days
    except:
        return 0

def clean_and_load_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, header=None)
    col_map = {}
    status_row_idx = -1
    
    for idx, row in df.iterrows():
        row_vals = row.dropna().astype(str).values
        if any('کدپرسنلی' in val.replace(' ', '') for val in row_vals):
            for col_idx, val in enumerate(row.values):
                if pd.notna(val):
                    v_str = str(val).strip().replace(' ', '')
                    if 'کدپرسنلی' in v_str: col_map['کدپرسنلی'] = col_idx
                    elif 'تاریخ' in v_str and 'تاریخ' not in col_map: col_map['تاریخ'] = col_idx
                    elif 'نامونامخانوادگی' in v_str: col_map['نام و نام خانوادگی'] = col_idx
                    elif 'روز' == v_str: col_map['روز'] = col_idx
                        
        if any('وضعیت' in val for val in row_vals):
            for col_idx, val in enumerate(row.values):
                if pd.notna(val) and 'وضعیت' in str(val).strip():
                    col_map['وضعیت'] = col_idx
                    status_row_idx = idx
                    break
    
    required_keys = ['وضعیت', 'روز', 'تاریخ', 'نام و نام خانوادگی', 'کدپرسنلی']
    if not all(k in col_map for k in required_keys) or status_row_idx == -1:
        return None
        
    data = df.iloc[status_row_idx + 1:].copy()
    data = data[list(col_map.values())]
    data.columns = list(col_map.keys())
    
    data['تاریخ'] = data['تاریخ'].ffill()
    data['کدپرسنلی'] = data['کدپرسنلی'].ffill()
    data['نام و نام خانوادگی'] = data['نام و نام خانوادگی'].ffill()
    data['روز'] = data['روز'].ffill()
    
    data = data.dropna(subset=['وضعیت'])
    return data

def is_pure_sick_leave(statuses):
    valid_statuses = [str(s) for s in statuses if pd.notna(s) and str(s).strip() not in ('', 'nan')]
    if not valid_statuses: return False
    
    combined_status = " ".join(valid_statuses)
    if "استعلاجی" not in combined_status: return False
        
    cleaned = combined_status.replace("مرخصی", "").replace("استعلاجی", "")
    cleaned = re.sub(r'\W+', '', cleaned)
    cleaned = re.sub(r'\d+', '', cleaned)
    
    if len(cleaned.strip()) > 0: return False
    return True

def generate_pdf_ready_html(results, filename):
    """ساخت فایل HTML با ظاهری مشابه Streamlit جهت پرینت به عنوان PDF"""
    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="fa">
    <head>
        <meta charset="UTF-8">
        <title>گزارش {filename}</title>
        <style>
            @import url('https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css');
            body {{ font-family: 'Vazirmatn', Tahoma, sans-serif; background-color: #f0f2f6; padding: 20px; color: #31333F; }}
            .container {{ max-width: 1000px; margin: auto; }}
            .header {{ text-align: center; margin-bottom: 30px; }}
            .card {{ background: white; border-radius: 10px; padding: 20px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: 1px solid #e6e6e6; page-break-inside: avoid; }}
            .title-sec {{ font-size: 1.2em; font-weight: bold; margin-bottom: 15px; border-bottom: 2px solid #f0f2f6; padding-bottom: 10px; color: #2c3e50; }}
            .badge {{ background-color: #00c2a8; color: white; padding: 3px 10px; border-radius: 12px; font-size: 0.8em; margin-right: 10px; }}
            .row {{ display: flex; flex-wrap: wrap; gap: 20px; }}
            .col {{ flex: 1; min-width: 200px; }}
            .box-start {{ background-color: #e8f4fd; color: #0056b3; padding: 12px; border-radius: 6px; margin-bottom: 8px; border-right: 4px solid #0056b3; font-size: 0.9em; }}
            .box-end {{ background-color: #fdf3f2; color: #c82333; padding: 12px; border-radius: 6px; border-right: 4px solid #c82333; font-size: 0.9em; }}
            .btn-print {{ display: block; width: 250px; margin: 0 auto 30px auto; padding: 12px; background: #ff4b4b; color: white; text-align: center; border-radius: 8px; cursor: pointer; border: none; font-family: 'Vazirmatn'; font-size: 16px; box-shadow: 0 4px 6px rgba(255,75,75,0.3); }}
            @media print {{
                body {{ background: white; }}
                .btn-print {{ display: none; }}
                .card {{ box-shadow: none; border: 1px solid #ccc; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <button class="btn-print" onclick="window.print()">🖨️ ذخیره به عنوان PDF</button>
            <h2 class="header">📊 گزارش تحلیل مرخصی‌های استعلاجی</h2>
    """
    
    for res in results:
        html += f"""
        <div class="card">
            <div class="title-sec">👤 {res['نام و نام خانوادگی']} | کد: {res['کد پرسنلی']} <span class="badge">بخش {res['بخش']}</span></div>
            <div class="row">
                <div class="col" style="display: flex; align-items: center; font-size: 1.1em;">
                    <strong>تعداد روز محاسبه شده:</strong> &nbsp; {res['تعداد روزهای استعلاجی']} روز
                </div>
                <div class="col">
                    <div class="box-start"><strong>شروع استعلاجی:</strong> {res['شروع استعلاجی']}</div>
                    <div class="box-end"><strong>پایان استعلاجی:</strong> {res['پایان استعلاجی']}</div>
                </div>
            </div>
        </div>
        """
        
    html += "</div></body></html>"
    return html

st.title("📊 سیستم تحلیل و شمارش مرخصی‌های استعلاجی")
st.markdown("اپلیکیشن هوشمند با قابلیت **بخش‌بندی اتوماتیک** استعلاجی‌ها و **خروجی PDF** با حفظ استایل.")
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
                    date_grouped = group.groupby(['تاریخ', 'روز'], sort=False)
                    
                    # استخراج تمام روزهای خالص استعلاجی
                    for (date, day), d_group in date_grouped:
                        statuses = d_group['وضعیت'].tolist()
                        if is_pure_sick_leave(statuses):
                            valid_sick_leaves.append({'تاریخ': date, 'روز': day})
                            
                    if valid_sick_leaves:
                        # مرتب سازی تاریخی برای اطمینان
                        valid_sick_leaves = sorted(valid_sick_leaves, key=lambda x: x['تاریخ'])
                        
                        blocks = []
                        current_block = []
                        
                        # الگوریتم بخش بندی بر اساس تشخیص گپ در تاریخ ها
                        for item in valid_sick_leaves:
                            if not current_block:
                                current_block.append(item)
                            else:
                                prev_item = current_block[-1]
                                diff = shamsi_to_days(item['تاریخ']) - shamsi_to_days(prev_item['تاریخ'])
                                
                                if diff == 1:
                                    current_block.append(item)
                                else:
                                    # یک گپ پیدا شد (فرد سر کار آمده یا روز دیگری بوده)، پس بلاک بسته میشود
                                    blocks.append(current_block)
                                    current_block = [item]
                                    
                        if current_block:
                            blocks.append(current_block)
                        
                        try: clean_emp_code = int(float(emp_code))
                        except ValueError: clean_emp_code = emp_code
                            
                        # ثبت بخش های مختلف برای یک شخص
                        for i, block in enumerate(blocks):
                            start_date = block[0]
                            end_date = block[-1]
                            
                            results.append({
                                'کد پرسنلی': clean_emp_code,
                                'نام و نام خانوادگی': emp_name,
                                'بخش': i + 1,
                                'شروع استعلاجی': f"{start_date['روز']} {start_date['تاریخ']}",
                                'پایان استعلاجی': f"{end_date['روز']} {end_date['تاریخ']}",
                                'تعداد روزهای استعلاجی': len(block)
                            })
                
                if results:
                    # ایجاد فایل خروجی آماده برای PDF
                    html_report = generate_pdf_ready_html(results, file.name)
                    b64 = base64.b64encode(html_report.encode('utf-8')).decode()
                    href = f'<a href="data:text/html;base64,{b64}" download="Report_{file.name}.html" style="text-decoration: none; padding: 10px 20px; background-color: #ff4b4b; color: white; border-radius: 5px; font-weight: bold; display: inline-block; margin-bottom: 20px;">📥 دانلود گزارش (برای ذخیره به صورت PDF)</a>'
                    st.markdown(href, unsafe_allow_html=True)
                    
                    # نمایش در Streamlit
                    for res in results:
                        with st.container():
                            st.markdown(f"#### 👤 {res['نام و نام خانوادگی']} <span style='color:gray; font-size:16px'>(کد: {res['کد پرسنلی']})</span> | <span style='color:#00c2a8'>بخش {res['بخش']}</span>", unsafe_allow_html=True)
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown(f"**تعداد روز محاسبه شده:** {res['تعداد روزهای استعلاجی']} روز")
                            with col2:
                                st.info(f"**شروع:** {res['شروع استعلاجی']}")
                                st.error(f"**پایان:** {res['پایان استعلاجی']}")
                            st.divider()
                else:
                    st.warning("هیچ مرخصی استعلاجی معتبری در این فایل یافت نشد.")
