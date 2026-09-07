import streamlit as st
import pandas as pd
import io

# --- Config & Setup ---
st.set_page_config(page_title="سیستم پردازش هوشمند گزارش تردد", layout="wide")

# --- Helper Functions ---
def parse_time_to_minutes(time_val):
    """زمان‌ها را به صورت کاملاً امن به دقیقه تبدیل می‌کند."""
    if pd.isna(time_val):
        return 0
    time_str = str(time_val).strip()
    if time_str in ['', '0', 'nan', 'None', 'NaT', '<NA>']:
        return 0
    
    try:
        parts = time_str.split(':')
        if len(parts) >= 2:
            hours = int(parts[0])
            minutes = int(parts[1])
            return (hours * 60) + minutes
        return 0
    except Exception:
        return 0

def minutes_to_time_str(total_minutes):
    """مجموع دقایق را مجددا به فرمت HH:MM تبدیل می‌کند."""
    if pd.isna(total_minutes) or total_minutes == 0:
        return "00:00"
    hours = int(total_minutes) // 60
    minutes = int(total_minutes) % 60
    return f"{hours:02d}:{minutes:02d}"

def apply_custom_styles(df):
    """استایل‌های رنگی و ظاهری را برای رابط کاربری و فایل اکسل اعمال می‌کند."""
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    
    for idx, row in df.iterrows():
        # استایل ردیف مجموع
        if idx == df.index[-1] and 'مجموع' in str(row.get('روز', '')):
            if 'غیبت روزانه' in df.columns:
                styles.loc[idx, 'غیبت روزانه'] = 'background-color: #FF9999; font-weight: bold;'
            if 'کسرکار روزانه' in df.columns:
                styles.loc[idx, 'کسرکار روزانه'] = 'background-color: #FFCC99; font-weight: bold;'
            if 'اضافه کار تعطیلی' in df.columns:
                styles.loc[idx, 'اضافه کار تعطیلی'] = 'background-color: #99CCFF; font-weight: bold;'
            if 'اضافه کار عادی نهایی' in df.columns:
                styles.loc[idx, 'اضافه کار عادی نهایی'] = 'background-color: #99FF99; font-weight: bold;'
            if 'جمعه کاری' in df.columns:
                styles.loc[idx, 'جمعه کاری'] = 'background-color: #FFFF99; font-weight: bold;'
            if 'اضافه کار نهایی روز تعطیل' in df.columns:
                styles.loc[idx, 'اضافه کار نهایی روز تعطیل'] = 'background-color: #CC99FF; font-weight: bold;'
            
            for col in df.columns:
                if not styles.loc[idx, col]:
                    styles.loc[idx, col] = 'font-weight: bold; background-color: #E8E8E8;'
        
        # استایل ردیف‌های داده
        else:
            status = str(row.get('وضعیت', ''))
            is_sub_row = '↳' in str(row.get('روز', ''))
            
            color = ''
            if 'مرخصی استحقاقی' in status:
                color = 'background-color: #D8BFD8;'
            elif 'مرخصی استعلاجی' in status:
                color = 'background-color: #90EE90;'
            elif 'عدم حضور' in status:
                color = 'background-color: #FFB6C1;'
            
            if color:
                if is_sub_row:
                    color += ' opacity: 0.85;'
                styles.loc[idx, :] = color
            
            if is_sub_row:
                styles.loc[idx, 'روز'] = f"{color} color: #666666; font-weight: bold;"
                
    return styles

def process_dataframe(df):
    """منطق اصلی پردازش داده‌ها، پاکسازی، محاسبات و قالب‌بندی UX"""
    
    # اطمینان از اینکه نام تمامی ستون‌ها استرینگ هستند (جلوگیری از خطای float has no len در Arrow)
    df.columns = df.columns.astype(str)
    
    # 1. حذف ستون‌های زائد
    cols_to_drop = [
        'اطلاعات تردد', 'عنوان گروه کاری', 'توضیحات', 'پیغام', 
        'حضور در روز', 'تاخیر نهایی', 'تعجیل نهایی', 'مرخصی بدون حقوق', 
        'اضافه کار قبل از شیفت روزانه', 'شروع شیفت', 'پایان شیفت', 
        'عنوان شیفت', 'مجوز اضافه کار', 'شبکاری'
    ]
    df = df.drop(columns=[c for c in cols_to_drop if c in df.columns], errors='ignore')
    
    # 2. محاسبه مجموع زمان‌ها
    time_columns = [
        'اضافه کار تعطیلی', 'کسرکار روزانه', 'اضافه کار عادی نهایی', 
        'جمعه کاری', 'اضافه کار نهایی روز تعطیل', 'غیبت روزانه'
    ]
    sums = {}
    for col in time_columns:
        if col in df.columns:
            total_mins = df[col].apply(parse_time_to_minutes).sum()
            sums[col] = minutes_to_time_str(total_mins)
            
    # 3. تبدیل ایمن داده‌ها به استرینگ
    df = df.astype(str)
    df = df.replace(['nan', 'None', '<NA>', 'NaT', 'nan.0'], '')
    
    # 4. بهینه‌سازی UX (معماری درختی روزها)
    if 'روز' in df.columns:
        previous_day = None
        for idx in df.index:
            current_day = str(df.at[idx, 'روز']).strip()
            
            if current_day != '' and current_day == previous_day:
                df.at[idx, 'روز'] = ' ↳ '
            else:
                if current_day != '':
                    previous_day = current_day

    # 5. اضافه کردن ردیف مجموع
    total_row = {col: '' for col in df.columns}
    total_row['روز'] = 'مجموع نهایی'
    total_row.update(sums)
    
    df_total = pd.DataFrame([total_row])
    df = pd.concat([df, df_total], ignore_index=True)
    
    # 6. اعمال استایل
    styled_df = df.style.apply(apply_custom_styles, axis=None)
    
    return styled_df, df

# --- Main App UI ---
st.title("📊 سیستم پردازش هوشمند گزارش تردد")
st.markdown("فایل **گزارش-تردد.xlsx** را آپلود کنید. ستون‌های زائد حذف شده، مجموع زمان‌ها محاسبه می‌گردد و ردیف‌های تکراری ساده‌سازی می‌شوند.")

uploaded_file = st.file_uploader("آپلود فایل اکسل", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        with st.spinner('در حال پردازش داده‌ها...'):
            df_raw = pd.read_excel(uploaded_file)
            styled_df, clean_df = process_dataframe(df_raw)
            
            st.success("✅ فایل با موفقیت پردازش شد!")
            
            # رفع هشدار مربوط به use_container_width با جایگزینی عرض کشسان
            st.dataframe(styled_df, width="stretch", hide_index=True, height=600)
            
            # آماده‌سازی خروجی اکسل
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                styled_df.to_excel(writer, index=False, sheet_name='گزارش نهایی')
                
                # تنظیم خودکار و امن عرض ستون‌ها (رفع خطای float has no len)
                worksheet = writer.sheets['گزارش نهایی']
                for i, col in enumerate(clean_df.columns):
                    col_str = str(col)
                    # محاسبه طولانی‌ترین رشته در ستون با استفاده از str.len() ایمن در پانداز
                    max_data_len = clean_df[col].astype(str).str.len().fillna(0).max()
                    max_data_len = int(max_data_len) if pd.notna(max_data_len) else 0
                    
                    final_width = max(max_data_len, len(col_str)) + 4
                    worksheet.set_column(i, i, final_width)
                    
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 دانلود فایل اکسل نهایی",
                data=processed_data,
                file_name="گزارش-تردد-هوشمند.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
    except Exception as e:
        st.error(f"⚠️ خطایی در پردازش فایل رخ داد: {str(e)}")
        st.info("لطفاً مطمئن شوید ساختار فایل با قالب استاندارد گزارش-تردد همخوانی دارد.")
