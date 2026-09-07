import streamlit as st
import pandas as pd
import io

# --- Config & Setup ---
st.set_page_config(page_title="Attendance Data Processor", layout="wide")

# --- Helper Functions ---
def parse_time_to_minutes(time_val):
    """Safely parses HH:MM or HH:MM:SS strings into total minutes."""
    if pd.isna(time_val) or str(time_val).strip() in ['', '0']:
        return 0
    
    time_str = str(time_val).strip()
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
    """Converts total minutes back to HH:MM format (supports >24 hours)."""
    if total_minutes == 0:
        return "00:00"
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{hours:02d}:{minutes:02d}"

def apply_custom_styles(df):
    """Applies row-level and cell-level styling based on the blueprint."""
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    
    for idx, row in df.iterrows():
        # Handle the Total Row
        if idx == df.index[-1] and row.get('روز') == 'مجموع':
            # Highlight specific sum cells in the total row
            if 'غیبت روزانه' in df.columns:
                styles.loc[idx, 'غیبت روزانه'] = 'background-color: #FF9999; font-weight: bold; font-size: 14px;' # Red
            if 'کسرکار روزانه' in df.columns:
                styles.loc[idx, 'کسرکار روزانه'] = 'background-color: #FFCC99; font-weight: bold; font-size: 14px;' # Orange
            if 'اضافه کار تعطیلی' in df.columns:
                styles.loc[idx, 'اضافه کار تعطیلی'] = 'background-color: #99CCFF; font-weight: bold; font-size: 14px;' # Blue
            if 'اضافه کار عادی نهایی' in df.columns:
                styles.loc[idx, 'اضافه کار عادی نهایی'] = 'background-color: #99FF99; font-weight: bold; font-size: 14px;' # Green
            if 'جمعه کاری' in df.columns:
                styles.loc[idx, 'جمعه کاری'] = 'background-color: #FFFF99; font-weight: bold; font-size: 14px;' # Yellow
            if 'اضافه کار نهایی روز تعطیل' in df.columns:
                styles.loc[idx, 'اضافه کار نهایی روز تعطیل'] = 'background-color: #CC99FF; font-weight: bold; font-size: 14px;' # Purple
            
            # Make the whole total row bold
            for col in df.columns:
                if not styles.loc[idx, col]:
                    styles.loc[idx, col] = 'font-weight: bold; background-color: #F0F0F0;'
        
        # Handle Data Rows
        else:
            status = str(row.get('وضعیت', ''))
            color = ''
            if 'مرخصی استحقاقی' in status:
                color = 'background-color: #D8BFD8;' # Plum / Purple
            elif 'مرخصی استعلاجی' in status:
                color = 'background-color: #90EE90;' # Light Green
            elif 'عدم حضور' in status:
                color = 'background-color: #FFB6C1;' # Light Red / Pink
            
            if color:
                styles.loc[idx, :] = color
                
    return styles

def process_dataframe(df):
    """Executes the core business logic on the dataframe."""
    
    # Task 2: Remove unnecessary columns
    cols_to_drop = [
        'اطلاعات تردد', 'عنوان گروه کاری', 'توضیحات', 'پیغام', 
        'حضور در روز', 'تاخیر نهایی', 'تعجیل نهایی', 'مرخصی بدون حقوق', 
        'اضافه کار قبل از شیفت روزانه', 'شروع شیفت', 'پایان شیفت', 
        'عنوان شیفت', 'مجوز اضافه کار', 'شبکاری'
    ]
    df = df.drop(columns=cols_to_drop, errors='ignore')
    
    # Task 3: Sum time columns
    time_columns = [
        'اضافه کار تعطیلی', 'کسرکار روزانه', 'اضافه کار عادی نهایی', 
        'جمعه کاری', 'اضافه کار نهایی روز تعطیل', 'غیبت روزانه'
    ]
    
    # Calculate sums
    sums = {}
    for col in time_columns:
        if col in df.columns:
            total_mins = df[col].apply(parse_time_to_minutes).sum()
            sums[col] = minutes_to_time_str(total_mins)
    
    # Create the total row
    total_row = {col: '' for col in df.columns}
    total_row['روز'] = 'مجموع'
    total_row.update(sums)
    
    # Append total row safely
    df_total = pd.DataFrame([total_row])
    df = pd.concat([df, df_total], ignore_index=True)
    
    # Apply Styling
    styled_df = df.style.apply(apply_custom_styles, axis=None)
    
    return styled_df

# --- Main App UI ---
st.title("📊 سیستم پردازش گزارش تردد")
st.markdown("فایل **گزارش-تردد.xlsx** خود را آپلود کنید تا پردازش، حذف ستون‌های زائد، محاسبه مجموع زمان‌ها و هایلایت وضعیت‌ها به صورت خودکار انجام شود.")

uploaded_file = st.file_uploader("آپلود فایل اکسل", type=['xlsx'])

if uploaded_file is not None:
    try:
        with st.spinner('در حال پردازش داده‌ها...'):
            # Load Data
            df = pd.read_excel(uploaded_file)
            
            # Process Data
            styled_df = process_dataframe(df)
            
            st.success("✅ فایل با موفقیت پردازش شد!")
            
            # Display Data in Streamlit
            st.dataframe(styled_df, use_container_width=True, height=600)
            
            # Export to Excel Buffer
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                styled_df.to_excel(writer, index=False, sheet_name='گزارش نهایی')
                # Auto-adjust column widths for better UX in downloaded file
                worksheet = writer.sheets['گزارش نهایی']
                for i, col in enumerate(df.columns):
                    column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, column_len)
                    
            processed_data = output.getvalue()
            
            # Download Button
            st.download_button(
                label="📥 دانلود فایل اکسل پردازش شده",
                data=processed_data,
                file_name="گزارش-تردد-پردازش-شده.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"⚠️ خطایی در پردازش فایل رخ داد: {str(e)}")
        st.info("لطفاً مطمئن شوید ساختار فایل با قالب استاندارد گزارش-تردد همخوانی دارد.")
