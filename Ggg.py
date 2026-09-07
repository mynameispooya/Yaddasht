import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="سیستم هوشمند عارضه‌یابی تردد", layout="wide")

# ==========================================
# 1. منطق پارس کردن شیفت‌ها
# ==========================================
def parse_shift_pattern(shift_df):
    """الگوی شیفت را از فایل اکسل استخراج می‌کند"""
    pattern = []
    # استانداردسازی نام ستون‌ها
    shift_df.columns = shift_df.columns.astype(str).str.strip()
    
    for idx, row in shift_df.iterrows():
        start_str = str(row.get('زمان شروع شیفت', '')).strip()
        end_str = str(row.get('زمان پایان شیفت', '')).strip()
        
        # استخراج ساعت (HH:MM) با Regex
        start_time_match = re.search(r'(\d{2}:\d{2})', start_str)
        end_time_match = re.search(r'(\d{2}:\d{2})', end_str)
        
        start_time = start_time_match.group(1) if start_time_match else "00:00"
        end_time = end_time_match.group(1) if end_time_match else "00:00"
        
        # تشخیص روز تعطیل (شروع و پایان یکسان)
        is_off = (start_time == end_time) and start_time != "00:00"
        
        # تشخیص شیفت شب (پایان در روز بعد)
        is_night = 'روز بعد' in end_str
        
        pattern.append({
            'shift_index': idx,
            'start': start_time,
            'end': end_time,
            'is_off': is_off,
            'is_night': is_night,
            'raw_start': start_str,
            'raw_end': end_str
        })
    return pattern

# ==========================================
# 2. منطق تجمیع داده‌های تردد یک روز
# ==========================================
def clean_time(val):
    if pd.isna(val) or str(val).strip() in ['', 'nan', 'None', 'NaT']:
        return None
    return str(val).strip()

def aggregate_daily_attendance(df):
    """ردیف‌های چندگانه یک روز را تجمیع می‌کند تا اولین ورود و آخرین خروج کشف شود"""
    daily_records = []
    
    # استخراج نام روز اصلی (حذف کاراکترهای ↳ و فاصله‌ها)
    df['روز_اصلی'] = df['روز'].astype(str).str.replace('↳', '').str.strip()
    df['روز_اصلی'].replace(['nan', 'None', ''], pd.NA, inplace=True)
    df['روز_اصلی'] = df['روز_اصلی'].ffill()
    
    grouped = df.groupby('روز_اصلی', sort=False)
    
    for day, group in grouped:
        if day == 'مجموع' or day == 'مجموع نهایی' or pd.isna(day):
            continue
            
        # یافتن همه مقادیر شروع و پایان
        starts = group['شروع'].apply(clean_time).dropna().tolist()
        ends = group['پایان'].apply(clean_time).dropna().tolist()
        
        actual_in = starts[0] if starts else None
        actual_out = ends[-1] if ends else None
        
        # بررسی وضعیت غیبت سیستم
        absent_flags = group.get('غیبت روزانه', pd.Series())
        is_absent = any(absent_flags.astype(str).str.contains('1'))
        
        daily_records.append({
            'تاریخ': day,
            'ورود_ثبت_شده': actual_in,
            'خروج_ثبت_شده': actual_out,
            'غیبت_سیستمی': is_absent,
            'ردیف_های_گروه': len(group)
        })
        
    return pd.DataFrame(daily_records)

# ==========================================
# 3. همگام‌سازی و مقایسه
# ==========================================
def sync_and_diagnose(daily_df, shift_pattern):
    """تطبیق شیفت با روزها و کشف مغایرت"""
    
    # تعیین فاز الگو (به صورت پیش‌فرض سعی می‌کنیم از منطق روز اول استفاده کنیم)
    # اگر روز اول ورود ندارد اما خروج 07:15 دارد -> پس امروز تعطیل است (فاز آف)
    pattern_len = len(shift_pattern)
    start_offset = 0
    
    if not daily_df.empty:
        first_in = daily_df.iloc[0]['ورود_ثبت_شده']
        first_out = daily_df.iloc[0]['خروج_ثبت_شده']
        
        # لاجیک مشخص شده توسط شما: خروج صبحگاهی بدون ورود یعنی شیفت استراحت
        if not first_in and first_out and '07' in first_out:
            # جستجو برای پیدا کردن ایندکس روز استراحت در الگو
            for i, p in enumerate(shift_pattern):
                if p['is_off']:
                    start_offset = i
                    break
        elif first_in:
            # پیدا کردن شیفتی که ساعتش نزدیک به ورود ثبت شده است
            hour_in = first_in[:2]
            for i, p in enumerate(shift_pattern):
                if not p['is_off'] and p['start'].startswith(hour_in):
                    start_offset = i
                    break

    results = []
    for i, row in daily_df.iterrows():
        # تخصیص شیفت به این روز بر اساس ایندکس
        current_shift = shift_pattern[(i + start_offset) % pattern_len]
        
        diagnosis = []
        is_error = False
        
        expected_in = current_shift['start'] if not current_shift['is_off'] else '-'
        expected_out = current_shift['end'] if not current_shift['is_off'] else '-'
        
        actual_in = row['ورود_ثبت_شده'] or '-'
        actual_out = row['خروج_ثبت_شده'] or '-'
        
        # لاجیک عارضه‌یابی
        if current_shift['is_off']:
            if actual_in != '-' or actual_out != '-':
                diagnosis.append("تردد در روز تعطیل (احتمال اضافه‌کار یا اشتباه)")
        else:
            if actual_in == '-' and actual_out == '-':
                if not row['غیبت_سیستمی']:
                    diagnosis.append("بدون تردد اما غیبت ثبت نشده!")
                    is_error = True
            else:
                if row['غیبت_سیستمی']:
                    diagnosis.append("ثبت غیبت با وجود داشتن تردد!")
                    is_error = True
                    
                if actual_in == '-':
                    diagnosis.append("فراموشی کارت ورود")
                    is_error = True
                if actual_out == '-':
                    diagnosis.append("فراموشی کارت خروج")
                    is_error = True

        if not diagnosis:
            diagnosis.append("عادی")
            
        results.append({
            'تاریخ': row['تاریخ'],
            'ورود (سیستم)': actual_in,
            'خروج (سیستم)': actual_out,
            'ورود (مورد انتظار)': expected_in,
            'خروج (مورد انتظار)': expected_out,
            'وضعیت غیبت': 'بله' if row['غیبت_سیستمی'] else 'خیر',
            'دلیل مغایرت / عارضه': ' | '.join(diagnosis),
            '_is_error': is_error
        })
        
    return pd.DataFrame(results)

def style_diagnosis_df(df):
    """هایلایت کردن خطاهای کشف شده"""
    def row_style(row):
        if row.get('_is_error', False):
            return ['background-color: #ffe6e6; color: #cc0000; font-weight: bold;'] * len(row)
        elif 'تعطیل' in row.get('دلیل مغایرت / عارضه', ''):
            return ['background-color: #e6f7ff; color: #0066cc;'] * len(row)
        return [''] * len(row)
    
    df_clean = df.drop(columns=['_is_error'])
    return df_clean.style.apply(row_style, axis=1)

# ==========================================
# رابط کاربری Streamlit
# ==========================================
st.title("🕵️‍♂️ سیستم هوشمند عارضه‌یابی شیفت و تردد")
st.markdown("ابتدا فایل **شیفت** را آپلود کنید تا سیستم الگوی کاری را یاد بگیرد، سپس فایل **گزارش تردد** را آپلود کنید تا مغایرت‌ها به صورت خودکار شناسایی شوند.")

col1, col2 = st.columns(2)
with col1:
    shift_file = st.file_uploader("1️⃣ آپلود فایل شیفت (مانند: شیفت (2).xlsx)", type=['xlsx', 'xls'])
with col2:
    attendance_file = st.file_uploader("2️⃣ آپلود فایل تردد (مانند: نمونه.xlsx)", type=['xlsx', 'xls'])

if shift_file and attendance_file:
    with st.spinner("در حال تحلیل الگوهای شیفت و تطبیق با ترددها..."):
        try:
            # 1. بارگذاری داده‌ها
            df_shift_raw = pd.read_excel(shift_file)
            df_attendance_raw = pd.read_excel(attendance_file)
            
            # 2. یادگیری الگو
            shift_pattern = parse_shift_pattern(df_shift_raw)
            
            st.success("✅ الگوی شیفت با موفقیت استخراج و یادگیری شد.")
            with st.expander("👀 مشاهده الگوی یادگیری شده"):
                st.json([{k: v for k, v in p.items() if not k.startswith('raw_')} for p in shift_pattern])
                
            # 3. تجمیع و عارضه‌یابی
            daily_attendance = aggregate_daily_attendance(df_attendance_raw)
            diagnosis_df = sync_and_diagnose(daily_attendance, shift_pattern)
            
            # 4. نمایش مقایسه‌ای
            st.subheader("🔍 جدول مقایسه‌ای و عارضه‌یابی")
            styled_diagnosis = style_diagnosis_df(diagnosis_df)
            
            st.dataframe(styled_diagnosis, width="stretch", hide_index=True, height=600)
            
            # 5. خروجی اکسل
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                styled_diagnosis.to_excel(writer, index=False, sheet_name='عارضه یابی')
                worksheet = writer.sheets['عارضه یابی']
                
                # تنظیم عرض ستون‌ها
                df_out = diagnosis_df.drop(columns=['_is_error'])
                for i, col in enumerate(df_out.columns):
                    col_str = str(col)
                    max_data_len = df_out[col].astype(str).str.len().fillna(0).max()
                    final_width = max(max_data_len, len(col_str)) + 4
                    worksheet.set_column(i, i, final_width)
                    
            st.download_button(
                label="📥 دانلود گزارش عارضه‌یابی",
                data=output.getvalue(),
                file_name="گزارش-عارضه-یابی.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        except Exception as e:
            st.error(f"⚠️ خطایی رخ داد: {str(e)}")
            st.exception(e)
