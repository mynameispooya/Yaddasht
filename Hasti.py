import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(page_title="سیستم هوشمند عارضه‌یابی تردد", layout="wide")

# ==========================================
# 1. پارسر هوشمند و پاکسازی داده‌ها
# ==========================================
def parse_shift_pattern(shift_df):
    pattern = []
    shift_df.columns = shift_df.columns.astype(str).str.strip()
    for idx, row in shift_df.iterrows():
        start_str = str(row.get('زمان شروع شیفت', '')).strip()
        end_str = str(row.get('زمان پایان شیفت', '')).strip()
        
        start_match = re.search(r'(\d{2}:\d{2})', start_str)
        end_match = re.search(r'(\d{2}:\d{2})', end_str)
        
        start_time = start_match.group(1) if start_match else "00:00"
        end_time = end_match.group(1) if end_match else "00:00"
        
        pattern.append({
            'shift_index': idx,
            'start': start_time,
            'end': end_time,
            'is_off': (start_time == end_time) and start_time != "00:00"
        })
    return pattern

def parse_attendance_file(file):
    """
    تابع هوشمند برای مسطح‌سازی فایل‌های اکسل دارای Multi-Index
    و پاکسازی کاراکترهای مخفی و نامرئی (مثل ZWNJ).
    """
    df_raw = pd.read_excel(file, header=None)
    
    # 1. پیدا کردن سطر هدر اصلی
    header_idx = -1
    for i, row in df_raw.iterrows():
        row_str = ' '.join(row.fillna('').astype(str).tolist())
        if 'تاریخ' in row_str and 'روز' in row_str:
            header_idx = i
            break
            
    if header_idx == -1:
        return pd.read_excel(file)
        
    main_headers = df_raw.iloc[header_idx].fillna('').astype(str).tolist()
    sub_headers = df_raw.iloc[header_idx + 1].fillna('').astype(str).tolist() if header_idx + 1 < len(df_raw) else []
    
    # 2. ترکیب هدرها
    raw_columns = []
    for i in range(len(main_headers)):
        mh = main_headers[i].strip()
        sh = sub_headers[i].strip() if i < len(sub_headers) else ''
        
        if mh in ['اطلاعات تردد', 'nan', '']:
            raw_columns.append(sh if sh and sh != 'nan' else (mh if mh and mh != 'nan' else f'Col_{i}'))
        else:
            raw_columns.append(mh)
            
    # -- رفع خطای ستون‌های تکراری (Deduplication) --
    seen = {}
    final_columns = []
    for col in raw_columns:
        if col in seen:
            seen[col] += 1
            final_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            final_columns.append(col)
            
    start_data_idx = header_idx + 1
    if any(x in sub_headers for x in ['ورود', 'خروج', 'وضعیت', 'گیت']):
        start_data_idx = header_idx + 2
        
    df_clean = df_raw.iloc[start_data_idx:].copy()
    df_clean.columns = final_columns
    
    # 3. پاکسازی کاراکترهای پنهان فارسی و مقادیر تهی
    for col in df_clean.columns:
        if df_clean[col].dtype == object:
            df_clean[col] = df_clean[col].astype(str).str.replace('\u200f', '').str.strip()
            df_clean[col] = df_clean[col].replace(['nan', 'None', '', '<NA>', 'NaN'], pd.NA)
            
    return df_clean

# ==========================================
# 2. توابع محاسباتی زمان
# ==========================================
def clean_time(val):
    if pd.isna(val): return None
    val_str = str(val).strip()
    if val_str in ['', 'nan', 'None', 'NaT']: return None
    return val_str

def time_to_minutes(t_str):
    if pd.isna(t_str) or not t_str or t_str == '-': return None
    try:
        t_str = str(t_str).replace('روز بعد', '').strip()
        t_str = t_str.translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789'))
        
        parts = t_str.split(':')
        if len(parts) >= 2:
            h = int(parts[0].strip())
            m = int(parts[1].strip()[:2])
            return h * 60 + m
    except:
        return None
    return None

def calc_time_diff(actual, expected, is_checkout=False):
    act_m = time_to_minutes(actual)
    exp_m = time_to_minutes(expected)
    if act_m is None or exp_m is None: return 0
    if is_checkout and exp_m < 720 and act_m > 720: act_m -= 1440 
    return act_m - exp_m

# ==========================================
# 3. تجمیع و عارضه‌یابی پیشرفته
# ==========================================
def aggregate_daily_attendance(df):
    daily_records = []
    date_col = 'تاریخ' if 'تاریخ' in df.columns else 'روز'
    
    df['روز_اصلی'] = df[date_col].astype(str).str.replace('↳', '').str.strip()
    df['روز_اصلی'] = df['روز_اصلی'].replace(['nan', 'None', '', 'NaT'], pd.NA).ffill()
    
    col_in = 'ورود' if 'ورود' in df.columns else 'شروع' if 'شروع' in df.columns else None
    col_out = 'خروج' if 'خروج' in df.columns else 'پایان' if 'پایان' in df.columns else None
    col_s_start = 'شروع شیفت' if 'شروع شیفت' in df.columns else None
    col_s_end = 'پایان شیفت' if 'پایان شیفت' in df.columns else None
    
    grouped = df.groupby('روز_اصلی', sort=False)
    for day, group in grouped:
        if pd.isna(day) or day in ['مجموع', 'مجموع نهایی', 'مجمــــــــوع', 'nan']: continue
            
        ins = group[col_in].apply(clean_time).dropna().tolist() if col_in else []
        outs = group[col_out].apply(clean_time).dropna().tolist() if col_out else []
        s_starts = group[col_s_start].apply(clean_time).dropna().tolist() if col_s_start else []
        s_ends = group[col_s_end].apply(clean_time).dropna().tolist() if col_s_end else []
        
        daily_records.append({
            'تاریخ': day,
            'ورود_ثبت_شده': ins[0] if ins else None,
            'خروج_ثبت_شده': outs[-1] if outs else None,
            'شروع_شیفت_سیستم': s_starts[0] if s_starts else None,
            'پایان_شیفت_سیستم': s_ends[0] if s_ends else None,
        })
    return pd.DataFrame(daily_records)

def sync_and_diagnose(daily_df, shift_pattern):
    results = []
    
    for i, row in daily_df.iterrows():
        actual_in = row['ورود_ثبت_شده'] or '-'
        actual_out = row['خروج_ثبت_شده'] or '-'
        sys_shift_in = row['شروع_شیفت_سیستم'] or '-'
        sys_shift_out = row['پایان_شیفت_سیستم'] or '-'
        
        diagnosis = []
        action = []
        is_error = False
        
        if sys_shift_in != '-' and 'تعطیل' not in sys_shift_in:
            if actual_in == '-':
                diagnosis.append("خطا: عدم ثبت ورود")
                action.append(f"ورود باید بر اساس شیفت {sys_shift_in[:5]} ثبت شود")
                is_error = True
            if actual_out == '-':
                diagnosis.append("خطا: عدم ثبت خروج")
                if "عدم ثبت ورود" not in diagnosis: 
                    action.append(f"خروج باید بر اساس شیفت {sys_shift_out[:5]} ثبت شود")
                is_error = True
                
        if actual_in != '-':
            if sys_shift_in == '-' or 'تعطیل' in sys_shift_in:
                diagnosis.append("آلارم: خروج از الگو (تردد خارج از شیفت تعریف شده)")
                action.append("بررسی تردد مازاد یا لزوم اصلاح شیفت")
                is_error = True
                
        if actual_in != '-' and sys_shift_in != '-' and 'تعطیل' not in sys_shift_in:
            in_diff = calc_time_diff(actual_in, sys_shift_in, is_checkout=False)
            if in_diff > 15:
                diagnosis.append(f"تاخیر در ورود ({in_diff} دقیقه)")
                is_error = True
                
        if actual_out != '-' and sys_shift_out != '-' and 'تعطیل' not in sys_shift_out:
            out_diff = calc_time_diff(actual_out, sys_shift_out, is_checkout=True)
            if out_diff < -15:
                diagnosis.append(f"تعجیل در خروج ({abs(out_diff)} دقیقه)")
                is_error = True
                
        if not diagnosis:
            diagnosis.append("عادی (بدون مغایرت)")
            action.append("-")
            
        results.append({
            'تاریخ': row['تاریخ'],
            'شیفت سیستم (ورود)': sys_shift_in,
            'شیفت سیستم (خروج)': sys_shift_out,
            'تردد (ورود)': actual_in,
            'تردد (خروج)': actual_out,
            'عارضه / آلارم سیستم': ' | '.join(diagnosis),
            'اقدام پیشنهادی': ' | '.join(action),
            '_is_error': is_error
        })
        
    return pd.DataFrame(results)

def style_diagnosis_df(df):
    def row_style(row):
        if row.get('_is_error', False):
            return ['background-color: #FFF0F0; color: #CC0000; font-weight: bold;'] * len(row)
        elif 'خروج از الگو' in row.get('عارضه / آلارم سیستم', ''):
            return ['background-color: #FFFACD; color: #B8860B; font-weight: bold;'] * len(row)
        return ['background-color: #F5FFFA; color: #006400;'] * len(row)
    
    df_clean = df.drop(columns=['_is_error'])
    return df_clean.style.apply(row_style, axis=1)

# ==========================================
# رابط کاربری Streamlit
# ==========================================
st.title("🕵️‍♂️ سیستم هوشمند عارضه‌یابی شیفت و تردد")
st.markdown("ابتدا فایل **شیفت** و سپس فایل **تردد** (با فرمت‌های چندسطحی جدید) را آپلود کنید.")

col1, col2 = st.columns(2)
with col1:
    shift_file = st.file_uploader("1️⃣ آپلود فایل شیفت", type=['xlsx', 'xls'])
with col2:
    attendance_file = st.file_uploader("2️⃣ آپلود فایل تردد", type=['xlsx', 'xls'])

if shift_file and attendance_file:
    with st.spinner("در حال تحلیل الگوهای تو‌در‌تو و تطبیق با سیستم..."):
        try:
            df_shift = pd.read_excel(shift_file)
            df_attendance_clean = parse_attendance_file(attendance_file)
            
            shift_pattern = parse_shift_pattern(df_shift)
            daily_attendance = aggregate_daily_attendance(df_attendance_clean)
            
            diagnosis_df = sync_and_diagnose(daily_attendance, shift_pattern)
            
            st.success("✅ ساختار پیچیده فایل اکسل مسطح شد و عارضه‌یابی با موفقیت پایان یافت.")
            
            st.subheader("🔍 جدول هوشمند مقایسه و عارضه‌یابی")
            styled_diagnosis = style_diagnosis_df(diagnosis_df)
            st.dataframe(styled_diagnosis, width="stretch", hide_index=True, height=600)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                styled_diagnosis.to_excel(writer, index=False, sheet_name='گزارش هوشمند')
                worksheet = writer.sheets['گزارش هوشمند']
                
                df_out = diagnosis_df.drop(columns=['_is_error'])
                for i, col in enumerate(df_out.columns):
                    col_str = str(col)
                    max_data_len = df_out[col].astype(str).str.len().fillna(0).max()
                    final_width = max(max_data_len, len(col_str)) + 4
                    worksheet.set_column(i, i, final_width)
                    
            st.download_button(
                label="📥 دانلود گزارش نهایی عارضه‌یابی",
                data=output.getvalue(),
                file_name="گزارش-هوشمند-تردد.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        except Exception as e:
            st.error(f"⚠️ خطایی در پردازش رخ داد: {str(e)}")
            st.exception(e)
