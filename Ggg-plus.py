import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(page_title="سیستم هوشمند عارضه‌یابی تردد", layout="wide")

# ==========================================
# 1. استخراج الگوی شیفت
# ==========================================
def parse_shift_pattern(shift_df):
    pattern = []
    shift_df.columns = shift_df.columns.astype(str).str.strip()
    
    for idx, row in shift_df.iterrows():
        start_str = str(row.get('زمان شروع شیفت', '')).strip()
        end_str = str(row.get('زمان پایان شیفت', '')).strip()
        
        start_time_match = re.search(r'(\d{2}:\d{2})', start_str)
        end_time_match = re.search(r'(\d{2}:\d{2})', end_str)
        
        start_time = start_time_match.group(1) if start_time_match else "00:00"
        end_time = end_time_match.group(1) if end_time_match else "00:00"
        
        is_off = (start_time == end_time) and start_time != "00:00"
        
        pattern.append({
            'shift_index': idx,
            'start': start_time,
            'end': end_time,
            'is_off': is_off
        })
    return pattern

# ==========================================
# 2. تجمیع و پاکسازی داده‌های تردد
# ==========================================
def clean_time(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    if val_str in ['', 'nan', 'None', 'NaT', 'NaN', '<NA>']:
        return None
    return val_str

def aggregate_daily_attendance(df):
    daily_records = []
    
    # استانداردسازی نام روزها
    df['روز_اصلی'] = df['روز'].astype(str).str.replace('↳', '').str.strip()
    df['روز_اصلی'] = df['روز_اصلی'].replace(['nan', 'None', '', 'NaT', '<NA>'], pd.NA)
    df['روز_اصلی'] = df['روز_اصلی'].ffill()
    
    grouped = df.groupby('روز_اصلی', sort=False)
    
    for day, group in grouped:
        if day in ['مجموع', 'مجموع نهایی'] or pd.isna(day):
            continue
            
        starts = group['شروع'].apply(clean_time).dropna().tolist()
        ends = group['پایان'].apply(clean_time).dropna().tolist()
        
        actual_in = starts[0] if starts else None
        actual_out = ends[-1] if ends else None
        
        # بررسی وضعیت‌ها برای یافتن غیبت یا عدم حضور ثبت شده توسط دستگاه
        statuses = group.get('وضعیت', pd.Series(dtype=str)).astype(str).tolist()
        is_absent_status = any('عدم حضور' in s or 'غیبت' in s for s in statuses)
        
        daily_records.append({
            'تاریخ': day,
            'ورود_ثبت_شده': actual_in,
            'خروج_ثبت_شده': actual_out,
            'وضعیت_عدم_حضور': is_absent_status
        })
        
    return pd.DataFrame(daily_records)

# ==========================================
# 3. توابع محاسباتی زمان
# ==========================================
def time_to_minutes(t_str):
    if not t_str or t_str == '-': return None
    try:
        h, m = map(int, t_str.split(':'))
        return h * 60 + m
    except:
        return None

def calc_time_diff(actual, expected, is_checkout=False):
    """محاسبه اختلاف زمان به دقیقه. برای خروج‌ها منطق برعکس است."""
    act_m = time_to_minutes(actual)
    exp_m = time_to_minutes(expected)
    if act_m is None or exp_m is None: return 0
    
    # مدیریت شیفت شب (عبور از نیمه‌شب)
    if is_checkout and exp_m < 720 and act_m > 720: 
        act_m -= 1440 # خروج شب قبل ثبت شده
        
    diff = act_m - exp_m
    return diff

# ==========================================
# 4. الگوریتم لنگرگیری و عارضه‌یابی
# ==========================================
def sync_and_diagnose(daily_df, shift_pattern):
    pattern_len = len(shift_pattern)
    
    # -- مرحله 1: یافتن لنگر (Anchor) --
    anchor_day_idx = 0
    anchor_pattern_idx = 0
    found_anchor = False
    
    for i, row in daily_df.iterrows():
        actual_in = row['ورود_ثبت_شده']
        if actual_in:
            in_mins = time_to_minutes(actual_in)
            # مقایسه با الگوهای شیفت برای یافتن نزدیک‌ترین تطابق
            for p_idx, p in enumerate(shift_pattern):
                if not p['is_off']:
                    exp_mins = time_to_minutes(p['start'])
                    # اگر ورود فرد حداکثر 2 ساعت با شروع شیفت اختلاف داشت، این روز لنگر ماست
                    if abs(in_mins - exp_mins) <= 120 or abs(in_mins - exp_mins) >= 1320:
                        anchor_day_idx = i
                        anchor_pattern_idx = p_idx
                        found_anchor = True
                        break
        if found_anchor:
            break

    # -- مرحله 2: اعمال الگو به کل روزها بر اساس لنگر --
    results = []
    for i, row in daily_df.iterrows():
        # محاسبه شیفت این روز نسبت به لنگر پیدا شده
        distance_from_anchor = i - anchor_day_idx
        current_pattern_idx = (anchor_pattern_idx + distance_from_anchor) % pattern_len
        current_shift = shift_pattern[current_pattern_idx]
        
        diagnosis = []
        is_error = False
        
        expected_in = current_shift['start'] if not current_shift['is_off'] else '-'
        expected_out = current_shift['end'] if not current_shift['is_off'] else '-'
        
        actual_in = row['ورود_ثبت_شده'] or '-'
        actual_out = row['خروج_ثبت_شده'] or '-'
        
        # -- مرحله 3: عارضه‌یابی هوشمند --
        if current_shift['is_off']:
            if actual_in != '-' or actual_out != '-':
                diagnosis.append("تردد در روز استراحت (اضافه‌کار خارج از شیفت)")
        else:
            # 1. بررسی عدم حضور / غیبت
            if row['وضعیت_عدم_حضور']:
                if actual_in == '-' and actual_out == '-':
                    diagnosis.append("غیبت کامل (ثبت عدم حضور)")
                    is_error = True
                else:
                    diagnosis.append("فراموشی ثبت چهره (عدم حضور در سیستم ثبت شده با وجود تردد ناقص)")
                    is_error = True
            
            # 2. بررسی نواقص تردد
            if actual_in == '-' and not row['وضعیت_عدم_حضور']:
                diagnosis.append("عدم ثبت چهره در ورود")
                is_error = True
            if actual_out == '-' and not row['وضعیت_عدم_حضور'] and actual_in != '-':
                diagnosis.append("عدم ثبت چهره در خروج")
                is_error = True
                
            # 3. بررسی تاخیر و تعجیل (فقط اگر هر دو زمان موجود باشند)
            if actual_in != '-' and expected_in != '-':
                in_diff = calc_time_diff(actual_in, expected_in, is_checkout=False)
                if in_diff > 15: # بیشتر از 15 دقیقه تاخیر
                    diagnosis.append(f"تاخیر در شروع کار ({in_diff} دقیقه)")
                    is_error = True
                    
            if actual_out != '-' and expected_out != '-':
                out_diff = calc_time_diff(actual_out, expected_out, is_checkout=True)
                if out_diff < -15: # بیشتر از 15 دقیقه تعجیل (زودتر رفتن)
                    diagnosis.append(f"تعجیل در خروج ({abs(out_diff)} دقیقه)")
                    is_error = True

        if not diagnosis:
            diagnosis.append("بدون مغایرت")
            
        results.append({
            'تاریخ': row['تاریخ'],
            'شیفت مورد انتظار': 'استراحت' if current_shift['is_off'] else f"از {expected_in} تا {expected_out}",
            'ورود (سیستم)': actual_in,
            'خروج (سیستم)': actual_out,
            'عارضه یابی سیستم': ' | '.join(diagnosis),
            '_is_error': is_error
        })
        
    return pd.DataFrame(results)

def style_diagnosis_df(df):
    def row_style(row):
        if row.get('_is_error', False):
            return ['background-color: #FFF0F0; color: #CC0000; font-weight: bold;'] * len(row)
        elif 'استراحت' in row.get('شیفت مورد انتظار', ''):
            return ['background-color: #F0F8FF; color: #0055A4;'] * len(row)
        return ['background-color: #F5FFFA; color: #006400;'] * len(row) # بدون مغایرت
    
    df_clean = df.drop(columns=['_is_error'])
    return df_clean.style.apply(row_style, axis=1)

# ==========================================
# رابط کاربری Streamlit
# ==========================================
st.title("🕵️‍♂️ سیستم هوشمند عارضه‌یابی شیفت و تردد")
st.markdown("ابتدا فایل **شیفت** را آپلود کنید تا سیستم الگوی کاری را یاد بگیرد، سپس فایل **گزارش تردد** را آپلود کنید تا مغایرت‌ها به صورت خودکار شناسایی شوند.")

col1, col2 = st.columns(2)
with col1:
    shift_file = st.file_uploader("1️⃣ آپلود فایل شیفت", type=['xlsx', 'xls'])
with col2:
    attendance_file = st.file_uploader("2️⃣ آپلود فایل تردد", type=['xlsx', 'xls'])

if shift_file and attendance_file:
    with st.spinner("در حال تحلیل الگوهای شیفت و تطبیق با ترددها..."):
        try:
            df_shift = pd.read_excel(shift_file)
            df_attendance = pd.read_excel(attendance_file)
            
            shift_pattern = parse_shift_pattern(df_shift)
            daily_attendance = aggregate_daily_attendance(df_attendance)
            
            # اجرای هسته هوشمند
            diagnosis_df = sync_and_diagnose(daily_attendance, shift_pattern)
            
            st.success("✅ الگوی شیفت استخراج و با موفقیت روی ترددها اعمال شد.")
            
            st.subheader("🔍 جدول هوشمند مقایسه و عارضه‌یابی")
            styled_diagnosis = style_diagnosis_df(diagnosis_df)
            
            st.dataframe(styled_diagnosis, width="stretch", hide_index=True, height=600)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                styled_diagnosis.to_excel(writer, index=False, sheet_name='عارضه یابی')
                worksheet = writer.sheets['عارضه یابی']
                
                df_out = diagnosis_df.drop(columns=['_is_error'])
                for i, col in enumerate(df_out.columns):
                    col_str = str(col)
                    max_data_len = df_out[col].astype(str).str.len().fillna(0).max()
                    final_width = max(max_data_len, len(col_str)) + 4
                    worksheet.set_column(i, i, final_width)
                    
            st.download_button(
                label="📥 دانلود گزارش نهایی",
                data=output.getvalue(),
                file_name="گزارش-هوشمند-عارضه-یابی.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        except Exception as e:
            st.error(f"⚠️ خطایی رخ داد: {str(e)}")
            st.exception(e)
