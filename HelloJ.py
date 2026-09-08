import streamlit as st
import pandas as pd
import numpy as np
import io
import json

st.set_page_config(page_title="سیستم هوشمند عارضه‌یابی تردد", layout="wide")

# ==========================================
# 1. داده‌های پیش‌فرض
# ==========================================
DEFAULT_JSON = """[
  {
    "نام": "روزکار",
    "شروع": "07:15",
    "پایان": "19:15",
    "تعطیل": false
  },
  {
    "نام": "شب‌کار",
    "شروع": "19:15",
    "پایان": "07:15",
    "تعطیل": false
  },
  {
    "نام": "استراحت",
    "شروع": "00:00",
    "پایان": "00:00",
    "تعطیل": true
  }
]"""

# ==========================================
# 2. پارسر و پاکسازی داده‌ها
# ==========================================
def parse_attendance_file(file):
    df_raw = pd.read_excel(file, header=None)
    
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
    
    raw_columns = []
    for i in range(len(main_headers)):
        mh = main_headers[i].strip()
        sh = sub_headers[i].strip() if i < len(sub_headers) else ''
        
        if mh in ['اطلاعات تردد', 'nan', '']:
            raw_columns.append(sh if sh and sh != 'nan' else (mh if mh and mh != 'nan' else f'Col_{i}'))
        else:
            raw_columns.append(mh)
            
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
    
    for col in df_clean.columns:
        if df_clean[col].dtype == object:
            df_clean[col] = df_clean[col].astype(str).str.replace('\u200f', '').str.strip()
            df_clean[col] = df_clean[col].replace(['nan', 'None', '', '<NA>', 'NaN'], pd.NA)
            
    return df_clean

# ==========================================
# 3. توابع محاسباتی زمان و تشخیص Null (حل مشکل None)
# ==========================================
def is_missing_time(val):
    """این تابع با دقت بالا تمام حالت‌های خالی بودن سلول را فیلتر می‌کند"""
    if pd.isna(val) or val is None: 
        return True
    val_str = str(val).strip().lower()
    if val_str in ['-', '', 'none', 'nan', 'nat', '<na>']: 
        return True
    return False

def time_to_minutes(t_str):
    if is_missing_time(t_str): return None
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
# 4. تجمیع و عارضه‌یابی
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
            
        ins = group[col_in].dropna().tolist() if col_in else []
        outs = group[col_out].dropna().tolist() if col_out else []
        s_starts = group[col_s_start].dropna().tolist() if col_s_start else []
        s_ends = group[col_s_end].dropna().tolist() if col_s_end else []
        
        daily_records.append({
            'تاریخ': day,
            'ورود_ثبت_شده': ins[0] if ins else None,
            'خروج_ثبت_شده': outs[-1] if outs else None,
            'شروع_شیفت_سیستم': s_starts[0] if s_starts else None,
            'پایان_شیفت_سیستم': s_ends[0] if s_ends else None,
        })
    return pd.DataFrame(daily_records)

def sync_and_diagnose(daily_df, shift_pattern_json):
    results = []
    
    for i, row in daily_df.iterrows():
        actual_in = row['ورود_ثبت_شده']
        actual_out = row['خروج_ثبت_شده']
        sys_shift_in = row['شروع_شیفت_سیستم']
        sys_shift_out = row['پایان_شیفت_سیستم']
        
        in_missing = is_missing_time(actual_in)
        out_missing = is_missing_time(actual_out)
        
        sys_shift_in_str = str(sys_shift_in).strip()
        sys_shift_out_str = str(sys_shift_out).strip()
        
        # آیا روز کاری است یا تعطیل؟
        shift_is_workday = not is_missing_time(sys_shift_in) and 'تعطیل' not in sys_shift_in_str
        
        diagnosis = []
        action = []
        is_error = False
        
        # 1. بررسی عدم ثبت ورود/خروج (سناریوی ۱۴۰۵/۰۶/۰۹)
        if shift_is_workday:
            if in_missing:
                diagnosis.append("خطا: عدم ثبت ورود")
                action.append(f"ورود باید بر اساس شیفت {sys_shift_in_str[:5]} ثبت شود")
                is_error = True
            if out_missing:
                diagnosis.append("خطا: عدم ثبت خروج")
                if "خطا: عدم ثبت ورود" not in diagnosis: 
                    action.append(f"خروج باید بر اساس شیفت {sys_shift_out_str[:5]} ثبت شود")
                is_error = True
                
        # 2. کشف خروج از الگو (تردد مازاد در روز استراحت - سناریوی ۱۴۰۵/۰۶/۱۳)
        if not in_missing:
            if not shift_is_workday:
                diagnosis.append("آلارم: خروج از الگو (تردد در روز استراحت)")
                action.append("بررسی تردد مازاد یا لزوم اصلاح شیفت")
                is_error = True
                
        # 3. محاسبه تاخیر و تعجیل
        if not in_missing and shift_is_workday:
            in_diff = calc_time_diff(actual_in, sys_shift_in, is_checkout=False)
            if in_diff > 15:
                diagnosis.append(f"تاخیر در ورود ({in_diff} دقیقه)")
                is_error = True
                
        if not out_missing and shift_is_workday:
            out_diff = calc_time_diff(actual_out, sys_shift_out, is_checkout=True)
            if out_diff < -15:
                diagnosis.append(f"تعجیل در خروج ({abs(out_diff)} دقیقه)")
                is_error = True
                
        if not diagnosis:
            diagnosis.append("عادی (بدون مغایرت)")
            action.append("-")
            
        results.append({
            'تاریخ': row['تاریخ'],
            'شیفت سیستم (ورود)': sys_shift_in_str if not is_missing_time(sys_shift_in) else '-',
            'شیفت سیستم (خروج)': sys_shift_out_str if not is_missing_time(sys_shift_out) else '-',
            'تردد (ورود)': str(actual_in) if not in_missing else '-',
            'تردد (خروج)': str(actual_out) if not out_missing else '-',
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
st.title("🕵️‍♂️ سیستم هوشمند عارضه‌یابی تردد (نسخه JSON)")

# منوی کناری برای تنظیمات JSON
st.sidebar.title("⚙️ تنظیمات شیفت")
json_input = st.sidebar.text_area("الگوی شیفت خود را با فرمت JSON وارد کنید:", value=DEFAULT_JSON, height=350)

try:
    shift_pattern = json.loads(json_input)
    st.sidebar.success("✅ فرمت JSON معتبر است.")
except Exception:
    st.sidebar.error("⚠️ خطای فرمت JSON! لطفا بررسی کنید.")
    shift_pattern = []

# صفحه اصلی آپلود
st.markdown("دیگر نیازی به فایل اکسل شیفت نیست! الگو از طریق منوی کناری (JSON) مدیریت می‌شود. فقط **گزارش تردد** را آپلود کنید.")
attendance_file = st.file_uploader("📥 آپلود فایل گزارش تردد", type=['xlsx', 'xls'])

if attendance_file and shift_pattern:
    with st.spinner("در حال تطبیق ترددها..."):
        try:
            df_attendance_clean = parse_attendance_file(attendance_file)
            daily_attendance = aggregate_daily_attendance(df_attendance_clean)
            
            # اجرای هسته عارضه‌یابی با منطق جدید و رفع باگ None
            diagnosis_df = sync_and_diagnose(daily_attendance, shift_pattern)
            
            st.success("✅ پردازش پایان یافت.")
            
            st.subheader("🔍 جدول هوشمند مقایسه و عارضه‌یابی")
            styled_diagnosis = style_diagnosis_df(diagnosis_df)
            st.dataframe(styled_diagnosis, width="stretch", hide_index=True, height=600)
            
            # آماده‌سازی اکسل خروجی
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
