import streamlit as st
import pandas as pd
import jdatetime
import holidays
import io
import base64

st.set_page_config(page_title="محاسبه کارکرد و تردد", layout="wide", page_icon="⏱️")

st.markdown('''
    <style>
        .stApp {
            direction: rtl;
            font-family: 'Tahoma', 'B Yekan', sans-serif;
        }
    </style>
''', unsafe_allow_html=True)

st.title("⏱️ سیستم هوشمند محاسبه کارکرد و تردد")
st.write("این اپلیکیشن فایل‌های اکسل تردد را دریافت کرده و کارکرد روزانه، اضافه کاری، اضافه کاری تعطیلات و جمعه کاری را محاسبه می‌کند.")
st.info("نکته: در محاسبات، ساعت شروع برای همه افراد از 07:15 در نظر گرفته می‌شود (تاخیر قبل از آن محاسبه نمی‌گردد) و خروج بعد از 16:00 به عنوان اضافه کاری لحاظ می‌شود.")

def normalize_date(d_str):
    """استانداردسازی فرمت تاریخ به شکل YYYY/MM/DD برای مقایسه دقیق"""
    try:
        parts = str(d_str).strip().split('/')
        if len(parts) == 3:
            return f"{int(parts[0]):04d}/{int(parts[1]):02d}/{int(parts[2]):02d}"
    except:
        pass
    return str(d_str).strip()

def safe_time_str(t_val):
    if pd.isna(t_val): return ""
    t_str = str(t_val).strip()
    if t_str == 'nan' or t_str == '': return ""
    try:
        parts = t_str.split(':')
        if len(parts) >= 2:
            return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
    except:
        pass
    return t_str

def parse_time(t_str):
    if pd.isna(t_str): return 0
    t_val = str(t_str).strip()
    if t_val == '' or t_val == 'nan': return 0
    try:
        parts = t_val.split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except:
        return 0

def format_time(mins):
    if mins <= 0:
        return "00:00"
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h:02d}:{m:02d}"

# توابع مربوط به استایل و هایلایت جدول
def highlight_cells(val, color):
    if val != "00:00" and val != 0 and pd.notna(val) and val != "-":
        return f'background-color: {color}; color: #000000'
    return ''

def highlight_day_type(val):
    if val == 'جمعه':
        return 'background-color: #f8d7da; font-weight: bold; color: #721c24'
    elif val in ['تعطیل رسمی', 'تعطیل دستی']:
        return 'background-color: #fff3cd; font-weight: bold; color: #856404'
    return ''

st.sidebar.header("تنظیمات تاریخ و تعطیلات")

start_date_filter = st.sidebar.text_input(
    "📅 تاریخ شروع محاسبات (اختیاری)",
    placeholder="مثال: 1405/06/07",
    help="اگر تاریخی وارد کنید، روزهای قبل از این تاریخ نادیده گرفته می‌شوند. روی تمامی فایل‌ها اعمال می‌شود."
)

manual_holidays_input = st.sidebar.text_area(
    "🏖️ تاریخ‌های تعطیل دستی (با کاما جدا کنید)", 
    placeholder="مثال: 1405/01/01, 1405/01/02",
    help="در صورتی که روز خاصی تعطیل است اما در تقویم رسمی لحاظ نشده، تاریخ آن را در این قسمت وارد کنید."
)

manual_holidays_list = [normalize_date(x) for x in manual_holidays_input.split(',') if x.strip()]

uploaded_files = st.file_uploader("فایل‌های اکسل تردد را انتخاب کنید", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("شروع محاسبه 🚀"):
        comprehensive_results = []
        all_individual_dfs = {}
        
        progress_bar = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            try:
                df = pd.read_excel(file, header=0)
            except Exception as e:
                st.error(f"خطا در خواندن فایل {file.name}: {e}")
                continue
                
            daily_stats = {}
            
            for index, row in df.iterrows():
                rooz = str(row.get('روز', ''))
                vazeiat = str(row.get('وضعیت', '')).strip()
                start_t = row.get('شروع', '')
                end_t = row.get('پایان', '')
                
                if pd.isna(rooz) or rooz.strip() == '' or rooz == 'nan':
                    continue
                    
                parts = rooz.split()
                if len(parts) < 2:
                    continue
                    
                day_name = parts[0]
                date_str = parts[-1]
                norm_date_str = normalize_date(date_str)
                
                # فیلتر کردن بر اساس تاریخ شروع وارد شده توسط کاربر
                if start_date_filter and start_date_filter.strip():
                    norm_start_filter = normalize_date(start_date_filter)
                    if norm_date_str < norm_start_filter:
                        continue
                
                # تعیین نوع روز (عادی، جمعه، تعطیل)
                if date_str not in daily_stats:
                    is_friday = 'جمعه' in day_name
                    is_manual_holiday = norm_date_str in manual_holidays_list
                    is_official_holiday = False
                    
                    if not is_manual_holiday and not is_friday:
                        try:
                            y, m, d = map(int, norm_date_str.split('/'))
                            g_date = jdatetime.date(y, m, d).togregorian()
                            ir_holidays = holidays.IR(years=[g_date.year])
                            if g_date in ir_holidays:
                                is_official_holiday = True
                        except:
                            pass
                            
                    day_type_str = "عادی"
                    if is_friday:
                        day_type_str = "جمعه"
                    elif is_manual_holiday:
                        day_type_str = "تعطیل دستی"
                    elif is_official_holiday:
                        day_type_str = "تعطیل رسمی"
                        
                    daily_stats[date_str] = {
                        'روز': rooz,
                        'نوع روز': day_type_str,
                        'شروع': [],
                        'پایان': [],
                        'کارکرد روزانه': 0,
                        'اضافه کاری': 0,
                        'اضافه کاری روز تعطیل': 0,
                        'جمعه کاری': 0
                    }
                    
                # محاسبات زمانی تنها برای وضعیت «حضور»
                if vazeiat == 'حضور' and not pd.isna(start_t) and not pd.isna(end_t):
                    t_start = parse_time(start_t)
                    t_end = parse_time(end_t)
                    
                    if t_start == 0 and t_end == 0:
                        continue
                        
                    # ذخیره ساعات شروع و پایان برای نمایش در جدول
                    s_str = safe_time_str(start_t)
                    e_str = safe_time_str(end_t)
                    if s_str: daily_stats[date_str]['شروع'].append(s_str)
                    if e_str: daily_stats[date_str]['پایان'].append(e_str)
                    
                    # قفل کردن زمان شروع روی 07:15 (معادل 435 دقیقه)
                    t_start = max(t_start, 435)
                    
                    if t_start >= t_end:
                        continue
                        
                    is_friday = daily_stats[date_str]['نوع روز'] == 'جمعه'
                    is_holiday = daily_stats[date_str]['نوع روز'] in ['تعطیل رسمی', 'تعطیل دستی']
                    
                    if is_friday:
                        daily_stats[date_str]['جمعه کاری'] += (t_end - t_start)
                    elif is_holiday:
                        daily_stats[date_str]['اضافه کاری روز تعطیل'] += (t_end - t_start)
                    else:
                        calc_end_reg = min(t_end, 960) # پایان ساعت کاری عادی روی 16:00 (معادل 960 دقیقه)
                        if calc_end_reg > t_start:
                            daily_stats[date_str]['کارکرد روزانه'] += (calc_end_reg - t_start)
                            
                        calc_start_ot = max(t_start, 960) # شروع اضافه کاری از 16:00
                        if t_end > calc_start_ot:
                            daily_stats[date_str]['اضافه کاری'] += (t_end - calc_start_ot)
            
            # خلاصه سازی نتایج برای هر فایل
            result_list = []
            total_reg = 0
            total_ot = 0
            total_hol_ot = 0
            total_fri = 0
            
            for date_str, stats in daily_stats.items():
                result_list.append({
                    'روز': stats['روز'],
                    'تاریخ': date_str,
                    'نوع روز': stats['نوع روز'],
                    'شروع': " ، ".join(stats['شروع']) if stats['شروع'] else "-",
                    'پایان': " ، ".join(stats['پایان']) if stats['پایان'] else "-",
                    'کارکرد روزانه': format_time(stats['کارکرد روزانه']),
                    'اضافه کاری': format_time(stats['اضافه کاری']),
                    'اضافه کاری روز تعطیل': format_time(stats['اضافه کاری روز تعطیل']),
                    'جمعه کاری': format_time(stats['جمعه کاری'])
                })
                total_reg += stats['کارکرد روزانه']
                total_ot += stats['اضافه کاری']
                total_hol_ot += stats['اضافه کاری روز تعطیل']
                total_fri += stats['جمعه کاری']
                
            indiv_df = pd.DataFrame(result_list)
            
            # اضافه کردن ردیف مجموع به فایل شخص
            totals_row = {
                'روز': 'مجموع',
                'تاریخ': '-',
                'نوع روز': '-',
                'شروع': '-',
                'پایان': '-',
                'کارکرد روزانه': format_time(total_reg),
                'اضافه کاری': format_time(total_ot),
                'اضافه کاری روز تعطیل': format_time(total_hol_ot),
                'جمعه کاری': format_time(total_fri)
            }
            indiv_df.loc[len(indiv_df)] = totals_row
            
            all_individual_dfs[file.name] = indiv_df
            
            comprehensive_results.append({
                'نام فایل (شخص)': file.name,
                'مجموع کارکرد روزانه': format_time(total_reg),
                'مجموع اضافه کاری': format_time(total_ot),
                'مجموع اضافه کاری روز تعطیل': format_time(total_hol_ot),
                'مجموع جمعه کاری': format_time(total_fri)
            })
            
            progress_bar.progress((i + 1) / len(uploaded_files))
            
        st.success("محاسبات با موفقیت انجام شد!")
        
        # ---------------- نمایش ریز گزارش هر شخص با هایلایت ----------------
        st.markdown("---")
        st.subheader("📋 ریز محاسبات اشخاص (جهت بررسی و تایید)")
        
        for fname, idf in all_individual_dfs.items():
            with st.expander(f"📁 بررسی ریز گزارش فایل: {fname}", expanded=False):
                styled_df = idf.style
                
                # بررسی نسخه Pandas و اعمال متد مناسب برای استایل دهی
                apply_method = getattr(styled_df, 'map', getattr(styled_df, 'applymap', None))
                
                if apply_method:
                    styled_df = apply_method(lambda x: highlight_cells(x, '#d4edda'), subset=['کارکرد روزانه'])
                    styled_df = apply_method(lambda x: highlight_cells(x, '#cce5ff'), subset=['اضافه کاری'])
                    styled_df = apply_method(lambda x: highlight_cells(x, '#ffeeba'), subset=['اضافه کاری روز تعطیل'])
                    styled_df = apply_method(lambda x: highlight_cells(x, '#e2d9f3'), subset=['جمعه کاری'])
                    styled_df = apply_method(highlight_day_type, subset=['نوع روز'])
                    styled_df = apply_method(lambda x: 'background-color: #d1ecf1; font-weight: bold' if x == 'مجموع' else '', subset=['روز'])
                
                st.dataframe(styled_df, use_container_width=True)
        
        st.markdown("---")
        
        # ---------------- نمایش گزارش جامع ----------------
        comp_df = pd.DataFrame(comprehensive_results)
        st.subheader("📊 گزارش جامع (خلاصه تمامی افراد)")
        st.dataframe(comp_df, use_container_width=True)
        
        # ایجاد فایل اکسل نهایی در حافظه موقت
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            comp_df.to_excel(writer, sheet_name='گزارش جامع', index=False)
            for fname, idf in all_individual_dfs.items():
                sheet_name = fname.replace('.xlsx', '')[:31]
                idf.to_excel(writer, sheet_name=sheet_name, index=False)
                
        st.download_button(
            label="📥 دانلود فایل اکسل گزارش نهایی",
            data=buffer.getvalue(),
            file_name="گزارش_تردد_جامع.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
