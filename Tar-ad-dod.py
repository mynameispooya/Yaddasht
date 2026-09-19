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

def parse_time(t_str):
    if pd.isna(t_str) or str(t_str).strip() == '' or str(t_str).strip() == 'nan':
        return 0
    try:
        parts = str(t_str).split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except:
        return 0

def format_time(mins):
    if mins <= 0:
        return "00:00"
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h:02d}:{m:02d}"

st.sidebar.header("تنظیمات تعطیلات")
manual_holidays_input = st.sidebar.text_area(
    "تاریخ‌های تعطیل دستی (با کاما جدا کنید)", 
    placeholder="مثال: 1402/01/01, 1402/01/02",
    help="در صورتی که روز خاصی تعطیل است اما در تقویم رسمی لحاظ نشده، تاریخ آن را در این قسمت وارد کنید."
)

manual_holidays_list = [x.strip() for x in manual_holidays_input.split(',') if x.strip()]

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
                vazeiat = str(row.get('وضعیت', ''))
                start_t = row.get('شروع', '')
                end_t = row.get('پایان', '')
                
                if pd.isna(rooz) or rooz.strip() == '' or rooz == 'nan':
                    continue
                    
                parts = rooz.split()
                if len(parts) < 2:
                    continue
                    
                day_name = parts[0]
                date_str = parts[-1]
                
                if date_str not in daily_stats:
                    daily_stats[date_str] = {
                        'روز': rooz,
                        'کارکرد روزانه': 0,
                        'اضافه کاری': 0,
                        'اضافه کاری روز تعطیل': 0,
                        'جمعه کاری': 0
                    }
                    
                if 'حضور' in vazeiat and not pd.isna(start_t) and not pd.isna(end_t):
                    t_start = parse_time(start_t)
                    t_end = parse_time(end_t)
                    
                    if t_start == 0 and t_end == 0:
                        continue
                        
                    # قفل کردن زمان شروع روی 07:15 (معادل 435 دقیقه)
                    t_start = max(t_start, 435)
                    
                    if t_start >= t_end:
                        continue
                        
                    is_friday = 'جمعه' in day_name
                    is_holiday = date_str in manual_holidays_list
                    
                    if not is_holiday and not is_friday:
                        try:
                            y, m, d = map(int, date_str.split('/'))
                            g_date = jdatetime.date(y, m, d).togregorian()
                            ir_holidays = holidays.IR(years=[g_date.year])
                            if g_date in ir_holidays:
                                is_holiday = True
                        except:
                            pass
                    
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
        
        comp_df = pd.DataFrame(comprehensive_results)
        
        st.subheader("📊 گزارش جامع")
        st.dataframe(comp_df, use_container_width=True)
        
        # ایجاد فایل اکسل نهایی در حافظه موقت (Buffer)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            comp_df.to_excel(writer, sheet_name='گزارش جامع', index=False)
            for fname, idf in all_individual_dfs.items():
                # محدود کردن نام شیت به ۳۱ کاراکتر (محدودیت اکسل)
                sheet_name = fname.replace('.xlsx', '')[:31]
                idf.to_excel(writer, sheet_name=sheet_name, index=False)
                
        st.download_button(
            label="📥 دانلود فایل اکسل گزارش نهایی (شامل گزارش جامع و ریز هر شخص)",
            data=buffer.getvalue(),
            file_name="گزارش_تردد_جامع.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
