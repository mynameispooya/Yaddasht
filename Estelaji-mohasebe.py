import streamlit as st
import jdatetime
from datetime import timedelta

# تنظیمات اولیه صفحه
st.set_page_config(page_title="محاسبه‌گر استعلاجی", layout="centered", page_icon="🏥")

st.title("🏥 محاسبه‌گر زمان پایان مرخصی استعلاجی")
st.write("تاریخ شروع و مدت زمان استعلاجی خود را وارد کنید تا سیستم تاریخ دقیق پایان مرخصی و زمان بازگشت به کار را بر اساس تقویم شمسی محاسبه کند.")

st.markdown("---")

# تابع کمکی برای محاسبه دقیق ماه‌ها در تقویم شمسی
def add_months_jalali(start_date, months):
    year = start_date.year + (start_date.month + months - 1) // 12
    month = (start_date.month + months - 1) % 12 + 1
    day = start_date.day
    
    # تنظیم روز در صورتی که بیشتر از ظرفیت ماه مقصد باشد (مثلا ۳۱ در ماه‌های نیمه دوم سال)
    if month <= 6:
        max_days = 31
    elif month <= 11:
        max_days = 30
    else:
        # بررسی سال کبیسه
        is_leap = jdatetime.date(year, 1, 1).isleap()
        max_days = 30 if is_leap else 29
        
    day = min(day, max_days)
    return jdatetime.date(year, month, day)

# دریافت تاریخ امروز برای مقادیر پیش‌فرض
today = jdatetime.date.today()

st.subheader("۱. انتخاب تاریخ شروع استعلاجی")
col1, col2, col3 = st.columns(3)

with col1:
    year = st.number_input("سال", min_value=1400, max_value=1450, value=today.year, step=1)
with col2:
    month = st.number_input("ماه", min_value=1, max_value=12, value=today.month, step=1)
with col3:
    day = st.number_input("روز", min_value=1, max_value=31, value=today.day, step=1)

# اعتبارسنجی تاریخ وارد شده
try:
    start_date = jdatetime.date(year, month, day)
except ValueError:
    st.error("⚠️ تاریخ وارد شده نامعتبر است. لطفاً تطابق روز و ماه را بررسی کنید (مثلاً ماه مهر ۳۰ روزه است).")
    st.stop()

st.subheader("۲. انتخاب مدت زمان استعلاجی")
col_unit, col_val = st.columns(2)
with col_unit:
    unit = st.selectbox("واحد زمان", ["روز", "هفته", "ماه"])
with col_val:
    duration = st.number_input("تعداد", min_value=1, value=1, step=1)

st.markdown("---")

# دکمه محاسبه
if st.button("محاسبه تاریخ‌ها", use_container_width=True, type="primary"):
    # محاسبه تاریخ بازگشت به کار (روز بعد از اتمام مرخصی)
    if unit == "روز":
        return_date = start_date + timedelta(days=duration)
    elif unit == "هفته":
        return_date = start_date + timedelta(weeks=duration)
    elif unit == "ماه":
        return_date = add_months_jalali(start_date, duration)
        
    # آخرین روز استعلاجی دقیقا یک روز قبل از تاریخ بازگشت به کار است
    end_date = return_date - timedelta(days=1)
    
    # نمایش نتایج
    st.markdown("### 📊 نتایج محاسبه:")
    
    st.info(f"🟢 **تاریخ شروع استعلاجی:** {start_date.strftime('%A %d %B %Y')}")
    st.warning(f"🔴 **آخرین روز استعلاجی:** {end_date.strftime('%A %d %B %Y')}")
    st.success(f"💼 **تاریخ شروع به کار:** {return_date.strftime('%A %d %B %Y')}")
