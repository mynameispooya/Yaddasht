import streamlit as st
import jdatetime
import pandas as pd
from datetime import timedelta

st.set_page_config(page_title="محاسبه‌گر روزهای کاری", layout="wide", page_icon="📅")

st.title("📅 محاسبه‌گر روزهای کاری تقویم شمسی")
st.write("بازه زمانی مورد نظر خود را انتخاب کنید. روزهای جمعه و تعطیلات ثابت شمسی به صورت خودکار لحاظ می‌شوند. برای سایر تعطیلات (مانند مناسبت‌های قمری یا تعطیلات ناگهانی)، کافیست تیک ستون **«تعطیل است؟»** را در جدول فعال کنید.")

st.markdown("---")

# لیست تعطیلات ثابت شمسی برای برچسب‌گذاری خودکار
FIXED_SOLAR_HOLIDAYS = {
    (1, 1): "عید نوروز", (1, 2): "عید نوروز", (1, 3): "عید نوروز", (1, 4): "عید نوروز",
    (1, 12): "روز جمهوری اسلامی", (1, 13): "روز طبیعت",
    (3, 14): "رحلت امام خمینی (ره)", (3, 15): "قیام ۱۵ خرداد",
    (11, 22): "پیروزی انقلاب اسلامی", (12, 29): "ملی شدن صنعت نفت"
}

WEEKDAYS_FA = {
    0: "شنبه", 1: "یکشنبه", 2: "دوشنبه", 
    3: "سه‌شنبه", 4: "چهارشنبه", 5: "پنج‌شنبه", 6: "جمعه"
}

# بخش انتخاب تاریخ
col1, col2 = st.columns(2)

today = jdatetime.date.today()

with col1:
    st.subheader("تاریخ شروع")
    c1, c2, c3 = st.columns(3)
    start_year = c1.number_input("سال شروع", 1300, 1500, today.year)
    start_month = c2.number_input("ماه شروع", 1, 12, today.month)
    start_day = c3.number_input("روز شروع", 1, 31, today.day)

with col2:
    st.subheader("تاریخ پایان")
    c4, c5, c6 = st.columns(3)
    # پیش‌فرض تاریخ پایان را یک ماه بعد در نظر می‌گیریم
    end_year = c4.number_input("سال پایان", 1300, 1500, today.year)
    end_month = c5.number_input("ماه پایان", 1, 12, (today.month % 12) + 1)
    end_day = c6.number_input("روز پایان", 1, 31, today.day)

try:
    start_date = jdatetime.date(start_year, start_month, start_day)
    end_date = jdatetime.date(end_year, end_month, end_day)
except ValueError:
    st.error("⚠️ تاریخ وارد شده نامعتبر است. (مثلاً ماه‌های نیمه دوم سال ۳۰ روزه هستند).")
    st.stop()

if start_date > end_date:
    st.error("⚠️ تاریخ شروع نمی‌تواند پس از تاریخ پایان باشد!")
    st.stop()

st.markdown("---")

# تولید لیست تمام روزهای بین تاریخ شروع و پایان
delta = end_date - start_date
dates_list = [start_date + timedelta(days=i) for i in range(delta.days + 1)]

# ساخت دیتافریم برای نمایش شماتیک و تعاملی
data = []
for d in dates_list:
    weekday_index = d.weekday()
    is_friday = (weekday_index == 6)
    
    # بررسی تعطیلات ثابت شمسی
    solar_event = FIXED_SOLAR_HOLIDAYS.get((d.month, d.day), "")
    
    # تعیین وضعیت پیش‌فرض تعطیلی (اگر جمعه باشد یا مناسبت ثابت شمسی داشته باشد)
    is_holiday_default = is_friday or bool(solar_event)
    
    status_text = "جمعه" if is_friday else solar_event if solar_event else "روز کاری"
    
    data.append({
        "تاریخ": d.strftime("%Y/%m/%d"),
        "روز هفته": WEEKDAYS_FA[weekday_index],
        "وضعیت / مناسبت": status_text,
        "تعطیل است؟": is_holiday_default  # این ستون چک‌باکس خواهد شد
    })

df = pd.DataFrame(data)

st.subheader("🗓️ تقویم شماتیک و اعمال تعطیلات")
st.info("💡 **راهنما:** جدول زیر بر اساس تاریخ‌های انتخابی شما ساخته شده است. جمعه‌ها و تعطیلات ثابت شمسی تیک خورده‌اند. اگر روزی تعطیل رسمی قمری است یا به هر دلیلی تعطیل اعلام شده، تیک ستون **«تعطیل است؟»** را برای آن روز روشن کنید.")

# نمایش جدول تعاملی
# فقط ستون آخر قابل ویرایش است
edited_df = st.data_editor(
    df,
    disabled=["تاریخ", "روز هفته", "وضعیت / مناسبت"],
    use_container_width=True,
    hide_index=True,
    height=400
)

# محاسبه نهایی
total_days = len(edited_df)
holiday_days = edited_df["تعطیل است؟"].sum()
working_days = total_days - holiday_days

st.markdown("---")
st.subheader("📊 نتیجه نهایی")

col_res1, col_res2, col_res3 = st.columns(3)
col_res1.metric("کل روزهای بازه", f"{total_days} روز")
col_res2.metric("تعداد روزهای تعطیل", f"{holiday_days} روز")
col_res3.metric("تعداد روزهای کاری", f"{working_days} روز", delta_color="normal")

if working_days > 0:
    st.success(f"✅ بر اساس تنظیمات شما، در این بازه زمانی **{working_days} روز کاری** وجود دارد.")
else:
    st.warning("⚠️ در این بازه هیچ روز کاری وجود ندارد!")
