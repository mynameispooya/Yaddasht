import streamlit as st
import pandas as pd

st.title("اپلیکیشن کنترل تردد پرسنل")

uploaded_file = st.file_uploader("فایل اکسل تردد را آپلود کنید", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # مرحله 1: پاکسازی ستون‌ها
    if "اطلاعات تردد" in df.columns:
        df = df.drop(columns=["اطلاعات تردد"], errors="ignore")

    # مرحله 2: گروه‌بندی بر اساس کد پرسنلی و تاریخ
    tardod_naqs = []
    tardod_mashkook = []

    grouped = df.groupby(["کدپرسنلی", "تاریخ"])
    for (code, date), group in grouped:
        name = group["نام و نام خانوادگی"].iloc[0]
        day = group["روز"].iloc[0]

        # شمارش ورود و خروج
        vorood = group["ورود"].dropna().tolist()
        khorooj = group["خروج"].dropna().tolist()

        # تردد ناقص
        if len(vorood) == 0 or len(khorooj) == 0:
            tardod_naqs.append([code, name, date, day, vorood, khorooj])

        # تردد مشکوک (بیش از یک ورود یا خروج)
        if len(vorood) > 1 or len(khorooj) > 1:
            tardod_mashkook.append([code, name, date, day, vorood, khorooj])

    # ساخت دیتافریم خروجی
    df_naqs = pd.DataFrame(tardod_naqs, columns=["کدپرسنلی","نام و نام خانوادگی","تاریخ","روز","ورود","خروج"])
    df_mashkook = pd.DataFrame(tardod_mashkook, columns=["کدپرسنلی","نام و نام خانوادگی","تاریخ","روز","ورود","خروج"])

    # ذخیره فایل‌ها
    df_naqs.to_excel("tardod_naqs.xlsx", index=False)
    df_mashkook.to_excel("tardod_mashkook.xlsx", index=False)

    st.success("پاکسازی و پردازش انجام شد ✅")
    st.write("📂 فایل‌های خروجی تولید شدند: tardod_naqs.xlsx و tardod_mashkook.xlsx")
