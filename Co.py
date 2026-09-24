import streamlit as st
import pandas as pd

st.title("اپلیکیشن کنترل تردد پرسنل")

uploaded_file = st.file_uploader("فایل اکسل تردد را آپلود کنید", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # پاکسازی نام ستون‌ها (حذف فاصله و نیم‌فاصله)
    df.columns = df.columns.str.strip()

    # تغییر نام ستون‌ها به انگلیسی برای راحتی
    rename_map = {
        "کدپرسنلی": "personnel_id",
        "نام و نام خانوادگی": "name",
        "تاریخ": "date",
        "روز": "day",
        "ورود": "entry",
        "خروج": "exit",
        "وضعیت": "status",
        "شروع شیفت": "shift_start",
        "پایان شیفت": "shift_end"
    }
    df = df.rename(columns=rename_map)

    # مرحله 2: گروه‌بندی بر اساس کد پرسنلی و تاریخ
    tardod_naqs = []
    tardod_mashkook = []

    grouped = df.groupby(["personnel_id", "date"])
    for (code, date), group in grouped:
        name = group["name"].iloc[0]
        day = group["day"].iloc[0]

        vorood = group["entry"].dropna().tolist()
        khorooj = group["exit"].dropna().tolist()

        # تردد ناقص
        if len(vorood) == 0 or len(khorooj) == 0:
            tardod_naqs.append([code, name, date, day, vorood, khorooj])

        # تردد مشکوک (بیش از یک ورود یا خروج)
        if len(vorood) > 1 or len(khorooj) > 1:
            tardod_mashkook.append([code, name, date, day, vorood, khorooj])

    # ساخت دیتافریم خروجی
    df_naqs = pd.DataFrame(tardod_naqs, columns=["کدپرسنلی","نام و نام خانوادگی","تاریخ","روز","ورود","خروج"])
    df_mashkook = pd.DataFrame(tardod_mashkook, columns=["کدپرسنلی","نام و نام خانوادگی","تاریخ","روز","ورود","خروج"])

    # امکان دانلود مستقیم از Streamlit
    st.download_button("دانلود فایل تردد ناقص", df_naqs.to_csv(index=False).encode("utf-8"), "tardod_naqs.csv")
    st.download_button("دانلود فایل تردد مشکوک", df_mashkook.to_csv(index=False).encode("utf-8"), "tardod_mashkook.csv")

    st.success("✅ پردازش کامل شد. حالا می‌توانید فایل‌ها را دانلود کنید.")
