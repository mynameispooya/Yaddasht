import sqlite3
import pandas as pd
import streamlit as st

# تنظیمات اولیه صفحه
st.set_page_config(page_title="مدیریت دیتابیس اکسل", layout="wide")

# نام فایل دیتابیس محلی (اطلاعات در این فایل ماندگار خواهد بود)
DB_FILE = "questionnaires.db"


def get_db_connection():
    """ایجاد اتصال به دیتابیس SQLite"""
    conn = sqlite3.connect(DB_FILE)
    return conn


def init_db():
    """ساخت جدول در دیتابیس در صورت عدم وجود"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # تعریف جدول با ستون responder_name به عنوان کلید اصلی (Primary Key)
    # توجه: ستون‌های دیگر را بر اساس نیاز اکسل خود تغییر دهید
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS responses (
            responder_name TEXT PRIMARY KEY,
            gender TEXT,
            age TEXT,
            data_json TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """
    )
    conn.commit()
    conn.close()


# مقداردهی اولیه دیتابیس
init_db()

st.title("📊 سیستم ثبت و به‌روزرسانی پرسشنامه‌ها")

# ۱. بخش آپلود فایل اکسل
uploaded_file = st.file_uploader(
    "فایل اکسل جدید را آپلود کنید", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        # نمایش پیش‌نمایش فایل آپلود شده
        st.write("پیش‌نمایش فایل آپلود شده:")
        st.dataframe(df.head())

        if st.button("ذخیره/به‌روزرسانی در دیتابیس"):
            conn = get_db_connection()
            cursor = conn.cursor()

            updated_count = 0

            # پیمایش ردیف‌های فایل اکسل و ذخیره در دیتابیس
            for _, row in df.iterrows():
                # فرضا ستون نام تکمیل کننده در اکسل name_column نام دارد
                # نام ستون را مطابق فایل اکسل خود درج کنید (مثلاً G1, G2, G20)
                responder_name = str(row["نام تکمیل کننده"]).strip()
                gender = str(row.get("جنسیت", ""))
                age = str(row.get("سن", ""))
                data_json = row.to_json(force_unicode=True)

                # دستور REPLACE INTO در صورت وجود کلید تکراری، سطر قبلی را جایگزین می‌کند
                cursor.execute(
                    """
                    INSERT INTO responses (responder_name, gender, age, data_json)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(responder_name) DO UPDATE SET
                        gender = excluded.gender,
                        age = excluded.age,
                        data_json = excluded.data_json,
                        created_at = CURRENT_TIMESTAMP
                """,
                    (responder_name, gender, age, data_json),
                )

                updated_count += 1

            conn.commit()
            conn.close()

            st.success(
                f"تعداد {updated_count} ردیف با موفقیت در دیتابیس ثبت/به‌روزرسانی شد!"
            )

    except Exception as e:
        st.error(f"خطا در پردازش فایل: {e}")

st.divider()

# ۲. نمایش اطلاعات موجود در دیتابیس ماندگار
st.subheader("📁 اطلاعات ثبت‌شده در دیتابیس")

conn = get_db_connection()
db_df = pd.read_sql_query("SELECT * FROM responses", conn)
conn.close()

if not db_df.empty:
    st.write(f"کل رکوردها: {len(db_df)}")
    st.dataframe(db_df)

    # امکان دانلود خروجی از دیتابیس فعلی
    csv = db_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="دانلود کل دیتابیس به صورت CSV",
        data=csv,
        file_name="database_export.csv",
        mime="text/csv",
    )
else:
    st.info("هنوز هیچ داده‌ای در دیتابیس ثبت نشده است.")
