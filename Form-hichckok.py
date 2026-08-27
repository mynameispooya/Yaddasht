import json
import sqlite3
import pandas as pd
import streamlit as st

st.set_page_config(page_title="مدیریت دیتابیس پرسشنامه‌ها", layout="wide")

DB_FILE = "questionnaires.db"


def get_db_connection():
    return sqlite3.connect(DB_FILE)


def init_db():
    """ساخت جدول با ساختار مشخص بر اساس پرسشنامه‌ها"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS main_responses (
            form_id TEXT PRIMARY KEY,
            gender TEXT,
            age TEXT,
            work_experience TEXT,
            current_service_experience TEXT,
            education TEXT,
            organizational_post TEXT,
            raw_json TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """
    )
    conn.commit()
    conn.close()


def flatten_json_form(form_data):
    """استخراج داده‌های شناسنامه‌ای و آماده‌سازی برای ثبت در دیتابیس"""
    form_id = form_data.get("form_id")
    pages = form_data.get("pages", {})

    demographics = {}
    # پیدا کردن اطلاعات شناسنامه‌ای از اولین صفحه‌ای که آن را دارد
    for page_key, page_val in pages.items():
        if "demographics" in page_val:
            demographics = page_val["demographics"]
            break

    return {
        "form_id": form_id,
        "gender": demographics.get("gender"),
        "age": demographics.get("age"),
        "work_experience": demographics.get("work_experience"),
        "current_service_experience": demographics.get(
            "current_service_experience"
        ),
        "education": demographics.get("education"),
        "organizational_post": demographics.get("organizational_post"),
        "raw_json": json.dumps(form_data, ensure_ascii=False),
    }


# مقداردهی اولیه دیتابیس
init_db()

st.title("📥 بارگذاری و ثبت فایل JSON در دیتابیس")

# ۱. بخش آپلود و پردازش فایل JSON
uploaded_json = st.file_uploader("فایل JSON پرسشنامه را آپلود کنید", type=["json"])

if uploaded_json is not None:
    try:
        raw_data = json.load(uploaded_json)

        # پشتیبانی از حالت تک‌فرمی یا لیستی از فرم‌ها
        forms_list = raw_data if isinstance(raw_data, list) else [raw_data]

        st.success(f"تعداد {len(forms_list)} فرم در فایل JSON شناسایی شد.")

        # پیش‌نمایش داده‌های استخراج‌شده
        parsed_records = [flatten_json_form(f) for f in forms_list if "form_id" in f]
        preview_df = pd.DataFrame(parsed_records).drop(columns=["raw_json"])

        st.write("📋 **پیش‌نمایش فرم‌های شناسایی‌شده:**")
        st.dataframe(preview_df)

        if st.button("💾 ذخیره / به‌روزرسانی در دیتابیس"):
            conn = get_db_connection()
            cursor = conn.cursor()
            saved_count = 0

            for rec in parsed_records:
                cursor.execute(
                    """
                    INSERT INTO main_responses (
                        form_id, gender, age, work_experience, 
                        current_service_experience, education, organizational_post, raw_json
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(form_id) DO UPDATE SET
                        gender = excluded.gender,
                        age = excluded.age,
                        work_experience = excluded.work_experience,
                        current_service_experience = excluded.current_service_experience,
                        education = excluded.education,
                        organizational_post = excluded.organizational_post,
                        raw_json = excluded.raw_json,
                        updated_at = CURRENT_TIMESTAMP
                """,
                    (
                        rec["form_id"],
                        rec["gender"],
                        rec["age"],
                        rec["work_experience"],
                        rec["current_service_experience"],
                        rec["education"],
                        rec["organizational_post"],
                        rec["raw_json"],
                    ),
                )
                saved_count += 1

            conn.commit()
            conn.close()
            st.success(
                f"✅ تعداد {saved_count} فرم با موفقیت در دیتابیس ذخیره و جایگزین شد!"
            )

    except Exception as e:
        st.error(f"خطا در خواندن فایل JSON: {e}")

st.divider()

# ۲. نمایش محتوای ماندگار دیتابیس
st.subheader("🗄️ اطلاعات ثبت‌شده در دیتابیس (ماندگار)")

conn = get_db_connection()
db_df = pd.read_sql_query(
    "SELECT form_id, gender, age, work_experience, current_service_experience, education, organizational_post, updated_at FROM main_responses",
    conn,
)
conn.close()

if not db_df.empty:
    st.write(f"تعداد کل کدهای یکتا ثبت‌شده: {len(db_df)}")
    st.dataframe(db_df)
else:
    st.info("دیتابیس در حال حاضر خالی است.")
