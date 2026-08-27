import json
import sqlite3
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="مدیریت دیتابیس (Excel + JSON)", layout="wide"
)

DB_FILE = "questionnaires.db"


def get_db_connection():
    return sqlite3.connect(DB_FILE)


def init_db():
    """ایجاد یا بازسازی جدول دیتابیس با ساختار جدید"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # بررسی وجود ستون form_id در جدول موجود
    cursor.execute("PRAGMA table_info(responses)")
    columns = [col[1] for col in cursor.fetchall()]

    # اگر جدول وجود نداشت یا ستون form_id در آن نبود، جدول را از نو می‌سازیم
    if not columns or "form_id" not in columns:
        cursor.execute("DROP TABLE IF EXISTS responses")
        cursor.execute(
            """
            CREATE TABLE responses (
                form_id TEXT PRIMARY KEY,
                gender TEXT,
                age TEXT,
                work_experience TEXT,
                current_service_experience TEXT,
                education TEXT,
                organizational_post TEXT,
                raw_data TEXT,
                source_type TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """
        )

    conn.commit()
    conn.close()


def flatten_json_form(form_data):
    """استخراج اطلاعات از فایل JSON"""
    form_id = form_data.get("form_id")
    pages = form_data.get("pages", {})

    demographics = {}
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
        "raw_data": json.dumps(form_data, ensure_ascii=False),
    }


# مقداردهی اولیه و اطمینان از ساختار درست جدول
init_db()

st.title("📊 سیستم مدیریت و به‌روزرسانی پرسشنامه‌ها")

# ==========================================
# ورودی ۱: بارگذاری اولیه از طریق Excel
# ==========================================
st.subheader("۱. بارگذاری اولیه (Excel)")
st.caption(
    "از این بخش فقط یک‌بار برای درج اولیه داده‌ها در دیتابیس استفاده کنید."
)

uploaded_excel = st.file_uploader(
    "فایل اکسل را آپلود کنید", type=["xlsx", "xls"], key="excel_uploader"
)

if uploaded_excel is not None:
    try:
        df_excel = pd.read_excel(uploaded_excel)
        df_excel = df_excel.astype(str).replace("nan", None)

        st.write("پیش‌نمایش فایل اکسل:")
        st.dataframe(df_excel.head())

        if st.button("ثبت اولیه داده‌های اکسل در دیتابیس"):
            conn = get_db_connection()
            cursor = conn.cursor()
            inserted_count = 0

            for _, row in df_excel.iterrows():
                # شناسه فرم بر اساس ستون‌های مرسوم
                form_id = str(
                    row.get(
                        "نام تکمیل کننده",
                        row.get("form_id", row.get("نام تکمیل‌کننده", "")),
                    )
                ).strip()

                if form_id and form_id != "None":
                    gender = str(row.get("جنسیت", ""))
                    age = str(row.get("سن", ""))
                    work_exp = str(row.get("سابقه کار", ""))
                    serv_exp = str(row.get("سابقه در محل خدمت", ""))
                    edu = str(row.get("مدرک تحصیلات", ""))
                    post = str(row.get("پست سازمانی", ""))
                    raw_str = row.to_json(force_unicode=True)

                    cursor.execute(
                        """
                        INSERT INTO responses (
                            form_id, gender, age, work_experience,
                            current_service_experience, education, organizational_post,
                            raw_data, source_type
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'EXCEL')
                        ON CONFLICT(form_id) DO UPDATE SET
                            gender = excluded.gender,
                            age = excluded.age,
                            work_experience = excluded.work_experience,
                            current_service_experience = excluded.current_service_experience,
                            education = excluded.education,
                            organizational_post = excluded.organizational_post,
                            raw_data = excluded.raw_data,
                            source_type = 'EXCEL',
                            updated_at = CURRENT_TIMESTAMP
                    """,
                        (
                            form_id,
                            gender,
                            age,
                            work_exp,
                            serv_exp,
                            edu,
                            post,
                            raw_str,
                        ),
                    )
                    inserted_count += 1

            conn.commit()
            conn.close()
            st.success(
                f"تعداد {inserted_count} ردیف با موفقیت از اکسل وارد دیتابیس شد."
            )
            st.rerun()

    except Exception as e:
        st.error(f"خطا در پردازش فایل اکسل: {e}")

st.divider()

# ==========================================
# ورودی ۲: آپدیت و جایگزینی داده‌ها (JSON)
# ==========================================
st.subheader("۲. آپدیت و جایگزینی داده‌ها (JSON)")
st.caption(
    "از این بخش برای به‌روزرسانی یا اضافه کردن فرم‌های جدید از طریق فایل JSON استفاده کنید."
)

uploaded_json = st.file_uploader(
    "فایل JSON را آپلود کنید", type=["json"], key="json_uploader"
)

if uploaded_json is not None:
    try:
        raw_json_data = json.load(uploaded_json)
        forms_list = (
            raw_json_data
            if isinstance(raw_json_data, list)
            else [raw_json_data]
        )

        parsed_records = [
            flatten_json_form(f) for f in forms_list if "form_id" in f
        ]

        if parsed_records:
            preview_json_df = pd.DataFrame(parsed_records).drop(
                columns=["raw_data"]
            )
            st.write("پیش‌نمایش فرم‌های JSON:")
            st.dataframe(preview_json_df)

            if st.button("به‌روزرسانی دیتابیس با فایل JSON"):
                conn = get_db_connection()
                cursor = conn.cursor()
                updated_count = 0

                for rec in parsed_records:
                    cursor.execute(
                        """
                        INSERT INTO responses (
                            form_id, gender, age, work_experience,
                            current_service_experience, education, organizational_post,
                            raw_data, source_type
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'JSON')
                        ON CONFLICT(form_id) DO UPDATE SET
                            gender = excluded.gender,
                            age = excluded.age,
                            work_experience = excluded.work_experience,
                            current_service_experience = excluded.current_service_experience,
                            education = excluded.education,
                            organizational_post = excluded.organizational_post,
                            raw_data = excluded.raw_data,
                            source_type = 'JSON',
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
                            rec["raw_data"],
                        ),
                    )
                    updated_count += 1

                conn.commit()
                conn.close()
                st.success(
                    f"تعداد {updated_count} فرم بر اساس فایل JSON در دیتابیس به‌روزرسانی/جایگزین گردید."
                )
                st.rerun()

    except Exception as e:
        st.error(f"خطا در پردازش فایل JSON: {e}")

st.divider()

# ==========================================
# ۳. نمایش دیتابیس ذخیره‌شده و ماندگار
# ==========================================
st.subheader("🗄️ محتوای فعلی دیتابیس")

conn = get_db_connection()
try:
    db_df = pd.read_sql_query(
        "SELECT form_id, gender, age, work_experience, current_service_experience, education, organizational_post, source_type, updated_at FROM responses",
        conn,
    )
    if not db_df.empty:
        st.write(f"تعداد کل رکوردهای موجود در دیتابیس: {len(db_df)}")
        st.dataframe(db_df)
    else:
        st.info(
            "دیتابیس در حال حاضر خالی است. ابتدا فایل اکسل اولیه را بارگذاری کنید."
        )
except Exception as e:
    st.error(f"خطا در خواندن دیتابیس: {e}")
finally:
    conn.close()
