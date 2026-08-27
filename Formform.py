import json
import sqlite3
import pandas as pd
import streamlit as st

st.set_page_config(page_title="مدیریت دیتابیس پرسشنامه‌ها", layout="wide")

DB_FILE = "questionnaires.db"


def get_db_connection():
    return sqlite3.connect(DB_FILE)


def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS responses (
            form_id TEXT PRIMARY KEY,
            data_json TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """
    )
    conn.commit()
    conn.close()


init_db()

st.title("📊 مدیریت دیتابیس پرسشنامه‌ها")

# بخش آپلود فایل (پشتیبانی همزمان از JSON و Excel)
uploaded_file = st.file_uploader(
    "فایل خود را آپلود کنید (JSON یا Excel)", type=["json", "xlsx", "xls"]
)

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1]

    if file_type == "json":
        try:
            json_data = json.load(uploaded_file)
            st.success("فایل JSON با موفقیت خوانده شد.")

            # اگر فایل یک لیست از فرم‌ها باشد یا یک تک فرم
            forms_list = (
                json_data if isinstance(json_data, list) else [json_data]
            )

            if st.button("ذخیره/به‌روزرسانی فایل JSON در دیتابیس"):
                conn = get_db_connection()
                cursor = conn.cursor()
                count = 0

                for item in forms_list:
                    form_id = item.get("form_id")
                    if form_id:
                        cursor.execute(
                            """
                            INSERT INTO responses (form_id, data_json)
                            VALUES (?, ?)
                            ON CONFLICT(form_id) DO UPDATE SET
                                data_json = excluded.data_json,
                                updated_at = CURRENT_TIMESTAMP
                        """,
                            (str(form_id).strip(), json.dumps(item, ensure_ascii=False)),
                        )
                        count += 1

                conn.commit()
                conn.close()
                st.success(f"تعداد {count} فرم با موفقیت در دیتابیس ثبت/جایگزین شد.")

        except Exception as e:
            st.error(f"خطا در پردازش فایل JSON: {e}")

    elif file_type in ["xlsx", "xls"]:
        try:
            df = pd.read_excel(uploaded_file)

            # رفع خطای PyArrow با تبدیل تمام ستون‌ها به رشته یا فرمت استاندارد
            df = df.astype(str).replace("nan", None)

            st.write("پیش‌نمایش فایل اکسل:")
            st.dataframe(df)

            if st.button("ذخیره/به‌روزرسانی فایل اکسل در دیتابیس"):
                conn = get_db_connection()
                cursor = conn.cursor()
                count = 0

                for _, row in df.iterrows():
                    # نام ستون مربوط به شناسه فرم را در صورت لزوم تغییر دهید
                    form_id = str(row.get("نام تکمیل کننده", row.get("form_id", ""))).strip()
                    if form_id and form_id != "None":
                        data_json = row.to_json(force_unicode=True)
                        cursor.execute(
                            """
                            INSERT INTO responses (form_id, data_json)
                            VALUES (?, ?)
                            ON CONFLICT(form_id) DO UPDATE SET
                                data_json = excluded.data_json,
                                updated_at = CURRENT_TIMESTAMP
                        """,
                            (form_id, data_json),
                        )
                        count += 1

                conn.commit()
                conn.close()
                st.success(f"تعداد {count} ردیف با موفقیت در دیتابیس ثبت/جایگزین شد.")

        except Exception as e:
            st.error(f"خطا در پردازش فایل اکسل: {e}")

st.divider()

# نمایش دیتابیس ماندگار
st.subheader("📁 اطلاعات ثبت‌شده در دیتابیس (ثابت و ماندگار)")

conn = get_db_connection()
db_df = pd.read_sql_query("SELECT * FROM responses", conn)
conn.close()

if not db_df.empty:
    st.write(f"تعداد کل رکوردهای موجود: {len(db_df)}")
    st.dataframe(db_df)
else:
    st.info("هیچ داده‌ای در دیتابیس یافت نشد.")
