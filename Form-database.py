import json
import sqlite3
import pandas as pd
import streamlit as st
import os

st.set_page_config(
    page_title="مدیریت دیتابیس (Excel + JSON)", layout="wide"
)

DB_FILE = "questionnaires.db"

# ==========================================
# توابع اصلی دیتابیس و پردازش
# ==========================================

def get_db_connection():
    """ایجاد اتصال به دیتابیس"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """ایجاد یا بازسازی جدول دیتابیس با ساختار درست"""
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("PRAGMA table_info(responses)")
        columns = [col[1] for col in cursor.fetchall()]
    except Exception:
        columns = []

    if not columns or "form_id" not in columns or "source_type" not in columns:
        st.warning("در حال ساخت یا بازسازی ساختار دیتابیس...")
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
        st.success("ساختار دیتابیس آماده شد.")

    conn.close()


def flatten_json_form(form_data):
    """استخراج اطلاعات از فایل JSON استاندارد"""
    form_id = form_data.get("form_id")
    pages = form_data.get("pages", {})

    demographics = {}
    for page_key, page_val in pages.items():
        if isinstance(page_val, dict) and "demographics" in page_val:
            demographics = page_val["demographics"]
            break

    return {
        "form_id": form_id,
        "gender": demographics.get("gender"),
        "age": demographics.get("age"),
        "work_experience": demographics.get("work_experience"),
        "current_service_experience": demographics.get("current_service_experience"),
        "education": demographics.get("education"),
        "organizational_post": demographics.get("organizational_post"),
        "raw_data": json.dumps(form_data, ensure_ascii=False),
    }


def find_column_normalize(df, target_names):
    """
    پیدا کردن نام واقعی ستون در دیتافریم بدون حساسیت به فاصله، نیم‌فاصله یا بزرگ‌کوچکی حروف.
    """
    import re

    norm_targets = [re.sub(r'\s+', '', name).lower() for name in target_names]

    for col in df.columns:
        norm_col = re.sub(r'\s+', '', str(col)).lower()
        if norm_col in norm_targets:
            return col
    return None


# ==========================================
# شروع برنامه
# ==========================================

init_db()

st.title("📊 سیستم مدیریت و به‌روزرسانی پرسشنامه‌ها")

# نمایش وضعیت فعلی دیتابیس در نوار کناری
conn = get_db_connection()
try:
    count_res = conn.execute("SELECT count(*) FROM responses").fetchone()
    current_count = count_res[0]
except Exception:
    current_count = 0
finally:
    conn.close()

st.sidebar.metric("تعداد فرم‌های ثبت شده در دیتابیس", current_count)
if current_count == 0:
    st.sidebar.info("دیتابیس خالی است. لطفا ابتدا فایل اکسل را از بخش ۱ آپلود کنید.")


# ==========================================
# ورودی ۱: بارگذاری اولیه از طریق Excel
# ==========================================
st.subheader("۱. بارگذاری اولیه داده‌ها (فقط Excel)")
st.markdown("""
<div style="background-color:#f0f2f6;padding:10px;border-radius:5px;border-right:5px solid #ff4b4b;margin-bottom:15px;">
<strong>دقت کنید:</strong> ستون‌های اصلی در فایل اکسل شما باید نام‌های مشخصی داشته باشند.<br>
نام ستون اصلی باید چیزی شبیه به <strong>"نام تکمیل کننده"</strong>، <strong>"نام تکمیل‌کننده"</strong> (با نیم‌فاصله) یا <strong>"form_id"</strong> باشد تا سیستم بتواند کدهای G1, G2, ... را شناسایی کند.
</div>
""", unsafe_allow_html=True)

uploaded_excel = st.file_uploader(
    "فایل اکسل اولیه را آپلود کنید", type=["xlsx", "xls"], key="excel_uploader"
)

if uploaded_excel is not None:
    try:
        df_excel = pd.read_excel(uploaded_excel)

        st.write("پیش‌نمایش ۵ سطر اول فایل اکسل:")
        st.dataframe(df_excel.head())

        if st.button("💾 ثبت اولیه داده‌های اکسل در دیتابیس"):
            col_form_id = find_column_normalize(df_excel, ["نام تکمیل کننده", "نام تکمیل‌کننده", "form_id", "id", "code"])
            
            if not col_form_id:
                st.error("❌ خطای اساسی: ستون مربوط به شناسه فرم (مانند 'نام تکمیل کننده' یا 'form_id') در فایل اکسل پیدا نشد. لطفا فایل را بررسی کرده و دوباره آپلود کنید.")
                st.stop()

            df_processed = df_excel.astype(str).replace("nan", None)
            
            col_gender = find_column_normalize(df_excel, ["جنسیت", "gender"])
            col_age = find_column_normalize(df_excel, ["سن", "age"])
            col_work_exp = find_column_normalize(df_excel, ["سابقه کار", "work_experience"])
            col_serv_exp = find_column_normalize(df_excel, ["سابقه در محل خدمت", "current_service_experience"])
            col_edu = find_column_normalize(df_excel, ["مدرک تحصیلات", "تحصیلات", "education"])
            col_post = find_column_normalize(df_excel, ["پست سازمانی", "سمت", "organizational_post"])

            missing_optional = []
            if not col_gender: missing_optional.append("جنسیت")
            if not col_age: missing_optional.append("سن")
            if missing_optional:
                st.warning(f"⚠️ ستون‌های اختیاری مقابل در اکسل پیدا نشدند و خالی رد می‌شوند: {', '.join(missing_optional)}")

            conn = get_db_connection()
            cursor = conn.cursor()
            inserted_count = 0
            error_count = 0

            progress_bar = st.progress(0)
            status_text = st.empty()
            total_rows = len(df_processed)

            for i, row in df_processed.iterrows():
                form_id_raw = row[col_form_id]
                
                if form_id_raw is None or str(form_id_raw).lower() == 'none':
                    form_id = None
                else:
                    form_id = str(form_id_raw).strip()

                if form_id and form_id != "":
                    try:
                        gender = row[col_gender] if col_gender else None
                        age = row[col_age] if col_age else None
                        work_exp = row[col_work_exp] if col_work_exp else None
                        serv_exp = row[col_serv_exp] if col_serv_exp else None
                        edu = row[col_edu] if col_edu else None
                        post = row[col_post] if col_post else None
                        
                        raw_data_json = row.to_json(force_unicode=True)

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
                                raw_data_json,
                            ),
                        )
                        inserted_count += 1
                    except sqlite3.Error:
                        error_count += 1
                
                if (i + 1) % 10 == 0 or (i + 1) == total_rows:
                    progress_bar.progress((i + 1) / total_rows)
                    status_text.text(f"در حال پردازش سطر {i+1} از {total_rows}...")

            conn.commit()
            conn.close()
            
            progress_bar.empty()
            status_text.empty()

            if inserted_count > 0:
                st.success("✅ عملیات با موفقیت انجام شد.")
                st.balloons()
                msg = f"تعداد **{inserted_count}** فرم با موفقیت از فایل اکسل استخراج و در دیتابیس ذخیره/جایگزین شد."
                if error_count > 0:
                    msg += f" (تعداد {error_count} سطر به دلیل خطا ذخیره نشدند)."
                st.info(msg)
                st.rerun()
            else:
                st.error("❌ هیچ داده‌ای در دیتابیس ذخیره نشد. احتمالاً ستون 'نام تکمیل کننده' در تمام سطرهای فایل اکسل شما خالی بوده است.")

    except Exception as e:
        st.error(f"❌ یک خطای غیرمنتظره در پردازش فایل اکسل رخ داد: {e}")

st.divider()

# ==========================================
# ورودی ۲: آپدیت و جایگزینی داده‌ها (JSON)
# ==========================================
st.subheader("۲. آپدیت و جایگزینی داده‌ها (فقط فایل‌های JSON جدید)")
st.caption(
    "از این بخش برای به‌روزرسانی داده‌های موجود یا اضافه کردن فرم‌های جدید از طریق فایل JSON استفاده کنید."
)

uploaded_json = st.file_uploader(
    "فایل JSON پرسشنامه را آپلود کنید", type=["json"], key="json_uploader"
)

if uploaded_json is not None:
    try:
        raw_json_data = json.load(uploaded_json)
        
        forms_list = (
            raw_json_data
            if isinstance(raw_json_data, list)
            else [raw_json_data]
        )

        parsed_records = []
        for f in forms_list:
            if isinstance(f, dict) and "form_id" in f:
                parsed_records.append(flatten_json_form(f))

        if parsed_records:
            preview_json_df = pd.DataFrame(parsed_records).drop(
                columns=["raw_data"]
            )
            st.write(f"پیش‌نمایش {len(parsed_records)} فرم شناسایی شده در JSON:")
            st.dataframe(preview_json_df.head())

            if st.button("🔄 به‌روزرسانی دیتابیس با فایل JSON"):
                conn = get_db_connection()
                cursor = conn.cursor()
                updated_count = 0

                json_prog = st.progress(0)
                total_json = len(parsed_records)

                for i, rec in enumerate(parsed_records):
                    form_id = str(rec["form_id"]).strip() if rec["form_id"] else None
                    
                    if form_id:
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
                                form_id,
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
                    
                    if (i+1) % 5 == 0 or (i+1) == total_json:
                        json_prog.progress((i+1)/total_json)

                conn.commit()
                conn.close()
                json_prog.empty()
                
                if updated_count > 0:
                    st.success(
                        f"✅ تعداد **{updated_count}** فرم بر اساس فایل JSON در دیتابیس به‌روزرسانی یا جایگزین گردید."
                    )
                    st.rerun()
                else:
                    st.warning("هیچ فرم معتبری (دارای form_id) در فایل JSON برای آپدیت پیدا نشد.")

    except Exception as e:
        st.error(f"❌ خطا در پردازش فایل JSON: {e}")

st.divider()

# ==========================================
# ۳. نمایش دیتابیس ذخیره‌شده و ماندگار
# ==========================================
st.subheader("🗄️ محتوای فعلی دیتابیس (نمای کلی)")

conn = get_db_connection()
try:
    db_df = pd.read_sql_query(
        """
        SELECT 
            form_id as "شناسه فرم", 
            gender as "جنسیت", 
            age as "سن", 
            work_experience as "سابقه کار", 
            current_service_experience as "سابقه در محل", 
            education as "تحصیلات", 
            organizational_post as "پست", 
            source_type as "منبع داده", 
            updated_at as "زمان ثبت/تغییر" 
        FROM responses
        ORDER BY updated_at DESC
        """,
        conn,
    )
    if not db_df.empty:
        st.write(f"تعداد کل رکوردهای موجود: **{len(db_df)}**")
        st.dataframe(db_df)
        
        if st.checkbox("آماده‌سازی لینک دانلود کل داده‌ها (به همراه داده خام JSON)"):
            df_full = pd.read_sql_query("SELECT * FROM responses", conn)
            
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_full.to_excel(writer, index=False, sheet_name='Data')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 دانلود کل دیتابیس (Excel)",
                data=processed_data,
                file_name="full_database_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    else:
        st.info(
            "دیتابیس در حال حاضر خالی است. لطفا ابتدا فایل اکسل اولیه را در بخش ۱ بارگذاری کنید."
        )
except Exception as e:
    st.error(f"❌ خطا در خواندن دیتابیس: {e}")
finally:
    conn.close()

# بخش پاکسازی دیتابیس
with st.expander("🛠️ تنظیمات پیشرفته (منطقه خطر)"):
    st.warning("عملیات زیر غیرقابل بازگشت است.")
    if st.button("💣 پاکسازی کامل دیتابیس"):
        conn = get_db_connection()
        conn.execute("DROP TABLE IF EXISTS responses")
        conn.commit()
        conn.close()
        st.success("دیتابیس کاملا پاک شد. فایل دیتابیس در اجرای بعدی دوباره ساخته می‌شود.")
        st.rerun()
