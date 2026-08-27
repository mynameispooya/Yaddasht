import json
import sqlite3
import re
from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="مدیریت جامع دیتابیس پرسشنامه‌ها", layout="wide"
)

DB_FILE = "questionnaires.db"


# ==========================================
# توابع پایه دیتابیس
# ==========================================

def get_db_connection():
    """ارتباط با دیتابیس SQLite"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def build_column_names_from_excel(df_raw):
    """
    استخراج نام‌های استاندارد ستون‌ها بر اساس ساختار ۲ سطری فایل اکسل:
    - ستون‌های اولیه (نام، جنسیت، سن و ...)
    - ابعاد پرسشنامه (شغل شما، مسئول مستقیم، همکار، ارتقا، حقوق و مزایا، شرایط کار) + شماره سوال یا 'میانگین'
    """
    cols = []
    current_category = ""

    for col_idx in range(df_raw.shape[1]):
        row0_val = str(df_raw.iloc[0, col_idx]).strip()
        row1_val = str(df_raw.iloc[1, col_idx]).strip()

        # اگر ارزش سطر اول موجود باشد، دسته‌بندی جدید یا ستون ثابت است
        if row0_val != "nan" and row0_val != "":
            if row1_val == "nan" or row1_val == "":
                cols.append(row0_val)
                continue
            else:
                current_category = row0_val

        # اگر در سطر دوم شماره سوال یا کلمه 'میانگین' باشد
        if row1_val != "nan" and row1_val != "":
            try:
                q_num = int(float(row1_val))
                col_name = f"{current_category}_Q{q_num}"
            except ValueError:
                col_name = f"{current_category}_{row1_val}"
            cols.append(col_name)
        else:
            cols.append(f"Unmapped_Col_{col_idx}")

    return cols


def init_db(columns_list=None):
    """ایجاد یا بازسازی جدول دیتابیس بر اساس تمام ستون‌های اکسل"""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='responses'"
    )
    table_exists = cursor.fetchone()

    if not table_exists and columns_list:
        # ساخت دستور SQL
        col_defs = ["form_id TEXT PRIMARY KEY"]
        for col in columns_list:
            # پاکسازی نام ستون‌ها برای SQL
            safe_col = (
                col.replace(" ", "_").replace("‌", "_").replace("-", "_")
            )
            if safe_col != "نام_تکمیل_کننده_فرم" and safe_col != "form_id":
                col_defs.append(f'"{col}" TEXT')

        col_defs.append("raw_data TEXT")
        col_defs.append("source_type TEXT")
        col_defs.append("updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

        create_sql = f"CREATE TABLE responses ({', '.join(col_defs)})"
        cursor.execute(create_sql)
        conn.commit()

    conn.close()


def clean_str(val):
    if pd.isna(val) or val is None or str(val).strip().lower() in ["nan", "none"]:
        return None
    return str(val).strip()


# ==========================================
# شروع رابط کاربری
# ==========================================

st.title("📊 سیستم مدیریت و به‌روزرسانی دیتابیس پرسشنامه‌ها")

# بررسی تعداد داده‌های فعلی
conn = get_db_connection()
cursor = conn.cursor()
cursor.execute(
    "SELECT name FROM sqlite_master WHERE type='table' AND name='responses'"
)
table_exists = cursor.fetchone()

current_count = 0
if table_exists:
    count_res = conn.execute("SELECT count(*) FROM responses").fetchone()
    current_count = count_res[0]
conn.close()

st.sidebar.metric("تعداد کل فرم‌ها در دیتابیس", current_count)

# ==========================================
# بخش ۱: آپلود اولیه اکسل
# ==========================================
st.subheader("۱. بارگذاری اولیه الگوی اکسل و داده‌ها")
st.markdown(
    """
<div style="background-color:#f0f2f6;padding:12px;border-radius:6px;border-right:5px solid #0066cc;margin-bottom:15px;">
<strong>راهنما:</strong> فایل اکسل اولیه ساختار جدول دیتابیس را شکل می‌دهد. شناسه اصلی فرم‌ها از ستون <strong>"نام تکمیل کننده فرم"</strong> (مانند G1, G2, ...) خوانده می‌شود.
</div>
""",
    unsafe_allow_html=True,
)

uploaded_excel = st.file_uploader(
    "فایل اکسل اولیه را آپلود کنید", type=["xlsx", "xls"], key="excel_uploader"
)

if uploaded_excel is not None:
    try:
        df_raw = pd.read_excel(uploaded_excel, header=None)
        cols = build_column_names_from_excel(df_raw)

        # دیتای واقعی از سطر ۲ به بعد آغاز می‌شود
        df_data = df_raw.iloc[2:].copy()
        df_data.columns = cols

        # پیدا کردن ستون کلیدی form_id
        form_id_col = None
        for c in df_data.columns:
            if "تکمیل کننده" in c or c.lower() == "form_id":
                form_id_col = c
                break

        if not form_id_col:
            form_id_col = df_data.columns[0]

        st.write(f"نمایش پیش‌نمایش ({len(df_data)} سطر و {len(cols)} ستون):")
        st.dataframe(df_data.head())

        if st.button("💾 ایجاد ساختار دیتابیس و ثبت داده‌های اکسل"):
            init_db(cols)

            conn = get_db_connection()
            cursor = conn.cursor()

            # استخراج اسامی ستون‌های فعلی دیتابیس
            cursor.execute("PRAGMA table_info(responses)")
            db_cols = [row[1] for row in cursor.fetchall()]

            inserted_count = 0
            for idx, row in df_data.iterrows():
                fid = clean_str(row[form_id_col])
                if not fid:
                    continue

                row_dict = {}
                for col in cols:
                    if col in db_cols:
                        row_dict[col] = clean_str(row[col])

                row_dict["raw_data"] = json.dumps(
                    row.to_dict(), ensure_ascii=False
                )
                row_dict["source_type"] = "EXCEL"

                # درج/آپدیت متناظر در SQL
                fields = ["form_id"] + list(row_dict.keys())
                placeholders = ["?"] * len(fields)
                values = [fid] + list(row_dict.values())

                update_clause = ", ".join(
                    [
                        f'"{k}"=excluded."{k}"'
                        for k in row_dict.keys()
                    ]
                )

                sql = f"""
                    INSERT INTO responses ({", ".join([f'"{f}"' for f in fields])})
                    VALUES ({", ".join(placeholders)})
                    ON CONFLICT(form_id) DO UPDATE SET
                    {update_clause},
                    updated_at = CURRENT_TIMESTAMP
                """
                cursor.execute(sql, values)
                inserted_count += 1

            conn.commit()
            conn.close()

            st.success(
                f"✅ تعداد {inserted_count} رکورد با تمام ستون‌های استاندارد در دیتابیس ذخیره شد."
            )
            st.rerun()

    except Exception as e:
        st.error(f"❌ خطا در پردازش فایل اکسل: {e}")

st.divider()

# ==========================================
# بخش ۲: آپدیت هوشمند از طریق JSON
# ==========================================
st.subheader("۲. به‌روزرسانی دیتابیس با فایل JSON")

uploaded_json = st.file_uploader(
    "فایل JSON پرسشنامه را آپلود کنید", type=["json"], key="json_uploader"
)

if uploaded_json is not None:
    if not table_exists:
        st.error(
            "❌ ابتدا باید در بخش ۱ فایل اکسل را آپلود کنید تا دیتابیس ساخته شود."
        )
    else:
        try:
            raw_json_data = json.load(uploaded_json)
            forms_list = (
                raw_json_data
                if isinstance(raw_json_data, list)
                else [raw_json_data]
            )

            if st.button("🔄 به‌روزرسانی داده‌ها بر اساس JSON"):
                conn = get_db_connection()
                cursor = conn.cursor()

                cursor.execute("PRAGMA table_info(responses)")
                db_cols = [row[1] for row in cursor.fetchall()]

                updated_cnt = 0
                for item in forms_list:
                    fid = item.get("form_id") or item.get(
                        "نام تکمیل کننده فرم"
                    )
                    if not fid:
                        continue

                    fid = str(fid).strip()
                    update_payload = {}

                    # نگاشت داده‌های JSON به ستون‌های دیتابیس
                    for col in db_cols:
                        if col in [
                            "form_id",
                            "raw_data",
                            "source_type",
                            "updated_at",
                        ]:
                            continue

                        # جستجو در کل ساختار JSON
                        if col in item:
                            update_payload[col] = clean_str(item[col])
                        elif "pages" in item:
                            # جستجو در صفحات JSON
                            for page_k, page_v in item["pages"].items():
                                if (
                                    isinstance(page_v, dict)
                                    and col in page_v
                                ):
                                    update_payload[col] = clean_str(
                                        page_v[col]
                                    )

                    update_payload["raw_data"] = json.dumps(
                        item, ensure_ascii=False
                    )
                    update_payload["source_type"] = "JSON"

                    fields = ["form_id"] + list(update_payload.keys())
                    placeholders = ["?"] * len(fields)
                    values = [fid] + list(update_payload.values())

                    update_clause = ", ".join(
                        [
                            f'"{k}"=excluded."{k}"'
                            for k in update_payload.keys()
                        ]
                    )

                    sql = f"""
                        INSERT INTO responses ({", ".join([f'"{f}"' for f in fields])})
                        VALUES ({", ".join(placeholders)})
                        ON CONFLICT(form_id) DO UPDATE SET
                        {update_clause},
                        updated_at = CURRENT_TIMESTAMP
                    """
                    cursor.execute(sql, values)
                    updated_cnt += 1

                conn.commit()
                conn.close()

                st.success(
                    f"✅ تعداد {updated_cnt} پرسشنامه از طریق JSON به‌روزرسانی شد."
                )
                st.rerun()

        except Exception as e:
            st.error(f"❌ خطا در پردازش فایل JSON: {e}")

st.divider()

# ==========================================
# بخش ۳: مشاهده، مدیریت و حذف ردیف‌ها
# ==========================================
st.subheader("🗄️ مدیریت و ویرایش ردیف‌های دیتابیس")

if table_exists and current_count > 0:
    conn = get_db_connection()
    df_db = pd.read_sql_query("SELECT * FROM responses", conn)
    conn.close()

    st.write(f"تعداد ردیف‌های ثبت‌شده: **{len(df_db)}**")

    # انتخاب ردیف‌ها برای حذف
    st.markdown("##### 🗑️ حذف ردیف‌ها از دیتابیس")

    selected_ids = st.multiselect(
        "فرم‌هایی که قصد حذف آن‌ها را دارید انتخاب کنید (بر اساس form_id/کد G):",
        options=df_db["form_id"].tolist(),
    )

    if selected_ids:
        st.warning(
            f"شما تعداد {len(selected_ids)} ردیف را برای حذف انتخاب کرده‌اید."
        )
        if st.button("❌ تایید و حذف ردیف‌های انتخاب شده"):
            conn = get_db_connection()
            cursor = conn.cursor()

            placeholders = ", ".join(["?"] * len(selected_ids))
            cursor.execute(
                f"DELETE FROM responses WHERE form_id IN ({placeholders})",
                selected_ids,
            )

            conn.commit()
            conn.close()

            st.success("ردیف‌های مورد نظر با موفقیت حذف شدند.")
            st.rerun()

    # نمایش دیتابیس
    st.markdown("##### 📋 جدول کامل اطلاعات:")
    st.dataframe(df_db)

    # خروجی گرفتن از اکسل
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_db.to_excel(writer, index=False, sheet_name="Questionnaires")

    st.download_button(
        label="📥 دانلود کامل دیتابیس به صورت اکسل",
        data=output.getvalue(),
        file_name="questionnaires_database.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("دیتابیس در حال حاضر خالی است.")
