import streamlit as st
import pandas as pd
import io

# تنظیمات اولیه صفحه
st.set_page_config(page_title="ابزار پردازش تردد و آمار", layout="wide")
st.title("📊 اپلیکیشن یکپارچه‌سازی تردد و آمار تفکیکی")
st.markdown("فایل‌های خود را بارگذاری کنید تا داده‌ها پردازش، ادغام و برای خروجی آماده شوند.")

# توابع کمکی
def flatten_columns(df):
    """
    مسطح‌سازی ساختار Multi-Index ستون‌ها
    """
    new_cols = []
    for col in df.columns:
        l1, l2 = str(col[0]).strip(), str(col[1]).strip()
        
        # اگر ستون والد "اطلاعات تردد" بود، فقط زیرستون را برمی‌گردانیم
        if "اطلاعات تردد" in l1:
            new_cols.append(l2)
        # رفع مشکل ستون‌های Unnamed اکسل
        elif "Unnamed" in l1:
            new_cols.append(l2)
        elif "Unnamed" in l2:
            new_cols.append(l1)
        elif l1 == l2:
            new_cols.append(l1)
        else:
            new_cols.append(l1) # در سایر موارد اولویت با سطر اول هدر است
    return new_cols

def deduplicate_columns(cols):
    """
    جلوگیری از تداخل نام ستون‌های تکراری (مثل دو ستون به نام گیت)
    """
    seen = {}
    new_cols = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    return new_cols

def normalize_personnel_code(df):
    """
    یکسان‌سازی نام ستون کد پرسنلی برای اطمینان از Merge صحیح
    """
    for col in df.columns:
        if str(col).replace(" ", "") == "کدپرسنلی":
            df.rename(columns={col: "کدپرسنلی"}, inplace=True)
            return

def to_excel(df):
    """
    تبدیل دیتافریم به فایل اکسل برای دانلود
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

# بخش آپلود فایل‌ها
col1, col2 = st.columns(2)
with col1:
    file_chehre = st.file_uploader("📂 آپلود فایل چهره-هر-روز.xlsx", type=["xlsx", "xls"])
with col2:
    file_amar = st.file_uploader("📂 آپلود فایل آمار-تفکیکی.xlsx", type=["xlsx", "xls"])

if file_chehre and file_amar:
    try:
        with st.spinner('در حال پردازش و ادغام داده‌ها...'):
            # 1. خواندن فایل اول (چهره) با هدرهای دو سطحی (Multi-Index)
            df_chehre = pd.read_excel(file_chehre, header=[0, 1])
            
            # اصلاح و مسطح‌سازی ستون‌ها
            df_chehre.columns = flatten_columns(df_chehre)
            df_chehre.columns = deduplicate_columns(df_chehre.columns)
            
            # حذف ستون واحد سازمانی
            cols_to_drop = [c for c in df_chehre.columns if "واحد سازمانی" in c]
            if cols_to_drop:
                df_chehre.drop(columns=cols_to_drop, inplace=True)
                
            normalize_personnel_code(df_chehre)

            # 2. خواندن فایل دوم (آمار تفکیکی)
            df_amar = pd.read_excel(file_amar)
            normalize_personnel_code(df_amar)
            
            # انتخاب فقط ستون‌های مورد نیاز از فایل دوم برای ادغام
            target_cols_amar = ["کدپرسنلی", "پست", "شغل", "مرکز هزینه ( عنوان تفصیلی )"]
            # بررسی وجود ستون‌ها در فایل دوم برای جلوگیری از کرش
            available_cols = [c for c in target_cols_amar if c in df_amar.columns]
            df_amar_subset = df_amar[available_cols].drop_duplicates(subset=["کدپرسنلی"])

            # 3. ادغام (Merge) داده‌ها بر اساس مقادیر مشترک کد پرسنلی (Inner Join)
            merged_df = pd.merge(df_chehre, df_amar_subset, on="کدپرسنلی", how="inner")
            
            # 4. مرتب‌سازی الفبایی بر اساس شغل (برای گروه‌بندی a, b, c و...)
            if "شغل" in merged_df.columns:
                merged_df.sort_values(by=["شغل"], inplace=True)

        st.success("✅ پردازش و ادغام با موفقیت انجام شد!")

        # نمایش پیش‌نمایش دیتافریم اصلی
        st.subheader("نمای کلی داده‌های ادغام شده")
        st.dataframe(merged_df, use_container_width=True)
        
        # دکمه دانلود کل داده‌ها
        st.download_button(
            label="📥 دانلود فایل اکسل کلی (تمام داده‌ها)",
            data=to_excel(merged_df),
            file_name="Merged_All_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # 5. بخش فیلترینگ و خروجی‌های تفکیک‌شده
        st.subheader("🔍 فیلتر و خروجی اختصاصی بر اساس شغل")
        if "شغل" in merged_df.columns:
            # دریافت لیست شغل‌های موجود (حذف مقادیر خالی)
            job_types = merged_df["شغل"].dropna().unique().tolist()
            
            selected_job = st.selectbox("یک شغل را برای فیلتر و خروجی انتخاب کنید:", ["انتخاب کنید..."] + job_types)
            
            if selected_job != "انتخاب کنید...":
                # اعمال فیلتر و مرتب‌سازی
                filtered_df = merged_df[merged_df["شغل"] == selected_job]
                
                st.write(f"نمایش ردیف‌های مربوط به: **{selected_job}** (تعداد: {len(filtered_df)})")
                st.dataframe(filtered_df, use_container_width=True)
                
                # دکمه دانلود داده‌های فیلتر شده
                st.download_button(
                    label=f"📥 دانلود اکسل فقط برای شغل ({selected_job})",
                    data=to_excel(filtered_df),
                    file_name=f"Filtered_{selected_job}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("ستون 'شغل' در فایل آمار تفکیکی یافت نشد.")

    except Exception as e:
        st.error(f"❌ خطایی در طول پردازش رخ داد. لطفاً ساختار فایل‌ها را بررسی کنید.\n\nجزئیات خطا: {e}")
