import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ابزار پردازش تردد و آمار", layout="wide")
st.title("📊 اپلیکیشن یکپارچه‌سازی تردد و آمار تفکیکی")
st.markdown("فایل‌های خود را بارگذاری کنید تا داده‌ها پردازش، ادغام و برای خروجی آماده شوند.")

# توابع پردازشی هوشمند
def parse_chehre_excel(file_buffer):
    """
    اسکن فایل اکسل تردد برای پیدا کردن اتوماتیک سطر هدرها و تمیز کردن داده‌ها
    """
    # خواندن خام فایل بدون هدر
    df_raw = pd.read_excel(file_buffer, header=None)
    
    start_idx = None
    # پیدا کردن سطری که شامل کلمات کلیدی هدر ماست
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(x).replace(" ", "") for x in row.dropna()]).lower()
        if "کدپرسنلی" in row_str or "نامونامخانوادگی" in row_str or "اطلاعاتتردد" in row_str:
            start_idx = idx
            break
            
    if start_idx is None:
        raise ValueError("سطر حاوی 'کد پرسنلی' در فایل چهره یافت نشد. لطفاً ساختار فایل را بررسی کنید.")
    
    # استخراج دو سطر مربوط به هدر (برای مدیریت Multi-Index به صورت دستی)
    row1 = df_raw.iloc[start_idx].fillna('').astype(str).str.strip()
    row2 = df_raw.iloc[start_idx + 1].fillna('').astype(str).str.strip()
    
    # ترکیب هوشمند زیرستون‌ها با ستون‌های اصلی
    headers = []
    for c1, c2 in zip(row1, row2):
        if c2 and c2 != 'nan':
            headers.append(c2)
        else:
            headers.append(c1 if c1 != 'nan' else '')
            
    # ساخت دیتافریم اصلی از سطرهای زیر هدر
    df_data = df_raw.iloc[start_idx + 2:].copy()
    df_data.columns = headers
    
    # حذف ستون‌های کاملاً خالی
    df_data = df_data.loc[:, df_data.columns != '']
    
    # جلوگیری از تداخل نام ستون‌های تکراری (مثل دو ستون گیت)
    seen = {}
    new_cols = []
    for c in df_data.columns:
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    df_data.columns = new_cols
    
    # استانداردسازی نام ستون کد پرسنلی برای اطمینان
    df_data.rename(columns=lambda x: "کدپرسنلی" if "کد" in str(x).replace(" ", "") and "پرسنلی" in str(x).replace(" ", "") else x, inplace=True)
    
    return df_data

def normalize_amar_excel(df):
    """استانداردسازی نام کد پرسنلی در فایل دوم"""
    for col in df.columns:
        if "کد" in str(col).replace(" ", "") and "پرسنلی" in str(col).replace(" ", ""):
            df.rename(columns={col: "کدپرسنلی"}, inplace=True)
            return

def to_excel(df):
    """تبدیل دیتافریم به فایل خروجی"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

# رابط کاربری آپلود فایل‌ها
col1, col2 = st.columns(2)
with col1:
    file_chehre = st.file_uploader("📂 آپلود فایل چهره-هر-روز.xlsx", type=["xlsx", "xls"])
with col2:
    file_amar = st.file_uploader("📂 آپلود فایل آمار-تفکیکی.xlsx", type=["xlsx", "xls"])

if file_chehre and file_amar:
    try:
        with st.spinner('در حال پردازش، یافتن هدرها و ادغام داده‌ها...'):
            
            # ۱. خواندن و استخراج هوشمند فایل تردد
            df_chehre = parse_chehre_excel(file_chehre)
            
            # حذف ستون واحد سازمانی در صورت وجود
            cols_to_drop = [c for c in df_chehre.columns if "واحد سازمانی" in str(c)]
            if cols_to_drop:
                df_chehre.drop(columns=cols_to_drop, inplace=True)
                
            # استاندارد کردن فرمت اعداد (حذف صفرهای اعشاری شناور در کد پرسنلی)
            df_chehre['کدپرسنلی'] = df_chehre['کدپرسنلی'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

            # ۲. خواندن فایل آمار تفکیکی
            df_amar = pd.read_excel(file_amar)
            normalize_amar_excel(df_amar)
            df_amar['کدپرسنلی'] = df_amar['کدپرسنلی'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            # ۳. آماده‌سازی ستون‌های مورد نیاز برای ادغام
            target_cols_amar = ["کدپرسنلی", "پست", "شغل", "مرکز هزینه ( عنوان تفصیلی )"]
            available_cols = [c for c in target_cols_amar if c in df_amar.columns]
            df_amar_subset = df_amar[available_cols].drop_duplicates(subset=["کدپرسنلی"])

            # ۴. ادغام (Merge) داده‌ها فقط برای پرسنل مشترک
            merged_df = pd.merge(df_chehre, df_amar_subset, on="کدپرسنلی", how="inner")
            
            # ۵. مرتب‌سازی الفبایی بر اساس شغل
            if "شغل" in merged_df.columns:
                merged_df.sort_values(by=["شغل"], inplace=True)

        st.success("✅ هدرها با موفقیت شناسایی و پردازش و ادغام انجام شد!")

        # نمایش داده‌ها
        st.subheader("نمای کلی داده‌های ادغام شده")
        st.dataframe(merged_df, use_container_width=True)
        
        st.download_button(
            label="📥 دانلود فایل اکسل کلی (تمام داده‌ها)",
            data=to_excel(merged_df),
            file_name="Merged_All_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # بخش فیلترینگ بر اساس شغل
        st.subheader("🔍 فیلتر و خروجی اختصاصی بر اساس شغل")
        if "شغل" in merged_df.columns:
            job_types = merged_df["شغل"].dropna().unique().tolist()
            selected_job = st.selectbox("یک شغل را برای فیلتر و خروجی انتخاب کنید:", ["انتخاب کنید..."] + job_types)
            
            if selected_job != "انتخاب کنید...":
                filtered_df = merged_df[merged_df["شغل"] == selected_job]
                st.write(f"نمایش ردیف‌های مربوط به: **{selected_job}** (تعداد: {len(filtered_df)})")
                st.dataframe(filtered_df, use_container_width=True)
                
                st.download_button(
                    label=f"📥 دانلود اکسل فقط برای شغل ({selected_job})",
                    data=to_excel(filtered_df),
                    file_name=f"Filtered_{selected_job}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("ستون 'شغل' در فایل آمار تفکیکی یافت نشد.")

    except Exception as e:
        st.error(f"❌ خطایی در طول پردازش رخ داد. لطفاً جزئیات را بررسی کنید:\n\n{e}")
