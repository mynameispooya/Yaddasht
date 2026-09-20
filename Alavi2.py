import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ابزار پردازش تردد و آمار", layout="wide")
st.title("📊 اپلیکیشن یکپارچه‌سازی تردد و آمار تفکیکی")
st.markdown("فایل‌های خود را بارگذاری کنید تا داده‌ها پردازش، ادغام و برای خروجی آماده شوند.")

# توابع پردازشی هوشمند
def parse_chehre_excel(file_buffer):
    """اسکن فایل اکسل تردد برای پیدا کردن اتوماتیک سطر هدرها و تمیز کردن داده‌ها"""
    df_raw = pd.read_excel(file_buffer, header=None)
    start_idx = None
    
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(x).replace(" ", "") for x in row.dropna()]).lower()
        if "کدپرسنلی" in row_str or "نامونامخانوادگی" in row_str or "اطلاعاتتردد" in row_str:
            start_idx = idx
            break
            
    if start_idx is None:
        raise ValueError("سطر حاوی 'کد پرسنلی' در فایل چهره یافت نشد. لطفاً ساختار فایل را بررسی کنید.")
    
    row1 = df_raw.iloc[start_idx].fillna('').astype(str).str.strip()
    row2 = df_raw.iloc[start_idx + 1].fillna('').astype(str).str.strip()
    
    headers = []
    for c1, c2 in zip(row1, row2):
        if c2 and c2 != 'nan':
            headers.append(c2)
        else:
            headers.append(c1 if c1 != 'nan' else '')
            
    df_data = df_raw.iloc[start_idx + 2:].copy()
    df_data.columns = headers
    
    # حذف ستون‌های کاملاً خالی
    df_data = df_data.loc[:, df_data.columns != '']
    
    # جلوگیری از تداخل نام ستون‌های تکراری
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
            
            cols_to_drop = [c for c in df_chehre.columns if "واحد سازمانی" in str(c)]
            if cols_to_drop:
                df_chehre.drop(columns=cols_to_drop, inplace=True)
                
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

        st.success("✅ هدرها با موفقیت شناسایی و پردازش و ادغام انجام شد!")

        st.subheader("نمای کلی داده‌های ادغام شده")
        st.dataframe(merged_df, use_container_width=True)
        
        st.download_button(
            label="📥 دانلود فایل اکسل کلی (تمام داده‌ها)",
            data=to_excel(merged_df),
            file_name="Merged_All_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # ۵. بخش فیلترینگ پویا و چندگانه
        st.subheader("🔍 فیلتر و خروجی اختصاصی (پویا و چندگانه)")
        
        # دریافت لیست تمام ستون‌های موجود برای فیلتر کردن
        all_columns = merged_df.columns.tolist()
        
        # انتخاب ستون توسط کاربر
        selected_column = st.selectbox(
            "ابتدا ستون مورد نظر برای فیلتر را انتخاب کنید (مثلاً مرکز هزینه یا شغل):", 
            ["انتخاب کنید..."] + all_columns
        )
        
        if selected_column != "انتخاب کنید...":
            # استخراج مقادیر یکتای آن ستون (حذف مقادیر خالی)
            unique_values = merged_df[selected_column].dropna().unique().tolist()
            
            # استفاده از multiselect برای انتخاب چند آیتم همزمان
            selected_values = st.multiselect(
                f"گزینه‌های مورد نظر از ستون '{selected_column}' را انتخاب کنید (می‌توانید چند مورد انتخاب کنید):", 
                options=unique_values
            )
            
            if selected_values:
                # اعمال فیلتر بر اساس لیست مقادیر انتخاب شده
                filtered_df = merged_df[merged_df[selected_column].isin(selected_values)]
                
                # اگر فیلتر روی شغل یا مرکز هزینه بود، مرتب‌سازی هم انجام دهیم تا لیست تمیزتر باشد
                filtered_df = filtered_df.sort_values(by=[selected_column])
                
                st.write(f"نمایش ردیف‌های مربوط به: **{', '.join(str(v) for v in selected_values)}** (تعداد: {len(filtered_df)} ردیف)")
                st.dataframe(filtered_df, use_container_width=True)
                
                # ایجاد نام مناسب برای فایل دانلودی
                file_suffix = "_".join([str(v).replace(" ", "") for v in selected_values])[:30] # محدود کردن طول نام فایل
                
                st.download_button(
                    label=f"📥 دانلود اکسل فیلتر شده",
                    data=to_excel(filtered_df),
                    file_name=f"Filtered_{file_suffix}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ خطایی در طول پردازش رخ داد. لطفاً جزئیات را بررسی کنید:\n\n{e}")
