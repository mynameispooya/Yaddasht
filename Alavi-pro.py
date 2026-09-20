import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ابزار پردازش تردد و آمار", layout="wide")
st.title("📊 اپلیکیشن یکپارچه‌سازی تردد و آمار تفکیکی")
st.markdown("فایل‌های خود را بارگذاری کنید تا داده‌ها پردازش، ادغام و برای خروجی آماده شوند.")

# --- توابع پردازشی ---
def parse_chehre_excel(file_buffer):
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
    df_data = df_data.loc[:, df_data.columns != '']
    
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
    for col in df.columns:
        if "کد" in str(col).replace(" ", "") and "پرسنلی" in str(col).replace(" ", ""):
            df.rename(columns={col: "کدپرسنلی"}, inplace=True)
            return

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

def is_empty(val):
    if pd.isna(val): return True
    if str(val).strip().lower() in ['nan', 'none', '', 'null']: return True
    return False

# --- رابط کاربری آپلود فایل‌ها ---
col1, col2 = st.columns(2)
with col1:
    file_chehre = st.file_uploader("📂 آپلود فایل چهره-هر-روز.xlsx", type=["xlsx", "xls"])
with col2:
    file_amar = st.file_uploader("📂 آپلود فایل آمار-تفکیکی.xlsx", type=["xlsx", "xls"])

if file_chehre and file_amar:
    try:
        with st.spinner('در حال پردازش، یافتن هدرها و ادغام داده‌ها...'):
            df_chehre = parse_chehre_excel(file_chehre)
            
            cols_to_drop = [c for c in df_chehre.columns if "واحد سازمانی" in str(c)]
            if cols_to_drop:
                df_chehre.drop(columns=cols_to_drop, inplace=True)
                
            df_chehre['کدپرسنلی'] = df_chehre['کدپرسنلی'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

            df_amar = pd.read_excel(file_amar)
            normalize_amar_excel(df_amar)
            df_amar['کدپرسنلی'] = df_amar['کدپرسنلی'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            target_cols_amar = ["کدپرسنلی", "پست", "شغل", "مرکز هزینه ( عنوان تفصیلی )"]
            available_cols = [c for c in target_cols_amar if c in df_amar.columns]
            df_amar_subset = df_amar[available_cols].drop_duplicates(subset=["کدپرسنلی"])

            merged_df = pd.merge(df_chehre, df_amar_subset, on="کدپرسنلی", how="inner")

        st.success("✅ داده‌ها با موفقیت ادغام شدند!")
        st.divider()

        # ==========================================
        # مرحله ۱: فیلتر اولیه بر اساس ورود و خروج
        # ==========================================
        st.subheader("⚙️ مرحله ۱: فیلتر وضعیت تردد (ورود و خروج)")
        
        if "ورود" in merged_df.columns and "خروج" in merged_df.columns:
            filter_option = st.radio(
                "وضعیت تردد را انتخاب کنید:",
                [
                    "نمایش همه (بدون فیلتر)",
                    "ورود و خروج هر دو ثبت شده باشند (دارای مقدار)",
                    "فقط ورود ثبت شده باشد (خروج خالی باشد)",
                    "فقط خروج ثبت شده باشد (ورود خالی باشد)",
                    "ورود و خروج هر دو خالی باشند"
                ],
                horizontal=False
            )

            has_vorood = ~merged_df['ورود'].apply(is_empty)
            has_khorooj = ~merged_df['خروج'].apply(is_empty)

            if filter_option == "ورود و خروج هر دو ثبت شده باشند (دارای مقدار)":
                step1_df = merged_df[has_vorood & has_khorooj].copy()
            elif filter_option == "فقط ورود ثبت شده باشد (خروج خالی باشد)":
                step1_df = merged_df[has_vorood & ~has_khorooj].copy()
            elif filter_option == "فقط خروج ثبت شده باشد (ورود خالی باشد)":
                step1_df = merged_df[~has_vorood & has_khorooj].copy()
            elif filter_option == "ورود و خروج هر دو خالی باشند":
                step1_df = merged_df[~has_vorood & ~has_khorooj].copy()
            else:
                step1_df = merged_df.copy()
        else:
            st.warning("ستون‌های 'ورود' و 'خروج' یافت نشدند.")
            step1_df = merged_df.copy()
            
        st.write(f"📊 **تعداد ردیف‌ها پس از مرحله ۱:** {len(step1_df)}")

        st.divider()

        # ==========================================
        # مرحله ۲: فیلتر گیت تردد
        # ==========================================
        st.subheader("🚪 مرحله ۲: فیلتر گیت تردد")
        st.markdown("گیت‌های مورد نظر را انتخاب کنید. (اگر موردی انتخاب نشود، همه گیت‌ها لحاظ می‌شوند)")
        
        target_gate_col = "گیت_1"
        if target_gate_col in step1_df.columns:
            unique_gates = step1_df[target_gate_col].dropna().unique().tolist()
            
            selected_gates = st.multiselect(
                f"انتخاب از ستون '{target_gate_col}':",
                options=unique_gates
            )
            
            if selected_gates:
                step2_df = step1_df[step1_df[target_gate_col].isin(selected_gates)].copy()
            else:
                # اگر کاربر چیزی انتخاب نکرد، کل داده‌های مرحله قبل را به عنوان مرجع این مرحله در نظر می‌گیریم
                step2_df = step1_df.copy()
                
            st.write(f"📊 **تعداد ردیف‌های مرجع جدید پس از مرحله ۲:** {len(step2_df)}")
            
            with st.expander("👁️ مشاهده جدول مرجع جدید (کلیک کنید)"):
                st.dataframe(step2_df, use_container_width=True)
                st.download_button(
                    label="📥 دانلود فایل اکسل مرجع (خروجی مرحله ۱ و ۲)",
                    data=to_excel(step2_df),
                    file_name="Reference_Data_Step2.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning(f"ستون '{target_gate_col}' یافت نشد.")
            step2_df = step1_df.copy()

        st.divider()

        # ==========================================
        # مرحله ۳: فیلترینگ پویا و چندگانه (مبتنی بر مرحله ۲)
        # ==========================================
        st.subheader("🔍 مرحله ۳: فیلتر و خروجی نهایی (پویا و چندگانه)")
        st.markdown("فیلترهای این بخش دقیقاً روی **جدول مرجع** که در مراحل بالا تنظیم کردید اعمال می‌شوند.")
        
        all_columns = step2_df.columns.tolist()
        
        selected_column = st.selectbox(
            "ستون مورد نظر برای فیلتر نهایی را انتخاب کنید:", 
            ["انتخاب کنید..."] + all_columns
        )
        
        if selected_column != "انتخاب کنید...":
            unique_values = step2_df[selected_column].dropna().unique().tolist()
            
            selected_values = st.multiselect(
                f"گزینه‌های مورد نظر از ستون '{selected_column}':", 
                options=unique_values
            )
            
            if selected_values:
                final_filtered_df = step2_df[step2_df[selected_column].isin(selected_values)].copy()
                final_filtered_df = final_filtered_df.sort_values(by=[selected_column])
                
                st.success(f"تعداد خروجی نهایی: {len(final_filtered_df)} ردیف")
                st.dataframe(final_filtered_df, use_container_width=True)
                
                file_suffix = "_".join([str(v).replace(" ", "") for v in selected_values])[:30] 
                
                st.download_button(
                    label=f"📥 دانلود اکسل فیلتر شده نهایی",
                    data=to_excel(final_filtered_df),
                    file_name=f"Final_{file_suffix}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ خطایی در طول پردازش رخ داد. لطفاً جزئیات را بررسی کنید:\n\n{e}")
