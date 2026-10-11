import streamlit as st
import pandas as pd
import io

# تنظیمات صفحه
st.set_page_config(page_title="ویرایشگر هوشمند اکسل", layout="wide")
st.title("فیلتر و ویرایش داده‌های اکسل")
st.markdown("فایل اکسل خود (مثلاً **سیمرغ.xlsx**) را آپلود کنید، ستون‌ها را انتخاب کرده و مقادیر را مستقیماً در جدول ویرایش کنید.")

# آپلود فایل
uploaded_file = st.file_uploader("فایل اکسل خود را اینجا آپلود کنید (xlsx, xls)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # برای جلوگیری از ریست شدن داده‌ها در هر بار کلیک، دیتای اصلی را در Session State ذخیره می‌کنیم
    if 'main_df' not in st.session_state or st.session_state.get('uploaded_filename') != uploaded_file.name:
        st.session_state.main_df = pd.read_excel(uploaded_file)
        st.session_state.uploaded_filename = uploaded_file.name

    df = st.session_state.main_df
    
    st.divider()
    
    # انتخاب دو ستون
    st.subheader("۱. انتخاب ستون‌ها")
    col1, col2 = st.columns(2)
    
    with col1:
        # ستون اول برای فیلتر کردن
        filter_col = st.selectbox("ستون اول (مبنای فیلتر - مثلاً مقطع تحصیلی) را انتخاب کنید:", df.columns)
    
    with col2:
        # ستون دوم برای نمایش متناظر
        target_col = st.selectbox("ستون دوم (نمایش متناظر - مثلاً رشته تحصیلی) را انتخاب کنید:", df.columns)

    if filter_col and target_col:
        st.divider()
        st.subheader("۲. فیلتر کردن داده‌ها")
        
        # استخراج مقادیر یکتا از ستون اول (حذف مقادیر خالی)
        unique_values = df[filter_col].dropna().unique()
        
        # انتخاب یک مقدار برای فیلتر
        selected_val = st.selectbox(f"یک مقدار یکتا از ستون '{filter_col}' را انتخاب کنید:", unique_values)
        
        if selected_val:
            # فیلتر کردن دیتافریم بر اساس مقدار انتخاب شده
            # فقط ستون‌های انتخاب شده را برای ویرایش نمایش می‌دهیم تا محیط شلوغ نشود
            filtered_df = df[df[filter_col] == selected_val][[filter_col, target_col]]
            
            st.divider()
            st.subheader("۳. ویرایش مقادیر")
            st.info("💡 روی هر سلولی که می‌خواهید کلیک کنید و مقدار آن را تغییر دهید.")
            
            # نمایش جدول قابل ویرایش (تغییرات در متغیر edited_df ذخیره می‌شود)
            edited_df = st.data_editor(filtered_df, use_container_width=True)
            
            # دکمه برای اعمال تغییرات در دیتای اصلی
            if st.button("ثبت تغییرات جدول بالا در فایل اصلی", type="primary"):
                # تابع update بر اساس ایندکس ردیف‌ها، مقادیر تغییر یافته را روی دیتای اصلی اعمال می‌کند
                st.session_state.main_df.update(edited_df)
                st.success("✅ تغییرات با موفقیت در حافظه ذخیره شد! می‌توانید فیلتر را عوض کنید یا فایل نهایی را دانلود کنید.")
                st.balloons()
            
            st.divider()
            st.subheader("۴. دریافت خروجی")
            
            # ساخت فایل اکسل خروجی در مموری
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.main_df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            st.download_button(
                label="📥 دانلود فایل اکسل بروزرسانی شده",
                data=output.getvalue(),
                file_name="Updated_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
