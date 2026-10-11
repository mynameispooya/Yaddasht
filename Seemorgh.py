import streamlit as st
import pandas as pd
import io

# تنظیمات صفحه
st.set_page_config(page_title="ویرایشگر هوشمند اکسل", layout="wide")
st.title("فیلتر و ویرایش داده‌های اکسل")
st.markdown("فایل اکسل خود را آپلود کنید، ستون مبنای فیلتر و هر تعداد ستون دیگر که نیاز دارید را انتخاب کرده و مقادیر را ویرایش کنید.")

# آپلود فایل
uploaded_file = st.file_uploader("فایل اکسل خود را اینجا آپلود کنید (xlsx, xls)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    if 'main_df' not in st.session_state or st.session_state.get('uploaded_filename') != uploaded_file.name:
        st.session_state.main_df = pd.read_excel(uploaded_file)
        st.session_state.uploaded_filename = uploaded_file.name

    df = st.session_state.main_df
    
    st.divider()
    
    st.subheader("۱. انتخاب ستون‌ها")
    col1, col2 = st.columns([1, 2])
    
    with col1:
        # ستون اول فقط برای فیلتر کردن (یک انتخاب)
        filter_col = st.selectbox("ستون مبنای فیلتر (مثلاً مقطع تحصیلی):", df.columns)
    
    with col2:
        # امکان انتخاب چندگانه برای ستون‌های نمایشی
        target_cols = st.multiselect(
            "سایر ستون‌ها برای نمایش و ویرایش (می‌توانید هر تعداد که می‌خواهید انتخاب کنید):", 
            options=df.columns
        )

    # بررسی اینکه کاربر حتماً ستون‌های نمایشی را انتخاب کرده باشد
    if filter_col and len(target_cols) > 0:
        st.divider()
        st.subheader("۲. فیلتر کردن داده‌ها")
        
        unique_values = df[filter_col].dropna().unique()
        selected_val = st.selectbox(f"یک مقدار از ستون '{filter_col}' انتخاب کنید:", unique_values)
        
        if selected_val:
            # ترکیب ستون فیلتر و ستون‌های انتخابی (جلوگیری از تکرار نام ستون در لیست)
            cols_to_show = [filter_col]
            for col in target_cols:
                if col not in cols_to_show:
                    cols_to_show.append(col)
            
            # فیلتر کردن دیتافریم و نمایش ستون‌های انتخابی
            filtered_df = df[df[filter_col] == selected_val][cols_to_show]
            
            st.divider()
            st.subheader("۳. ویرایش مقادیر")
            st.info("💡 روی هر سلولی که می‌خواهید کلیک کنید و مقدار آن را تغییر دهید. می‌توانید ستون‌های راهنما (مثل نام) را ببینید و فقط ستون هدف (مثل رشته) را ویرایش کنید.")
            
            edited_df = st.data_editor(filtered_df, use_container_width=True)
            
            if st.button("ثبت تغییرات جدول بالا در فایل اصلی", type="primary"):
                st.session_state.main_df.update(edited_df)
                st.success("✅ تغییرات با موفقیت اعمال شد! می‌توانید فیلتر را عوض کنید یا خروجی بگیرید.")
                st.balloons()
            
            st.divider()
            st.subheader("۴. دریافت خروجی")
            
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
    elif filter_col and len(target_cols) == 0:
        st.warning("لطفاً از کادر بالا، حداقل یک ستون دیگر را برای نمایش انتخاب کنید.")
