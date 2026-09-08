import pandas as pd
import numpy as np

def clean_multiindex_headers(df):
    """
    ستون‌های دارای ساختار Multi-Index را یکپارچه می‌کند.
    اگر ستون زیرمجموعه 'اطلاعات تردد' باشد، نام زیرمجموعه جایگزین می‌شود.
    """
    new_cols = []
    for col in df.columns:
        level_0, level_1 = str(col[0]).strip(), str(col[1]).strip()
        
        # حذف عبارت "اطلاعات تردد" و جایگزینی با نام‌های ردیف پایین (مثل ورود، خروج، گیت)
        if 'اطلاعات تردد' in level_0:
            new_cols.append(level_1)
        # برای ستون‌هایی که ردیف دومشان خالی یا Unnamed است، همان ردیف اول را نگه می‌داریم
        elif 'Unnamed' in level_1:
            new_cols.append(level_0)
        else:
            new_cols.append(f"{level_0} {level_1}".strip())
            
    df.columns = new_cols
    return df

def analyze_attendance(file_path):
    # ۱. خواندن فایل با در نظر گرفتن دو سطر اول به عنوان هدر (Multi-Index)
    df = pd.read_excel(file_path, header=[0, 1])
    df = clean_multiindex_headers(df)
    
    # اضافه کردن ستون‌های جدید برای وضعیت سیستم و مقادیر پیشنهادی
    df['آلارم سیستم'] = ''
    df['اقدام پیشنهادی'] = ''
    df['نیازمند_هایلایت'] = False
    
    for index, row in df.iterrows():
        تاریخ = row.get('تاریخ')
        وضعیت = row.get('وضعیت')
        ورود = row.get('ورود')
        شروع_شیفت = row.get('شروع شیفت')
        پایان_شیفت = row.get('پایان شیفت')
        
        # سناریوی اول (مثال ۱۴۰۵/۰۶/۰۹): وضعیت خالی است، ورود ثبت نشده، اما شیفت تعریف شده است
        if pd.isna(وضعیت) and pd.isna(ورود) and pd.notna(شروع_شیفت):
            df.at[index, 'آلارم سیستم'] = 'خطا: عدم ثبت ورود'
            df.at[index, 'اقدام پیشنهادی'] = f'ورود باید بر اساس شیفت {شروع_شیفت} ثبت شود'
            df.at[index, 'نیازمند_هایلایت'] = True
            
        # سناریوی دوم (مثال ۱۴۰۵/۰۶/۱۳ در فایل جدید-2): ثبت ورود در روزی که الگوی شیفت انتظار آن را ندارد
        # (فرض بر این است که شروع_شیفت برای این روز در الگو خالی یا متفاوت است)
        elif pd.notna(ورود):
            # بررسی تطابق با الگو (اگر شیفت خالی باشد یا فرد در زمان استراحت وارد شود)
            if pd.isna(شروع_شیفت): 
                df.at[index, 'آلارم سیستم'] = 'آلارم: خروج از الگو'
                df.at[index, 'اقدام پیشنهادی'] = 'بررسی تردد خارج از شیفت تعریف شده'
                df.at[index, 'نیازمند_هایلایت'] = True
                
    return df

def highlight_rows(row):
    """تابع استایل‌دهی برای هایلایت کردن سطرهای دارای خطا در خروجی نهایی اکسل"""
    if row['نیازمند_هایلایت']:
        return ['background-color: #ffb3b3'] * len(row) # هایلایت قرمز روشن
    return [''] * len(row)

# نحوه اجرای اپلیکیشن و گرفتن خروجی:
# df_processed_1 = analyze_attendance('گزارش-تردد-جدید.xlsx')
# df_processed_2 = analyze_attendance('گزارش-تردد-جدید-2.xlsx')

# استخراج فایل نهایی با هایلایت سلول‌ها
# (df_processed_2.style.apply(highlight_rows, axis=1)
#                .to_excel('خروجی_هوشمند_تردد.xlsx', index=False))
