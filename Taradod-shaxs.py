# -*- coding: utf-8 -*-
"""
=============================================================================
 اپلیکیشن محاسبه تردد و کارکرد پرسنل  —  taradod-shaxs.py
=============================================================================
این اپ چند فایل اکسل «تردد» (با ساختاری مشابه فایل نمونه‌ی ارسالی کاربر،
شامل ستون‌های «روز»، «شروع»، «پایان»، «مدت» و «وضعیت») را می‌خواند و بر
اساس تاریخ شمسی و تقویم تعطیلات رسمی ایران، مقادیر زیر را برای هر روزِ هر
فرد محاسبه می‌کند:

    ۱) کارکرد روزانه          : زمان حضور در بازه‌ی ۰۷:۱۵ تا ۱۶:۰۰
    ۲) اضافه‌کاری عادی        : زمان حضور بعد از ۱۶:۰۰ در روزهای عادی
    ۳) اضافه‌کاری روز تعطیل   : کل زمان حضور (از ۰۷:۱۵ به بعد، بدون سقف)
                                 در روزهای تعطیل رسمی (به‌جز جمعه‌ها)
    ۴) جمعه‌کاری              : کل زمان حضور (از ۰۷:۱۵ به بعد، بدون سقف)
                                 در روزهای جمعه

قوانین دقیق پیاده‌سازی‌شده (طبق درخواست کاربر):
    - هر ساعتی قبل از ۰۷:۱۵ در محاسبات نادیده گرفته می‌شود.
    - تشخیص تعطیلات رسمی به‌صورت خودکار از روی تقویم ایران انجام می‌شود
      و کاربر می‌تواند تعطیلات را به‌صورت دستی هم اضافه/حذف کند.
    - محاسبه بر مبنای «هر بازه‌ی ورود/خروج ثبت‌شده» انجام می‌شود، نه صرفاً
      اولین ورود و آخرین خروج روز؛ به این ترتیب اگر در طول روز وقفه‌ای
      (مثل ساعت استراحت) بین دو بازه‌ی ثبت‌شده وجود داشته باشد، آن وقفه
      به‌اشتباه به‌عنوان کارکرد حساب نمی‌شود. (در فایل نمونه بازه‌ها پیوسته
      بودند، اما این منطق برای فایل‌های دیگر هم درست کار می‌کند.)

اجرای اپ:
    pip install -r requirements.txt
    streamlit run taradod-shaxs.py
=============================================================================
"""

import io
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional, Set, Tuple

import jdatetime
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import holidays as pyholidays
except ImportError:  # pragma: no cover
    pyholidays = None


# =============================================================================
# بخش ۱: ثابت‌های کسب‌وکار
# =============================================================================
SHIFT_START_STR = "07:15"   # شروع رسمی شیفت برای همه‌ی افراد
SHIFT_END_STR = "16:00"     # پایان رسمی شیفت برای همه‌ی افراد
FRIDAY_LABEL = "جمعه"

REQUIRED_HEADERS = ["روز", "شروع", "پایان"]
OPTIONAL_HEADERS = ["مدت", "وضعیت"]

APP_TITLE = "محاسبه‌گر تردد و کارکرد پرسنل"


# =============================================================================
# بخش ۲: توابع کمکی زمان (رشته‌ی hh:mm  <->  دقیقه)
# =============================================================================
def clean_text(value) -> str:
    """حذف فاصله‌های نامرئی (nbsp، zero-width) و یکدست‌سازی فاصله‌ها."""
    if value is None:
        return ""
    s = str(value)
    s = s.replace("\xa0", " ").replace("\u200c", " ").replace("\u200f", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def hhmm_to_minutes(value) -> Optional[int]:
    """تبدیل رشته‌ی hh:mm به تعداد دقیقه از نیمه‌شب. مقدار نامعتبر -> None."""
    s = clean_text(value)
    if not s or ":" not in s:
        return None
    parts = s.split(":")
    if len(parts) != 2:
        return None
    try:
        h = int(parts[0])
        m = int(parts[1])
    except ValueError:
        return None
    if h < 0 or m < 0 or m > 59:
        return None
    return h * 60 + m


def minutes_to_hhmm(total_minutes: Optional[float]) -> str:
    """
    تبدیل دقیقه به رشته‌ی H:MM.
    عمداً سقف ۲۴ ساعته اعمال نمی‌شود تا جمع‌های ماهانه (مثلاً ۱۳۴:۰۵) درست
    نمایش داده شوند.
    """
    if total_minutes is None:
        total_minutes = 0
    sign = "-" if total_minutes < 0 else ""
    total_minutes = int(round(abs(total_minutes)))
    h, m = divmod(total_minutes, 60)
    return f"{sign}{h}:{m:02d}"


def minutes_to_decimal_hours(total_minutes: Optional[float]) -> float:
    if total_minutes is None:
        total_minutes = 0
    return round(total_minutes / 60.0, 3)


SHIFT_START_MIN = hhmm_to_minutes(SHIFT_START_STR)
SHIFT_END_MIN = hhmm_to_minutes(SHIFT_END_STR)


# =============================================================================
# بخش ۳: تشخیص تاریخ شمسی و روز هفته از متن ستون «روز»
# =============================================================================
_DAY_RE = re.compile(r"^(?P<weekday>[^\d]+?)\s+(?P<date>\d{2,4}/\d{1,2}/\d{1,2})\s*$")


def parse_day_cell(raw_value) -> Optional[Tuple[str, str, "jdatetime.date"]]:
    """
    از متن ستون «روز» (مثال: «دوشنبه 1405/05/26») نام روز هفته، رشته‌ی
    تاریخ شمسی و شیء jdatetime.date را استخراج می‌کند.
    """
    text = clean_text(raw_value)
    if not text:
        return None
    m = _DAY_RE.match(text)
    if not m:
        return None
    weekday = m.group("weekday").strip()
    date_str = m.group("date").strip()
    try:
        y, mo, d = (int(x) for x in date_str.split("/"))
        jdate = jdatetime.date(y, mo, d)
    except Exception:
        return None
    return weekday, date_str, jdate


# =============================================================================
# بخش ۴: خواندن فایل اکسل ورودی
# =============================================================================
@dataclass
class RawDay:
    """داده‌ی خامِ یک روز، پیش از دسته‌بندی عادی/تعطیل/جمعه."""
    weekday: str
    jalali_date: str
    jdate: "jdatetime.date"
    order: int
    statuses: List[str] = field(default_factory=list)
    segments: List[Tuple[int, int]] = field(default_factory=list)  # (start_min, end_min)


def find_header_row(ws) -> Optional[Dict[str, int]]:
    """
    در ۱۰ سطر ابتدایی شیت به‌دنبال سطرِ عنوانی می‌گردد که هم‌زمان ستون‌های
    «روز»، «شروع» و «پایان» را داشته باشد.
    خروجی: دیکشنری {نام‌ستون: شماره‌ستون(۱-پایه)} به‌همراه کلید 'header_row'.
    """
    max_scan_row = min(10, ws.max_row)
    for r in range(1, max_scan_row + 1):
        headers = {}
        for c in range(1, ws.max_column + 1):
            val = clean_text(ws.cell(row=r, column=c).value)
            if val:
                headers[val] = c
        if all(h in headers for h in REQUIRED_HEADERS):
            return {"header_row": r, **headers}
    return None


def parse_attendance_workbook(file_obj, filename: str = "") -> List[RawDay]:
    """
    یک فایل اکسل تردد را می‌خواند و لیستی از RawDay (به ترتیب ظاهرشدن در
    فایل) برمی‌گرداند. هر سطرِ دارای مقدار معتبر «شروع» و «پایان» به‌عنوان
    یک بازه‌ی ورود/خروج به روزِ مربوطه افزوده می‌شود.
    """
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.worksheets[0]

    header_info = find_header_row(ws)
    if header_info is None:
        raise ValueError(
            f"ساختار فایل «{filename}» شناسایی نشد؛ ستون‌های «روز»، «شروع» و "
            f"«پایان» پیدا نشدند. لطفاً از یکسان بودن ساختار فایل با نمونه "
            f"اطمینان حاصل کنید."
        )

    header_row = header_info["header_row"]
    col_day = header_info["روز"]
    col_start = header_info["شروع"]
    col_end = header_info["پایان"]
    col_status = header_info.get("وضعیت")

    days: Dict[str, RawDay] = {}
    order_counter = 0

    for r in range(header_row + 1, ws.max_row + 1):
        day_raw = ws.cell(row=r, column=col_day).value
        parsed = parse_day_cell(day_raw)
        if parsed is None:
            continue
        weekday, jalali_date, jdate = parsed
        key = f"{weekday} {jalali_date}"

        if key not in days:
            days[key] = RawDay(
                weekday=weekday,
                jalali_date=jalali_date,
                jdate=jdate,
                order=order_counter,
            )
            order_counter += 1
        day_rec = days[key]

        if col_status:
            status_val = clean_text(ws.cell(row=r, column=col_status).value)
            if status_val:
                day_rec.statuses.append(status_val)

        start_val = ws.cell(row=r, column=col_start).value
        end_val = ws.cell(row=r, column=col_end).value
        start_min = hhmm_to_minutes(start_val)
        end_min = hhmm_to_minutes(end_val)
        if start_min is not None and end_min is not None and end_min > start_min:
            day_rec.segments.append((start_min, end_min))

    ordered_days = sorted(days.values(), key=lambda d: d.order)
    return ordered_days


# =============================================================================
# بخش ۵: محاسبه‌ی کارکرد هر روز
# =============================================================================
@dataclass
class DayResult:
    weekday: str
    jalali_date: str
    gregorian_date: str
    day_type: str  # 'عادی' | 'تعطیل رسمی' | 'جمعه'
    status: str
    first_in: str
    last_out: str
    daily_work_min: int
    overtime_normal_min: int
    holiday_overtime_min: int
    friday_work_min: int
    needs_review: bool
    review_note: str


def classify_day_type(weekday: str, is_holiday_calendar_day: bool) -> str:
    if clean_text(weekday) == FRIDAY_LABEL:
        return "جمعه"
    if is_holiday_calendar_day:
        return "تعطیل رسمی"
    return "عادی"


def compute_day_result(raw_day: RawDay, holiday_gregorian_dates: Set[date]) -> DayResult:
    """
    منطق محاسبه (به ازای هر بازه‌ی ورود/خروج ثبت‌شده در آن روز، جداگانه):
      - روز عادی:
            کارکرد روزانه   += بخشی از بازه که داخل [07:15 , 16:00] است
            اضافه‌کاری عادی += بخشی از بازه که بعد از 16:00 است
      - روز تعطیل رسمی (غیر جمعه) / جمعه:
            کل بخشِ بازه که از max(07:15, شروع بازه) به بعد است، بدون سقفِ
            16:00، در ستون مربوطه (اضافه‌کاری روز تعطیل / جمعه‌کاری) جمع می‌شود.
    هر زمانِ قبل از ۰۷:۱۵ در همه‌ی حالت‌ها نادیده گرفته می‌شود.
    """
    gdate = raw_day.jdate.togregorian()
    is_holiday = gdate in holiday_gregorian_dates
    day_type = classify_day_type(raw_day.weekday, is_holiday)

    segments = sorted(raw_day.segments)
    status = raw_day.statuses[0] if raw_day.statuses else ""

    review_notes = []
    first_in_min = segments[0][0] if segments else None
    last_out_min = segments[-1][1] if segments else None

    if segments:
        for i in range(1, len(segments)):
            if segments[i][0] < segments[i - 1][1]:
                review_notes.append("همپوشانی بین بازه‌های ورود/خروج ثبت‌شده")
                break
        if first_in_min is not None and first_in_min > SHIFT_END_MIN:
            review_notes.append("ورود ثبت‌شده بعد از پایان شیفت رسمی است")
    else:
        normal_absence_labels = ("عدم حضور", "غیبت", "")
        if status not in normal_absence_labels and "مرخص" not in status:
            review_notes.append("بدون بازه‌ی ورود/خروج معتبر ولی وضعیت نامشخص")

    daily_work = 0
    overtime_normal = 0
    holiday_overtime = 0
    friday_work = 0

    for seg_start, seg_end in segments:
        eff_start = max(seg_start, SHIFT_START_MIN)
        if eff_start >= seg_end:
            continue  # کل بازه قبل از ۰۷:۱۵ بوده -> نادیده گرفته می‌شود

        if day_type == "عادی":
            reg_part = max(0, min(seg_end, SHIFT_END_MIN) - eff_start)
            over_part = max(0, seg_end - max(eff_start, SHIFT_END_MIN))
            daily_work += reg_part
            overtime_normal += over_part
        elif day_type == "تعطیل رسمی":
            holiday_overtime += max(0, seg_end - eff_start)
        else:  # جمعه
            friday_work += max(0, seg_end - eff_start)

    return DayResult(
        weekday=raw_day.weekday,
        jalali_date=raw_day.jalali_date,
        gregorian_date=gdate.isoformat(),
        day_type=day_type,
        status=status,
        first_in=minutes_to_hhmm(first_in_min) if first_in_min is not None else "",
        last_out=minutes_to_hhmm(last_out_min) if last_out_min is not None else "",
        daily_work_min=daily_work,
        overtime_normal_min=overtime_normal,
        holiday_overtime_min=holiday_overtime,
        friday_work_min=friday_work,
        needs_review=bool(review_notes),
        review_note="؛ ".join(review_notes),
    )


# =============================================================================
# بخش ۶: مدیریت تعطیلات رسمی (تشخیص خودکار + افزودن/حذف دستی)
# =============================================================================
@st.cache_data(show_spinner=False)
def fetch_official_holidays(years: Tuple[int, ...]) -> Dict[str, str]:
    """
    دریافت تعطیلات رسمی ایران برای سال‌های میلادیِ داده‌شده.
    خروجی: {تاریخ میلادی به‌صورت ISO: عنوان تعطیلی}
    توجه: تاریخ‌های مذهبیِ قمری، تخمینی محاسبه می‌شوند و ممکن است در واقعیت
    یک روز اختلاف داشته باشند (چون تاریخ رسمی با رؤیت هلال اعلام می‌شود)؛
    به همین دلیل امکان ویرایش دستی این لیست در اپ فراهم شده است.
    """
    if pyholidays is None or not years:
        return {}
    try:
        hol = pyholidays.Iran(years=list(years))
        return {d.isoformat(): str(label) for d, label in hol.items()}
    except Exception:
        return {}


def jalali_year_range_for_gdate(gdate: date) -> int:
    return gdate.year


def build_default_holiday_table(raw_days_by_person: Dict[str, List[RawDay]]) -> pd.DataFrame:
    """
    از روی بازه‌ی تاریخیِ فایل‌های آپلودشده، سال‌های میلادیِ لازم را تشخیص
    می‌دهد، تعطیلات رسمی را می‌گیرد و آن‌ها را به‌صورت یک DataFrame قابل
    ویرایش در Streamlit برمی‌گرداند.
    """
    years: Set[int] = set()
    for raw_days in raw_days_by_person.values():
        for d in raw_days:
            g = d.jdate.togregorian()
            years.add(g.year)
    if not years:
        return pd.DataFrame(
            columns=["فعال", "تاریخ میلادی", "تاریخ شمسی", "روز هفته", "عنوان", "منبع"]
        )
    years_expanded = tuple(sorted({y - 1 for y in years} | years | {y + 1 for y in years}))
    hol_map = fetch_official_holidays(years_expanded)

    rows = []
    weekday_fa = ["دوشنبه", "سه شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه", "یکشنبه"]
    for iso_date, label in sorted(hol_map.items()):
        g = date.fromisoformat(iso_date)
        j = jdatetime.date.fromgregorian(date=g)
        rows.append(
            {
                "فعال": True,
                "تاریخ میلادی": iso_date,
                "تاریخ شمسی": j.strftime("%Y/%m/%d"),
                "روز هفته": weekday_fa[g.weekday()],
                "عنوان": label,
                "منبع": "خودکار",
            }
        )
    return pd.DataFrame(rows)


def jalali_str_to_gregorian_iso(jalali_str: str) -> Optional[str]:
    """تبدیل رشته‌ی تاریخ شمسی 'YYYY/MM/DD' به تاریخ میلادی ISO."""
    s = clean_text(jalali_str)
    m = re.match(r"^(\d{3,4})/(\d{1,2})/(\d{1,2})$", s)
    if not m:
        return None
    try:
        y, mo, d = (int(x) for x in m.groups())
        jd = jdatetime.date(y, mo, d)
        return jd.togregorian().isoformat()
    except Exception:
        return None


def holiday_set_from_table(df: pd.DataFrame) -> Set[date]:
    """استخراج مجموعه‌ی تاریخ‌های میلادیِ «فعال» از جدول تعطیلات برای محاسبه."""
    if df is None or df.empty:
        return set()
    active = df[df["فعال"] == True]  # noqa: E712
    result = set()
    for iso in active["تاریخ میلادی"]:
        try:
            result.add(date.fromisoformat(str(iso)))
        except Exception:
            continue
    return result


# =============================================================================
# بخش ۷: پردازش کامل یک فایل (شخص) و ساخت جدول روزانه + خلاصه
# =============================================================================
DETAIL_COLUMNS = [
    "ردیف",
    "روز هفته",
    "تاریخ شمسی",
    "تاریخ میلادی",
    "نوع روز",
    "وضعیت ثبت‌شده",
    "اولین ورود",
    "آخرین خروج",
    "کارکرد روزانه",
    "کارکرد روزانه (ساعت)",
    "اضافه‌کاری عادی",
    "اضافه‌کاری عادی (ساعت)",
    "اضافه‌کاری روز تعطیل",
    "اضافه‌کاری روز تعطیل (ساعت)",
    "جمعه‌کاری",
    "جمعه‌کاری (ساعت)",
    "نیاز به بررسی",
    "توضیح بررسی",
]


def build_daily_dataframe(raw_days: List[RawDay], holiday_set: Set[date]) -> pd.DataFrame:
    rows = []
    for i, rd in enumerate(raw_days, start=1):
        res = compute_day_result(rd, holiday_set)
        rows.append(
            {
                "ردیف": i,
                "روز هفته": res.weekday,
                "تاریخ شمسی": res.jalali_date,
                "تاریخ میلادی": res.gregorian_date,
                "نوع روز": res.day_type,
                "وضعیت ثبت‌شده": res.status,
                "اولین ورود": res.first_in,
                "آخرین خروج": res.last_out,
                "کارکرد روزانه": minutes_to_hhmm(res.daily_work_min),
                "کارکرد روزانه (ساعت)": minutes_to_decimal_hours(res.daily_work_min),
                "اضافه‌کاری عادی": minutes_to_hhmm(res.overtime_normal_min),
                "اضافه‌کاری عادی (ساعت)": minutes_to_decimal_hours(res.overtime_normal_min),
                "اضافه‌کاری روز تعطیل": minutes_to_hhmm(res.holiday_overtime_min),
                "اضافه‌کاری روز تعطیل (ساعت)": minutes_to_decimal_hours(res.holiday_overtime_min),
                "جمعه‌کاری": minutes_to_hhmm(res.friday_work_min),
                "جمعه‌کاری (ساعت)": minutes_to_decimal_hours(res.friday_work_min),
                "نیاز به بررسی": "بله" if res.needs_review else "خیر",
                "توضیح بررسی": res.review_note,
            }
        )
    return pd.DataFrame(rows, columns=DETAIL_COLUMNS)


def build_summary_dict(person_name: str, df: pd.DataFrame) -> Dict[str, object]:
    if df.empty:
        return {
            "نام": person_name,
            "تعداد روز ثبت‌شده": 0,
            "مجموع کارکرد روزانه": "0:00",
            "مجموع اضافه‌کاری عادی": "0:00",
            "مجموع اضافه‌کاری روز تعطیل": "0:00",
            "مجموع جمعه‌کاری": "0:00",
            "تعداد روزهای نیازمند بررسی": 0,
        }
    total_work = df["کارکرد روزانه (ساعت)"].sum() * 60
    total_over = df["اضافه‌کاری عادی (ساعت)"].sum() * 60
    total_hol = df["اضافه‌کاری روز تعطیل (ساعت)"].sum() * 60
    total_fri = df["جمعه‌کاری (ساعت)"].sum() * 60
    return {
        "نام": person_name,
        "تعداد روز ثبت‌شده": len(df),
        "مجموع کارکرد روزانه": minutes_to_hhmm(total_work),
        "مجموع اضافه‌کاری عادی": minutes_to_hhmm(total_over),
        "مجموع اضافه‌کاری روز تعطیل": minutes_to_hhmm(total_hol),
        "مجموع جمعه‌کاری": minutes_to_hhmm(total_fri),
        "تعداد روزهای نیازمند بررسی": int((df["نیاز به بررسی"] == "بله").sum()),
    }


def process_person_file(file_bytes: bytes, filename: str, holiday_set: Set[date]):
    """پردازش کامل یک فایل: خواندن + محاسبه؛ خروجی (raw_days, df, summary)."""
    raw_days = parse_attendance_workbook(io.BytesIO(file_bytes), filename)
    df = build_daily_dataframe(raw_days, holiday_set)
    return raw_days, df


# =============================================================================
# بخش ۸: ساخت خروجی اکسل استایل‌دار
# =============================================================================
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FRIDAY_FILL = PatternFill("solid", fgColor="DDEBF7")
HOLIDAY_FILL = PatternFill("solid", fgColor="FCE4D6")
REVIEW_FONT = Font(color="C00000", bold=True)
THIN_BORDER = Border(*(Side(style="thin", color="B7B7B7"),) * 4)
DECIMAL_COLS = {
    "کارکرد روزانه (ساعت)",
    "اضافه‌کاری عادی (ساعت)",
    "اضافه‌کاری روز تعطیل (ساعت)",
    "جمعه‌کاری (ساعت)",
}


def sanitize_sheet_name(name: str) -> str:
    bad_chars = r"[]:\\/?*"
    cleaned = "".join(c for c in name if c not in bad_chars)
    cleaned = cleaned.strip() or "Sheet"
    return cleaned[:31]


def _write_header(ws, headers: List[str], row: int, start_col: int = 1) -> None:
    for j, col_name in enumerate(headers, start=start_col):
        cell = ws.cell(row=row, column=j, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def _write_dataframe_block(ws, df: pd.DataFrame, start_row: int) -> int:
    """جدول df را از start_row می‌نویسد و شماره‌ی آخرین سطر نوشته‌شده را برمی‌گرداند."""
    headers = list(df.columns)
    _write_header(ws, headers, start_row)

    day_type_col = headers.index("نوع روز") + 1 if "نوع روز" in headers else None
    review_col = headers.index("نیاز به بررسی") + 1 if "نیاز به بررسی" in headers else None

    for i, (_, row_data) in enumerate(df.iterrows()):
        r = start_row + 1 + i
        for j, col_name in enumerate(headers, start=1):
            val = row_data[col_name]
            cell = ws.cell(row=r, column=j, value=val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_name in DECIMAL_COLS:
                cell.number_format = "0.00"

        if day_type_col:
            day_type_val = row_data.get("نوع روز")
            if day_type_val == "جمعه":
                for j in range(1, len(headers) + 1):
                    ws.cell(row=r, column=j).fill = FRIDAY_FILL
            elif day_type_val == "تعطیل رسمی":
                for j in range(1, len(headers) + 1):
                    ws.cell(row=r, column=j).fill = HOLIDAY_FILL

        if review_col and row_data.get("نیاز به بررسی") == "بله":
            ws.cell(row=r, column=review_col).font = REVIEW_FONT

    last_row = start_row + len(df)
    for j, col_name in enumerate(headers, start=1):
        try:
            max_len = max([len(str(col_name))] + [len(str(v)) for v in df[col_name].astype(str)])
        except Exception:
            max_len = len(str(col_name))
        ws.column_dimensions[get_column_letter(j)].width = min(max(10, max_len + 2), 32)

    ws.freeze_panes = ws.cell(row=start_row + 1, column=1).coordinate
    ws.auto_filter.ref = (
        f"A{start_row}:{get_column_letter(len(headers))}{last_row}" if last_row >= start_row else None
    )
    return last_row


def _write_summary_block(ws, summary: Dict[str, object], start_row: int, title: str) -> int:
    cell = ws.cell(row=start_row, column=1, value=title)
    cell.font = Font(bold=True, size=13)
    r = start_row + 1
    for key, val in summary.items():
        if key == "نام":
            continue
        ws.cell(row=r, column=1, value=key).font = Font(bold=True)
        ws.cell(row=r, column=2, value=val)
        r += 1
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20
    return r


def build_person_workbook(person_name: str, df: pd.DataFrame, summary: Dict[str, object]) -> bytes:
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = sanitize_sheet_name("خلاصه")
    ws_summary.sheet_view.rightToLeft = True
    _write_summary_block(ws_summary, summary, 1, f"خلاصه‌ی کارکرد: {person_name}")

    ws_detail = wb.create_sheet(sanitize_sheet_name("جزئیات روزانه"))
    ws_detail.sheet_view.rightToLeft = True
    _write_dataframe_block(ws_detail, df, 1)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_combined_workbook(persons: Dict[str, Tuple[pd.DataFrame, Dict[str, object]]]) -> bytes:
    wb = openpyxl.Workbook()
    ws_overview = wb.active
    ws_overview.title = sanitize_sheet_name("خلاصه کلی")
    ws_overview.sheet_view.rightToLeft = True

    overview_rows = [summary for (_, summary) in persons.values()]
    overview_df = pd.DataFrame(overview_rows)
    if not overview_df.empty:
        _write_dataframe_block(ws_overview, overview_df, 1)

    used_names: Set[str] = set()
    for person_name, (df, _summary) in persons.items():
        base_name = sanitize_sheet_name(person_name)
        sheet_name = base_name
        suffix = 1
        while sheet_name in used_names:
            suffix += 1
            sheet_name = sanitize_sheet_name(f"{base_name[:28]}_{suffix}")
        used_names.add(sheet_name)

        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.rightToLeft = True
        _write_dataframe_block(ws, df, 1)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =============================================================================
# بخش ۹: رابط کاربری Streamlit
# =============================================================================
RTL_CSS = """
<style>
html, body, [class*="css"]  { direction: rtl; text-align: right; }
.stDataFrame { direction: ltr; }
div[data-testid="stMetricValue"] { direction: ltr; }
</style>
"""


def _init_session_state():
    if "holiday_df" not in st.session_state:
        st.session_state.holiday_df = None
    if "holiday_years_key" not in st.session_state:
        st.session_state.holiday_years_key = None
    if "person_names" not in st.session_state:
        st.session_state.person_names = {}


def _sync_holiday_table(raw_days_by_person: Dict[str, List[RawDay]]):
    """اگر بازه‌ی سالی فایل‌ها تغییر کرده، جدول تعطیلات خودکار را می‌سازد یا
    گسترش می‌دهد، بدون این‌که ویرایش‌های دستیِ قبلیِ کاربر پاک شود."""
    years_key = tuple(
        sorted(
            {
                rd.jdate.togregorian().year
                for raw_days in raw_days_by_person.values()
                for rd in raw_days
            }
        )
    )
    if st.session_state.holiday_df is None or st.session_state.holiday_years_key != years_key:
        new_default = build_default_holiday_table(raw_days_by_person)
        if st.session_state.holiday_df is None:
            st.session_state.holiday_df = new_default
        else:
            existing = st.session_state.holiday_df
            existing_dates = set(existing["تاریخ میلادی"].astype(str))
            extra_rows = new_default[~new_default["تاریخ میلادی"].astype(str).isin(existing_dates)]
            if not extra_rows.empty:
                st.session_state.holiday_df = pd.concat([existing, extra_rows], ignore_index=True)
        st.session_state.holiday_years_key = years_key


def render_holiday_editor():
    st.subheader("📅 تعطیلات رسمی")
    st.caption(
        "تعطیلات رسمی به‌صورت خودکار از روی تقویم ایران تشخیص داده شده‌اند. "
        "تاریخ‌های مذهبیِ قمری تخمینی هستند و ممکن است یک روز اختلاف داشته "
        "باشند — در صورت نیاز، تیک هر ردیف را بردارید یا ردیف جدید اضافه کنید. "
        "جمعه‌ها به‌صورت جداگانه و خودکار شناسایی می‌شوند و نیازی به افزودن آن‌ها نیست."
    )

    edited = st.data_editor(
        st.session_state.holiday_df,
        key="holiday_editor",
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "فعال": st.column_config.CheckboxColumn("فعال (تعطیل حساب شود؟)"),
            "تاریخ میلادی": st.column_config.TextColumn("تاریخ میلادی (YYYY-MM-DD)"),
            "تاریخ شمسی": st.column_config.TextColumn("تاریخ شمسی (YYYY/MM/DD)"),
        },
        hide_index=True,
    )
    st.session_state.holiday_df = edited

    with st.expander("➕ افزودن تعطیل دستی با تاریخ شمسی"):
        c1, c2, c3 = st.columns([2, 3, 1])
        with c1:
            new_jalali = st.text_input("تاریخ شمسی (مثال: 1405/07/01)", key="new_holiday_jalali")
        with c2:
            new_label = st.text_input("عنوان تعطیلی", key="new_holiday_label")
        with c3:
            st.write("")
            st.write("")
            add_clicked = st.button("افزودن", key="add_holiday_btn")
        if add_clicked:
            iso = jalali_str_to_gregorian_iso(new_jalali)
            if iso is None:
                st.error("تاریخ شمسی نامعتبر است. فرمت صحیح: 1405/07/01")
            else:
                g = date.fromisoformat(iso)
                weekday_fa = ["دوشنبه", "سه شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه", "یکشنبه"]
                new_row = pd.DataFrame(
                    [
                        {
                            "فعال": True,
                            "تاریخ میلادی": iso,
                            "تاریخ شمسی": clean_text(new_jalali),
                            "روز هفته": weekday_fa[g.weekday()],
                            "عنوان": new_label or "تعطیل دستی",
                            "منبع": "دستی",
                        }
                    ]
                )
                st.session_state.holiday_df = pd.concat(
                    [st.session_state.holiday_df, new_row], ignore_index=True
                )
                st.rerun()


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🕘", layout="wide")
    st.markdown(RTL_CSS, unsafe_allow_html=True)
    _init_session_state()

    st.title("🕘 " + APP_TITLE)
    st.caption(
        f"محاسبه‌ی کارکرد روزانه، اضافه‌کاری، اضافه‌کاری روز تعطیل و جمعه‌کاری "
        f"بر اساس ساعت شروعِ رسمی «{SHIFT_START_STR}» و پایانِ رسمی «{SHIFT_END_STR}»."
    )

    with st.expander("ℹ️ منطق محاسبه (برای اطمینان از صحت قبل از استفاده در حقوق)", expanded=False):
        st.markdown(
            f"""
- هر بازه‌ی ورود/خروجِ ثبت‌شده در فایل، **جداگانه** محاسبه می‌شود (نه فقط
  اولین ورود و آخرین خروج روز)؛ بنابراین اگر بین دو بازه وقفه‌ای وجود داشته
  باشد، آن وقفه جزو کارکرد حساب نمی‌شود.
- هر زمانی **قبل از {SHIFT_START_STR}** در همه‌ی محاسبات نادیده گرفته می‌شود.
- **روز عادی:** بخشِ هر بازه که در محدوده‌ی {SHIFT_START_STR} تا {SHIFT_END_STR}
  است → «کارکرد روزانه»؛ بخشِ بعد از {SHIFT_END_STR} → «اضافه‌کاری عادی».
- **تعطیل رسمی (غیر جمعه) و جمعه:** کل زمان حضور از {SHIFT_START_STR} به بعد
  (بدون سقف {SHIFT_END_STR}) به‌طور کامل در ستون «اضافه‌کاری روز تعطیل» یا
  «جمعه‌کاری» جمع می‌شود.
- روزهایی که وضعیتِ غیرعادی دارند (مثلاً ورودِ بعد از پایان شیفت، یا
  همپوشانیِ بازه‌ها) با ستون «نیاز به بررسی» علامت‌گذاری می‌شوند تا پیش از
  استفاده در فیش حقوقی، به‌صورت دستی بازبینی شوند.
"""
        )

    uploaded_files = st.file_uploader(
        "فایل‌های اکسل تردد را آپلود کنید (هر فایل معادل یک نفر، با ساختار مشابه فایل نمونه)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("برای شروع، یک یا چند فایل اکسل تردد را آپلود کنید.")
        return

    file_bytes_map: Dict[str, bytes] = {}
    for f in uploaded_files:
        file_bytes_map[f.name] = f.getvalue()

    st.subheader("👤 نام‌گذاری افراد")
    st.caption("پیش‌فرض، نام فایل به‌عنوان نام فرد در نظر گرفته می‌شود؛ در صورت نیاز ویرایش کنید.")
    person_names: Dict[str, str] = {}
    cols = st.columns(min(3, len(uploaded_files)) or 1)
    for i, fname in enumerate(file_bytes_map):
        default_name = st.session_state.person_names.get(fname, fname.rsplit(".", 1)[0])
        with cols[i % len(cols)]:
            name = st.text_input(f"نام برای «{fname}»", value=default_name, key=f"name_{fname}")
        person_names[fname] = name.strip() or fname
    st.session_state.person_names = person_names

    raw_days_by_person: Dict[str, List[RawDay]] = {}
    parse_errors = []
    for fname, fbytes in file_bytes_map.items():
        person_name = person_names[fname]
        try:
            raw_days = parse_attendance_workbook(io.BytesIO(fbytes), fname)
            raw_days_by_person[person_name] = raw_days
        except Exception as e:  # noqa: BLE001
            parse_errors.append(f"«{fname}»: {e}")

    for err in parse_errors:
        st.error(err)

    if not raw_days_by_person:
        st.warning("هیچ فایل قابل پردازشی وجود ندارد.")
        return

    _sync_holiday_table(raw_days_by_person)
    render_holiday_editor()

    holiday_set = holiday_set_from_table(st.session_state.holiday_df)

    st.divider()
    st.subheader("📊 نتایج محاسبه")

    persons_results: Dict[str, Tuple[pd.DataFrame, Dict[str, object]]] = {}
    for person_name, raw_days in raw_days_by_person.items():
        df = build_daily_dataframe(raw_days, holiday_set)
        summary = build_summary_dict(person_name, df)
        persons_results[person_name] = (df, summary)

    overview_df = pd.DataFrame([s for (_, s) in persons_results.values()])
    st.markdown("**خلاصه‌ی کلی همه‌ی افراد**")
    st.dataframe(overview_df, use_container_width=True, hide_index=True)

    review_total = int(overview_df["تعداد روزهای نیازمند بررسی"].sum()) if not overview_df.empty else 0
    if review_total > 0:
        st.warning(
            f"⚠️ در مجموع {review_total} روز وجود دارد که به‌دلیل ثبت غیرعادی "
            f"(مثل ورود دیرهنگام یا همپوشانی بازه‌ها) نیاز به بررسی دستی دارد. "
            f"جزئیات را در ستون «نیاز به بررسی» هر فرد ببینید."
        )

    tabs = st.tabs(list(persons_results.keys()))
    for tab, (person_name, (df, summary)) in zip(tabs, persons_results.items()):
        with tab:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("کارکرد روزانه", summary["مجموع کارکرد روزانه"])
            m2.metric("اضافه‌کاری عادی", summary["مجموع اضافه‌کاری عادی"])
            m3.metric("اضافه‌کاری روز تعطیل", summary["مجموع اضافه‌کاری روز تعطیل"])
            m4.metric("جمعه‌کاری", summary["مجموع جمعه‌کاری"])

            st.dataframe(df, use_container_width=True, hide_index=True)

            person_xlsx = build_person_workbook(person_name, df, summary)
            st.download_button(
                f"⬇️ دانلود فایل جداگانه‌ی «{person_name}» (اکسل)",
                data=person_xlsx,
                file_name=f"کارکرد-{person_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{person_name}",
            )

    st.divider()
    st.subheader("⬇️ خروجی‌های نهایی")

    c1, c2 = st.columns(2)
    with c1:
        combined_xlsx = build_combined_workbook(persons_results)
        st.download_button(
            "📘 دانلود گزارش جامع (همه‌ی افراد در یک فایل)",
            data=combined_xlsx,
            file_name="گزارش-جامع-کارکرد.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with c2:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for person_name, (df, summary) in persons_results.items():
                zf.writestr(f"کارکرد-{person_name}.xlsx", build_person_workbook(person_name, df, summary))
        st.download_button(
            "🗂️ دانلود همه‌ی فایل‌های جداگانه (ZIP)",
            data=zip_buf.getvalue(),
            file_name="کارکرد-همه-افراد.zip",
            mime="application/zip",
        )


if __name__ == "__main__":
    main()
