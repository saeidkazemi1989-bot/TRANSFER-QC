#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
سازندهٔ رپورت QC راهکاران
==========================
سه فایل سایت + جدول گروه‌بندی محصولات را می‌گیرد و فایل
«QC Report راهکاران» را با 7 شیت تولید می‌کند:

    FULT EMS | QC EMS | ICT | QC ELE | SMD REPORT | FULTELE | qv
    (به‌علاوه شیت «گروه بندی محصولات» در پایان، همان فایل گروه‌بندی)

منطق هر شیت:
  شیت        | کد محصول | منبع عیب‌ها        | فیلتر عیب‌ها                                          | تعداد تولید (مرکز کاری)
  -----------|----------|---------------------|---------------------------------------------------------|------------------------------------------
  qv         | 121*     | اطلاعات جامع کیفیت | همه عیب‌های 121                                           | 3 مرکز مونتاژ/وان (بدون «قبل از وان»)
  SMD REPORT | 120*     | اطلاعات جامع کیفیت | همه عیب‌های 120، ایستگاه «-»                            | سامسونگ 1/2 + ماشین میرایی
  ICT        | 122*     | اطلاعات جامع کیفیت | فقط کد عیب ICT_01، ایستگاه «ICT»                        | فقط «تکمیل کاری نود ها» (همه ردیف‌ها)
  QC ELE     | 123*     | اطلاعات جامع کیفیت | همه عیب‌های 123، ایستگاه «کنترل نهایی»                  | 3 مرکز کنترل نهایی
  FULTELE    | 121+122* | عیب‌های سند بازرسی  | همه فرآیندهایی که QV نیستند (وان قلع، قطعه‌گذاری و     | 6 مرکز مونتاژ/وان/تکمیل کاری
             |          |                     | لحیم‌کاری، تکمیل کاری و ...)؛ ICT_01 مستثنا است           |
  FULT EMS   | 320*     | عیب‌های سند بازرسی  | همه عیب‌های 320 در 4 ایستگاه تست/مونتاژ،                 | فقط «تکمیل کاری ECU» و «تست و کنترل ECU»
             |          |                     | به‌جز ECU-07 دو کد ECU (آن‌ها می‌روند QC EMS)             |
  QC EMS     | 331+332* | عیب‌های سند بازرسی  | همه عیب‌های 33 + ECU-07 دو کد 3206133/3206134،           | همه کدهای 33
             |          |                     | ایستگاه «کنترل نهایی»                                    |

نکات مشترک:
  - کد گروه محصول = کد کالای منبع؛ کد محصول = نام کالای منبع
  - نام محصول/خانواده/ترکیبی/نهایی/برنچ از جدول گروه‌بندی (اولین ردیف تکراری،
    مثل VLOOKUP اکسل)؛ اگر کد نبود → #N/A
  - ردیف عیب: تعداد کل = 0؛ ردیف تولید: کد/شرح ایراد = «-»، تعداد ایراد = 0،
    تعداد کل = مقدار سالم
  - کد عیب نرمال‌سازی می‌شود: TS-08(ف) / TS-08(01) / TS-01.  →  TS-08 / TS-01
  - شیفت: عیب‌های سند بازرسی از ستون شیفت آن گزارش؛ بقیه «-»
  - در سند بازرسی، عیب‌ها معمولاً دو نسخه دارند (عنوان «QV» + عنوان عملیات واقعی)؛
    نسخهٔ غیر-QV استفاده می‌شود.
  - ردیف‌های تولید فقط برای کدهای موجود در جدول گروه‌بندی می‌آیند
    (شیت ICT مستثنا است و همه ردیف‌های مرکز خودش را می‌گیرد).
  - ردیف‌ها بر اساس کد محصول (ترتیب حضور) و داخل هر کد بر اساس تاریخ مرتب می‌شوند.

قالب و نکات خاص شیت‌ها (دقیقاً مطابق رپورت ماه قبل):
  - شیت SMD REPORT قالب 68 ستونه دارد: بعد از تاریخ یک ستون «ماه»، بعد از
    «نام قطعه» ستون «جانمایی قطعه معیوب در فرآیند SMD» (از ستون جانمایی
    گزارش کیفیت) و در انتها 13 ستون اپراتور/ماشین SMD. در ردیف‌های تولید
    SMD هزینه تعمیرات 0 و «ماشین Paste SMD» = مرکز کاری است؛ در ردیف‌های
    عیب، «رفع عیب»/«تست مجدد» خالی با «-» نوشته می‌شود.
  - ردیف‌های تولید ICT: هزینه و سه زمان تعمیرات «-» است.
  - در FULTELE شیفت ردیف‌های تولید مراکز «تکمیل کاری» خالی و بقیه «-» است.
  - مقدار خالی در ستون‌های متنی عیب، «-» نوشته می‌شود.

حالت «ادامه از رپورت ماه قبل» (مهم):
  --history "رپورت ماه قبل.xlsx"
    1) جدول گروه‌بندی داخل رپورت قبلی را هم می‌خواند و برای کدهایی که در فایل
       گروه‌بندی فعلی نیستند، اطلاعاتشان را از آنجا می‌گیرد.
    2) ردیف‌هایی که قبلاً در رپورت قبل ثبت شده‌اند (عیب یا تولید) را تکرار
       نمی‌کند؛ یعنی اگر فایل جدید شما برای یک محصول تا 28 و برای محصول دیگر
       فقط تا 20 داده دارد، دقیقاً همان ردیف‌های جدیدی که گزارش نداشته‌ایم
       اضافه می‌شوند و بقیه نادیده گرفته می‌شوند.

نحوه اجرا:
  python build_qc_report.py \
      --quality "اطلاعات جامع کیفیت حین تولید.xlsx" \
      --defect  "گزارش عیب های سند بازرسی.xlsx" \
      --prod    "گزارش تعداد تولید به تفکیک سند عملکرد.xlsx" \
      --grouping "گروه بندی محصولات.xlsx" \
      --history "QC Report راهکاران - ماه قبل.xlsx" \
      --out "QC Report راهکاران - 1405/06.xlsx"

    --month 1405/06        (اختیاری) فقط ردیف‌های یک ماه خاص
    --validate فایل.xlsx   (اختیاری) مقایسه خروجی با یک رپورت موجود
    --dedup                (اختیاری) رکوردهای کاملاً تکراری (همهٔ ستون‌ها
                           یکسان) را جدا می‌کند: یکی نگه داشته می‌شود و
                           بقیه به‌عنوان تکراری در لاگ فهرست می‌شوند
    --gui                  حالت گرافیکی: فایل‌ها و مسیر خروجی از روی
                           پنجره انتخاب می‌شوند (بدون بقیهٔ پارامترها)
"""

import argparse
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# پیکربندی
# ---------------------------------------------------------------------------

# مقادیر ثابت ستون‌های پایانی (مطابق فایل «گزارش مرجع»)
CONST_COST = 44000           # هزینه ریالی تعمیرات
CONST_LIMIT = 14000          # حد قابل قبول
CONST_TARGET = 12000         # هدف
CONST_PART_TYPE = "فرآیندی"   # نوع قطعات تولیدی و تامینی

# کدهای ECU که عیب ECU-07 آن‌ها در QC EMS ثبت می‌شود (بقیه 320 در FULT EMS)
ECU_FINAL_CONTROL_CODES = {"3206133", "3206134"}
ECU_FINAL_CONTROL_DEFECT = "ECU-07"

# ایستگاه‌های سند بازرسی که عیب‌هایشان در FULT EMS می‌آیند
FULT_EMS_STATIONS = {
    "تست و کنترل ECU",
    "مونتاژ پدال گاز برقی",
    "تکمیل کاری ECU",
    "مونتاژ انتن و یونیت ایمو بلایزر",
}

# عنوان عملیات «QV» در سند بازرسی نسخهٔ تکراری همان عیب است
QV_TITLE = "QV"

# در نمونه، ایستگاه «وان قلع و کنترل ماشینی پس از وان» در ردیف‌های عیبِ
# «اطلاعات جامع کیفیت» به «وان قلع» کوتاه نوشته شده
STATION_RENAME_QUALITY = {"وان قلع و کنترل ماشینی پس از وان": "وان قلع"}

SHEETS = [
    # ترتیب شیت‌ها مطابق رپورت ماه قبل
    dict(
        name="FULT EMS",
        defect_source="defect",
        code_prefix=("320",),
        defect_filter="fult_ems",
        station="source",
        prod_prefix=("320",),
        prod_centers={"تکمیل کاری ECU", "تست و کنترل ECU"},
        defect_decl="-",
        prod_shift="",
        branch_const="EMS",       # برنچ شیت‌های EMS ثابت است
    ),
    dict(
        name="QC EMS",
        defect_source="defect",
        code_prefix=("331", "332"),
        defect_filter="qc_ems",   # عیب‌های 33 + ECU-07 دو کد ECU
        station="کنترل نهایی",
        prod_prefix=("331", "332"),
        prod_centers=None,        # همه مراکز
        defect_decl="-",
        prod_shift="-",
        branch_const="EMS",
    ),
    dict(
        name="ICT",
        defect_source="quality",
        code_prefix=("122",),
        defect_filter="ICT",
        station="ICT",
        prod_prefix=("122",),
        prod_centers={"تکمیل کاری نود ها"},
        defect_decl="-",
        prod_shift="-",
        prod_branch="-",                 # در نمونه برنچ ردیف‌های تولید ICT «-» است
        prod_requires_grouping=False,    # ICT همه ردیف‌های مرکز خودش را می‌گیرد
        prod_cnt="-",                    # تعداد ایراد ردیف‌های تولید ICT «-» است
        prod_repair_dash=True,           # هزینه/زمان‌های ردیف‌های تولید ICT «-» است
    ),
    dict(
        name="QC ELE",
        defect_source="quality",
        code_prefix=("123",),
        defect_filter="any",
        station="کنترل نهایی",
        prod_prefix=("123",),
        prod_centers={"کنترل نهایی خط D", "کنترل نهایی جلوآمپر ها",
                      "کنترل نهایی خط نود ها"},
        defect_decl="-",
        prod_shift="-",
    ),
    dict(
        name="SMD REPORT",
        defect_source="quality",
        code_prefix=("120",),
        defect_filter="any",
        station="-",
        prod_prefix=("120",),
        prod_centers={"سامسونگ 1", "سامسونگ 2", "ماشین میرایی"},
        defect_decl="کنترل ماشینی",
        prod_shift="-",
        extra_month_col=True,   # شیت SMD یک ستون «ماه» بعد از تاریخ دارد
        time_empty_dash=True,   # در شیت SMD، «رفع عیب»/«تست مجدد» خالی با «-» نوشته می‌شود
    ),
    dict(
        name="FULTELE",
        defect_source="defect",
        code_prefix=("121", "122"),
        defect_filter="non_qv",   # همه فرآیندهای غیر-QV (ICT_01 مستثنا)
        station="title",
        prod_prefix=("121", "122"),
        prod_centers={"مونتاژ دستی خطوط جلوآمپر", "مونتاژ دستی کلیدها",
                      "وان قلع و کنترل ماشینی پس از وان", "تکمیل کاری خط D",
                      "تکمیل کاری جلوآمپر", "تکمیل کاری نود ها"},
        defect_decl="-",
        prod_shift="-",
        # در نمونه شیفت ردیف‌های تولید مراکز «تکمیل کاری» خالی و بقیه «-» است
        prod_shift_tamamil_empty=True,
    ),
    dict(
        name="qv",
        defect_source="quality",
        code_prefix=("121",),
        defect_filter="any",
        station="source",
        prod_prefix=("121",),
        prod_centers={"مونتاژ دستی خطوط جلوآمپر", "مونتاژ دستی کلیدها",
                      "وان قلع و کنترل ماشینی پس از وان"},
        defect_decl="کنترل چشمی",
        prod_shift="-",
    ),
]

# ---------------------------------------------------------------------------
# ابزارها
# ---------------------------------------------------------------------------


def s(v):
    return "" if v is None else str(v).strip()


def n(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return 0
    return int(f) if f == int(f) else f


def norm_dc(x):
    """نرمال‌سازی کد عیب: حذف پرانتزها (مثل (ف) و (01)) و نقطهٔ آخر"""
    return re.sub(r"\(.*?\)", "", s(x)).rstrip(".").strip()


def date_key(d):
    d = s(d)
    parts = d.replace("٠", "0").split("/")
    try:
        return tuple(int(p) for p in parts)
    except ValueError:
        return (9999, 99, 99)


def load_sheet(path, sheet=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    hdr = [str(h).strip() if h is not None else f"col{i}" for i, h in enumerate(rows[0])]
    out = []
    for r in rows[1:]:
        if not any(v is not None for v in r):
            continue
        out.append(dict(zip(hdr, r)))
    return out


# ---------------------------------------------------------------------------
# جدول گروه‌بندی
# ---------------------------------------------------------------------------

GRP_KEYS = ["کد راهکاران", "نام محصول", "خانواده محصول",
            "نام محصول ترکیبی", "گروه محصول نهایی", "برنچ"]
GRP_CODE_COL = "کد راهکاران"
GRP_CODE_COL_ALT = "کد گروه محصول"   # در نسخهٔ جدیدتر جدول، نام ستون کد این است


def load_grouping(path):
    """خواندن جدول گروه‌بندی:
      ردیف هدر از بین ۱۵ ردیف اول هر شیت پیدا می‌شود (پس ردیف عنوان/خالی
      بالای جدول و جابه‌جایی جدول بین شیت‌ها خرابی نمی‌سازد).
      ستون کد: «کد راهکاران» یا «کد گروه محصول».
      برگشت: (ردیف‌ها, هدرها, توضیح)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        allr = list(ws.iter_rows(values_only=True))
        for hi, row in enumerate(allr[:15]):
            hdr = [str(h).strip() if h is not None else "" for h in row]
            code_col = next((c for c in (GRP_CODE_COL, GRP_CODE_COL_ALT)
                             if c in hdr), None)
            if code_col:
                out = []
                for r in allr[hi + 1:]:
                    if not any(v is not None for v in r):
                        continue
                    d = dict(zip(hdr, r))
                    if code_col != GRP_CODE_COL:
                        d.setdefault(GRP_CODE_COL, d.get(code_col))
                    out.append(d)
                return out, hdr, f"شیت «{ws.title}»، ردیف {hi + 1}"
    # ستون پیدا نشد — هدر شیت اول را برگردان تا بتواند تشخيص داده شود
    first = wb.worksheets[0]
    row1 = next(first.iter_rows(min_row=1, max_row=1, values_only=True), None) or ()
    hdr = [str(h).strip() if h is not None else "" for h in row1]
    found = ", ".join(h for h in hdr if h)[:180]
    return [], hdr, (f"ستون «{GRP_CODE_COL}» یا «{GRP_CODE_COL_ALT}» پیدا نشد "
                     f"(هدرهای شیت اول: {found or '—'})")


def build_grouping(grp_rows, history_grp_rows):
    """
    ساخت جدول گروه‌بندی:
      - فایل گروه‌بندی فعلی اولویت دارد (اولین ردیف تکراری، مثل VLOOKUP)
      - کدهای موجود در رپورت قبلی ولی نه در فایل فعلی، از رپورت قبلی می‌آیند
    """
    grp = {}
    for rows in (grp_rows, history_grp_rows or []):
        for r in rows:
            code = s(r.get("کد راهکاران"))
            if not code or code in grp:
                continue
            grp[code] = {
                "site_name": s(r.get("نام محصول")),
                "family": s(r.get("خانواده محصول")),
                "combo": s(r.get("نام محصول ترکیبی")),
                "final": s(r.get("گروه محصول نهایی")),
                "branch": s(r.get("برنچ")),
            }
    return grp


def read_history(path):
    """خواندن جدول گروه‌بندی و ردیف‌های ثبت‌شده از رپورت ماه قبل"""
    wb = openpyxl.load_workbook(path, data_only=True)

    # 1) شیت گروه‌بندی: شیتی که هدرش «کد راهکاران» دارد
    grp_rows = []
    for ws in wb.worksheets:
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if first and any(s(h) == "کد راهکاران" for h in first):
            grp_rows = load_sheet(path, ws.title)
            break

    # 2) ردیف‌های ثبت‌شده در هر شیت رپورت (برای حذف تکراری‌ها)
    reported = {}
    for cfg in SHEETS:
        sheet_name = cfg["name"] + (" " if cfg["name"] == "qv" else "")
        rows = sheet_rows_raw(path, sheet_name)
        defect_keys = Counter()
        prod_keys = Counter()
        for r in rows:
            vals = [s(v) for v in r]
            off = 1 if len(vals) > 54 else 0   # شیت SMD یک ستون «ماه» اضافه دارد
            date = vals[0]
            code = vals[2 + off]
            dc = vals[9 + off]
            cnt = n(vals[11 + off])
            total = n(vals[12 + off])
            if dc not in ("", "-"):
                defect_keys[(date, code, norm_dc(dc), cnt)] += 1
            else:
                prod_keys[(date, code, total)] += 1
        reported[cfg["name"]] = (defect_keys, prod_keys)
    return grp_rows, reported


def sheet_rows_raw(path, name):
    wb = openpyxl.load_workbook(path, data_only=True)
    for sh in wb.sheetnames:
        if sh.strip() == name.strip():
            ws = wb[sh]
            return list(ws.iter_rows(values_only=True))[1:]
    return []


# ---------------------------------------------------------------------------
# ساخت ردیف‌ها
# ---------------------------------------------------------------------------

BASE_HEADERS = [
    "تاریخ", "شیفت کاری", "کد گروه محصول", "کد محصول", "نام محصول",
    "خانواده محصول", "نام محصول ترکیبی", "گروه محصول نهایی", "ایستگاه",
    "کد ایراد", "شرح ایراد", "تعداد ایراد", "تعداد کل", "برنچ",
    "شرح فعالیت انجام شده توسط تعمیرات", "عامل مسبب ایراد 6M", "کد قطعه",
    "نام قطعه", "نام تامین کننده", "کد و نام تجهیزات و ابزارآلات",
    "کد تجهیزات و ابزارآلات", "کد و نام تجهیزات و ابزارآلات1",
    "کد فرآیند (OPC)", "نام فرآیند (OPC)", "حالت خرابی بالقوه",
    "نوع حالت خرابی بالقوه", "توضیحات تعمیرات", "ستون مشترک حالت خرابی",
    "شدت حالت خرابی", "تشخیص حالت خرابی", "وقوع قدیمی",
    "مدت زمان عیب یابی\n( دقیقه )", "مدت زمان رفع عیب\n( دقیقه )",
    "مدت زمان تست مجدد\n( دقیقه )", "هزینه ریالی تعمیرات", "خانواده قطعات",
    "شماره BOM", "شماره برگ ارسال", "کد اپراتور", "نام اپراتور",
    "کد بازرس چشمی", "نام بازرس", "AOI", "نام بازرس AOI",
    "نوع قطعات تولیدی و تامینی", "وضعیت اعلام ایرادات", "توضیحات اضافه",
    "توضیحات", "Control Station", "حد قابل قبول", "هدف",
    "علت افت یا بهبود PPM", "حد قابل قبول عظام", "هدف عظام",
]


# شیت SMD در قالب کاربر 68 ستونه است: علاوه بر ستون «ماه»، ستون
# «جانمایی قطعه معیوب در فرآیند SMD» بعد از «نام قطعه» و ۱۳ ستون
# مخصوص SMD (اپراتور/ماشین‌ها) در انتها دارد (مطابق رپورت ماه قبل)
SMD_HEADERS = [
    "تاریخ", "ماه", "شیفت کاری", "کد گروه محصول", "کد محصول", "نام محصول",
    "خانواده محصول", "نام محصول ترکیبی", "گروه محصول نهایی", "ایستگاه",
    "کد ایراد", "شرح ایراد", "تعداد ایراد", "تعداد کل", "برنچ",
    "شرح فعالیت انجام شده توسط تعمیرات", "عامل مسبب ایراد 6M", "کد قطعه",
    "نام قطعه", "جانمایی قطعه معیوب در فرآیند SMD", "نام تامین کننده",
    "کد و نام تجهیزات و ابزارآلات", "کد تجهیزات و ابزارآلات",
    "کد و نام تجهیزات و ابزارآلات1", "کد فرآیند (OPC)", "نام فرآیند (OPC)",
    "حالت خرابی بالقوه", "نوع حالت خرابی بالقوه", "توضیحات تعمیرات",
    "ستون مشترک حالت خرابی", "شدت حالت خرابی", "تشخیص حالت خرابی",
    "وقوع قدیمی", "مدت زمان عیب یابی\n( دقیقه )", "مدت زمان رفع عیب\n( دقیقه )",
    "مدت زمان تست مجدد\n( دقیقه )", "هزینه ریالی تعمیرات", "خانواده قطعات",
    "شماره BOM", "شماره برگ ارسال", "کد اپراتور", "نام اپراتور",
    "کد بازرس چشمی", "نام بازرس", "AOI", "نام بازرس AOI",
    "نوع قطعات تولیدی و تامینی", "وضعیت اعلام ایرادات", "توضیحات اضافه",
    "توضیحات", "Control Station", "حد قابل قبول", "هدف",
    "علت افت یا بهبود PPM", "کد اپراتور\nچشمی SMD", "نام اپراتور\nچشمی SMD",
    "کد اپراتور\nPaste SMD", "نام اپراتور\nPaste SMD", "ماشین مونتاژ SMD",
    "ماشین \nPaste SMD", "ماشین\nOven SMD", "Control Station",
    "نام اپراتور بازرس SMD", "کد \nکنترلر بازرس SMD", "شماره \nBOM SMD",
    "توضیحات", "حد قابل قبول عظام", "هدف عظام",
]


def headers_for(cfg):
    if cfg.get("extra_month_col"):
        return list(SMD_HEADERS)
    return list(BASE_HEADERS)


def to_smd_row(row, placement, center):
    """تبدیل ردیف 55 ستونه (با ماه) به قالب 68 ستونه شیت SMD"""
    r = list(row)
    is_prod = s(r[10]) in ("", "-")
    out = r[:19]                    # 0-18
    out.append(s(placement) or "-") # 19 جانمایی
    out += r[19:37]                 # 20-37: تامین‌کننده تا خانواده قطعات
    if is_prod:
        out[36] = 0                 # هزینه ردیف‌های تولید SMD: 0
    out += [0, 0]                   # 38-39: BOM / برگ ارسال
    out += r[39:45]                 # 40-45: اپراتور و بازرس
    out += [r[45], r[46], r[47], r[48]]  # 46-49: نوع قطعات تا توضیحات
    out.append(r[46])               # 50: Control Station = وضعیت اعلام
    out += [r[50], r[51], r[52]]    # 51-53: حد / هدف / PPM
    out += ["-", "-", "-", "-", "-"]  # 54-58: اپراتور و ماشین مونتاژ
    out.append(s(center) or "-")    # 59: ماشین Paste SMD = مرکز کاری
    out += ["-", "-", "-", "-"]     # 60-63: اوون / Control Station / بازرس
    out.append(0)                   # 64: BOM SMD
    out.append("-")                 # 65: توضیحات
    out += [r[53], r[54]]           # 66-67: حد و هدف عظام
    # 33=عیب یابی, 34=رفع عیب, 35=تست مجدد
    if is_prod:
        out[33] = out[34] = out[35] = 0
    else:
        if out[33] is None:
            out[33] = 0
        if out[34] is None:
            out[34] = "-"
        if out[35] is None:
            out[35] = "-"
    return out


def insert_month(row_list, month_value="-"):
    return [row_list[0], month_value] + row_list[1:]


def make_row(cfg, grp, date, shift, code, prod_name, station,
             dc, desc, cnt, total, extra):
    g = grp.get(code)
    if g is not None:
        # کد در جدول گروه‌بندی پیدا شد (مثل VLOOKUP: مقدار خالی می‌ماند خالی)
        name = g["site_name"]
        family = g["family"]
        combo = g["combo"]
        final = g["final"]
        branch = g["branch"]
    else:
        # کد پیدا نشد → #N/A (همان رفتار VLOOKUP اکسل)
        name = family = combo = final = "#N/A"
        branch = "#N/A"

    if cfg.get("branch_const"):
        branch = cfg["branch_const"]

    row = [date, shift, code, prod_name, name, family, combo, final,
           station, dc, desc, cnt, total, branch]

    ex = extra or {}
    # شیت SMD خالی بودن «رفع عیب»/«تست مجدد» را با «-» نشان می‌دهد؛
    # برای آن خالی بودن (None) حفظ می‌شود و در to_smd_row تبدیل می‌شود
    is_smd = bool(cfg.get("extra_month_col"))
    # در برخی شیت‌ها (ICT) هزینه و زمان‌های تعمیرات ردیف‌های تولید «-» است
    repair_dash = cfg.get("prod_repair_dash") and s(dc) in ("", "-")

    def tv(key):
        if repair_dash:
            return "-"
        v = ex.get(key)
        if v is not None:
            return v
        return 0 if not is_smd else None

    row += [
        ex.get("activity") or "-",
        ex.get("m6") or "-",
        ex.get("part_code") or "-",
        ex.get("part_name") or "-",
        ex.get("supplier") or "-",
        "-",
        "-",
        ex.get("equipment") or "-",
        ex.get("opc_code") or "-",
        ex.get("opc_name") or "-",
        ex.get("latent") or "-",
        ex.get("latent_type") or "-",
        ex.get("repair_note") or "-",
        "-",
        ex.get("severity", 0),
        ex.get("detection", 0),
        ex.get("occurrence", 0),
        tv("t_time"),
        tv("f_time"),
        tv("r_time"),
        CONST_COST if not repair_dash else "-",
        ex.get("part_family") or "-",
        "-",
        "-",
        "-",
        ex.get("operator") or "-",
        "-",
        ex.get("inspector") or "-",
        "-",
        "-",
        CONST_PART_TYPE,
        cfg.get("defect_decl", "-"),
        "-",
        "-",
        "-",
        CONST_LIMIT,
        CONST_TARGET,
        "-",
        CONST_LIMIT,
        CONST_TARGET,
    ]
    return row


def quality_extra(r):
    def g(k):
        return s(r.get(k)) or None

    def t(k):
        # مقدار عددی؛ خالی بودن None می‌ماند تا شیت‌ها بتوانند تصمیم بگیرند
        v = r.get(k)
        if v is None or s(v) in ("", "-"):
            return None
        return n(v)

    return {
        "activity": g("شرح فعالیت انجام شده توسط تعمیرات"),
        "m6": g("Mعامل مسبب 6"),
        "part_code": g("کد قطعه"),
        "part_name": g("نام قطعه"),
        "supplier": g("نام تامین کننده"),
        "equipment": g("کد و نام ابزارآلات و تجهیزات"),
        "opc_code": g("کد فرایند"),
        "opc_name": g("نام فرآیند OPC"),
        "latent": g("حالت خرابی بلقوه"),
        "latent_type": g("نوع حالت خرابی بلقوه"),
        "repair_note": g("توضیحات تعمیرات"),
        "severity": n(r.get("شدت حالت خرابی")),
        "detection": n(r.get("تشخیص حالت خرابی")),
        "occurrence": n(r.get("وقوع قدیمی")),
        "t_time": t("مدت زمان عیب یابی دقیقه"),
        "f_time": t("مدت زمان رفع عیب دقیقه"),
        "r_time": t("مدت زمان تست مجدد دقیقه"),
        "part_family": g("خانواده قطعات"),
        "operator": g("اپراتور مسبب ایراد"),
        "inspector": g("بازرس مسبب ایراد"),
        "placement": g("جانمایی"),
    }


# ---------------------------------------------------------------------------
# فیلترها
# ---------------------------------------------------------------------------


def in_prefix(code, prefixes):
    code = s(code)
    return any(code.startswith(p) for p in prefixes)


def code_allowed(cfg, code):
    if in_prefix(code, cfg["code_prefix"]):
        return True
    # کدهای ECU با ECU-07 به QC EMS می‌روند هرچند پیشوند 320 دارند
    if cfg["defect_filter"] == "qc_ems" and s(code) in ECU_FINAL_CONTROL_CODES:
        return True
    return False


def defect_kept(cfg, dc_norm, source_row):
    f = cfg["defect_filter"]
    if f == "any":
        return True
    if f == "ICT":
        return dc_norm.upper().startswith("ICT")
    if f == "non_qv":
        if dc_norm.upper() == "ICT_01":
            return False
        if s(source_row.get("title")) == "ICT":
            return False
        return True
    if f == "fult_ems":
        if s(source_row["station"]) not in FULT_EMS_STATIONS:
            return False
        if (s(source_row["code"]) in ECU_FINAL_CONTROL_CODES
                and dc_norm == ECU_FINAL_CONTROL_DEFECT):
            return False
        return True
    if f == "qc_ems":
        if (s(source_row["code"]) in ECU_FINAL_CONTROL_CODES
                and dc_norm == ECU_FINAL_CONTROL_DEFECT):
            return True
        return in_prefix(source_row["code"], cfg["code_prefix"])
    return True


def prod_station_for(cfg, center):
    """ایستگاه ردیف‌های تولید:
      - شیت‌های ایستگاه ثابت (ICT، SMD، QC ELE، QC EMS): همان مقدار ثابت
      - «source» (qv و FULT EMS): نام مرکز کاری همان‌طور که هست
      - «title» (FULTELE): نام مرکز کاری با همان کوتاه‌نویسی ایستگاه (وان قلع)"""
    st = cfg["station"]
    if st not in ("source", "title"):
        return st
    c = s(center) or "-"
    if st == "title":
        c = STATION_RENAME_QUALITY.get(c, c)
    return c


def station_for(cfg, source_row, from_quality):
    st = cfg["station"]
    if st == "source":
        val = s(source_row["station"]) or "-"
        if from_quality:
            val = STATION_RENAME_QUALITY.get(val, val)
        return val
    if st == "title":
        title = s(source_row.get("title", ""))
        if title and title not in ("تکمیل کاری", QV_TITLE, "ICT"):
            return title
        return s(source_row.get("station", "")) or "-"
    return st


# ---------------------------------------------------------------------------
# تولید شیت‌ها
# ---------------------------------------------------------------------------


def build_sheet(cfg, quality, defect, prod, grp, history, month, dedup=False):
    rows = []
    seq = 0
    hist_defect, hist_prod = (history.get(cfg["name"]) or (Counter(), Counter()))
    skipped = 0
    dropped_nogroup = 0

    month_key = (int(month[:4]), int(month[5:7])) if month else None

    # ---- ردیف‌های عیب
    src = quality if cfg["defect_source"] == "quality" else defect
    if cfg["defect_source"] == "defect":
        # نسخهٔ غیر-QV را نگه می‌داریم؛ اگر نبود، نسخهٔ QV
        has_non_qv = set()
        for r in src:
            if s(r.get("کد ایراد")) and s(r.get("عنوان عملیات آزمایش")) != QV_TITLE:
                has_non_qv.add((s(r.get("تاریخ سفارش")), s(r.get("کد محصول")),
                                norm_dc(r.get("کد ایراد"))))
        src = [r for r in src
               if s(r.get("عنوان عملیات آزمایش")) != QV_TITLE
               or (s(r.get("تاریخ سفارش")), s(r.get("کد محصول")),
                   norm_dc(r.get("کد ایراد"))) not in has_non_qv]

    for r in src:
        code = s(r.get("کد محصول") or r.get("کد کالا"))
        if not code or not code_allowed(cfg, code):
            continue
        raw_dc = s(r.get("کد ایراد") or r.get("کد عیب"))
        if not raw_dc:
            continue
        dc = norm_dc(raw_dc)
        date = s(r.get("تاریخ سفارش"))
        if month_key and date_key(date)[:2] != month_key:
            continue
        source_row = {
            "code": code,
            "station": r.get("ایستگاه"),
            "title": r.get("عنوان عملیات آزمایش", ""),
        }
        if not defect_kept(cfg, dc, source_row):
            continue
        desc = s(r.get("شرح ایراد") or r.get("شرح عیب")) or "-"
        cnt = n(r.get("تعداد ایراد") or r.get("تعداد عیب مربوطه") or r.get("تعداد عیب"))
        # ردیفی که قبلاً ثبت شده تکرار نمی‌شود
        hkey = (date, code, dc, cnt)
        if hist_defect.get(hkey, 0) > 0:
            hist_defect[hkey] -= 1
            skipped += 1
            continue
        shift = s(r.get("شیفت")) if cfg["defect_source"] == "defect" else "-"
        extra = quality_extra(r) if cfg["defect_source"] == "quality" else None
        row = make_row(cfg, grp, date, shift, code,
                       s(r.get("نام محصول") or r.get("نام کالا")),
                       station_for(cfg, source_row, cfg["defect_source"] == "quality"),
                       dc, desc, cnt, 0, extra)
        if cfg.get("extra_month_col"):
            row = insert_month(row)
            row = to_smd_row(row, extra.get("placement") if extra else None,
                             s(r.get("ایستگاه")))
        rows.append((code, date_key(date), seq, row))
        seq += 1

    # ---- ردیف‌های تولید
    for r in prod:
        code = s(r.get("کد کالا"))
        if not code or not in_prefix(code, cfg["prod_prefix"]):
            continue
        center = s(r.get("مرکز کاری"))
        if cfg["prod_centers"] is not None and center not in cfg["prod_centers"]:
            continue
        if cfg.get("prod_requires_grouping", True) and code not in grp:
            dropped_nogroup += 1
            continue
        date = s(r.get("تاریخ تولید"))
        if month_key and date_key(date)[:2] != month_key:
            continue
        total = n(r.get("مقدار سالم"))
        hkey = (date, code, total)
        if hist_prod.get(hkey, 0) > 0:
            hist_prod[hkey] -= 1
            skipped += 1
            continue
        station_val = prod_station_for(cfg, center)
        pshift = cfg["prod_shift"]
        if cfg.get("prod_shift_tamamil_empty") and station_val.startswith("تکمیل کاری"):
            pshift = ""
        row = make_row(cfg, grp, date, pshift, code,
                       s(r.get("نام کالا")), station_val,
                       "-", "-", cfg.get("prod_cnt", 0), total, None)
        if cfg.get("prod_branch") == "-":
            row[13] = "-"
        if cfg.get("extra_month_col"):
            row = insert_month(row)
            row = to_smd_row(row, None, center)
        rows.append((code, date_key(date), seq, row))
        seq += 1

    order = {}
    for i, (code, dk, seq_i, row) in enumerate(rows):
        order.setdefault(code, len(order))
    rows.sort(key=lambda t: (order[t[0]], t[1], t[2]))
    data = [t[3] for t in rows]

    # گزینهٔ حذف تکراری کامل: رکوردی که همهٔ ستون‌هایش با رکورد(های)
    # دیگر یکسان است → یکی نگه داشته می‌شود و بقیه جدا می‌شوند
    dupes = []
    if dedup:
        seen = set()
        kept = []
        for row in data:
            key = tuple(row)
            if key in seen:
                dupes.append(row)
                continue
            seen.add(key)
            kept.append(row)
        data = kept
    return data, skipped, dropped_nogroup, dupes


def style_sheet(ws, ncols):
    for c in range(1, ncols + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.sheet_view.rightToLeft = True


def norm_val(v):
    t = s(v)
    if re.fullmatch(r"-?\d+\.0", t):
        t = t[:-2]
    return t


def sheet_rows(path, name):
    wb = openpyxl.load_workbook(path, data_only=True)
    for sh in wb.sheetnames:
        if sh.strip() == name.strip():
            ws = wb[sh]
            rows = list(ws.iter_rows(values_only=True))
            return [tuple(norm_val(v) for v in r) for r in rows[1:]
                    if any(v is not None for v in r)]
    return []


def validate(gen_path, ref_path, log=print):
    log("\n========== مقایسه با رپورت مرجع ==========")
    for cfg in SHEETS:
        name = cfg["name"] + (" " if cfg["name"] == "qv" else "")
        gen = sheet_rows(gen_path, name)
        ref = sheet_rows(ref_path, name)
        g = [r[:14] for r in gen]
        r_ = [r[:14] for r in ref]
        cg, cr = Counter(g), Counter(r_)
        only_gen = cg - cr
        only_ref = cr - cg
        both = cg & cr
        log(f"  {name:>12}: مشترک {sum(both.values()):>4} | "
            f"فقط خروجی {sum(only_gen.values()):>3} | فقط مرجع {sum(only_ref.values()):>3}")
        for k, v in list(only_ref.most_common(4)):
            log(f"        - فقط مرجع : {k[0]} | {k[2]} | {k[8]} | {k[9]} | cnt={k[11]} | tot={k[12]}")
        for k, v in list(only_gen.most_common(4)):
            log(f"        + فقط خروجی: {k[0]} | {k[2]} | {k[8]} | {k[9]} | cnt={k[11]} | tot={k[12]}")


def run_build(quality, defect, prod, grouping, history=None, month=None,
              out=None, validate_path=None, dedup=False, log=print):
    """اجرای ساخت رپورت؛ log تابع خروجی‌دهنده (در حالت گرافیکی به لاگ فرستاده می‌شود)"""
    quality_rows = load_sheet(quality)
    defect_rows = load_sheet(defect)
    prod_rows = load_sheet(prod)
    grp_rows, grp_hdr, grp_info = load_grouping(grouping)

    history_grp = []
    hist_reported = None
    if history:
        log(f"خواندن رپورت قبل: {history}")
        history_grp, hist_reported = read_history(history)

    grp = build_grouping(grp_rows, history_grp)
    cur_codes = {s(r.get("کد راهکاران")) for r in grp_rows if s(r.get("کد راهکاران"))}
    log(f"جدول گروه‌بندی: {len(grp)} کد (فایل فعلی: {len(cur_codes)}; {grp_info}"
        + (f" + رپورت قبل: {len(grp) - len(cur_codes)}" if len(grp) > len(cur_codes) else "") + ")")
    if cur_codes:
        log(f"  نمونه کدها: {', '.join(list(cur_codes)[:3])}")
    missing_cols = [c for c in ("نام محصول", "خانواده محصول", "نام محصول ترکیبی",
                                "گروه محصول نهایی", "برنچ") if c not in grp_hdr]
    if missing_cols:
        log(f"  ⚠ ستون‌های موجود نبود در جدول گروه‌بندی: {', '.join(missing_cols)}")
    if not grp:
        log("  ⚠ هشدار: هیچ کدی از جدول گروه‌بندی خوانده نشد؛ همه ردیف‌های رپورت "
            "#N/A می‌شوند. مطمئن شوید ستونی با هدر «کد راهکاران» یا "
            "«کد گروه محصول» در جدول هست و کدهای آن با «کد کالا» "
            "فایل‌های سایت یکی‌اند (مثل 3206133).")

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    for cfg in SHEETS:
        data, skipped, nogroup, dupes = build_sheet(cfg, quality_rows, defect_rows,
                                                    prod_rows, grp, hist_reported or {},
                                                    month, dedup=dedup)
        title = cfg["name"] + (" " if cfg["name"] == "qv" else "")
        ws = out_wb.create_sheet(title=title)
        ws.append(headers_for(cfg))
        for row in data:
            ws.append(row)
        style_sheet(ws, len(headers_for(cfg)))
        ndef = sum(1 for row in data
                   if s(row[9] if not cfg.get("extra_month_col") else row[10])
                   not in ("", "-"))
        extra = f" | تکراری‌های حذف‌شده: {skipped}" if skipped else ""
        if nogroup:
            extra += f" | بدون گروه‌بندی: {nogroup}"
        if dupes:
            extra += f" | ردیف تکراری کامل: {len(dupes)}"
        log(f"  شیت {title:>12}: {len(data):>4} ردیف (عیب: {ndef} | "
            f"تولید: {len(data)-ndef}){extra}")
        if dupes:
            off = 1 if cfg.get("extra_month_col") else 0
            for row in dupes[:8]:
                log(f"      - جدا شد به‌عنوان تکراری: {s(row[0])} | "
                    f"کد={s(row[2 + off])} | عیب={s(row[9 + off])} | cnt={s(row[11 + off])}")
            if len(dupes) > 8:
                log(f"      - ...و {len(dupes) - 8} تکراری کامل دیگر")

    ws = out_wb.create_sheet(title="گروه بندی محصولات")
    if grp_rows:
        hdr = list(grp_rows[0].keys())
        ws.append(hdr)
        for r in grp_rows:
            ws.append([r.get(k) for k in hdr])
        style_sheet(ws, len(hdr))

    out_wb.save(out)
    log(f"\nخروجی ذخیره شد: {out}")

    if validate_path:
        validate(out, validate_path, log=log)


def run_gui():
    """حالت گرافیکی: بارگذاری فایل‌ها و انتخاب مسیر خروجی"""
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox

    app = tk.Tk()
    app.title("سازندهٔ رپورت QC راهکاران")
    app.geometry("720x680")

    tk.Label(app, text="فایل‌ها را انتخاب کنید و روی «تولید رپورت» بزنید",
             font=("", 11, "bold")).pack(pady=(10, 2))
    tk.Label(app, text="رپورت ماه قبل و ماه، اختیاری‌اند", fg="#555").pack(pady=(0, 6))

    fields = {}
    labels = {}

    def add_file(label, key, save=False):
        row = tk.Frame(app)
        row.pack(fill="x", padx=12, pady=3)
        ent = tk.Entry(row)
        ent.pack(side="right", fill="x", expand=True)
        btn = tk.Button(row, text="انتخاب...", width=10)
        btn.pack(side="left", padx=(0, 6))
        tk.Label(row, text=label, width=34, anchor="e").pack(side="left")

        def pick():
            if save:
                m = month_ent.get().strip()
                name = (f"QC Report راهکاران - {m}.xlsx" if m
                        else "QC Report راهکاران.xlsx")
                path = filedialog.asksaveasfilename(
                    title="فایل خروجی", defaultextension=".xlsx",
                    initialfile=name, filetypes=[("فایل اکسل", "*.xlsx")])
            else:
                path = filedialog.askopenfilename(
                    title=label, filetypes=[("فایل اکسل", "*.xlsx *.xlsm"),
                                             ("همه فایل‌ها", "*.*")])
            if path:
                ent.delete(0, "end")
                ent.insert(0, path)
        btn.config(command=pick)
        fields[key] = ent
        labels[key] = label
        return ent

    add_file("اطلاعات جامع کیفیت حین تولید:", "quality")
    add_file("گزارش عیب های سند بازرسی:", "defect")
    add_file("گزارش تعداد تولید به تفکیک سند عملکرد:", "prod")
    add_file("گروه بندی محصولات:", "grouping")
    add_file("رپورت ماه قبل (برای ادامه بدون تکرار):", "history")

    mrow = tk.Frame(app)
    mrow.pack(fill="x", padx=12, pady=3)
    month_ent = tk.Entry(mrow, width=14)
    month_ent.pack(side="right")
    tk.Label(mrow, text="ماه (مثلا 1405/06):", width=34, anchor="e").pack(side="left")

    dedup_var = tk.BooleanVar(value=False)
    drow = tk.Frame(app)
    drow.pack(fill="x", padx=12, pady=3)
    tk.Checkbutton(drow, variable=dedup_var,
                   text="ردیف‌های کاملاً تکراری را جدا کن (همهٔ ستون‌ها یکسان؛ "
                        "یکی نگه داشته می‌شود، بقیه در لاگ فهرست می‌شوند)"
                   ).pack(side="right")

    add_file("فایل خروجی:", "out", save=True)

    run_btn = tk.Button(app, text="تولید رپورت", font=("", 12, "bold"),
                        command=None)
    run_btn.pack(pady=12, ipadx=24, ipady=4)

    log_frame = tk.Frame(app)
    log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 6))
    log_text = tk.Text(log_frame, height=14, state="disabled", wrap="word")
    sb = tk.Scrollbar(log_frame, command=log_text.yview)
    log_text.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    log_text.pack(side="left", fill="both", expand=True)

    status = tk.Label(app, text="", fg="#0a600a")
    status.pack(pady=(0, 8))

    q = queue.Queue()
    state = {"running": False}

    def append_log(msg):
        log_text.config(state="normal")
        log_text.insert("end", str(msg) + "\n")
        log_text.see("end")
        log_text.config(state="disabled")

    def set_status(msg, color="#0a600a"):
        status.config(text=msg, fg=color)

    def start():
        if state["running"]:
            return
        month = month_ent.get().strip()
        missing = [labels[k] for k in ("quality", "defect", "prod", "grouping", "out")
                   if not fields[k].get().strip()]
        if missing:
            messagebox.showwarning(
                "اطلاعات ناقص",
                "این فیلدها خالی‌اند:\n" + "\n".join(missing))
            return
        out = fields["out"].get().strip()
        if not out.lower().endswith((".xlsx", ".xlsm")):
            out += ".xlsx"
        state["running"] = True
        run_btn.config(state="disabled")
        set_status("در حال تولید...", "#1a4f9c")
        log_text.config(state="normal")
        log_text.delete("1.0", "end")
        log_text.config(state="disabled")

        def worker():
            try:
                run_build(
                    quality=fields["quality"].get().strip(),
                    defect=fields["defect"].get().strip(),
                    prod=fields["prod"].get().strip(),
                    grouping=fields["grouping"].get().strip(),
                    history=fields["history"].get().strip() or None,
                    month=month or None,
                    out=out,
                    dedup=bool(dedup_var.get()),
                    log=lambda m: q.put(m),
                )
                q.put(("__done__", out))
            except Exception as exc:
                q.put(("__error__", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def poll():
        try:
            while True:
                item = q.get_nowait()
                if isinstance(item, tuple):
                    tag, msg = item
                    state["running"] = False
                    run_btn.config(state="normal")
                    if tag == "__done__":
                        set_status(f"تمام شد: {msg}", "#0a600a")
                    else:
                        set_status("خطا!", "#b00020")
                        messagebox.showerror("خطا", msg)
                else:
                    append_log(item)
        except queue.Empty:
            pass
        app.after(100, poll)

    run_btn.config(command=start)
    poll()
    app.mainloop()


def main():
    cli_args = sys.argv[1:]
    # اجرا بدون هیچ پارامتری (مثلاً دابل‌کلیک) یا با --gui → حالت گرافیکی
    if not cli_args or "--gui" in cli_args:
        run_gui()
        return
    ap = argparse.ArgumentParser(description="سازندهٔ رپورت QC راهکاران")
    ap.add_argument("--quality", required=True)
    ap.add_argument("--defect", required=True)
    ap.add_argument("--prod", required=True)
    ap.add_argument("--grouping", required=True)
    ap.add_argument("--history", default=None,
                    help="رپورت ماه قبل؛ برای جدول گروه‌بندی و حذف تکراری‌ها")
    ap.add_argument("--month", default=None, help="مثلا 1405/06 (اختیاری)")
    ap.add_argument("--out", required=True)
    ap.add_argument("--validate", default=None)
    ap.add_argument("--dedup", action="store_true",
                    help="رکوردهای کاملاً تکراری (همهٔ ستون‌ها یکسان) را جدا کند؛ "
                         "یکی نگه داشته می‌شود")
    ap.add_argument("--gui", action="store_true",
                    help="حالت گرافیکی (بدون بقیهٔ پارامترها اجرا شود)")
    args = ap.parse_args()

    run_build(quality=args.quality,
              defect=args.defect,
              prod=args.prod,
              grouping=args.grouping,
              history=args.history,
              month=args.month,
              out=args.out,
              validate_path=args.validate,
              dedup=args.dedup)


if __name__ == "__main__":
    main()