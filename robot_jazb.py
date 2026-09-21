# -*- coding: utf-8 -*-
"""
ربات گزارش جذب  ----  robot_jazb.py  (هستهٔ محاسبات)

شیت «لیست جذب» را می‌خواند و یک فایل اکسل خروجی با جدول‌های خلاصه و
نمودارهای داخلی اکسل می‌سازد:
    - نمودار میله‌ای برای «تعداد»ها (نفرات / جذب / پست / واحد / پروژه / ماه)
    - نمودار دایره‌ای برای «درصد»ها (نرخ جذب)
    - فاصلهٔ «تاریخ درخواست» تا «شروع به کار» برای هر نفر (زمان جذب)

خط فرمان:
    python robot_jazb.py "مسیر فایل ورودی.xlsx"
    python robot_jazb.py            (جدیدترین xlsx کنار اسکریپت)
پنجره (UI):  robot_jazb_ui.py  یا  جذب.exe
"""

import sys, os, glob, re, datetime, tempfile, statistics
from collections import Counter

import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.drawing.text import RegularTextRun
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (Paragraph, ParagraphProperties,
                                   CharacterProperties, RichTextProperties)
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter

# ---- پالت رنگ گزارش ----
NAVY, BLUE, TEAL, AMBER = "1F4E78", "2E75B6", "2FB4A6", "F2A900"
LIGHT, PAPER, INK = "EAF1F8", "F7F9FC", "33404A"
BAND = "EEF3F9"        # رنگ ردیف‌های یک‌درمیانِ جدول‌ها

# پالت دسته‌ای با رنگ‌های کاملاً متفاوت (برای میله‌ها و تکه‌های دایره)
PALETTE = ["2E75B6", "E15759", "59A14F", "F28E2B", "8E44AD",
           "17BECF", "E377C2", "8C564B", "BAB0AC", "EDC948",
           "1F77B4", "D62728", "2CA02C", "FF7F0E", "9467BD"]
PIE_COLORS = PALETTE

# ستون‌های helperِ مخفیِ نمودارها روی شیت‌های تفکیکی (T..X) — خارج از ناحیهٔ چاپ
HB_CAT, HB_VAL = 20, 21      # نمودار میله‌ای: «نام (عدد)» و عدد
HP_CAT, HP_VAL = 23, 24      # نمودار دایره‌ای: نام و درصد
PAGE_PT = 784                # ارتفاعِ قابل‌چاپِ یک برگهٔ A4 عمودی (point)

try:
    import jdatetime
except ImportError:
    jdatetime = None

SRC_SHEET = "لیست جذب"

# ترتیب مالی ماه‌ها: مهر تا شهریور
MONTH_ORDER = [7, 8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6]
MONTH_NAME = {1: "فروردین", 2: "اردیبهشت", 3: "خرداد", 4: "تیر", 5: "مرداد", 6: "شهریور",
              7: "مهر", 8: "آبان", 9: "آذر", 10: "دی", 11: "بهمن", 12: "اسفند"}


# ---------------------------------------------------------------- ابزارها
def fa2en(s):
    """ارقام فارسی/عربی → لاتین، فاصله‌های تکراری و NBSP → یک فاصله."""
    if s is None:
        return ""
    s = str(s)
    for i, d in enumerate("۰۱۲۳۴۵۶۷۸۹"):
        s = s.replace(d, str(i))
    for i, d in enumerate("٠١٢٣٤٥٦٧٨٩"):
        s = s.replace(d, str(i))
    return re.sub(r"\s+", " ", s.replace(" ", " ")).strip()


def en2fa(s):
    """ارقام لاتین → فارسی (برای متن‌های گزارش)."""
    return str(s).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


def metrics_pct(n, d):
    return en2fa(f"{(n / d * 100) if d else 0:.0f}") + "٪"


_ZW = re.compile("[​‌‍‎‏﻿ـً-ْ]")


def norm_key(s):
    """کلید یکسان‌سازی برای گروه‌بندی: یِ/کِ عربی، نیم‌فاصله، اعراب و همهٔ فاصله‌ها حذف."""
    s = fa2en(s)
    s = (s.replace("ي", "ی").replace("ك", "ک").replace("أ", "ا")
          .replace("إ", "ا").replace("آ", "ا").replace("ة", "ه").replace("ۀ", "ه"))
    s = _ZW.sub("", s)
    return re.sub(r"\s+", "", s)


def parse_jalali(val):
    """رشتهٔ تاریخ شمسی مثل 1404/11/01 → تاریخ میلادی. خطا → None."""
    s = fa2en(val)
    if not s:
        return None
    for sep in ("/", "-", ".", " "):
        s = s.replace(sep, "/")
    parts = [p for p in s.split("/") if p != ""]
    if len(parts) < 3:
        return None
    try:
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
    except ValueError:
        return None
    if not (1300 <= y <= 1500) or m < 1 or m > 12 or d < 1:
        return None
    if d > 31:                       # مثل 015 یا 0606 خرابِ فایل
        d = int(str(d)[:2]) if int(str(d)[:2]) <= 31 else 1
    if d > 31:
        d = 1
    try:
        if jdatetime:
            return jdatetime.date(y, m, d).togregorian()
        return _jalali_to_gregorian(y, m, d)
    except Exception:
        return None


def _jalali_to_gregorian(jy, jm, jd):
    jy += 1595
    days = -355668 + (365 * jy) + ((jy // 33) * 8) + (((jy % 33) + 3) // 4) + jd
    days += (jm - 1) * 31 if jm < 7 else ((jm - 7) * 30 + 186)
    gy = 400 * (days // 146097)
    days %= 146097
    if days > 36524:
        days -= 1
        gy += 100 * (days // 36524)
        days %= 36524
        if days >= 365:
            days += 1
    gy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        gy += (days - 1) // 365
        days = (days - 1) % 365
    gd = days + 1
    leap = (gy % 4 == 0 and gy % 100 != 0) or (gy % 400 == 0)
    months = [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    gm = 0
    while gm < 12 and gd > months[gm]:
        gd -= months[gm]
        gm += 1
    return datetime.date(gy, gm + 1, gd)


def find_input():
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        return sys.argv[1]
    here = os.path.dirname(os.path.abspath(__file__))
    cands = [f for f in glob.glob(os.path.join(here, "*.xlsx"))
             if "خروجی" not in os.path.basename(f) and not os.path.basename(f).startswith("~$")]
    if not cands:
        print("!! هیچ فایل xlsx ورودی پیدا نشد. مسیر فایل را جلوی دستور بنویسید.")
        sys.exit(1)
    return max(cands, key=os.path.getmtime)


# ---------------------------------------------------------------- خواندن داده
def read_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if SRC_SHEET not in wb.sheetnames:
        raise ValueError(f"شیت «{SRC_SHEET}» در فایل نیست. شیت‌ها: {wb.sheetnames}")
    ws = wb[SRC_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    header = [fa2en(c) for c in rows[0]]

    def col(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None

    idx = {
        "radif": col("ردیف"),
        "name": col("نام", "نام "),
        "family": col("نام خانوادگی"),
        "unit": col("واحد"),
        "post": col("پست"),
        "project": col("محل خدمت"),
        "req_date": col("تاریخ درخواست"),
        "start_date": col("شروع به کار"),
        "month": col("ماه"),
    }
    data = []
    seq = 0
    for r in rows[1:]:
        # ردیفِ خالی را دیگر نادیده نمی‌گیریم — خیلی از ردیف‌های واقعی (اسم/پست/تاریخ دارند)
        # فقط عدد ترتیبی جلویشان نوشته نشده. سطر را فقط وقتی رد می‌کنیم که واقعاً خالی باشد
        # (نه ردیف، نه نام، نه نام خانوادگی، نه تاریخ شروع).
        has_name = idx["name"] is not None and r[idx["name"]] not in (None, "")
        has_family = idx["family"] is not None and r[idx["family"]] not in (None, "")
        has_radif = idx["radif"] is not None and r[idx["radif"]] not in (None, "")
        has_start = idx["start_date"] is not None and r[idx["start_date"]] not in (None, "")
        if not (has_name or has_family or has_radif or has_start):
            continue
        seq += 1
        rec = {k: (r[i] if i is not None else None) for k, i in idx.items()}
        if not has_radif:
            rec["radif"] = seq        # شمارهٔ جایگزین، فقط برای نمایش در شیت «زمان جذب»
        data.append(rec)
    return data


# ---------------------------------------------------------------- محاسبات
def _start_year(rec):
    m = re.match(r"(\d{4})", fa2en(rec["start_date"]))
    return int(m.group(1)) if m else None


def available_years(data):
    """سال‌های موجود در ستون «شروع به کار» (برای فیلترِ UI)."""
    ys = {y for y in (_start_year(r) for r in data) if y and 1300 <= y <= 1500}
    return sorted(ys)


def _start_month(rec):
    """ماهِ شمسیِ شروع به کار (۱..۱۲). از خودِ تاریخ می‌خوانیم تا با فیلترِ سال هم‌منبع باشد؛
    اگر تاریخ خراب بود، از ستون «ماه» کمک می‌گیریم."""
    m = re.match(r"\d{4}\s*[/\-.]\s*(\d{1,2})", fa2en(rec["start_date"]))
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(1))
    try:
        v = int(fa2en(rec["month"]))
        return v if 1 <= v <= 12 else None
    except (ValueError, TypeError):
        return None


def _month_label(mset):
    """برچسبِ فارسیِ ماه‌های انتخاب‌شده (مثلاً «شهریور» یا «تیر تا شهریور» یا «تیر، آذر»)."""
    ms = sorted(mset)
    if len(ms) == 1:
        return MONTH_NAME[ms[0]]
    if ms == list(range(ms[0], ms[-1] + 1)):
        return f"{MONTH_NAME[ms[0]]} تا {MONTH_NAME[ms[-1]]}"
    return "، ".join(MONTH_NAME[m] for m in ms)


def build_metrics(data, years=None, months=None):
    """years : مجموعهٔ سال‌های شمسیِ «شروع به کار» (None = همهٔ سال‌ها).
    months: مجموعهٔ ماه‌های شمسی ۱..۱۲ (None = همهٔ ماه‌ها). با years «و» می‌شود."""
    yset = set(years) if years else None
    mset = set(months) if months else None
    if yset is not None:
        data = [r for r in data if _start_year(r) in yset]
    if mset is not None:
        data = [r for r in data if _start_month(r) in mset]
    if (yset is not None or mset is not None) and not data:
        raise ValueError("در سال/ماهِ انتخاب‌شده هیچ رکوردی پیدا نشد.")

    total = len(data)
    hired, not_hired = [], 0
    for rec in data:
        g_start = parse_jalali(rec["start_date"])
        rec["_g_start"] = g_start
        if g_start is not None:
            hired.append(rec)
        else:
            not_hired += 1

    # زمان جذب = شروع به کار − تاریخ درخواست. هیچ سقفِ بالایی نداریم (۶ سال هم حساب می‌شود).
    # اگر شروع «قبل از» درخواست ثبت شده باشد عددِ منفی در شیت نشان داده می‌شود
    # (تا خطای ثبت پیدا شود) ولی وارد میانگین نمی‌شود.
    durations, neg_days = [], 0
    for rec in data:
        g_req = parse_jalali(rec["req_date"])
        g_start = rec["_g_start"]
        if g_req and g_start:
            days = (g_start - g_req).days
            rec["_days"] = days
            if days >= 0:
                durations.append(days)
            else:
                neg_days += 1
        else:
            rec["_days"] = None

    def breakdown(key, pool):
        """گروه‌بندی با کلید یکسان‌سازی‌شده تا نوشته‌های هم‌معنی با هم جمع شوند."""
        groups = {}
        for rec in pool:
            raw = fa2en(rec[key])
            if not raw:
                continue
            g = groups.setdefault(norm_key(raw), {"count": 0, "labels": Counter()})
            g["count"] += 1
            g["labels"][raw] += 1
        out = [(g["labels"].most_common(1)[0][0], g["count"]) for g in groups.values()]
        out.sort(key=lambda x: -x[1])
        return out

    by_project = breakdown("project", hired)
    by_post = breakdown("post", hired)
    by_unit = breakdown("unit", hired)

    mc = Counter()
    for rec in hired:
        try:
            mn = int(fa2en(rec["month"]))
        except ValueError:
            mn = None
        if mn in MONTH_NAME:
            mc[mn] += 1
    by_month = [(MONTH_NAME[m], mc[m]) for m in MONTH_ORDER if mc[m] > 0]

    # بازهٔ زمانی گزارش (از روی تاریخ‌های شروع به کار)
    years = set()
    for rec in hired:
        m = re.match(r"(\d{4})", fa2en(rec["start_date"]))
        if m:
            years.add(int(m.group(1)))
    span = ""
    if by_month and years:
        y1, y2 = min(years), max(years)
        if len(by_month) == 1:                      # فقط یک ماه (مثلاً فیلترِ ماه)
            span = (f"{by_month[0][0]} {y1}" if y1 == y2
                    else f"{by_month[0][0]} {y1} تا {y2}")
        else:
            span = (f"{by_month[0][0]} {y1} تا {by_month[-1][0]} {y2}"
                    if y1 != y2 else f"{by_month[0][0]} تا {by_month[-1][0]} {y1}")

    if yset:
        ys = sorted(yset)
        year_label = f"سال {ys[0]}" if len(ys) == 1 else f"سال‌های {ys[0]} تا {ys[-1]}"
    else:
        year_label = "همهٔ سال‌ها"
    if mset:
        # مثلاً «سال ۱۴۰۵ • شهریور»؛ اگر سال انتخاب نشده «همهٔ سال‌ها • شهریور»
        year_label += f" • {_month_label(mset)}"

    nh = len(hired)
    A = []
    if span:
        A.append(f"در بازهٔ {span}، مجموعاً {nh} نفر جذب شده‌اند؛ "
                 f"از {total} متقاضیِ ثبت‌شده، نرخ جذب {metrics_pct(nh, total)} بوده است.")
    else:
        A.append(f"مجموعاً {nh} نفر از {total} متقاضیِ ثبت‌شده جذب شده‌اند "
                 f"(نرخ جذب {metrics_pct(nh, total)}).")
    if len(by_month) > 1:                           # با یک ماه، «پرتراکم‌ترین ماه» بی‌معنی است
        pk = max(by_month, key=lambda x: x[1])
        A.append(f"پرتراکم‌ترین ماهِ جذب، {pk[0]} با {pk[1]} نفر "
                 f"({metrics_pct(pk[1], nh)} کل) بوده است.")
    if by_project:
        tp = by_project[0]
        s = f"بیشترین جذب در پروژهٔ «{tp[0]}» با {tp[1]} نفر ({metrics_pct(tp[1], nh)}) انجام شده"
        if len(by_project) > 1:
            s += f"؛ پس از آن «{by_project[1][0]}» با {by_project[1][1]} نفر."
        else:
            s += "."
        A.append(s)
    if by_post:
        A.append(f"پرتکرارترین پستِ جذب‌شده «{by_post[0][0]}» با {by_post[0][1]} نفر است.")
    if by_unit:
        A.append(f"در سطح واحد، «{by_unit[0][0]}» با {by_unit[0][1]} نفر بیشترین سهم را دارد.")
    if durations:
        avg = sum(durations) / len(durations)
        med = statistics.median(durations)
        s = (f"میانگین زمان جذب (از تاریخ درخواست تا شروع به کار) {avg:.0f} روز "
             f"(میانه {med:.0f} روز) است")
        s += (f" — بر پایهٔ {len(durations)} رکوردی که هر دو تاریخ را دارند."
              if len(durations) < total * 0.5 else ".")
        A.append(s)
    if neg_days:
        A.append(f"در {neg_days} رکورد تاریخِ شروع به کار قبل از تاریخِ درخواست ثبت شده "
                 f"(احتمالاً اشتباهِ ثبت)؛ این موارد در شیت «زمان جذب» قرمز شده‌اند و در میانگین نیامده‌اند.")
    if not_hired:
        A.append(f"{not_hired} نفر هنوز جذب نشده یا فرایندشان در جریان است.")
    A = [en2fa(s) for s in A]
    span = en2fa(span)
    year_label = en2fa(year_label)

    return {
        "year_label": year_label,
        "month_filtered": bool(mset),
        "total": total,
        "hired": nh,
        "not_hired": not_hired,
        "rate": (nh / total * 100) if total else 0,
        "avg_days": (sum(durations) / len(durations)) if durations else None,
        "n_days": len(durations),
        "n_project": len(by_project),
        "n_post": len(by_post),
        "n_unit": len(by_unit),
        "by_project": by_project,
        "by_post": by_post,
        "by_unit": by_unit,
        "by_month": by_month,
        "span": span,
        "analysis": A,
    }, data


# ---------------------------------------------------------------- ساخت اکسل خروجی
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=13, color="1F4E78")


def _fill_range(ws, r1, c1, r2, c2, color):
    f = PatternFill("solid", fgColor=color)
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            ws.cell(rr, cc).fill = f


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")


def write_table(ws, title, headers, rows, start_row=1, start_col=1, big=False):
    """جدولِ آراسته: بنر عنوان، هدر رنگی، ردیف‌های یک‌درمیان، قاب نازک، بدون خط شبکه.

    start_col : ستونِ شروعِ جدول (پیش‌فرض A؛ برای شیت‌های تفکیکی C).
    big       : فونتِ درشت.
    """
    ncol = len(headers)
    c0 = start_col
    cN = start_col + ncol - 1
    ws.sheet_view.showGridLines = False
    t_sz = 20 if big else 14
    h_sz = 14 if big else 11
    b_sz = 14 if big else 11

    # بنرِ عنوان
    ws.merge_cells(start_row=start_row, start_column=c0, end_row=start_row, end_column=cN)
    _fill_range(ws, start_row, c0, start_row, cN, NAVY)
    tc = ws.cell(start_row, c0, title)
    tc.font = Font(bold=True, size=t_sz, color="FFFFFF")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[start_row].height = 34 if big else 22

    hr = start_row + 1
    for j, h in enumerate(headers):
        cell = ws.cell(hr, c0 + j, h)
        cell.fill = HDR_FILL
        cell.font = Font(bold=True, size=h_sz, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[hr].height = 28 if big else 18

    thin = Side("thin", color="DCE4EC")
    for i, row in enumerate(rows, hr + 1):
        for j, v in enumerate(row):
            if isinstance(v, bool):
                disp = v
            elif isinstance(v, int):
                disp = en2fa(v)
            elif isinstance(v, float):
                disp = en2fa(f"{v:g}").replace(".", "٫")
            else:
                disp = v
            c = ws.cell(i, c0 + j, disp)
            c.border = Border(bottom=thin, left=thin, right=thin)
            c.font = Font(size=b_sz)
            c.alignment = Alignment(horizontal=("right" if j == 0 else "center"),
                                    vertical="center")
        if big:
            ws.row_dimensions[i].height = 24
    last = hr + len(rows)

    for j in range(ncol):
        ci = c0 + j
        letter = get_column_letter(ci)
        if big and ncol <= 3:
            # جدولِ ۲–۳ ستونه را پهن می‌کنیم تا محدودهٔ C..G را پر کند
            ws.column_dimensions[letter].width = 30 if ncol == 2 else 22
        else:
            width = max((len(str(ws.cell(ri, ci).value)) for ri in range(hr, last + 1)
                         if ws.cell(ri, ci).value is not None), default=10)
            ws.column_dimensions[letter].width = min(max(width + (6 if big else 3),
                                                         16 if big else 12), 52)

    if start_row == 1:
        ws.freeze_panes = ws.cell(hr + 1, 1).coordinate
    return hr, last


def _clean_labels(dl, show_val=False, show_pct=False):
    dl.showVal = show_val
    dl.showPercent = show_pct
    dl.showSerName = False
    dl.showCatName = False
    dl.showLegendKey = False
    dl.showBubbleSize = False


def _axis_text(sz=900, rot=-2700000):
    return RichText(
        bodyPr=RichTextProperties(rot=rot, vert="horz", anchor="ctr"),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz)))],
    )


def _frame(ch):
    """قاب نمودار: بدون خط دور، گوشهٔ گرد، پس‌زمینهٔ سفید."""
    ch.roundedCorners = True
    ch.visible_cells_only = False        # داده از ستون‌های مخفیِ helper هم خوانده شود
    ch.graphical_properties = GraphicalProperties(solidFill="FFFFFF")
    ch.graphical_properties.line = LineProperties(noFill=True)


def bar_chart(ws, title, min_row, max_row, cat_col=1, val_col=2, width=20, height=11,
              values=None):
    ch = BarChart()
    ch.type = "col"
    ch.title = title
    ch.height = height
    ch.width = width
    ch.legend = None
    ch.gapWidth = 45
    ch.overlap = 0
    data = Reference(ws, min_col=val_col, min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=cat_col, min_row=min_row + 1, max_row=max_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    n = max_row - min_row
    # برچسب داده روی میله نمی‌گذاریم؛ عددِ فارسی داخل خودِ برچسبِ محورِ دسته می‌آید
    ch.dLbls = DataLabelList()
    _clean_labels(ch.dLbls)
    ch.dLbls.delete = True
    # هر میله یک رنگِ کاملاً متفاوت از پالت
    pts = []
    for i in range(n):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=PALETTE[i % len(PALETTE)])
        dp.graphicalProperties.line = LineProperties(noFill=True)
        pts.append(dp)
    ch.series[0].data_points = pts
    # محورها — محور عمودی حذف می‌شود تا عددِ انگلیسی نماند (برچسبِ فارسیِ هر میله کافی است)
    ch.x_axis.delete = False
    ch.y_axis.delete = True
    ch.x_axis.tickLblPos = "low"
    ch.x_axis.txPr = _axis_text(900, -2700000)
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="C9D6E4"))
    ch.y_axis.majorGridlines = ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(solidFill="EDF1F6")))
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    _frame(ch)
    return ch


def line_chart(ws, title, min_row, max_row, cat_col=1, val_col=2, width=18, height=6.5,
               color=NAVY):
    """نمودار خطی با نشانه (Excel: «2‑D Line with Markers») از روی همان helper نمودار میله‌ای.
    عددِ فارسیِ هر نقطه داخل برچسبِ محور (نام (عدد)) است؛ محورِ عمودی حذف می‌شود."""
    ch = LineChart()
    ch.title = title
    ch.height = height
    ch.width = width
    ch.legend = None
    data = Reference(ws, min_col=val_col, min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=cat_col, min_row=min_row + 1, max_row=max_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    s = ch.series[0]
    s.smooth = False
    s.graphicalProperties = GraphicalProperties()
    s.graphicalProperties.line = LineProperties(solidFill=color, w=31750)     # ۲٫۵pt
    s.marker = Marker(symbol="circle", size=8)
    s.marker.graphicalProperties = GraphicalProperties(solidFill=color)
    s.marker.graphicalProperties.line = LineProperties(solidFill="FFFFFF", w=12700)
    ch.dLbls = DataLabelList()
    _clean_labels(ch.dLbls)
    ch.dLbls.delete = True
    ch.x_axis.delete = False
    ch.y_axis.delete = True
    ch.x_axis.tickLblPos = "low"
    ch.x_axis.txPr = _axis_text(900, -2700000)
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="C9D6E4"))
    ch.y_axis.majorGridlines = ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(solidFill="EDF1F6")))
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    _frame(ch)
    return ch


def pie_chart(ws, title, min_row, max_row, cat_col=1, val_col=2, width=13, height=9,
              values=None):
    ch = PieChart()
    ch.title = title
    ch.height = height
    ch.width = width
    data = Reference(ws, min_col=val_col, min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=cat_col, min_row=min_row + 1, max_row=max_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.dLbls = DataLabelList()
    _clean_labels(ch.dLbls, show_pct=True)   # درصدِ محاسبه‌شدهٔ اکسل با ارقام فارسی نمایش می‌یابد
    ch.legend.position = "r"
    ch.legend.overlay = False
    # رنگ هر تکه از پالت
    n = max_row - min_row
    pts = []
    for i in range(n):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(
            solidFill=PIE_COLORS[i % len(PIE_COLORS)])
        dp.graphicalProperties.line = LineProperties(solidFill="FFFFFF", w=12700)
        pts.append(dp)
    ch.series[0].data_points = pts
    _frame(ch)
    return ch


def split_small(rows, min_pct=3.0, keep_max=8):
    """(اسلایس‌های اصلی ، مواردی که در «سایر» جمع می‌شوند).

    اگر فقط «یک» مورد کوچک باشد، جمع‌کردنش در «سایر» بی‌معنی است؛ همان با اسمِ خودش می‌ماند.
    """
    total = sum(v for _, v in rows) or 1
    big, small = [], []
    for name, v in rows:
        if (v / total * 100) >= min_pct and len(big) < keep_max:
            big.append((name, v))
        else:
            small.append((name, v))
    if len(small) == 1:
        big.append(small.pop())
    return big, small


def group_small(rows, min_pct=3.0, keep_max=8):
    """اسلایس‌های کوچک‌تر از ۳٪ نمودار دایره‌ای در «سایر» جمع می‌شوند."""
    big, small = split_small(rows, min_pct, keep_max)
    if small:
        big.append(("سایر", sum(v for _, v in small)))
    return big


def _fa_pct(x):
    """۱٫۵ (یک رقم اعشار، اعشارِ فارسی) — برای درصدهای کوچک."""
    return en2fa(f"{x:.1f}").replace(".", "٫")


def others_caption(rows, min_pct=3.0, keep_max=8, max_items=8):
    """متنِ توضیحِ «سایر»: چه مواردی داخلش است، هرکدام چند نفر و چند درصد، و جمعشان.
    اگر «سایر»ی در نمودار نباشد رشتهٔ خالی برمی‌گرداند."""
    _, small = split_small(rows, min_pct, keep_max)
    if not small:
        return ""
    total = sum(v for _, v in rows) or 1
    sm_total = sum(v for _, v in small)
    shown = small if len(small) <= max_items + 1 else small[:max_items]    # یک موردِ اضافه را «و …» نکن
    parts = [f"{nm}: {en2fa(v)} نفر ({_fa_pct(v / total * 100)}٪)" for nm, v in shown]
    if len(shown) < len(small):
        rest = small[len(shown):]
        parts.append(f"و {en2fa(len(rest))} موردِ دیگر ({en2fa(sum(v for _, v in rest))} نفر)")
    return (f"«سایر» = {_fa_pct(sm_total / total * 100)}٪ ({en2fa(sm_total)} نفر) شامل:  "
            + "  •  ".join(parts))


def banner(ws, row, text, c1, c2, sub=None):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row + 1, end_column=c2)
    _fill_range(ws, row, c1, row + 1, c2, NAVY)
    cell = ws.cell(row, c1, text)
    cell.font = Font(bold=True, size=17, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    if sub:
        ws.merge_cells(start_row=row + 2, start_column=c1, end_row=row + 2, end_column=c2)
        sc = ws.cell(row + 2, c1, sub)
        sc.font = Font(size=10, color="6B7683")
        sc.alignment = Alignment(horizontal="center")


def kpi_card(ws, r, c, label, value, accent):
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 2)
    ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 3, end_column=c + 2)
    _fill_range(ws, r, c, r, c + 2, LIGHT)
    _fill_range(ws, r + 1, c, r + 3, c + 2, "FFFFFF")
    lc = ws.cell(r, c, label)
    lc.font = Font(size=9, bold=True, color="55606B")
    lc.alignment = Alignment(horizontal="center", vertical="center")
    vc = ws.cell(r + 1, c, value)
    vc.font = Font(size=26, bold=True, color=NAVY)
    vc.alignment = Alignment(horizontal="center", vertical="center")
    top = Side("thick", color=accent)
    thin = Side("thin", color="E3E9F0")
    for cc in range(c, c + 3):
        ws.cell(r, cc).border = Border(top=top, left=thin, right=thin)
        ws.cell(r + 3, cc).border = Border(bottom=thin, left=thin, right=thin)
    ws.row_dimensions[r].height = 18
    for rr in (r + 1, r + 2, r + 3):
        ws.row_dimensions[rr].height = 20


def _prep_logos(logo_path, wm_path=None):
    """(بنرِ کوچکِ واضح ، تصویرِ واترمارکِ تمام‌صفحه) را برمی‌گرداند.

    اگر wm_path داده شود، همان تصویرِ کاربر بدونِ هیچ تغییری برای پس‌زمینه به کار می‌رود.
    در غیر این صورت، نسخهٔ کم‌رنگِ خودکار از روی لوگو ساخته می‌شود.
    """
    bp = None
    if logo_path and os.path.isfile(logo_path):
        try:
            from PIL import Image as PILImage
            im = PILImage.open(logo_path).convert("RGBA")
            h = 90
            bp = os.path.join(tempfile.gettempdir(), "_jazb_logo_banner.png")
            im.resize((max(1, int(im.width * h / im.height)), h),
                      PILImage.LANCZOS).save(bp)
        except Exception:
            bp = None

    if wm_path and os.path.isfile(wm_path):
        return bp, wm_path          # تصویرِ آمادهٔ کاربر — دست‌نخورده

    wp = None
    if logo_path and os.path.isfile(logo_path):
        try:
            from PIL import Image as PILImage
            im = PILImage.open(logo_path).convert("RGBA")
            W = 1500
            wm = im.resize((W, max(1, int(im.height * W / im.width))), PILImage.LANCZOS)
            a = wm.split()[3]
            if a.getextrema()[0] > 200:
                # لوگو زمینهٔ سفیدِ تو‌پُر دارد → سفیدها را شفاف کن، وگرنه با کم‌رنگ‌شدن
                # یک «پردهٔ سفید» مستطیلی روی نمودارها و کارت‌ها می‌افتد.
                from PIL import ImageChops
                a = ImageChops.darker(a, wm.convert("L").point(lambda p: 0 if p > 235 else 255))
            wm.putalpha(a.point(lambda p: int(p * 0.16)))
            wp = os.path.join(tempfile.gettempdir(), "_jazb_logo_wm.png")
            wm.save(wp)
        except Exception:
            wp = None
    return bp, wp


def _watermark(ws, wp, anchors, height_px):
    if not wp:
        return
    for a in anchors:
        try:
            ws.add_image(_xl_img(wp, height_px), a)
        except Exception:
            pass


def _one_page(ws, area, landscape=False, fit_h=1, vcenter=True):
    """چاپ روی برگهٔ A4."""
    ws.print_area = area
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_h
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = vcenter
    for m_ in ("left", "right", "top", "bottom"):
        setattr(ws.page_margins, m_, 0.4)


def _row_at_pt(y, tstart, n):
    """شمارهٔ ردیفی که y پوینت پایین‌تر از بالای صفحه است (ارتفاعِ واقعیِ ردیف‌های جدول لحاظ می‌شود)."""
    acc, r = 0.0, 1
    while True:
        if r == tstart:
            h = 34
        elif r == tstart + 1:
            h = 28
        elif tstart + 2 <= r <= tstart + 1 + n:
            h = 24
        else:
            h = 15
        if acc + h > y:
            return r
        acc += h
        r += 1


def _xl_img(path, height_px):
    from openpyxl.drawing.image import Image as XLImage
    img = XLImage(path)
    img.width = int(img.width * height_px / img.height)
    img.height = height_px
    return img


def build_output(metrics, data, out_path, logo_path=None, wm_path=None):
    wb = openpyxl.Workbook()
    banner_logo, wm_logo = _prep_logos(logo_path, wm_path)

    def brand_banner(ws, row, text, c1, c2, sub=None):
        banner(ws, row, text, c1, c2, sub)
        if banner_logo:
            try:
                ws.add_image(_xl_img(banner_logo, 34), f"B{row}")
            except Exception:
                pass

    # ============================ شیت گزارش مدیریتی ============================
    ws = wb.active
    ws.title = "گزارش مدیریتی"
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for col in "ABCDEFGHIJKLM":
        ws.column_dimensions[col].width = 12
    ws.column_dimensions["A"].width = 3
    _watermark(ws, wm_logo, ["A2"], 1000)     # پس‌زمینهٔ تمام‌صفحه

    parts = [metrics.get("year_label") or "همهٔ سال‌ها"]
    if metrics.get("span") and not metrics.get("month_filtered"):
        parts.append(f"دورهٔ {metrics['span']}")     # با فیلترِ ماه، برچسب خودش ماه را دارد
    parts.append("واحد منابع انسانی")
    brand_banner(ws, 1, "گزارش مدیریتی جذب نیروی انسانی", 2, 12,
                 sub="  •  ".join(parts))

    cards = [
        ("تعداد نفرات", en2fa(metrics["total"]), NAVY),
        ("تعداد جذب", en2fa(metrics["hired"]), BLUE),
        ("نرخ جذب", metrics_pct(metrics["hired"], metrics["total"]), TEAL),
        ("میانگین زمان جذب (روز)",
         en2fa(f'{metrics["avg_days"]:.0f}') if metrics["avg_days"] is not None else "—", AMBER),
        ("جذب‌نشده / در جریان", en2fa(metrics["not_hired"]), "9E9E9E"),
        ("تعداد پروژه", en2fa(metrics["n_project"]), TEAL),
        ("تعداد پست", en2fa(metrics["n_post"]), NAVY),
        ("تعداد واحد", en2fa(metrics["n_unit"]), BLUE),
    ]
    starts = [(6, 11), (6, 8), (6, 5), (6, 2), (12, 11), (12, 8), (12, 5), (12, 2)]
    for (label, value, accent), (r, c) in zip(cards, starts):
        kpi_card(ws, r, c, label, value, accent)

    # ============================ شیت داشبورد ============================
    dash = wb.create_sheet("داشبورد")
    dash.sheet_view.rightToLeft = True
    dash.sheet_view.showGridLines = False
    dash.sheet_properties.tabColor = NAVY
    dash.column_dimensions["A"].width = 3
    for col in "BCDEFGHIJKLMNOPQRS":
        dash.column_dimensions[col].width = 11
    brand_banner(dash, 1, "داشبورد گزارش جذب", 2, 18)
    # آرمِ بزرگ روی هر صفحهٔ چاپ (هر ~۶۲ ردیف ≈ یک برگهٔ A4)
    _watermark(dash, wm_logo, [f"A{r}" for r in range(2, 320, 52)], 950)
    dash.page_setup.orientation = "portrait"
    dash.page_setup.paperSize = dash.PAPERSIZE_A4
    dash.page_setup.fitToWidth = 1
    dash.page_setup.fitToHeight = 0
    dash.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    dash.print_options.horizontalCentered = True
    anchor = [5]

    def strip(row, text):
        dash.merge_cells(start_row=row, start_column=2, end_row=row, end_column=18)
        _fill_range(dash, row, 2, row, 18, LIGHT)
        c = dash.cell(row, 2, text)
        c.font = Font(bold=True, size=12, color=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        dash.row_dimensions[row].height = 20

    def add_block(sheet_name, title, headers, rows, chart_kind, bar_top=None):
        s = wb.create_sheet(sheet_name)
        s.sheet_view.rightToLeft = True
        s.sheet_properties.tabColor = "9AA7B4"
        n = len(rows)
        long_tbl = bool(bar_top) or n > 20          # پست: فهرست بلند
        bar_rows = rows if bar_top is None else rows[:bar_top]

        # ---------- چیدمان: [جدول + نمودارِ میله‌ای + نمودارِ خطی] به‌صورت یک بلوک ----------
        # (ارتفاع‌ها بر حسب point؛ ردیفِ معمولی = ۱۵pt ؛ صفحهٔ A4 عمودی ≈ ۷۸۴pt)
        SH_CH_W = 18.2                               # پهنای نمودارها روی خودِ شیت (cm)
        table_pt = 34 + 28 + 24 * n
        if long_tbl:
            # فهرست بلند: دو نمودار بالای صفحهٔ اول، جدول زیرشان (چند صفحه می‌شود)
            BAR_H, LINE_H = 8.0, 7.5
            chart_row, line_row, tstart, fit_h = 2, 19, 35, 0
            parea = f"B1:J{tstart + n + 4}"
        else:
            BAR_H, LINE_H = 6.8, 6.4
            stack_pt = (BAR_H + LINE_H) / 2.54 * 72 + 12
            block_pt = table_pt + 30 + stack_pt
            tstart = max(2, 1 + round((PAGE_PT - block_pt) / 2 / 15))
            chart_row = tstart + n + 4               # ۲ ردیف فاصله زیر آخرین سطر جدول
            line_row = chart_row + 14                # نمودار میله‌ای ≈ ۱۳ ردیف
            after_pt = max(0, PAGE_PT - (tstart - 1) * 15 - table_pt)
            parea = (f"B1:J{max(tstart + n + 1 + max(int(after_pt // 15), 20), line_row + 14)}")
            fit_h = 1                                # اگر بلوک بلندتر از برگه شد، کمی کوچک می‌شود
        hr, last = write_table(s, title, headers, rows,
                               start_row=tstart, start_col=5, big=True)
        # ستون‌ها را جوری تنظیم می‌کنیم که پهنای چاپ ≈ یک برگهٔ A4 شود
        s.column_dimensions["E"].width = 26
        s.column_dimensions["F"].width = 26
        for cc in ("A", "B", "C", "D", "G", "H", "I", "J"):
            s.column_dimensions[cc].width = 6
        if banner_logo:
            try:
                s.add_image(_xl_img(banner_logo, 28), "A1")
            except Exception:
                pass
        # آرمِ کاملِ KPE — وسطِ صفحه و کاملاً داخل برگه (با نوشتهٔ KPE)
        wm_row = 4 if long_tbl else _row_at_pt((PAGE_PT - 640 * 0.75) / 2, tstart, n)
        _watermark(s, wm_logo, [f"B{wm_row}"], 640)
        _one_page(s, parea, fit_h=fit_h, vcenter=False)

        # ---- helperِ عددیِ مخفی برای نمودارها (ثابت از ردیف ۲؛ مستقل از جای جدول).
        # عمداً دور از ناحیهٔ چاپ (ستون‌های T..X) تا پهنای چاپ کم نشود.
        h0 = 2
        s.cell(h0, HB_CAT, headers[0])
        s.cell(h0, HB_VAL, headers[1])
        for i, (nm, v) in enumerate(bar_rows, h0 + 1):
            s.cell(i, HB_CAT, f"{nm} ({en2fa(v)})")   # نام + عددِ فارسی داخل برچسبِ محور
            s.cell(i, HB_VAL, v)
        bh_last = h0 + len(bar_rows)
        for ci in range(HB_CAT, HP_VAL + 1):
            s.column_dimensions[get_column_letter(ci)].hidden = True
        bar_vals = [v for _, v in bar_rows]

        # ---- روی خودِ شیت، از روی همین جدول: نمودارِ میله‌ای + نمودارِ خطی
        if bar_rows:
            sh_bar = bar_chart(s, "تعداد " + title, h0, bh_last, cat_col=HB_CAT,
                               val_col=HB_VAL, width=SH_CH_W, height=BAR_H, values=bar_vals)
            s.add_chart(sh_bar, f"B{chart_row}")
            sh_line = line_chart(s, "نمودار خطیِ " + title, h0, bh_last, cat_col=HB_CAT,
                                 val_col=HB_VAL, width=SH_CH_W, height=LINE_H)
            s.add_chart(sh_line, f"B{line_row}")

        # ---- داشبورد
        strip(anchor[0], title)
        anchor[0] += 2

        wide = len(bar_rows) > 12
        bc = bar_chart(s, "تعداد " + title, h0, bh_last, cat_col=HB_CAT, val_col=HB_VAL,
                       width=32 if wide else 20, height=13 if wide else 11,
                       values=bar_vals)
        dash.add_chart(bc, f"B{anchor[0]}")
        anchor[0] += 30 if wide else 24

        if chart_kind == "bar+pie":
            pie_rows = group_small(rows)
            s.cell(h0, HP_CAT, headers[0])
            s.cell(h0, HP_VAL, "درصد")
            tot = sum(v for _, v in pie_rows) or 1
            pcts = [round(v / tot * 100, 1) for _, v in pie_rows]
            for i, (nm, pc_) in enumerate(zip([p[0] for p in pie_rows], pcts), h0 + 1):
                s.cell(i, HP_CAT, nm)
                s.cell(i, HP_VAL, pc_)
            pc = pie_chart(s, "درصد " + title, h0, h0 + len(pie_rows),
                           cat_col=HP_CAT, val_col=HP_VAL, values=pcts)
            dash.add_chart(pc, f"B{anchor[0]}")
            anchor[0] += 19
            cap = others_caption(rows)            # «سایر» شامل چه مواردی است
            if cap:
                # سه ردیفِ ادغام‌شده با فونتِ درشت (صفحهٔ داشبورد موقع چاپ کوچک می‌شود)
                dash.merge_cells(start_row=anchor[0], start_column=2,
                                 end_row=anchor[0] + 2, end_column=18)
                _fill_range(dash, anchor[0], 2, anchor[0] + 2, 18, "FFFFFF")
                cc = dash.cell(anchor[0], 2, cap)
                cc.font = Font(size=13, color=INK)
                cc.alignment = Alignment(horizontal="right", vertical="center",
                                         wrap_text=True, readingOrder=2)
                for k in range(3):
                    dash.row_dimensions[anchor[0] + k].height = 24
            anchor[0] += 4

        anchor[0] += 3

    add_block("جذب بر اساس ماه", "جذب بر اساس ماه", ["ماه", "تعداد جذب"],
              metrics["by_month"], "bar+pie")
    add_block("جذب بر اساس پروژه", "جذب بر اساس پروژه", ["محل خدمت", "تعداد جذب"],
              metrics["by_project"], "bar+pie")
    add_block("جذب بر اساس واحد", "جذب بر اساس واحد", ["واحد", "تعداد جذب"],
              metrics["by_unit"], "bar+pie")
    add_block("جذب بر اساس پست", "جذب بر اساس پست (۲۰ ردیف نخست)", ["پست", "تعداد جذب"],
              metrics["by_post"], "bar", bar_top=20)

    zs = wb.create_sheet("زمان جذب")
    zs.sheet_view.rightToLeft = True
    zs.sheet_view.showGridLines = False
    zs.sheet_properties.tabColor = "9AA7B4"
    zs.freeze_panes = "A3"
    def _days_txt(d):
        if d is None:
            return "—"
        return en2fa(d) if d >= 0 else "‎−" + en2fa(-d)       # منفی: «−۷» (با LRM تا سمتِ چپ نیفتد)

    prows = [[
        en2fa(fa2en(rec["radif"])), fa2en(rec["name"]), fa2en(rec["family"]),
        en2fa(fa2en(rec["req_date"])), en2fa(fa2en(rec["start_date"])),
        _days_txt(rec["_days"]),
    ] for rec in data]
    hr_z, _last_z = write_table(zs, "فاصلهٔ تاریخ درخواست تا شروع به کار (روز)",
                                ["ردیف", "نام", "نام خانوادگی", "تاریخ درخواست",
                                 "شروع به کار", "زمان جذب (روز)"],
                                prows, start_col=5)
    # ردیف‌هایی که شروع به کار «قبل از» درخواست ثبت شده → قرمز (احتمالاً اشتباهِ ثبت)
    red = PatternFill("solid", fgColor="FDE9E7")
    for i, rec in enumerate(data):
        if rec["_days"] is not None and rec["_days"] < 0:
            for cc in range(5, 11):
                cell = zs.cell(hr_z + 1 + i, cc)
                cell.fill = red
                cell.font = Font(size=cell.font.size or 11, color="C0392B", bold=(cc == 10))
    zs.page_setup.orientation = "landscape"
    zs.page_setup.paperSize = zs.PAPERSIZE_A4
    zs.page_setup.fitToWidth = 1
    zs.page_setup.fitToHeight = 0
    zs.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    zs.print_options.horizontalCentered = True
    if banner_logo:
        try:
            zs.add_image(_xl_img(banner_logo, 28), "A1")
        except Exception:
            pass
    _watermark(zs, wm_logo, [f"A{r}" for r in range(2, 320, 38)], 680)

    # ===== تکمیل شیت گزارش مدیریتی: دو نمودار کلیدی + تحلیل + تنظیم چاپ =====
    _fill_range(ws, 19, 2, 19, 12, LIGHT)
    ws.merge_cells("B19:M19")
    tc = ws.cell(19, 2, "نمودارهای کلیدی")
    tc.font = Font(bold=True, size=12, color=NAVY)
    tc.alignment = Alignment(horizontal="center", vertical="center")

    sh_m = wb["جذب بر اساس ماه"]
    m_bar = bar_chart(sh_m, "جذب ماهانه", 2, 2 + len(metrics["by_month"]),
                      cat_col=HB_CAT, val_col=HB_VAL, width=16, height=9,
                      values=[v for _, v in metrics["by_month"]])
    ws.add_chart(m_bar, "B21")

    sh_p = wb["جذب بر اساس پروژه"]
    p_pie_rows = group_small(metrics["by_project"])
    _ptot = sum(v for _, v in p_pie_rows) or 1
    p_pie = pie_chart(sh_p, "سهم پروژه‌ها از جذب", 2, 2 + len(p_pie_rows),
                      cat_col=HP_CAT, val_col=HP_VAL, width=12, height=9,
                      values=[round(v / _ptot * 100, 1) for _, v in p_pie_rows])
    ws.add_chart(p_pie, "I21")

    # توضیحِ «سایر» درست زیرِ نمودار دایره‌ای: چه مواردی، هرکدام چند نفر/چند درصد، و جمعشان
    cap = others_caption(metrics["by_project"])
    if cap:
        ws.merge_cells("I39:M42")
        thin_ = Side("thin", color="D6DEE8")
        for rr in range(39, 43):
            for cc in range(9, 14):
                ws.cell(rr, cc).fill = PatternFill("solid", fgColor="F6F8FB")
                ws.cell(rr, cc).border = Border(
                    top=thin_ if rr == 39 else None, bottom=thin_ if rr == 42 else None,
                    left=thin_ if cc == 9 else None, right=thin_ if cc == 13 else None)
        cc_ = ws.cell(39, 9, cap)
        cc_.font = Font(size=11, color=INK)
        cc_.alignment = Alignment(horizontal="right", vertical="center",
                                  wrap_text=True, readingOrder=2)

    A0 = 44                                            # شروعِ بخشِ تحلیل (زیرِ توضیحِ «سایر»)
    _fill_range(ws, A0, 2, A0, 12, LIGHT)
    ws.merge_cells(start_row=A0, start_column=2, end_row=A0, end_column=13)
    ac = ws.cell(A0, 2, "تحلیل و نتیجه‌گیری")
    ac.font = Font(bold=True, size=12, color=NAVY)
    ac.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=A0 + 1, start_column=2, end_row=A0 + 12, end_column=13)
    body = ws.cell(A0 + 1, 2, "\n".join("•  " + s for s in metrics.get("analysis", [])))
    body.font = Font(size=11, color=INK)
    body.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
    for rr in range(A0 + 1, A0 + 13):
        ws.row_dimensions[rr].height = 18

    _one_page(ws, f"A1:M{A0 + 14}")

    wb.save(out_path)


def find_logo(near):
    """فایل لوگو کنار ورودی/اسکریپت/دسکتاپ: kpe.png / logo.png …"""
    dirs = {os.path.dirname(os.path.abspath(near)),
            os.path.dirname(os.path.abspath(__file__)),
            os.path.dirname(os.path.abspath(sys.argv[0])),
            os.path.join(os.path.expanduser("~"), "Desktop")}
    for d in dirs:
        for n in ("kpe.png", "KPE.png", "kpe.jpg", "kpe.jpeg",
                  "logo.png", "logo.jpg", "logo.jpeg"):
            p = os.path.join(d, n)
            if os.path.isfile(p):
                return p
    return None


def parse_years(text):
    """'1405' یا '1404-1405' یا '1403,1405' → مجموعهٔ سال‌ها. خالی → None."""
    if not text:
        return None
    t = fa2en(text).replace("تا", "-")
    nums = [int(x) for x in re.findall(r"\d{4}", t)]
    if not nums:
        return None
    if "-" in t and len(nums) >= 2:
        return set(range(min(nums), max(nums) + 1))
    return set(nums)


_MONTH_KEYS = {norm_key(n): i for i, n in MONTH_NAME.items()}


def _month_token(tok):
    """یک ماه (عدد ۱..۱۲ یا نام فارسی) → عدد. نامعتبر → ValueError."""
    tok = tok.strip()
    if tok.isdigit():
        n = int(tok)
    else:
        n = _MONTH_KEYS.get(norm_key(tok))
    if n is None or not 1 <= n <= 12:
        raise ValueError(f"ماهِ نامعتبر: «{tok}» (عدد ۱ تا ۱۲ یا نام ماه، مثل شهریور)")
    return n


def parse_months(text):
    """ورودیِ ماه → مجموعهٔ عددهای ۱..۱۲؛ خالی/«همه» → None.

    نمونه‌ها:  6  |  شهریور  |  4-6  |  تیر تا شهریور  |  4,5,6  |  تیر، مرداد
    بازهٔ معکوس دورِ سال می‌چرخد:  دی تا فروردین = ۱۰،۱۱،۱۲،۱
    """
    if text is None:
        return None
    t = fa2en(text)
    if not t or t in ("همه", "all", "*"):
        return None
    t = re.sub(r"\s+تا\s+|(?<=\d)تا(?=\d)", "-", t)     # «تا» → خط‌تیره
    t = re.sub(r"\s*[-–—]\s*", "-", t)
    out = set()
    for tok in re.split(r"[,،;؛\s]+", t):
        if not tok:
            continue
        if "-" in tok:
            a, b = (_month_token(x) for x in tok.split("-", 1))
            out |= set(range(a, b + 1)) if a <= b else set(range(a, 13)) | set(range(1, b + 1))
        else:
            out.add(_month_token(tok))
    return out or None


def find_watermark(near):
    """تصویرِ آمادهٔ پس‌زمینه کنار ورودی/برنامه/دسکتاپ."""
    dirs = {os.path.dirname(os.path.abspath(near)),
            os.path.dirname(os.path.abspath(__file__)),
            os.path.dirname(os.path.abspath(sys.argv[0])),
            os.path.join(os.path.expanduser("~"), "Desktop")}
    for d in dirs:
        for n in ("watermark.png", "watermark.jpg", "arme.png", "arme.jpg",
                  "pass-zamine.png", "background.png"):
            p = os.path.join(d, n)
            if os.path.isfile(p):
                return p
    return None


def make_report(src, folder=None, logo_path=None, years=None, wm_path=None, months=None):
    """کل مسیر: خواندن → محاسبه → ساخت خروجی. مسیر فایل خروجی را برمی‌گرداند.
    years : مجموعهٔ سال‌های شمسیِ «شروع به کار» (یا None برای همه).
    months: مجموعهٔ ماه‌های شمسی ۱..۱۲ (یا None برای همه).
    wm_path: تصویرِ آمادهٔ پس‌زمینه (بدون تغییر استفاده می‌شود)."""
    data = read_rows(src)
    metrics, data = build_metrics(data, years, months)
    base = os.path.splitext(os.path.basename(src))[0]
    folder = folder or os.path.dirname(os.path.abspath(src))
    wm_path = wm_path or find_watermark(src)
    tag = ""
    if years:
        ys = sorted(years)
        tag += f" - {ys[0]}" if len(ys) == 1 else f" - {ys[0]} تا {ys[-1]}"
    if months:
        tag += f" - {_month_label(set(months))}"
    out = os.path.join(folder, f"خروجی گزارش جذب - {base}{tag}.xlsx")
    try:
        build_output(metrics, data, out, logo_path, wm_path)
    except PermissionError:
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        out = out[:-5] + f" [{stamp}].xlsx"
        build_output(metrics, data, out, logo_path, wm_path)
    return out, metrics


def main():
    argv = sys.argv[1:]
    # --month / -m  <مقدار>   مثل:  --month شهریور   یا   -m 4-6
    mtext = None
    for flag in ("--month", "-m"):
        if flag in argv:
            i = argv.index(flag)
            mtext = argv[i + 1] if i + 1 < len(argv) else None
            argv = argv[:i] + argv[i + 2:]
    mons = parse_months(mtext)
    args = [a for a in argv if not re.fullmatch(r"\d{4}(\s*-\s*\d{4})?", a)]
    yrs = next((parse_years(a) for a in argv
                if re.fullmatch(r"\d{4}(\s*-\s*\d{4})?", a)), None)
    src = args[0] if args and os.path.isfile(args[0]) else find_input()
    print(f"ورودی : {src}   |   سال: {yrs or 'همه'}   |   ماه: "
          f"{_month_label(mons) if mons else 'همه'}")
    out, m = make_report(src, years=yrs, months=mons)
    print(f"خروجی: {out}")
    print(f"  تعداد نفرات = {m['total']}  |  تعداد جذب = {m['hired']}  |  نرخ جذب = {m['rate']:.1f}%")
    if m["avg_days"] is not None:
        print(f"  میانگین زمان جذب = {m['avg_days']:.1f} روز  (روی {m['n_days']} رکورد)")
    print("تمام شد. نمودارها در شیت «داشبورد».")


if __name__ == "__main__":
    main()
