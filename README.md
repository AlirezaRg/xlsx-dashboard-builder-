# xlsx-dashboard-builder — Excel → Styled Report & Dashboard Generator

Turn a plain Persian data sheet into a polished, print‑ready Excel workbook:
KPI cards, bar charts for **counts**, pie charts for **percentages**, an
auto‑written management one‑pager, per‑row date‑gap calculations, a full
RTL layout with Persian digits, an optional company‑logo watermark, and a
year filter — all from one click.

It ships as a single Windows `.exe` with a small GUI; no Python install
needed for end users.

---

## Features

- **Management one‑pager** (`گزارش مدیریتی`) — banner with the reporting
  period, 8 KPI cards, two key charts, and a **“Analysis & Conclusions”**
  section the tool writes automatically from the numbers (peak month, top
  category share, overall rate, average duration, …).
- **Dashboard** (`داشبورد`) — every breakdown as a bar chart (counts) and a
  pie chart (percentages), stacked one section per printed page.
- **Per‑breakdown sheets** — one styled table per dimension
  (by month / by project / by unit / by role), big fonts, zebra rows.
- **Duration sheet** — for every row, the gap in days between a “request
  date” column and a “start date” column.
- **All numbers rendered as Persian digits** — tables, chart labels, dates,
  KPI values, the analysis text. No Latin digits in the output.
- **Grouping is normalized** — half‑spaces (ZWNJ), Arabic vs. Persian
  ي/ك, diacritics and extra spaces are folded, so
  `کمک انباردار` / `کمک انبار دار` count as one.
- **Company‑logo watermark** — supply a transparent PNG and it is placed,
  **unmodified**, as a full‑page background on every sheet and every
  printed page (logo included at the bottom, not clipped).
- **Year filter** — type a single Jalali year (`1405`) or a range
  (`1404-1405`); every metric, chart and title is recomputed for that
  window, and the year is added to the output filename.
- **Print‑ready** — each sheet is set to A4, fit‑to‑width, short tables
  vertically centered.

---

## Quick start (end users)

1. Put **`جذب.exe`** anywhere (e.g. the Desktop).
2. Double‑click it.
3. **Choose Excel file** → pick your data workbook.
4. *(optional)* **Choose logo** and/or type a **year**.
5. **Build report.**
6. **Open output file** — it is written next to the input as
   `خروجی گزارش جذب - <name>.xlsx`.

If the output file is already open in Excel, a timestamped copy is written
instead of failing.

---

## Input format

The tool reads **one sheet named `لیست جذب`**. It locates columns by their
header text (order does not matter). Recognized headers:

| Header (in the sheet) | Used for |
|---|---|
| `ردیف` | row id (required — rows without it are skipped) |
| `نام`, `نام خانوادگی` | person name (duration sheet) |
| `شروع به کار` | “start” date — a row counts as *completed* when this is a valid Jalali date; its year drives the year filter |
| `تاریخ درخواست` | “request” date — used with the start date for the duration (days) |
| `ماه` | month number (1–12), fiscal order مهر→شهریور |
| `محل خدمت` | breakdown: by project |
| `واحد` | breakdown: by unit |
| `پست` | breakdown: by role |

Dates may be written with Persian or Latin digits and `/`, `-`, `.` or
space separators; malformed day values (`10/015`, `0606`) are tolerated.

### Optional files (auto‑detected next to the exe / input / Desktop)

| File name | Effect |
|---|---|
| `watermark.png` (or `arme.png`, `background.png`) | used **verbatim** as the page background — must be a transparent PNG |
| `kpe.png` (or `logo.png`) | small crisp logo in each sheet banner; if no `watermark.png` exists, a faded copy of this is used as the background |

Excel places pictures **on top of** cell text, so a watermark with an
opaque (white) background will hide the tables. Export it from your editor
as a **transparent** PNG.

---

## Output workbook

| Sheet | Contents |
|---|---|
| `گزارش مدیریتی` | management one‑pager (KPIs + 2 charts + analysis) |
| `داشبورد` | all bar + pie charts |
| `جذب بر اساس ماه` / `... پروژه` / `... واحد` / `... پست` | one styled table each |
| `زمان جذب` | per‑row request→start gap in days |

---

## Command line

```bat
python robot_jazb.py "path\to\input.xlsx"
python robot_jazb.py "path\to\input.xlsx" 1405
python robot_jazb.py "path\to\input.xlsx" 1404-1405
```

With no path it picks the most recent `*.xlsx` next to the script.

---

## Run from source

```bat
pip install openpyxl jdatetime pillow
python robot_jazb_ui.py
```

Python 3.10+ on Windows (the GUI uses Tkinter, which ships with Python).

## Build the exe

```bat
pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed ^
  --icon kpe.ico --name "جذب" --add-data "kpe.ico;." ^
  --hidden-import jdatetime --hidden-import openpyxl ^
  --hidden-import PIL --hidden-import PIL.ImageTk ^
  robot_jazb_ui.py
```

The exe appears in `dist\جذب.exe`.

---

## Project layout

| File | Role |
|---|---|
| `robot_jazb.py` | core — read sheet, compute metrics, build the workbook |
| `robot_jazb_ui.py` | Tkinter GUI wrapper |
| `make_icon.py` | builds `icon.ico` (fallback app icon) |
| `kpe.ico` | app icon bundled into the exe |

---

## Notes / limitations

- Built for the specific `لیست جذب` sheet layout described above; other
  sheets/headers are ignored.
- Chart data labels are embedded in the category axis text
  (`دی (۱۶)`) because Excel cannot render custom per‑point label text via
  the library used.
- A true “behind everything, prints too” page background does not exist in
  `.xlsx`; the watermark is a floating picture, so its transparent areas
  are what let the tables show through.
