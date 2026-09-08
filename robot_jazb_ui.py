# -*- coding: utf-8 -*-
"""
ربات گزارش جذب – نسخهٔ پنجره‌ای (UI)
اجرا:  جذب.exe   یا   pythonw robot_jazb_ui.py
"""

import os
import sys
import threading
import traceback
import subprocess

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import robot_jazb as core

# ---------------- تم رنگی ----------------
NAVY   = "#1F4E78"
NAVY_D = "#163A5C"
ACCENT = "#2E75B6"
BG     = "#EEF2F7"
CARD   = "#FFFFFF"
LINE   = "#D6DEE8"
MUTED  = "#6B7683"
INK    = "#2A333C"
FONT   = ("Tahoma", 10)
FONT_B = ("Tahoma", 10, "bold")

LOGO_NAMES = ("kpe.png", "KPE.png", "kpe.jpg", "kpe.jpeg", "logo.png", "logo.jpg")


def _res(name):
    """مسیر فایلِ همراهِ برنامه (چه اجرای معمولی چه exeی PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def _find_logo():
    dirs = [
        os.path.dirname(os.path.abspath(sys.argv[0])),
        os.getcwd(),
        os.path.join(os.path.expanduser("~"), "Desktop"),
        os.path.dirname(os.path.abspath(__file__)),
    ]
    for d in dirs:
        for n in LOGO_NAMES:
            p = os.path.join(d, n)
            if os.path.isfile(p):
                return p
    return None


def _load_img(path, h):
    """تصویر را با ارتفاع h برای نمایش در Tk برمی‌گرداند (نیازمند Pillow)."""
    try:
        from PIL import Image, ImageTk
        im = Image.open(path).convert("RGBA")
        w = max(1, int(im.width * h / im.height))
        return ImageTk.PhotoImage(im.resize((w, h), Image.LANCZOS))
    except Exception:
        try:
            img = tk.PhotoImage(file=path)
            f = max(1, img.height() // h)
            return img.subsample(f, f)
        except Exception:
            return None


class Btn(tk.Button):
    """دکمهٔ تخت با افکت hover."""
    def __init__(self, master, primary=False, **kw):
        bg = NAVY if primary else "#E3EAF2"
        fg = "white" if primary else INK
        self._bg, self._hover = bg, (NAVY_D if primary else "#D2DDEA")
        base = dict(bg=bg, fg=fg, activebackground=self._hover, activeforeground=fg,
                    relief="flat", bd=0, cursor="hand2",
                    font=FONT_B if primary else FONT, padx=14, pady=7)
        base.update(kw)
        super().__init__(master, **base)
        self.bind("<Enter>", lambda e: self.configure(bg=self._hover))
        self.bind("<Leave>", lambda e: self.configure(bg=self._bg))

    def set_bg(self, bg, hover):
        self._bg, self._hover = bg, hover
        self.configure(bg=bg)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ربات گزارش جذب")
        self.configure(bg=BG)
        self.geometry("660x670")
        self.minsize(620, 620)
        self.src_path = tk.StringVar(value="")
        self.logo_path = tk.StringVar(value="")
        self.wm_path = tk.StringVar(value="")
        self.out_path = None
        self._imgs = []
        # آیکون ثابتِ برنامه = لوگوی KPE (مستقل از انتخاب لوگوی گزارش)
        for setter in ("iconbitmap", "wm_iconbitmap"):
            try:
                getattr(self, setter)(_res("kpe.ico"))
                break
            except Exception:
                pass
        try:
            from PIL import Image, ImageTk
            self._app_icon = ImageTk.PhotoImage(Image.open(_res("kpe.ico")))
            self.iconphoto(True, self._app_icon)
        except Exception:
            pass
        self._build()

    # ---------------- ساخت ظاهر ----------------
    def _card(self, parent):
        c = tk.Frame(parent, bg=CARD, highlightbackground=LINE,
                     highlightthickness=1, bd=0)
        c.pack(fill="x", pady=6)
        return c

    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("J.Horizontal.TProgressbar", troughcolor="#DCE4EC",
                        background=ACCENT, bordercolor="#DCE4EC",
                        lightcolor=ACCENT, darkcolor=ACCENT)

        # ---- هدر
        head = tk.Frame(self, bg=NAVY)
        head.pack(fill="x")
        inner = tk.Frame(head, bg=NAVY)
        inner.pack(fill="x", padx=18, pady=14)
        self.hdr_logo = tk.Label(inner, bg=NAVY)   # اگر لوگو انتخاب شد اینجا نشان داده می‌شود
        txt = tk.Frame(inner, bg=NAVY)
        txt.pack(side="right", fill="x", expand=True)
        tk.Label(txt, text="ربات گزارش جذب نیروی انسانی", bg=NAVY, fg="white",
                 font=("Tahoma", 15, "bold"), anchor="e").pack(fill="x")
        tk.Label(txt, text="فایل اکسل «لیست جذب» را بده، فایل گزارش با نمودار را بگیر",
                 bg=NAVY, fg="#C7D6E6", font=("Tahoma", 9), anchor="e").pack(fill="x")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=(12, 0))

        # ---- کارت فایل اکسل
        c1 = self._card(body)
        row1 = tk.Frame(c1, bg=CARD)
        row1.pack(fill="x", padx=12, pady=12)
        Btn(row1, text="انتخاب فایل اکسل…", primary=True,
            command=self.pick).pack(side="right")
        self.lbl_file = tk.Label(row1, text="فایلی انتخاب نشده", font=FONT, bg="#F6F8FB",
                                 fg=MUTED, anchor="e", relief="flat")
        self.lbl_file.pack(side="right", fill="x", expand=True, padx=(0, 10), ipady=6)

        # ---- کارت لوگو
        c2 = self._card(body)
        row2 = tk.Frame(c2, bg=CARD)
        row2.pack(fill="x", padx=12, pady=12)
        Btn(row2, text="انتخاب لوگو…", command=self.pick_logo).pack(side="right")
        self.lbl_logo = tk.Label(row2, font=FONT, bg="#F6F8FB", anchor="e", relief="flat",
                                 fg=MUTED, text="بدون لوگو  •  با دکمهٔ «انتخاب لوگو» انتخاب کن")
        self.lbl_logo.pack(side="right", fill="x", expand=True, padx=(0, 10), ipady=6)

        # ---- کارت تصویر پس‌زمینه (واترمارک آماده)
        c2b = self._card(body)
        row2b = tk.Frame(c2b, bg=CARD)
        row2b.pack(fill="x", padx=12, pady=12)
        Btn(row2b, text="انتخاب تصویر پس‌زمینه…", command=self.pick_wm).pack(side="right")
        self.lbl_wm = tk.Label(row2b, font=FONT, bg="#F6F8FB", anchor="e", relief="flat",
                               fg=MUTED,
                               text="بدون پس‌زمینه  •  یا فایل watermark.png را کنار برنامه بگذار")
        self.lbl_wm.pack(side="right", fill="x", expand=True, padx=(0, 10), ipady=6)

        # ---- کارت سال
        c3 = self._card(body)
        row3 = tk.Frame(c3, bg=CARD)
        row3.pack(fill="x", padx=12, pady=12)
        tk.Label(row3, text="سال (اختیاری):", font=FONT_B, bg=CARD, fg=INK).pack(side="right")
        self.ent_year = tk.Entry(row3, font=FONT, width=18, justify="center",
                                 relief="solid", bd=1, bg="#F6F8FB")
        self.ent_year.pack(side="right", padx=10, ipady=4)
        tk.Label(row3, text="خالی = همه   •   یک سال: 1405   •   بازه: 1404-1405",
                 font=("Tahoma", 8), bg=CARD, fg=MUTED).pack(side="right")
        self.lbl_year = tk.Label(c3, text="", font=("Tahoma", 8), bg=CARD, fg=MUTED,
                                 anchor="e")
        self.lbl_year.pack(fill="x", padx=12, pady=(0, 8))

        # ---- دکمهٔ اصلی
        self.btn_run = Btn(body, text="ساخت گزارش", primary=True, command=self.run,
                           font=("Tahoma", 12, "bold"))
        self.btn_run.configure(state="disabled", pady=10)
        self.btn_run.pack(fill="x", pady=(10, 4))

        # ---- فوتر (از پایین)
        foot = tk.Frame(self, bg=BG)
        foot.pack(side="bottom", fill="x", padx=18, pady=(6, 14))
        self.btn_open = Btn(foot, text="باز کردن فایل خروجی", command=self.open_out)
        self.btn_open.configure(state="disabled")
        self.btn_open.pack(side="right")
        self.btn_dir = Btn(foot, text="باز کردن پوشه", command=self.open_dir)
        self.btn_dir.configure(state="disabled")
        self.btn_dir.pack(side="right", padx=8)

        self.pb = ttk.Progressbar(self, mode="indeterminate",
                                  style="J.Horizontal.TProgressbar")
        self.pb.pack(side="bottom", fill="x", padx=18)

        # ---- لاگ
        wrap = tk.Frame(self, bg=LINE, bd=0)
        wrap.pack(side="bottom", fill="both", expand=True, padx=18, pady=(10, 8))
        self.log = tk.Text(wrap, height=8, font=("Consolas", 9), state="disabled",
                           bg="#1B2733", fg="#D6E0EA", wrap="word", relief="flat",
                           padx=10, pady=8, insertbackground="#D6E0EA")
        self.log.pack(fill="both", expand=True, padx=1, pady=1)

    # ---------------- کارها ----------------
    def say(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def pick(self):
        p = filedialog.askopenfilename(title="فایل اکسل لیست جذب",
                                       filetypes=[("Excel", "*.xlsx"), ("همه", "*.*")])
        if not p:
            return
        self.src_path.set(p)
        self.lbl_file.configure(text=os.path.basename(p), fg=INK)
        self.btn_run.configure(state="normal")
        self.btn_open.configure(state="disabled")
        self.btn_dir.configure(state="disabled")
        self.say("فایل انتخاب شد: " + p)
        try:
            found = [str(y) for y in core.available_years(core.read_rows(p))]
            if found:
                self.lbl_year.configure(text="سال‌های موجود در این فایل: " + "، ".join(found))
                self.say("سال‌های موجود: " + "، ".join(found))
            else:
                self.lbl_year.configure(text="سالی در ستون «شروع به کار» پیدا نشد")
        except Exception as e:
            self.say("خواندن سال‌ها ناموفق: " + str(e))

    def pick_logo(self):
        p = filedialog.askopenfilename(
            title="فایل لوگوی شرکت",
            filetypes=[("تصویر", "*.png *.jpg *.jpeg *.bmp *.gif"), ("همه", "*.*")])
        if not p:
            return
        self.logo_path.set(p)
        self.lbl_logo.configure(text=os.path.basename(p), fg=INK)
        self._show_hdr_logo(p)
        self.say("لوگو انتخاب شد: " + p)

    def pick_wm(self):
        p = filedialog.askopenfilename(
            title="تصویر پس‌زمینهٔ صفحه (واترمارک آماده)",
            filetypes=[("تصویر", "*.png *.jpg *.jpeg"), ("همه", "*.*")])
        if not p:
            return
        self.wm_path.set(p)
        self.lbl_wm.configure(text=os.path.basename(p), fg=INK)
        self.say("تصویر پس‌زمینه انتخاب شد (بدون تغییر استفاده می‌شود): " + p)

    def _show_hdr_logo(self, p):
        img = _load_img(p, 44)
        if img:
            self._imgs = [img]
            self.hdr_logo.configure(image=img)
            self.hdr_logo.pack(side="right", padx=(0, 12))

    def run(self):
        self.btn_run.configure(state="disabled")
        self.btn_open.configure(state="disabled")
        self.btn_dir.configure(state="disabled")
        self.pb.start(12)
        threading.Thread(target=self._work, daemon=True).start()

    def _work(self):
        try:
            self.say("در حال خواندن شیت «لیست جذب» …")
            years = core.parse_years(self.ent_year.get())
            if years:
                self.say(f"فیلتر سال: {sorted(years)}")
            out, m = core.make_report(self.src_path.get(),
                                      logo_path=self.logo_path.get() or None,
                                      years=years,
                                      wm_path=self.wm_path.get() or None)
            self.out_path = out
            self.say("")
            self.say(f"تعداد نفرات = {m['total']}   |   تعداد جذب = {m['hired']}"
                     f"   |   نرخ جذب = {m['rate']:.1f}%")
            if m["avg_days"] is not None:
                self.say(f"میانگین زمان جذب = {m['avg_days']:.1f} روز "
                         f"(روی {m['n_days']} رکورد دارای تاریخ درخواست)")
            self.say("")
            self.say("✔ خروجی ساخته شد:")
            self.say(out)
            self.after(0, self._done_ok)
        except Exception as e:
            self.say("")
            self.say("✖ خطا: " + str(e))
            self.say(traceback.format_exc())
            self.after(0, lambda: self._done_err(e))

    def _done_ok(self):
        self.pb.stop()
        self.btn_run.configure(state="normal")
        self.btn_open.configure(state="normal")
        self.btn_dir.configure(state="normal")
        messagebox.showinfo("انجام شد", "فایل گزارش ساخته شد.\nنمودارها در شیت «داشبورد» هستند.")

    def _done_err(self, e):
        self.pb.stop()
        self.btn_run.configure(state="normal")
        messagebox.showerror("خطا", f"ساخت گزارش انجام نشد:\n{e}")

    def open_out(self):
        if self.out_path and os.path.exists(self.out_path):
            os.startfile(self.out_path)

    def open_dir(self):
        if self.out_path:
            subprocess.run(["explorer", "/select,", os.path.normpath(self.out_path)])


if __name__ == "__main__":
    App().mainloop()
