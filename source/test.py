import tkinter as tk
from tkinter import messagebox, ttk, simpledialog, colorchooser, filedialog
import csv
import tkinter.font as tkfont
import sqlite3
from category_list_prices import LIST_PRICE, category_list_prices, ensure_list_price_columns, list_prices_for_sale
_OPENPYXL_OK = True
try:
    import openpyxl
except ImportError:
    _OPENPYXL_OK = False
from datetime import datetime, timedelta
from tkcalendar import Calendar


def normalize_date_text(value):
    """Normalize common Malaysian date input styles to YYYY-MM-DD."""
    raw = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError("Invalid date")


def item_picker_categories(items):
    """Return stable category names for the Edit Order item picker."""
    categories = {
        str(item["category"] or "Uncategorized").strip() or "Uncategorized"
        for item in items
    }
    return sorted(categories, key=str.casefold)


def filter_item_picker_items(items, query="", category=None):
    """Filter item-picker rows by the selected category and search text."""
    needle = str(query or "").strip().casefold()
    selected = str(category or "").strip().casefold()
    matches = []
    for item in items:
        name = str(item["name"] or "")
        item_category = (
            str(item["category"] or "Uncategorized").strip()
            or "Uncategorized"
        )
        if selected and item_category.casefold() != selected:
            continue
        if needle and needle not in name.casefold() and needle not in item_category.casefold():
            continue
        matches.append(item)
    return matches


class DatePicker(tk.Frame):
    """Consistent, keyboard-friendly date input used throughout the desktop POS."""

    def __init__(self, parent, textvariable, callback=None, **kwargs):
        try:
            parent_bg = parent.cget("bg")
        except Exception:
            parent_bg = "white"
        bg = kwargs.get("bg", parent_bg)
        super().__init__(parent, bg=bg)

        self.var = textvariable
        self.callback = callback
        self.language = str(kwargs.get("language", "en") or "en").lower()
        self._popup = None
        font = kwargs.get("font", ("Arial", 11))
        width = int(kwargs.get("width", 12))

        self.entry = tk.Entry(
            self, textvariable=self.var, font=font, width=width,
            bg="white", fg="#1F2937", insertbackground="#1F2937",
            relief="solid", bd=1, justify="center"
        )
        self.entry.pack(side="left", ipady=4)
        self.entry.bind("<Return>", self._apply_typed_date)
        self.entry.bind("<Control-a>", self._select_all)

        self.btn = tk.Button(
            self, text="📅", command=self.on_click,
            font=("Arial", 11), bg="#2563EB", fg="white",
            activebackground="#1D4ED8", activeforeground="white",
            relief="flat", bd=0, padx=9, pady=4, cursor="hand2"
        )
        self.btn.pack(side="left", padx=(3, 0))

    def _t(self, english, chinese):
        return chinese if self.language.startswith("zh") else english

    def _select_all(self, _event=None):
        self.entry.selection_range(0, "end")
        return "break"

    def _close_popup(self):
        top = self._popup
        self._popup = None
        self.top = None
        if top is not None:
            try:
                top.grab_release()
            except Exception:
                pass
            try:
                top.destroy()
            except Exception:
                pass

    def _commit_date(self, value, close_popup=True):
        try:
            normalized = normalize_date_text(value)
        except ValueError:
            self.entry.configure(bg="#FEF2F2")
            messagebox.showerror(
                self._t("Invalid Date", "日期无效"),
                self._t(
                    "Use YYYY-MM-DD, DD-MM-YYYY or DD/MM/YYYY.",
                    "可输入 YYYY-MM-DD、DD-MM-YYYY 或 DD/MM/YYYY。"
                ),
                parent=self.winfo_toplevel()
            )
            self.entry.focus_set()
            self.entry.selection_range(0, "end")
            return False

        self.var.set(normalized)
        self.entry.configure(bg="white")
        if close_popup:
            self._close_popup()
        if self.callback:
            self.callback()
        return True

    def _apply_typed_date(self, _event=None):
        self._commit_date(self.var.get(), close_popup=False)
        return "break"

    def on_click(self, _event=None):
        if self._popup is not None:
            try:
                if self._popup.winfo_exists():
                    self._popup.lift()
                    self._popup.focus_force()
                    return
            except Exception:
                self._popup = None

        try:
            selected = datetime.strptime(normalize_date_text(self.var.get()), "%Y-%m-%d")
        except ValueError:
            selected = datetime.now()

        top = tk.Toplevel(self)
        self._popup = top
        self.top = top
        top._skip_auto_center = True
        top.title(self._t("Choose Date", "选择日期"))
        top.configure(bg="white")
        top.resizable(False, False)
        top.transient(self.winfo_toplevel())
        top.protocol("WM_DELETE_WINDOW", self._close_popup)

        header = tk.Frame(top, bg="#1E3A5F", padx=14, pady=10)
        header.pack(fill="x")
        tk.Label(
            header, text=self._t("📅 Choose a date", "📅 选择日期"),
            font=("Arial", 13, "bold"), bg="#1E3A5F", fg="white"
        ).pack(side="left")
        tk.Button(
            header, text="✕", command=self._close_popup,
            font=("Arial", 10, "bold"), bg="#1E3A5F", fg="white",
            activebackground="#B91C1C", activeforeground="white",
            relief="flat", bd=0, cursor="hand2"
        ).pack(side="right")

        quick = tk.Frame(top, bg="white", padx=10, pady=9)
        quick.pack(fill="x")

        def quick_button(label, offset, color):
            tk.Button(
                quick, text=label,
                command=lambda: self._commit_date(
                    (datetime.now() + timedelta(days=offset)).strftime("%Y-%m-%d")
                ),
                font=("Arial", 9, "bold"), bg=color, fg="white",
                activebackground=color, activeforeground="white",
                relief="flat", bd=0, padx=11, pady=5, cursor="hand2"
            ).pack(side="left", padx=3, expand=True, fill="x")

        quick_button(self._t("Yesterday", "昨天"), -1, "#64748B")
        quick_button(self._t("Today", "今天"), 0, "#16A34A")
        quick_button(self._t("Tomorrow", "明天"), 1, "#2563EB")

        nav = tk.Frame(top, bg="#F8FAFC", padx=10, pady=7)
        nav.pack(fill="x", padx=10)
        tk.Label(
            nav, text=self._t("Jump to:", "快速跳转："),
            font=("Arial", 9, "bold"), bg="#F8FAFC", fg="#475569"
        ).pack(side="left", padx=(0, 6))

        month_names = (
            [f"{month}月" for month in range(1, 13)]
            if self.language.startswith("zh")
            else [
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            ]
        )
        month_var = tk.StringVar(value=month_names[selected.month - 1])
        year_var = tk.StringVar(value=str(selected.year))
        month_box = ttk.Combobox(
            nav, textvariable=month_var, values=month_names,
            state="readonly", width=10, font=("Arial", 9)
        )
        month_box.pack(side="left", padx=3)
        year_spin = tk.Spinbox(
            nav, textvariable=year_var, from_=2000, to=datetime.now().year + 10,
            width=6, font=("Arial", 9), justify="center"
        )
        year_spin.pack(side="left", padx=3)

        cal_options = dict(
            selectmode="day",
            date_pattern="yyyy-mm-dd",
            showweeknumbers=False,
            firstweekday="monday",
            year=selected.year,
            month=selected.month,
            day=selected.day,
            font=("Arial", 11),
            background="#2563EB",
            foreground="white",
            headersbackground="#E2E8F0",
            headersforeground="#1E293B",
            normalbackground="white",
            normalforeground="#1E293B",
            weekendbackground="#F8FAFC",
            weekendforeground="#B91C1C",
            selectbackground="#16A34A",
            selectforeground="white",
            bordercolor="#CBD5E1"
        )
        try:
            self.cal = Calendar(
                top,
                locale="zh_CN" if self.language.startswith("zh") else "en_US",
                **cal_options
            )
        except Exception:
            self.cal = Calendar(top, **cal_options)
        self.cal.pack(fill="both", expand=True, padx=10, pady=(8, 4))

        def jump_month(_event=None):
            try:
                month = month_names.index(month_var.get()) + 1
                year = int(year_var.get())
                self.cal.see(datetime(year, month, 1))
            except (ValueError, IndexError):
                top.bell()

        def sync_month_controls(_event=None):
            try:
                month, year = self.cal.get_displayed_month()
                month_var.set(month_names[int(month) - 1])
                year_var.set(str(year))
            except Exception:
                pass

        tk.Button(
            nav, text=self._t("Go", "前往"), command=jump_month,
            font=("Arial", 9, "bold"), bg="#475569", fg="white",
            activebackground="#334155", activeforeground="white",
            relief="flat", bd=0, padx=9, pady=3, cursor="hand2"
        ).pack(side="left", padx=(5, 0))
        month_box.bind("<<ComboboxSelected>>", jump_month)
        year_spin.bind("<Return>", jump_month)
        self.cal.bind("<<CalendarMonthChanged>>", sync_month_controls)
        self.cal.bind(
            "<<CalendarSelected>>",
            lambda _event: top.after_idle(
                lambda: self._commit_date(self.cal.get_date())
            )
        )

        footer = tk.Frame(top, bg="white", padx=10, pady=8)
        footer.pack(fill="x")
        tk.Label(
            footer,
            text=self._t(
                "Click a day to use it immediately",
                "点击日期即可立即使用"
            ),
            font=("Arial", 9), bg="white", fg="#64748B"
        ).pack(side="left")
        tk.Button(
            footer, text=self._t("Cancel", "取消"), command=self._close_popup,
            font=("Arial", 9, "bold"), bg="#E2E8F0", fg="#334155",
            activebackground="#CBD5E1", relief="flat", bd=0,
            padx=12, pady=4, cursor="hand2"
        ).pack(side="right")

        top.update_idletasks()
        popup_width = max(390, top.winfo_reqwidth())
        popup_height = top.winfo_reqheight()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 4
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        x = max(8, min(x, screen_width - popup_width - 8))
        if y + popup_height + 50 > screen_height:
            y = self.winfo_rooty() - popup_height - 36
        y = max(8, y)
        top.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
        top.bind("<Escape>", lambda _event: self._close_popup())
        top.grab_set()
        top.focus_force()


from decimal import Decimal, ROUND_HALF_UP
import json
import os
import sys
import uuid
from contextlib import contextmanager
from wcwidth import wcswidth
import win32print


class StartupLoadingScreen:
    """Small painted splash screen used while the synchronous Tk UI is built."""

    def __init__(self, root):
        self.root = root
        self.window = tk.Toplevel(root)
        self.window.overrideredirect(True)
        self.window.configure(bg="#0F172A")
        try:
            self.window.attributes("-topmost", True)
        except tk.TclError:
            pass

        width, height = 520, 250
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

        tk.Label(
            self.window, text="919", font=("Arial", 34, "bold"),
            bg="#0F172A", fg="white"
        ).pack(pady=(28, 0))
        tk.Label(
            self.window, text="Restaurant POS", font=("Arial", 13, "bold"),
            bg="#0F172A", fg="#94A3B8"
        ).pack(pady=(0, 20))

        self.status_var = tk.StringVar(value="Loading / 加载中…")
        tk.Label(
            self.window, textvariable=self.status_var,
            font=("Arial", 10), bg="#0F172A", fg="#E2E8F0"
        ).pack()

        progress_style = ttk.Style(self.window)
        progress_style.configure(
            "Startup.Horizontal.TProgressbar",
            troughcolor="#334155", background="#22C55E",
            bordercolor="#334155", lightcolor="#22C55E",
            darkcolor="#22C55E"
        )
        self.progress = ttk.Progressbar(
            self.window, mode="determinate", maximum=100,
            style="Startup.Horizontal.TProgressbar", length=420
        )
        self.progress.pack(pady=(12, 5))
        self.percent_var = tk.StringVar(value="5%")
        tk.Label(
            self.window, textvariable=self.percent_var,
            font=("Arial", 9, "bold"), bg="#0F172A", fg="#94A3B8"
        ).pack()
        self.update_progress(5, "Starting / 正在启动…")

    def update_progress(self, value, message):
        value = max(0, min(100, int(value)))
        self.progress["value"] = value
        self.percent_var.set(f"{value}%")
        self.status_var.set(str(message))
        try:
            self.window.update_idletasks()
            self.window.update()
        except tk.TclError:
            pass

    def close(self):
        try:
            self.window.destroy()
        except tk.TclError:
            pass


# ── Rounded Button — smooth polygon, no jagged edges ───────────────────────
import math as _math

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk as _ImageTk
    _PIL_OK = True
except ImportError:
    _PIL_OK = False


class RoundedButton(tk.Canvas):
    _img_cache: dict = {}

    def __init__(self, parent, text="", command=None,
                 bg="#16A34A", fg="#ffffff", font=("Arial", 11, "bold"),
                 padx=18, pady=8, radius=10, width=0, height=0,
                 hover_bg=None, hover_fg=None, **kwargs):
        for k in ("relief", "cursor", "activebackground", "activeforeground",
                  "borderwidth", "bd"):
            kwargs.pop(k, None)

        try:
            parent_bg = parent.cget("bg")
        except Exception:
            parent_bg = "#F0F2F5"

        super().__init__(parent, highlightthickness=0, bg=parent_bg, **kwargs)

        self._text     = text
        self._command  = command
        self._font_spec = font
        self._padx     = padx
        self._pady     = pady
        self._radius   = radius
        self._bg_orig  = bg
        self._bg       = bg
        self._hover_bg = hover_bg if hover_bg else self._darken(bg, 22)
        self._fg_orig  = fg
        self._fg       = fg
        self._hover_fg = hover_fg if hover_fg else fg
        self._parent_bg = parent_bg
        self._photo    = None

        tmp = tk.Label(parent, text=text, font=font)
        tw, th = tmp.winfo_reqwidth(), tmp.winfo_reqheight()
        tmp.destroy()
        cw = width  if width  > 0 else tw + padx * 2
        ch = height if height > 0 else th + pady * 2
        self.config(width=max(80, cw), height=max(36, ch))

        self.bind("<Configure>",     self._on_resize)
        self.bind("<ButtonPress-1>", self._click)
        self.bind("<Enter>",  lambda e: self._switch_state(self._hover_bg, self._hover_fg))
        self.bind("<Leave>",  lambda e: self._switch_state(self._bg_orig, self._fg_orig))

    @staticmethod
    def _darken(hex_color, amount=22):
        try:
            h = hex_color.lstrip("#")
            r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
            return f"#{max(0,r-amount):02x}{max(0,g-amount):02x}{max(0,b-amount):02x}"
        except Exception:
            return hex_color

    @staticmethod
    def _hex_to_rgb(hex_color):
        h = hex_color.lstrip("#")
        return int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)

    def _switch_state(self, bg_color, fg_color):
        self._bg = bg_color
        self._fg = fg_color
        self._render()

    def _on_resize(self, event=None):
        self._render()

    def _render(self, event=None):
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 4 or h < 4:
            self.after(50, self._render)
            return

        # Optimization: Don't re-render if nothing changed
        current_state = (w, h, self._text, self._bg, self._fg, self._radius)
        if hasattr(self, "_last_rendered_state") and self._last_rendered_state == current_state:
            return
        self._last_rendered_state = current_state

        if _PIL_OK:
            self._render_pil(w, h)
        else:
            self._render_canvas(w, h)

    def _render_pil(self, w, h):
        key = (w, h, self._radius, self._bg, self._parent_bg)
        if key in self._img_cache:
            photo = self._img_cache[key]
        else:
            SCALE = 3
            W, H  = w * SCALE, h * SCALE
            r     = min(self._radius * SCALE, W // 2, H // 2)
            bg_rgb  = self._hex_to_rgb(self._bg)
            pbg_rgb = self._hex_to_rgb(self._parent_bg)

            img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            draw.rounded_rectangle([0, 0, W-1, H-1], radius=r, fill=bg_rgb + (255,))

            bg_img = Image.new("RGBA", (W, H), pbg_rgb + (255,))
            bg_img.paste(img, mask=img)
            final = bg_img.resize((w, h), Image.LANCZOS).convert("RGB")
            photo = _ImageTk.PhotoImage(final)
            self._img_cache[key] = photo

        self.delete("all")
        self._photo = photo  # Keep a reference to prevent GC
        self.create_image(0, 0, anchor="nw", image=photo)

        self.create_text(w // 2, h // 2,
                         text=self._text,
                         font=self._font_spec,
                         fill=self._fg,
                         justify="center")

    def _render_canvas(self, w, h):
        self.delete("all")
        r  = min(self._radius, w//2, h//2)
        x1, y1, x2, y2 = 1, 1, w-1, h-1
        c  = self._bg
        self.create_arc(x1,      y1,      x1+2*r, y1+2*r, start=90,  extent=90,  fill=c, outline=c)
        self.create_arc(x2-2*r,  y1,      x2,     y1+2*r, start=0,   extent=90,  fill=c, outline=c)
        self.create_arc(x1,      y2-2*r,  x1+2*r, y2,     start=180, extent=90,  fill=c, outline=c)
        self.create_arc(x2-2*r,  y2-2*r,  x2,     y2,     start=270, extent=90,  fill=c, outline=c)
        self.create_rectangle(x1+r, y1, x2-r, y2, fill=c, outline=c)
        self.create_rectangle(x1, y1+r, x2, y2-r, fill=c, outline=c)
        self.create_text(w//2, h//2, text=self._text,
                         font=self._font_spec, fill=self._fg, justify="center")

    def _click(self, event=None):
        if self._command:
            self._command()

    def pack(self, **kw):
        super().pack(**kw)
        self.after(15, self._render)

    def grid(self, **kw):
        super().grid(**kw)
        self.after(15, self._render)

    def configure(self, **kw):
        if "text"    in kw: self._text     = kw.pop("text")
        if "command" in kw: self._command  = kw.pop("command")
        if "bg"      in kw:
            self._bg = self._bg_orig = kw.pop("bg")
            self._hover_bg = self._darken(self._bg_orig)
        super().configure(**kw)


PRINTER_NAME = "POS-58-Series"

def _send_raw_printer_job(printer_name, document_name, data):
    handle = None
    job_id = None
    stage = "open printer"
    try:
        handle = win32print.OpenPrinter(printer_name)
        stage = "start job"
        job_id = win32print.StartDocPrinter(handle, 1, (document_name, None, "RAW"))
        stage = "start page"
        win32print.StartPagePrinter(handle)
        stage = "send data"
        written = win32print.WritePrinter(handle, data)
        if written != len(data):
            raise OSError(f"Windows accepted {written} of {len(data)} bytes")
        stage = "end page"
        win32print.EndPagePrinter(handle)
        stage = "finish job"
        win32print.EndDocPrinter(handle)
    except Exception as error:
        if job_id is not None:
            try:
                win32print.SetJob(handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
            except Exception:
                pass  # A completed or removed job may no longer exist.
        raise OSError(f"{printer_name} · {document_name} · {stage}\n{error}") from error
    finally:
        if handle is not None:
            try:
                win32print.ClosePrinter(handle)
            except Exception:
                pass  # The job may already be queued. Avoid reporting it as unsent.


def print_to_thermal(printer_name, text):
    data = b'\x1b\x40\x1c\x26' + text.encode("gb18030", errors="ignore")
    _send_raw_printer_job(printer_name, "Receipt", data)


def open_cash_drawer(printer_name=None):
    """Send ESC/POS cash-drawer kick command to the receipt printer.
    The drawer must be connected via RJ11 to the printer (USB printer to PC).
    Return the Windows error on failure so callers can report it without
    blocking a transaction."""
    try:
        _send_raw_printer_job(printer_name or PRINTER_NAME, "CashDrawer", b'\x1b\x70\x00\x19\xfa')
    except OSError as error:
        return str(error)
    return None



# ================= BUSINESS DAY HELPER =================
def get_business_date(dt=None):
    if dt is None:
        dt = datetime.now()
    if dt.hour < 6:
        return (dt - timedelta(days=1)).strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d")

def business_day_range(biz_date_str):
    start = datetime.strptime(biz_date_str, "%Y-%m-%d").replace(hour=6, minute=0, second=0)
    end   = start + timedelta(hours=23, minutes=59, seconds=59)
    return start, end

def today_biz_date():
    return get_business_date(datetime.now())

def today_biz_range():
    biz = today_biz_date()
    s, e = business_day_range(biz)
    return s.strftime("%Y-%m-%d %H:%M:%S"), e.strftime("%Y-%m-%d %H:%M:%S")

def biz_date_of_sale(datetime_str):
    try:
        dt = datetime.strptime(datetime_str, "%d-%m-%Y %H:%M:%S")
    except Exception:
        try:
            dt = datetime.strptime(datetime_str[:19], "%Y-%m-%d %H:%M:%S")
        except Exception:
            return datetime_str[:10]
    return get_business_date(dt)

def normalize_biz_date_text(raw, fallback_dt=None):
    """Normalize legacy date text into YYYY-MM-DD business date format."""
    txt = str(raw or '').strip()
    if not txt:
        return get_business_date(fallback_dt or datetime.now())
    try:
        return datetime.strptime(txt[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
    except Exception:
        pass
    try:
        dt = datetime.strptime(txt[:10], "%d-%m-%Y")
        return get_business_date(dt)
    except Exception:
        return get_business_date(fallback_dt or datetime.now())


_CASH_TRANSACTION_EXPENSE_TYPES = frozenset({
    'expense',
    'stock_purchase',
    'staff_advance',
    'cash_transfer',
    'kuih_cost',
})
_CASH_TRANSACTION_INCOME_TYPES = frozenset({
    'income',
    'kuih_sales',
})


def cash_transaction_flow_type(raw_type):
    """Map a cash transaction subtype to its cashflow direction.

    Unknown subtypes are deliberately ignored instead of being treated as
    income. Defaulting an unfamiliar value to income can inflate the drawer
    balance, which is what the old legacy backfill did for kuih_cost.
    """
    tx_type = str(raw_type or '').strip().lower()
    if tx_type in _CASH_TRANSACTION_EXPENSE_TYPES:
        return 'Expense'
    if tx_type in _CASH_TRANSACTION_INCOME_TYPES:
        return 'Income'
    return None


def find_sale_item_matches(items_json, item_query):
    """Return item-name and quantity pairs matching a name substring."""
    query = str(item_query or '').strip().casefold()
    if not query:
        return []
    try:
        raw_items = json.loads(items_json or '[]')
    except Exception:
        return []

    matches = {}
    for item in raw_items:
        if not isinstance(item, (list, tuple)) or len(item) < 3:
            continue
        name = str(item[0] or '').strip()
        if not name or query not in name.casefold():
            continue
        try:
            quantity = int(item[2] or 0)
        except Exception:
            quantity = 0
        matches[name] = matches.get(name, 0) + quantity
    return list(matches.items())


def compute_cash_summary_components(conn, biz_date):
    """Return physical-cash components with legacy fallback support.

    QR income remains revenue, but it must not increase the amount expected in
    the cash drawer at end of day.
    """
    c = conn.cursor()

    # Primary POS cash source: net cash sales (includes amend/reversal/refund handling).
    _total_sales, pos_cash, _qr_sales, _card_sales = compute_sales_payment_breakdown(conn, biz_date)

    c.execute("""SELECT COALESCE(SUM(amount),0) FROM cashflow
                 WHERE type='Income' AND source NOT LIKE 'Sales%' AND date=?
                   AND UPPER(COALESCE(NULLIF(TRIM(payment_method),''),'CASH'))='CASH'""", (biz_date,))
    other_income = float(c.fetchone()[0])

    c.execute("""SELECT COALESCE(SUM(amount),0) FROM cashflow
                 WHERE type='Expense' AND source!='Refund' AND date=?
                   AND UPPER(COALESCE(NULLIF(TRIM(payment_method),''),'CASH'))='CASH'""", (biz_date,))
    expenses = float(c.fetchone()[0])

    # Existing transaction identities in cashflow, independent of accounting direction.
    c.execute("SELECT source, reference_id FROM cashflow WHERE reference_id IS NOT NULL")
    cf_keys = {
        (str(r['source'] or ''), int(r['reference_id']))
        for r in c.fetchall()
        if r['reference_id'] is not None
    }

    # NOTE: no sales fallback here. pos_cash already comes from
    # compute_sales_payment_breakdown(), which includes its own legacy fallback logic.

    # Fallback 1: manual cash transactions not yet represented in cashflow.
    c.execute("SELECT id, type, category, amount, date, payment_method FROM cash_transactions")
    for r in c.fetchall():
        try:
            ref_id = int(r['id'])
        except Exception:
            continue
        t_raw = str(r['type'] or '').strip().lower()
        cf_type = cash_transaction_flow_type(t_raw)
        if not cf_type:
            continue
        source = (r['category'] or 'Other').strip() or 'Other'
        if (source, ref_id) in cf_keys:
            continue
        if normalize_biz_date_text(r['date']) != biz_date:
            continue
        if str(r['payment_method'] or 'CASH').strip().upper() != 'CASH':
            continue
        amt = float(r['amount'] or 0)
        if cf_type == 'Expense':
            expenses += amt
        else:
            other_income += amt

    # Fallback 2: rental payments not yet represented in cashflow.
    c.execute("SELECT id, amount, pay_date, voided, payment_method FROM rent_payments")
    for r in c.fetchall():
        try:
            ref_id = int(r['id'])
        except Exception:
            continue
        if int(r['voided'] or 0) != 0:
            continue
        if ('Stall Rental', ref_id) in cf_keys:
            continue
        if normalize_biz_date_text(r['pay_date']) != biz_date:
            continue
        if str(r['payment_method'] or 'CASH').strip().upper() != 'CASH':
            continue
        other_income += float(r['amount'] or 0)

    return pos_cash, other_income, expenses


def compute_qr_rent_total(conn, biz_date):
    """QR rent for closing, excluding voids and retaining legacy date support."""
    return sum(float(row['amount'] or 0) for row in conn.execute(
        "SELECT amount, pay_date FROM rent_payments WHERE COALESCE(voided,0)=0 "
        "AND UPPER(TRIM(payment_method))='QR'"
    ) if normalize_biz_date_text(row['pay_date']) == biz_date)


def compute_sales_payment_breakdown(conn, biz_date):
    """Return (total_sales, cash_sales, qr_sales, card_sales) for one business date."""
    c = conn.cursor()

    c.execute(
        """SELECT
               COALESCE(SUM(amount),0) AS total_sales,
               COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,''))='CASH' THEN amount ELSE 0 END),0) AS cash_sales,
               COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,''))='QR' THEN amount ELSE 0 END),0) AS qr_sales,
               COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,''))='CARD' THEN amount ELSE 0 END),0) AS card_sales
           FROM cashflow
           WHERE type='Income' AND source LIKE 'Sales%' AND date=?""",
        (biz_date,)
    )
    row = c.fetchone()
    total_sales = float(row[0] or 0)
    cash_sales = float(row[1] or 0)
    qr_sales = float(row[2] or 0)
    card_sales = float(row[3] or 0)

    # Deduct refunds recorded in cashflow so reported sales are net sales.
    c.execute(
        """SELECT
               COALESCE(SUM(amount),0) AS total_refunds,
               COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,''))='CASH' THEN amount ELSE 0 END),0) AS cash_refunds,
               COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,''))='QR' THEN amount ELSE 0 END),0) AS qr_refunds,
               COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,''))='CARD' THEN amount ELSE 0 END),0) AS card_refunds
           FROM cashflow
           WHERE type='Expense' AND source='Refund' AND date=?""",
        (biz_date,)
    )
    ref_row = c.fetchone()
    total_sales -= float(ref_row[0] or 0)
    cash_sales -= float(ref_row[1] or 0)
    qr_sales -= float(ref_row[2] or 0)
    card_sales -= float(ref_row[3] or 0)

    # Fallback: include legacy sales rows that are not represented in cashflow yet.
    c.execute(
        """SELECT DISTINCT reference_id
           FROM cashflow
           WHERE type='Income' AND source LIKE 'Sales%' AND reference_id IS NOT NULL"""
    )
    cf_sale_ids = {
        int(r['reference_id'])
        for r in c.fetchall()
        if r['reference_id'] is not None
    }

    c.execute("SELECT id, datetime, total, payment FROM sales WHERE amended = 0")
    for r in c.fetchall():
        try:
            sale_id = int(r['id'])
        except Exception:
            continue
        if sale_id in cf_sale_ids:
            continue
        if biz_date_of_sale(r['datetime'] or '') != biz_date:
            continue

        amt = float(r['total'] or 0)
        pm = str(r['payment'] or 'CASH').strip().upper()
        total_sales += amt
        if pm == 'CASH':
            cash_sales += amt
        elif pm == 'QR':
            qr_sales += amt
        elif pm == 'CARD':
            card_sales += amt

    # Fallback: legacy refunds rows not represented in cashflow yet.
    c.execute("SELECT DISTINCT reference_id FROM cashflow WHERE type='Expense' AND source='Refund' AND reference_id IS NOT NULL")
    cf_refund_ids = {
        int(r['reference_id'])
        for r in c.fetchall()
        if r['reference_id'] is not None
    }

    # Also track which sales have already been refunded in cashflow (to avoid double counting fallback)
    c.execute("SELECT DISTINCT notes FROM cashflow WHERE type='Expense' AND source='Refund' AND notes LIKE 'Full Refund: Sale #%'")
    cf_refunded_sale_ids = set()
    for r in c.fetchall():
        try:
            sid = int(r['notes'].split('#')[-1].split(':')[0])
            cf_refunded_sale_ids.add(sid)
        except Exception: pass

    c.execute("SELECT id, datetime, refund_amount, payment_method FROM refunds")
    for r in c.fetchall():
        try:
            refund_id = int(r['id'])
            sale_id_ref = int(r['sale_id'])
        except Exception:
            continue
        # If this refund ID is already in cashflow, don't subtract again.
        if refund_id in cf_refund_ids:
            continue
        # If this sale ID was used as a reference instead of the refund ID (legacy bug fix), don't subtract again.
        if sale_id_ref in cf_refund_ids:
            continue
        if biz_date_of_sale(r['datetime'] or '') != biz_date:
            continue

        amt = float(r['refund_amount'] or 0)
        pm = str(r['payment_method'] or 'CASH').strip().upper()
        total_sales -= amt
        if pm == 'CASH':
            cash_sales -= amt
        elif pm == 'QR':
            qr_sales -= amt
        elif pm == 'CARD':
            card_sales -= amt

    return total_sales, cash_sales, qr_sales, card_sales


def calculate_income_summary(net_sales, cash_sales, qr_sales, card_sales,
                             other_income, operating_expenses):
    """Return clearly named accounting totals for a period.

    ``net_sales`` is sales after refunds across every payment method. Refunds
    must not also be included in ``operating_expenses``.
    """
    total_income = float(net_sales) + float(other_income)
    net_profit = total_income - float(operating_expenses)
    return {
        'net_sales': float(net_sales),
        'cash_sales': float(cash_sales),
        'qr_sales': float(qr_sales),
        'card_sales': float(card_sales),
        'other_income': float(other_income),
        'total_income': total_income,
        'operating_expenses': float(operating_expenses),
        'net_profit': net_profit,
    }


def calculate_kuih_batch_summary(cost_amount, actual_sales, full_sale_value=0):
    """Return the simple daily cash result for one kuih batch.

    Kuih is purchased with cash in the morning and the actual cash collected is
    recorded when the batch is settled.  Unsold pieces are consumed internally,
    so they create no extra cash transaction and are already included in the
    batch cost.
    """
    cost = float(cost_amount or 0)
    sales = float(actual_sales or 0)
    target = float(full_sale_value or 0)
    return {
        'cost_amount': cost,
        'actual_sales': sales,
        'full_sale_value': target,
        'expected_margin': target - cost,
        'actual_margin': sales - cost,
        'unsold_value': max(0.0, target - sales),
        'drawer_net': sales - cost,
    }


# ================= DATABASE MANAGER =================
class DatabaseManager:
    def __init__(self, db_filename="pos.db"):
        # Get the directory where the .exe or .py file is located
        if getattr(sys, 'frozen', False):
            # If running as an EXE
            self.base_dir = os.path.dirname(sys.executable)
        else:
            # If running as a script
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
            
        # Check if a 'data' subfolder exists, if so use it to keep things even cleaner
        data_dir = os.path.join(self.base_dir, "data")
        if os.path.exists(data_dir):
            self.db_path = os.path.join(data_dir, db_filename)
        else:
            self.db_path = os.path.join(self.base_dir, db_filename)
            
        self.init_database()

    @contextmanager
    def get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        # SQLite does not enforce declared foreign keys unless each connection
        # enables them explicitly. A short busy timeout also avoids immediate
        # "database is locked" failures during overlapping POS operations.
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA busy_timeout = 5000")
        
        # Register custom to_iso function to convert DD-MM-YYYY HH:MM:SS to YYYY-MM-DD HH:MM:SS
        def to_iso(dt_str):
            if not dt_str:
                return ""
            try:
                if len(dt_str) >= 19 and dt_str[2] == '-' and dt_str[5] == '-':
                    return f"{dt_str[6:10]}-{dt_str[3:5]}-{dt_str[0:2]}{dt_str[10:19]}"
            except Exception:
                pass
            return dt_str
        conn.create_function("to_iso", 1, to_iso)
        
        try:
            yield conn
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()

    def init_database(self):
        with self.get_connection() as conn:
            c = conn.cursor()

            c.execute("""
            CREATE TABLE IF NOT EXISTS items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE,
                price REAL,
                category TEXT DEFAULT 'Kopi',
                stock INTEGER DEFAULT 0,
                low_stock_threshold INTEGER DEFAULT 5,
                active INTEGER DEFAULT 1,
                track_stock INTEGER DEFAULT 0,
                created_at TEXT,
                barcode TEXT UNIQUE
            )""")
            try:
                c.execute("ALTER TABLE items ADD COLUMN track_stock INTEGER DEFAULT 0")
            except Exception:
                pass

            c.execute("""
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datetime TEXT,
                total REAL,
                payment TEXT,
                items TEXT,
                discount REAL DEFAULT 0,
                tax REAL DEFAULT 0,
                cashier TEXT DEFAULT 'Admin',
                table_code TEXT DEFAULT NULL,
                amended INTEGER DEFAULT 0,
                amended_from INTEGER DEFAULT NULL
            )""")

            # Migrate: add amended columns if missing
            try:
                c.execute("ALTER TABLE sales ADD COLUMN table_code TEXT DEFAULT NULL")
            except Exception: pass

            c.execute("CREATE INDEX IF NOT EXISTS idx_sales_datetime ON sales(datetime)")

            c.execute("""
            CREATE TABLE IF NOT EXISTS daily_summary (
                date TEXT PRIMARY KEY,
                total_sales REAL,
                transaction_count INTEGER,
                cash_total REAL,
                qr_total REAL,
                card_total REAL,
                total_discount REAL DEFAULT 0
            )""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                action TEXT,
                details TEXT,
                user TEXT DEFAULT 'Admin'
            )""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS staff (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                staff_id TEXT UNIQUE,
                name TEXT,
                pin TEXT,
                role TEXT DEFAULT 'cashier',
                active INTEGER DEFAULT 1,
                created_at TEXT
            )""")
            c.execute("""INSERT OR IGNORE INTO staff (staff_id, name, pin, role, created_at)
                         VALUES ('ADMIN', 'Admin', '919888', 'admin', datetime('now'))""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS item_variants (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER,
                variant_name TEXT,
                price_adjustment REAL DEFAULT 0,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS stored_beverages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_id INTEGER,
                item_name TEXT,
                quantity INTEGER,
                stored_by TEXT,
                stored_staff TEXT,
                stored_at DATETIME,
                pickup_deadline DATETIME,
                status TEXT DEFAULT 'stored',
                picked_qty INTEGER DEFAULT 0,
                picked_by TEXT,
                picked_staff TEXT,
                picked_at DATETIME
            )""")
            for col_sql in [
                "ALTER TABLE stored_beverages ADD COLUMN stored_staff TEXT",
                "ALTER TABLE stored_beverages ADD COLUMN picked_qty INTEGER DEFAULT 0",
                "ALTER TABLE stored_beverages ADD COLUMN picked_by TEXT",
                "ALTER TABLE stored_beverages ADD COLUMN picked_staff TEXT",
                "ALTER TABLE stored_beverages ADD COLUMN picked_at DATETIME",
            ]:
                try:
                    c.execute(col_sql)
                except Exception:
                    pass

            c.execute("""
            CREATE TABLE IF NOT EXISTS beverage_customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                phone TEXT NOT NULL,
                notes TEXT,
                beer_balance INTEGER DEFAULT 0,
                active INTEGER DEFAULT 1,
                created_at TEXT,
                updated_at TEXT
            )""")
            c.execute("""CREATE INDEX IF NOT EXISTS idx_beverage_customers_phone
                         ON beverage_customers(phone)""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS beverage_transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id INTEGER NOT NULL,
                tx_time TEXT NOT NULL,
                tx_type TEXT NOT NULL,
                item_name TEXT,
                quantity INTEGER NOT NULL,
                balance_before INTEGER DEFAULT 0,
                balance_after INTEGER DEFAULT 0,
                staff TEXT,
                notes TEXT
            )""")
            for col_sql in [
                "ALTER TABLE beverage_customers ADD COLUMN notes TEXT",
                "ALTER TABLE beverage_customers ADD COLUMN beer_balance INTEGER DEFAULT 0",
                "ALTER TABLE beverage_customers ADD COLUMN active INTEGER DEFAULT 1",
                "ALTER TABLE beverage_customers ADD COLUMN created_at TEXT",
                "ALTER TABLE beverage_customers ADD COLUMN updated_at TEXT",
                "ALTER TABLE beverage_transactions ADD COLUMN item_name TEXT",
                "ALTER TABLE beverage_transactions ADD COLUMN balance_before INTEGER DEFAULT 0",
                "ALTER TABLE beverage_transactions ADD COLUMN balance_after INTEGER DEFAULT 0",
                "ALTER TABLE beverage_transactions ADD COLUMN notes TEXT",
            ]:
                try:
                    c.execute(col_sql)
                except Exception:
                    pass

            c.execute("""
            CREATE TABLE IF NOT EXISTS cash_transactions (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                datetime    TEXT NOT NULL,
                type        TEXT NOT NULL,
                category    TEXT NOT NULL,
                amount      REAL NOT NULL,
                description TEXT,
                recorded_by TEXT,
                date        TEXT NOT NULL,
                payment_method TEXT NOT NULL DEFAULT 'CASH',
                supplier    TEXT,
                quantity    REAL,
                item_name   TEXT
            )""")

            try:
                c.execute("ALTER TABLE cash_transactions ADD COLUMN supplier TEXT")
            except Exception: pass
            try:
                c.execute("ALTER TABLE cash_transactions ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
            except Exception: pass
            try:
                c.execute("ALTER TABLE cash_transactions ADD COLUMN quantity REAL")
            except Exception: pass
            try:
                c.execute("ALTER TABLE cash_transactions ADD COLUMN item_name TEXT")
            except Exception: pass
            try:
                c.execute("""UPDATE cash_transactions SET payment_method='CASH'
                             WHERE payment_method IS NULL OR TRIM(payment_method)=''""")
            except Exception: pass

            # Admin-managed ordinary expense choices shown in Cash Out.  System
            # workflows such as tracked stock purchases, staff advances and cash
            # transfers remain protected in code because they have extra ledger
            # side effects.
            c.execute("""
            CREATE TABLE IF NOT EXISTS expense_options (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                option_key    TEXT NOT NULL UNIQUE,
                name_en       TEXT NOT NULL,
                name_zh       TEXT,
                category      TEXT NOT NULL,
                detail_mode   TEXT NOT NULL DEFAULT 'fixed',
                quantity_mode TEXT NOT NULL DEFAULT 'none',
                stock_item_id INTEGER,
                help_en       TEXT,
                help_zh       TEXT,
                sort_order    INTEGER NOT NULL DEFAULT 0,
                active        INTEGER NOT NULL DEFAULT 1,
                created_at    TEXT,
                updated_at    TEXT
            )""")
            expense_option_columns = {
                str(row['name'])
                for row in c.execute("PRAGMA table_info(expense_options)").fetchall()
            }
            expense_stock_link_migration = "stock_item_id" not in expense_option_columns
            if expense_stock_link_migration:
                c.execute("ALTER TABLE expense_options ADD COLUMN stock_item_id INTEGER")
            c.execute(
                "CREATE INDEX IF NOT EXISTS idx_expense_options_active_sort "
                "ON expense_options(active, sort_order, id)"
            )
            expense_option_defaults = [
                (
                    "ice", "Ice", "冰块", "Ingredients", "fixed", "required",
                    "Enter the number of bags and the total cash actually paid. Ice is an ingredient, not lasting stock.",
                    "填写冰块袋数与实际支付总额；冰块归入食材，不作为长期库存。", 10,
                ),
                (
                    "ingredients", "Ingredients", "食材／原料", "Ingredients", "fixed", "none",
                    "Write what was bought, such as coffee powder or condensed milk, in Notes and enter the total paid.",
                    "在备注填写咖啡粉、炼奶等实际购买物品，再输入实付总额。", 20,
                ),
                (
                    "small_tools", "Small Tools & Utensils", "小工具／餐具",
                    "Small Tools & Utensils", "fixed", "none",
                    "Write the item, such as spoon, tong or kettle, in Notes. Record the full purchase cost once.",
                    "在备注填写汤匙、夹子、水壶等名称，并一次记录全部实付金额。", 30,
                ),
                (
                    "equipment_furniture", "Equipment & Furniture", "设备／家具",
                    "Equipment & Furniture", "fixed", "none",
                    "Use only when paid from this drawer. Write the item, such as chair, in Notes and enter the total paid.",
                    "只记录由本钱箱付款的设备或家具；在备注填写椅子等物品并输入实付总额。", 40,
                ),
                (
                    "daily_supplies", "Daily Supplies", "日常用品", "Daily Supplies", "fixed", "none",
                    "For cleaning and general shop supplies. Write the item in Notes and enter the total cash paid.",
                    "用于清洁及店内日用品；在备注填写物品名称并输入实际支付总额。", 50,
                ),
                (
                    "packaging", "Packaging", "包装用品", "Packaging", "fixed", "none",
                    "Write the packaging item in Notes and enter the total cash paid.",
                    "在备注填写包装用品名称并输入实际支付总额。", 60,
                ),
                (
                    "repair", "Repair & Maintenance", "维修保养", "Repair & Maintenance", "fixed", "none",
                    "Describe the repair and enter the total paid.", "填写维修内容与实付总额。", 70,
                ),
                (
                    "transport", "Transport", "交通运输", "Transport", "fixed", "none",
                    "Describe the transport cost and enter the total paid.", "填写交通用途与实付总额。", 80,
                ),
                (
                    "utilities", "Utilities & Licenses", "水电与执照", "Utilities & Licenses", "fixed", "none",
                    "Describe the bill or licence and enter the total paid.", "填写账单或执照名称与实付总额。", 90,
                ),
                (
                    "staff_welfare", "Staff Welfare", "员工福利", "Staff Welfare", "fixed", "none",
                    "Describe the staff expense and enter the total paid.", "填写员工福利内容与实付总额。", 100,
                ),
                (
                    "other_expense", "Other Expense", "其他营业开销", "Other Expense", "fixed", "none",
                    "Describe the cash expense clearly and enter the total paid.",
                    "清楚填写现金开销内容与实付总额。", 110,
                ),
            ]
            expense_options_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            c.executemany(
                """INSERT OR IGNORE INTO expense_options
                   (option_key,name_en,name_zh,category,detail_mode,quantity_mode,
                    help_en,help_zh,sort_order,active,created_at,updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,1,?,?)""",
                [
                    (*option_row, expense_options_now, expense_options_now)
                    for option_row in expense_option_defaults
                ],
            )
            if expense_stock_link_migration:
                # Existing untouched built-in choices used the old form's Item
                # Details/Quantity fields.  Migrate them once to the simpler
                # Notes-first form; admin-edited rows keep their settings.
                c.execute(
                    """UPDATE expense_options
                       SET detail_mode='fixed', quantity_mode='none',
                           help_en=NULL, help_zh=NULL, updated_at=?
                       WHERE option_key IN (
                           'ingredients','small_tools','equipment_furniture',
                           'daily_supplies','packaging','repair','transport',
                           'utilities','staff_welfare','other_expense'
                       )
                         AND (updated_at IS NULL OR created_at IS NULL OR updated_at=created_at)""",
                    (expense_options_now,),
                )

            # One simple batch per business day: morning cash cost followed by
            # the actual cash collected at settlement.  Leftovers are consumed
            # internally and therefore do not generate another money movement.
            c.execute("""
            CREATE TABLE IF NOT EXISTS kuih_batches (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                date                 TEXT NOT NULL UNIQUE,
                opened_at            TEXT NOT NULL,
                opened_by            TEXT,
                cost_amount          REAL NOT NULL,
                full_sale_value      REAL DEFAULT 0,
                actual_sales         REAL,
                status               TEXT NOT NULL DEFAULT 'OPEN',
                leftover_action      TEXT DEFAULT 'INTERNAL_CONSUMPTION',
                purchase_notes       TEXT,
                settlement_notes     TEXT,
                settled_at           TEXT,
                settled_by           TEXT,
                cost_transaction_id  INTEGER,
                sales_transaction_id INTEGER
            )""")
            c.execute("CREATE INDEX IF NOT EXISTS idx_kuih_batches_date ON kuih_batches(date)")

            # Salary advances reduce the drawer but are deliberately separate
            # from operating expenses so they are not presented as shop costs.
            c.execute("""
            CREATE TABLE IF NOT EXISTS staff_advances (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                date                TEXT NOT NULL,
                datetime            TEXT NOT NULL,
                staff_id            TEXT,
                staff_name          TEXT NOT NULL,
                amount              REAL NOT NULL,
                deduction_month     TEXT,
                status              TEXT NOT NULL DEFAULT 'OUTSTANDING',
                notes               TEXT,
                recorded_by         TEXT,
                cash_transaction_id INTEGER,
                settled_at          TEXT,
                settled_by          TEXT
            )""")
            c.execute("CREATE INDEX IF NOT EXISTS idx_staff_advances_staff ON staff_advances(staff_id, status)")

            c.execute("""
            CREATE TABLE IF NOT EXISTS product_categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE,
                icon TEXT,
                color TEXT,
                active INTEGER DEFAULT 1,
                sort_order INTEGER DEFAULT 0,
                exclude_export INTEGER DEFAULT 0
            )""")

            try:
                c.execute("ALTER TABLE product_categories ADD COLUMN exclude_export INTEGER DEFAULT 0")
            except Exception:
                pass

            # Backfill default categories if empty
            ensure_list_price_columns(conn)
            c.execute("SELECT COUNT(*) FROM product_categories")
            if c.fetchone()[0] == 0:
                defaults = [
                    ("Food", "🍛", "#78716C", 1),
                    ("Kopi", "☕", "#57534E", 2),
                    ("Teh", "🍵", "#92400E", 3),
                    ("Beer", "🍺", "#854D0E", 4),
                    ("Cans", "🥤", "#1E40AF", 5),
                    ("Drinks", "🧃", "#0F766E", 6),
                ]
                c.executemany("INSERT INTO product_categories (name, icon, color, sort_order) VALUES (?,?,?,?)", defaults)

            c.execute("""
            CREATE TABLE IF NOT EXISTS cash_closing (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                closing_date TEXT NOT NULL,
                pos_cash     REAL DEFAULT 0,
                other_income REAL DEFAULT 0,
                expenses     REAL DEFAULT 0,
                expected_cash REAL DEFAULT 0,
                actual_cash  REAL DEFAULT 0,
                difference   REAL DEFAULT 0,
                cashier      TEXT,
                notes        TEXT,
                closed_at    TEXT
            )""")
            for _col in [
                "ALTER TABLE cash_closing ADD COLUMN pos_cash REAL DEFAULT 0",
                "ALTER TABLE cash_closing ADD COLUMN other_income REAL DEFAULT 0",
                "ALTER TABLE cash_closing ADD COLUMN expenses REAL DEFAULT 0",
                "ALTER TABLE cash_closing ADD COLUMN expected_cash REAL DEFAULT 0",
                "ALTER TABLE cash_closing ADD COLUMN opening_cash REAL DEFAULT 0",
            ]:
                try: c.execute(_col)
                except Exception: pass

            c.execute("""
            CREATE TABLE IF NOT EXISTS refunds (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_id        INTEGER NOT NULL,
                datetime       TEXT NOT NULL,
                items          TEXT NOT NULL,
                refund_amount  REAL NOT NULL,
                reason         TEXT,
                processed_by   TEXT,
                payment_method TEXT
            )""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS shift_log (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                staff_id     TEXT NOT NULL,
                staff_name   TEXT NOT NULL,
                action       TEXT NOT NULL,
                datetime     TEXT NOT NULL,
                date         TEXT NOT NULL,
                opening_cash REAL DEFAULT 0,
                notes        TEXT
            )""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS cash_opening (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                date         TEXT NOT NULL,
                opening_cash REAL DEFAULT 0,
                opened_by    TEXT,
                opened_at    TEXT
            )""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS rent_tenants (
                id                INTEGER PRIMARY KEY AUTOINCREMENT,
                name              TEXT NOT NULL,
                stall_no          TEXT,
                rental_type       TEXT DEFAULT 'Daily',
                daily_rate        REAL DEFAULT 0,
                total_days_paid   REAL DEFAULT 0,
                paid_until_date   TEXT,
                last_payment_date TEXT,
                start_date        TEXT,
                active            INTEGER DEFAULT 1,
                notes             TEXT
            )""")

            c.execute("""
            CREATE TABLE IF NOT EXISTS rent_payments (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                tenant_id   INTEGER,
                tenant_name TEXT,
                stall_no    TEXT,
                days        REAL,
                rate        REAL,
                amount      REAL,
                pay_date    TEXT,
                datetime    TEXT,
                recorded_by TEXT,
                payment_method TEXT NOT NULL DEFAULT 'CASH',
                notes       TEXT,
                voided      INTEGER DEFAULT 0
            )""")

            # Legacy tables kept for compatibility
            try:
                c.execute("""CREATE TABLE IF NOT EXISTS rent_stalls (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT,
                    daily_rate REAL DEFAULT 0, monthly_rate REAL DEFAULT 0,
                    active INTEGER DEFAULT 1, notes TEXT)""")
            except Exception: pass

            # ── Cashflow: single source of truth for all money movement ──
            c.execute("""
            CREATE TABLE IF NOT EXISTS cashflow (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                date           TEXT NOT NULL,
                datetime       TEXT NOT NULL,
                type           TEXT NOT NULL,
                source         TEXT NOT NULL,
                reference_id   INTEGER,
                amount         REAL NOT NULL,
                payment_method TEXT NOT NULL,
                recorded_by    TEXT,
                notes          TEXT,
                supplier       TEXT
            )""")

            try:
                c.execute("ALTER TABLE cashflow ADD COLUMN supplier TEXT")
            except Exception: pass

            c.execute("""
            CREATE TABLE IF NOT EXISTS held_orders (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at   TEXT NOT NULL,
                updated_at   TEXT NOT NULL,
                cashier      TEXT,
                items        TEXT NOT NULL,
                subtotal     REAL NOT NULL,
                tax          REAL NOT NULL,
                total        REAL NOT NULL,
                note         TEXT,
                status       TEXT DEFAULT 'HELD'
            )""")
            
            c.execute("""
            CREATE TABLE IF NOT EXISTS item_quantity_sets (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id      INTEGER NOT NULL,
                quantity     INTEGER NOT NULL,
                set_price    REAL NOT NULL,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )""")
            try:
                c.execute("ALTER TABLE held_orders ADD COLUMN table_code TEXT")
            except Exception:
                pass

            c.execute("CREATE INDEX IF NOT EXISTS idx_held_orders_status ON held_orders(status)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_held_orders_table_code ON held_orders(table_code)")
            # A table can have multiple independent unpaid bills. Older web
            # versions installed this constraint in the shared database.
            c.execute("DROP INDEX IF EXISTS idx_held_orders_one_active_table")

            # Immutable inventory ledger. ``items.stock`` remains the fast current
            # balance, while this table explains exactly how that balance moved.
            c.execute("""
            CREATE TABLE IF NOT EXISTS stock_adjustment_batches (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                datetime       TEXT NOT NULL,
                source_type    TEXT NOT NULL,
                item_count     INTEGER NOT NULL,
                total_quantity INTEGER NOT NULL,
                recorded_by    TEXT,
                notes          TEXT,
                request_token  TEXT NOT NULL UNIQUE
            )""")
            c.execute(
                "CREATE INDEX IF NOT EXISTS idx_stock_batches_datetime "
                "ON stock_adjustment_batches(datetime, id)"
            )
            c.execute("""
            CREATE TABLE IF NOT EXISTS stock_movements (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                datetime       TEXT NOT NULL,
                item_id        INTEGER NOT NULL,
                item_name      TEXT NOT NULL,
                quantity_delta INTEGER NOT NULL,
                stock_before   INTEGER NOT NULL,
                stock_after    INTEGER NOT NULL,
                source_type    TEXT NOT NULL,
                reference_id   INTEGER,
                recorded_by    TEXT,
                notes          TEXT,
                FOREIGN KEY (item_id) REFERENCES items(id)
            )""")
            c.execute("CREATE INDEX IF NOT EXISTS idx_stock_movements_item ON stock_movements(item_id, id)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_stock_movements_source ON stock_movements(source_type, reference_id)")
            # Establish a trustworthy starting point for existing installations.
            # A zero-delta baseline records the balance without altering it.
            c.execute("""INSERT INTO stock_movements
                         (datetime,item_id,item_name,quantity_delta,stock_before,stock_after,
                          source_type,reference_id,recorded_by,notes)
                         SELECT ?,i.id,i.name,0,i.stock,i.stock,'MIGRATION_BASELINE',NULL,'System',
                                'Initial ledger balance'
                         FROM items i
                         WHERE i.track_stock=1
                           AND NOT EXISTS (SELECT 1 FROM stock_movements sm WHERE sm.item_id=i.id)""",
                      (datetime.now().strftime("%d-%m-%Y %H:%M:%S"),))

            c.execute("""
            CREATE TABLE IF NOT EXISTS cashflow_categories (
                id   INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                type TEXT NOT NULL
            )""")

            # Seed default flow categories if empty
            c.execute("SELECT COUNT(*) FROM cashflow_categories")
            if c.fetchone()[0] == 0:
                flow_defaults = [
                    ("Ingredients", "Expense"), 
                    ("Utilities", "Expense"),
                    ("Staff Wages", "Expense"),
                    ("Repairs", "Expense"),
                    ("Packaging", "Expense"),
                    ("Transport", "Expense"),
                    ("Beverage Cost", "Expense"),
                    ("Maintenance", "Expense"),
                    ("License", "Expense"),
                    ("Stall Rental", "Income"),
                    ("Kuih Sales", "Income"),
                    ("Other", "Expense")
                ]
                c.executemany("INSERT INTO cashflow_categories (name, type) VALUES (?,?)", flow_defaults)

            # Keep every admin-managed expense report group available to filters.
            # This follows expense_options dynamically so an admin rename is not
            # undone by re-seeding an old hard-coded name on the next startup.
            c.execute(
                """INSERT INTO cashflow_categories (name,type)
                   SELECT DISTINCT eo.category,'Expense'
                   FROM expense_options eo
                   WHERE TRIM(COALESCE(eo.category,''))<>''
                     AND NOT EXISTS (
                         SELECT 1 FROM cashflow_categories cc
                         WHERE cc.name=eo.category AND cc.type='Expense'
                     )"""
            )

            # System workflow groups remain protected because application logic
            # records them by these canonical names.
            required_flow_categories = [
                ("Stock Purchase", "Expense"),
                ("Staff Advance", "Expense"),
                ("Cash Transfer", "Expense"),
                ("Kuih Purchase", "Expense"),
                ("Kuih Sales", "Income"),
            ]
            for flow_name, flow_type in required_flow_categories:
                c.execute(
                    """INSERT INTO cashflow_categories (name,type)
                       SELECT ?,? WHERE NOT EXISTS (
                           SELECT 1 FROM cashflow_categories WHERE name=? AND type=?
                       )""",
                    (flow_name, flow_type, flow_name, flow_type)
                )
            c.execute(
                """DELETE FROM cashflow_categories
                   WHERE type='Income'
                     AND name IN ('Gas Sales','Beverage Agency','Other')
                     AND NOT EXISTS (
                         SELECT 1 FROM cashflow
                         WHERE cashflow.type='Income'
                           AND cashflow.source=cashflow_categories.name
                     )"""
            )

            c.execute("""
            CREATE TABLE IF NOT EXISTS dining_tables (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                table_code  TEXT UNIQUE NOT NULL,
                sort_order  INTEGER DEFAULT 0,
                active      INTEGER DEFAULT 1,
                created_at  TEXT
            )""")
            c.execute("SELECT COUNT(*) FROM dining_tables")
            tcount = int(c.fetchone()[0] or 0)
            defaults = [
                "A1", "A2", "A3", "A4", "A5", "A6",
                "B1", "B2", "B3",
                "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8",
                "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9",
            ]
            if tcount == 0:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for idx, code in enumerate(defaults, start=1):
                    c.execute(
                        "INSERT INTO dining_tables (table_code, sort_order, active, created_at) VALUES (?,?,1,?)",
                        (code, idx, now_str)
                    )
            else:
                # Backfill newly introduced standard tables for existing DBs.
                c.execute("SELECT table_code FROM dining_tables")
                existing_codes = {str(r[0]) for r in c.fetchall()}
                c.execute("SELECT COALESCE(MAX(sort_order), 0) FROM dining_tables")
                next_sort = int(c.fetchone()[0] or 0)
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for code in defaults:
                    if code not in existing_codes:
                        next_sort += 1
                        c.execute(
                            "INSERT INTO dining_tables (table_code, sort_order, active, created_at) VALUES (?,?,1,?)",
                            (code, next_sort, now_str)
                        )
            # Migrate pay_date column onto rent_payments if missing (legacy DB compatibility)
            try: c.execute("ALTER TABLE rent_payments ADD COLUMN pay_date TEXT")
            except Exception: pass
            # Backfill pay_date for old rows where possible
            try:
                c.execute("""UPDATE rent_payments
                             SET pay_date = CASE
                                 WHEN pay_date IS NOT NULL AND TRIM(pay_date) != '' THEN pay_date
                                 WHEN datetime IS NOT NULL AND LENGTH(datetime) >= 10 THEN SUBSTR(datetime, 1, 10)
                                 ELSE DATE('now')
                             END
                             WHERE pay_date IS NULL OR TRIM(pay_date) = ''""")
            except Exception:
                pass
            # Migrate voided column onto rent_payments if missing
            try: c.execute("ALTER TABLE rent_payments ADD COLUMN voided INTEGER DEFAULT 0")
            except Exception: pass
            # Persist how rent was received. Existing records were cash-only.
            try: c.execute("ALTER TABLE rent_payments ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
            except Exception: pass
            try:
                c.execute("""UPDATE rent_payments SET payment_method='CASH'
                             WHERE payment_method IS NULL OR TRIM(payment_method)=''""")
            except Exception: pass

            # Migrate start_date column onto rent_tenants if missing
            try: c.execute("ALTER TABLE rent_tenants ADD COLUMN start_date TEXT")
            except Exception: pass
            try:
                c.execute(
                    """UPDATE rent_tenants
                       SET start_date = CASE
                           WHEN start_date IS NOT NULL AND TRIM(start_date) != '' THEN start_date
                           WHEN last_payment_date IS NOT NULL AND TRIM(last_payment_date) != '' THEN last_payment_date
                           ELSE DATE('now')
                       END
                       WHERE start_date IS NULL OR TRIM(start_date) = ''"""
                )
            except Exception:
                pass

            # Ensure legacy records appear in cashflow-based summaries.
            self._migrate_legacy_cashflow(conn)

    def _migrate_legacy_cashflow(self, conn):
        c = conn.cursor()

        def _normalize_biz_date(raw, fallback_dt):
            val = str(raw or '').strip()
            if not val:
                return get_business_date(fallback_dt)
            try:
                return datetime.strptime(val[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
            except Exception:
                pass
            try:
                return get_business_date(datetime.strptime(val[:10], "%d-%m-%Y"))
            except Exception:
                return get_business_date(fallback_dt)

        # Backfill legacy sales rows.
        try:
            c.execute("SELECT id, datetime, total, payment, cashier FROM sales")
            for r in c.fetchall():
                ref_id = int(r['id'])
                c.execute("SELECT 1 FROM cashflow WHERE source='Sales' AND reference_id=? LIMIT 1", (ref_id,))
                if c.fetchone():
                    continue

                dt_raw = (r['datetime'] or '').strip()
                try:
                    dt_obj = datetime.strptime(dt_raw, "%d-%m-%Y %H:%M:%S")
                except Exception:
                    try:
                        dt_obj = datetime.strptime(dt_raw[:19], "%Y-%m-%d %H:%M:%S")
                    except Exception:
                        dt_obj = datetime.now()

                c.execute(
                    """INSERT INTO cashflow
                       (date, datetime, type, source, reference_id, amount, payment_method, recorded_by, notes)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (
                        biz_date_of_sale(dt_raw),
                        dt_obj.strftime("%Y-%m-%d %H:%M:%S"),
                        'Income',
                        'Sales',
                        ref_id,
                        float(r['total'] or 0),
                        str(r['payment'] or 'CASH').upper(),
                        r['cashier'] or 'Legacy',
                        'Legacy sales backfill'
                    )
                )
        except Exception:
            pass

        # Backfill legacy cash income/expense rows.
        try:
            c.execute("SELECT id, datetime, date, type, category, amount, recorded_by, description, payment_method FROM cash_transactions")
            for r in c.fetchall():
                ref_id = int(r['id'])
                tx_type = str(r['type'] or '').strip().lower()
                cf_type = cash_transaction_flow_type(tx_type)
                if not cf_type:
                    # Never guess the accounting direction of an unknown type.
                    continue
                source = (r['category'] or 'Other').strip() or 'Other'

                # source + reference_id identify the underlying transaction.
                # Direction is accounting data, not part of its identity.
                c.execute(
                    "SELECT 1 FROM cashflow WHERE source=? AND reference_id=? LIMIT 1",
                    (source, ref_id)
                )
                if c.fetchone():
                    continue

                dt_raw = (r['datetime'] or '').strip()
                try:
                    dt_obj = datetime.strptime(dt_raw[:19], "%Y-%m-%d %H:%M:%S")
                except Exception:
                    try:
                        dt_obj = datetime.strptime(dt_raw, "%d-%m-%Y %H:%M:%S")
                    except Exception:
                        dt_obj = datetime.now()

                biz_date = _normalize_biz_date(r['date'], dt_obj)
                c.execute(
                    """INSERT INTO cashflow
                       (date, datetime, type, source, reference_id, amount, payment_method, recorded_by, notes)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (
                        biz_date,
                        dt_obj.strftime("%Y-%m-%d %H:%M:%S"),
                        cf_type,
                        source,
                        ref_id,
                        float(r['amount'] or 0),
                        str(r['payment_method'] or 'CASH').strip().upper(),
                        r['recorded_by'] or 'Legacy',
                        r['description'] or 'Legacy cash transaction backfill'
                    )
                )
        except Exception:
            pass

        # Backfill legacy rent payments.
        try:
            c.execute("SELECT id, datetime, pay_date, amount, recorded_by, tenant_name, payment_method FROM rent_payments WHERE COALESCE(voided,0)=0")
            for r in c.fetchall():
                # Unified naming to 'Stall Rental' to match Cashier and prevent duplication.
                ref_id = int(r['id'])
                c.execute("SELECT 1 FROM cashflow WHERE source='Stall Rental' AND reference_id=? LIMIT 1", (ref_id,))
                if c.fetchone():
                    # Cleanup old duplicate 'Rental' records created by previous naming bug
                    # only if we found the correct 'Stall Rental' entry.
                    c.execute("DELETE FROM cashflow WHERE source='Rental' AND reference_id=? AND type='Income'", (ref_id,))
                    continue

                dt_raw = (r['datetime'] or '').strip()
                try:
                    dt_obj = datetime.strptime(dt_raw, "%d-%m-%Y %H:%M:%S")
                except Exception:
                    try:
                        dt_obj = datetime.strptime(dt_raw[:19], "%Y-%m-%d %H:%M:%S")
                    except Exception:
                        dt_obj = datetime.now()

                biz_date = _normalize_biz_date(r['pay_date'], dt_obj)
                c.execute(
                    """INSERT INTO cashflow
                       (date, datetime, type, source, reference_id, amount, payment_method, recorded_by, notes)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (
                        biz_date,
                        dt_obj.strftime("%Y-%m-%d %H:%M:%S"),
                        'Income',
                        'Stall Rental',
                        ref_id,
                        float(r['amount'] or 0),
                        str(r['payment_method'] or 'CASH').strip().upper(),
                        r['recorded_by'] or 'Legacy',
                        (r['tenant_name'] or 'Legacy rent') + ' backfill'
                    )
                )
        except Exception:
            pass


# ================= POS APP =================
class POS:
    def __init__(self, root, startup_progress=None):
        def report_startup(value, message):
            if startup_progress is None:
                return
            try:
                startup_progress(value, message)
            except Exception:
                # Loading feedback must never prevent the POS from starting.
                pass

        report_startup(10, "Preparing workspace / 正在准备工作区…")
        self.root = root
        self.root.title("919 Restaurant")
        self.root.configure(bg="#F0F2F5")
        self.root.resizable(True, True)
        root_w, root_h = self._center_window(self.root, 1450, 850)
        self.root.minsize(min(960, root_w), min(620, root_h))

        report_startup(22, "Opening database / 正在打开数据库…")
        self.db = DatabaseManager()
        try:
            self._repaired_orphan_cashflow_count = self._repair_orphan_cashflow_refs()
        except Exception:
            self._repaired_orphan_cashflow_count = 0
        report_startup(38, "Loading settings / 正在读取设置…")

        self.order = []
        self._active_held_order_id = None
        self.current_table_code = None
        self._force_table_reload_once = False
        self.total = Decimal('0.00')
        self.payment = tk.StringVar(value="CASH")
        self.lang = "en"
        self.i18n = {
            "en": {
                "brand_subtitle": "Point of Sale",
                "hub": "⚡ Hub",
                "admin": "🔒 Admin",
                "admin_on": "🔓 Admin ON",
                "lang_btn": "中文",
                "main_tables": "🏠\nMain Tables",
                "current_order": " CURRENT ORDER",
                "table_label": "Table: {table}",
                "new_order": "New Order",
                "unpaid_order": "Unpaid",
                "editing_hold": "Editing Unpaid Order",
                "hold_f6": "🕒   HOLD   (F6)",
                "update_hold_f6": "🕒   UPDATE HOLD   (F6)",
                "checkout": "✅   CHECKOUT   (F1)",
                "table_flow_hint": "🪑 Choose Hold or Checkout for the selected table",
                "table_layout": "🪑  TABLE LAYOUT",
                "table_layout_hint": "Select a table to start order",
                "recent_orders": "Orders",
                "cash": "Cash Ops",
                "rent": "Collect Rent",
                "bev_storage": "Bev Storage",
                "stock": "Stock Adjust",
                "admin_short": "Admin",
                "no_active_tables": "No active tables",
                "hold_suffix": "{count} HOLD",
                "switch_table_title": "Switch Table",
                "switch_table_msg": "Current order has items. Switch to table {table}?",
                "select_order_title": "Select Order - Table {table}",
                "table_has_orders": "Table {table} has {count} orders",
                "select_order_hint": "Select order to open",
                "new_order_btn": "➕ New Order",
                "cancel": "✖ Cancel",
                "drawer": "🗄️ Drawer",
            },
            "zh": {
                "brand_subtitle": "收银系统",
                "hub": "⚡ 功能",
                "admin": "🔒 管理",
                "admin_on": "🔓 管理已开",
                "lang_btn": "EN",
                "main_tables": "🏠\n主桌位",
                "current_order": " 当前订单",
                "table_label": "桌号：{table}",
                "new_order": "新订单",
                "unpaid_order": "未结",
                "editing_hold": "编辑未结订单",
                "hold_f6": "🕒   挂单   (F6)",
                "update_hold_f6": "🕒   更新挂单   (F6)",
                "checkout": "✅   直接结账   (F1)",
                "table_flow_hint": "🪑 所选桌位可选择挂单或直接结账",
                "table_layout": "🪑  桌位布局",
                "table_layout_hint": "请选择桌位开始点单",
                "recent_orders": "订单",
                "cash": "现金操作",
                "rent": "收取租金",
                "bev_storage": "酒水寄存",
                "stock": "库存调整",
                "admin_short": "管理",
                "no_active_tables": "没有可用桌位",
                "hold_suffix": "{count} 挂单",
                "switch_table_title": "切换桌位",
                "switch_table_msg": "当前订单有商品，切换到桌位 {table} 吗？",
                "select_order_title": "选择订单 - 桌位 {table}",
                "table_has_orders": "桌位 {table} 有 {count} 个订单",
                "select_order_hint": "请选择要打开的订单",
                "new_order_btn": "➕ 新订单",
                "cancel": "✖ 取消",
                "drawer": "🗄️ 钱箱",
            },
        }

        # ── CHANGE 1: admin mode toggle ──────────────────────────────────
        self.admin_mode = False
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.search_items)

        self.tax_rate = Decimal('0.00')
        self.currency = "RM"

        self.categories = []
        self.category_colors = {}
        self.category_icons = {}
        self._cat_btns = {}
        self._cat_active_bg = {}
        self.load_categories()
        report_startup(52, "Loading categories / 正在读取商品分类…")

        self.colors = {
            'bg':         '#F0F2F5',
            'panel':      '#FFFFFF',
            'panel2':     '#F8F9FA',
            'border':     '#E2E6EA',
            'border2':    '#CBD2D9',
            'text':       '#1A202C',
            'text_sub':   '#4A5568',
            'text_muted': '#9E9E9E',
            'text_light': '#9E9E9E',
            'header':     '#1E293B',
            'header2':    '#334155',
            'success':    '#16A34A',
            'success_lt': '#DCFCE7',
            'danger':     '#DC2626',
            'danger_lt':  '#FEE2E2',
            'warning':    '#D97706',
            'warning_lt': '#FEF3C7',
            'neutral':    '#475569',
            'neutral_lt': '#F1F5F9',
            'cat_food':   '#78716C',
            'cat_kopi':   '#57534E',
            'cat_teh':    '#92400E',
            'cat_cans':   '#1D4ED8',
            'cat_beer':   '#854D0E',
            'cat_roti':   '#6B7280',
            'primary':    '#1E293B',
            'secondary':  '#475569',
            'info':       '#475569',
            'dark':       '#0F172A',
            'light':      '#F0F2F5',
            'white':      '#FFFFFF',
            'accent':     '#475569',
            'teal':       '#475569',
            'orange':     '#D97706',
            'card':       '#1E293B',
            'dim':        '#9E9E9E',
        }

        self.current_staff = None
        self._auth_fail_tracker = {}
        self._auto_localize_job = None
        self._last_staff_login_id = None

        # Generic UI text mapping for hardcoded labels/titles not yet wired to tr().
        self._ui_text_map_zh = {
            "Action Hub": "功能中心",
            "Search:": "搜索：",
            "Open": "打开",
            "Close": "关闭",
            "Admin Panel": "管理面板",
            "Staff Management": "员工管理",
            "Add Staff": "新增员工",
            "Role:": "角色：",
            "Table Maintenance": "桌位维护",
            "No active tables": "没有可用桌位",
            "ITEM": "项目",
            "QTY": "数量",
            "PRICE": "价格",
            "Subtotal:  RM 0.00": "小计：RM 0.00",
            "Tax (0%):  -": "税额 (0%)：-",
            "TOTAL  RM 0.00": "总计  RM 0.00",
            "Quantity": "数量",
            "QUANTITY": "数量",
            "Select an option:": "请选择选项：",
            "No items found\n\nAdd items in Item Management": "未找到商品\n\n请到商品管理新增商品",
            "Held Orders": "未结订单",
            "Select Payment": "选择付款方式",
            "Choose Payment Method": "请选择付款方式",
            "Cash Balance": "找零计算",
            "Cash Received": "实收现金",
            "Confirm": "确认",
            "Cancel": "取消",
            "Print": "打印",
            "Not Print": "不打印",
            "Print Receipt?": "打印小票？",
            "Confirm Print": "确认打印",
            "Recent Orders": "近期订单",
            "No action found.": "没有找到对应功能。",
            "Select Table": "选择桌位",
            "No Tables": "没有桌位",
            "No Selection": "未选择",
            "Loaded": "已载入",
            "Delete Unpaid Order": "删除未结订单",
            "Delete Held Order": "删除未结订单",
            "Empty Order": "空订单",
            "Select Table": "选择桌位",
            "Out of Stock": "缺货",
            "Insufficient Stock": "库存不足",
            "Invalid Amount": "金额无效",
            "Insufficient Cash": "现金不足",
            "Error": "错误",
            "Success": "成功",
            "Done": "完成",
            "No Records": "无记录",
            "Show:": "显示：",
            "1 Hour": "1小时",
            "3 Hours": "3小时",
            "Today": "今天",
            "Time": "时间",
            "Total": "总额",
            "Payment": "付款",
            "Cashier": "收银员",
            "Status": "状态",
            "Change Payment": "更改付款",
            "Actions": "操作",
            "Details": "详情",
            "Reprint": "重打",
            "Edit": "编辑",
            "Expense": "支出",
            "Income": "收入",
            "Amount  RM:": "金额 RM：",
            "Category:": "类别：",
            "Notes:": "备注：",
            "Save Expense": "保存支出",
            "Save Income": "保存收入",
            "No tenants configured.\nAsk admin to add tenants in Cash (Admin) -> Rent.": "尚未设置租户。\n请联系管理员在现金(管理) -> 租金中新增。",
            "Days:": "天数：",
            "Close Day": "结账日结",
            "End of Day Cash Check": "日结现金核对",
            "Expected Cash:": "应有现金：",
            "Actual Cash:": "实点现金：",
            "Difference:": "差额：",
            "Confirm Close Day": "确认结账日结",
            "Store Beverages": "酒水寄存",
            "Beverage Storage": "酒水寄存",
            "Quick Guide: 1) Search/Create customer  2) Choose Beer Item  3) Enter qty  4) Add or Deduct": "快速指南：1) 查找/创建客户  2) 选择啤酒  3) 输入数量  4) 增加或扣减",
            "No need to find order anymore. Use customer phone for fastest lookup.": "无需再找订单，使用客户电话可快速查询。",
            "Step 1: Find or Create Customer": "步骤1：查找或创建客户",
            "Phone:": "电话：",
            "Name:": "姓名：",
            "Find Customer": "查找客户",
            "Create Customer": "创建客户",
            "Customer List": "客户列表",
            "(Use phone to search quickly. If not found, create customer once and reuse.)": "（建议先用电话搜索；找不到时创建一次即可重复使用。）",
            "No customer selected": "尚未选择客户",
            "Current Beer Balance: 0": "当前啤酒余额：0",
            "Step 2: Add or Deduct Beer": "步骤2：增加或扣减啤酒",
            "Beer Item:": "啤酒项目：",
            "Qty:": "数量：",
            "(Add increases customer stored beers. Deduct when customer picks up.)": "（增加=寄存增加；扣减=客户取酒时扣除。）",
            "Step 3: Confirm Transaction": "步骤3：确认交易",
            "Find or create customer first, then add/deduct quantity.": "请先查找或创建客户，然后再增加/扣减数量。",
            "Confirm  (PIN Required)": "确认（需要PIN）",
            "Pick Up Existing Record": "客户取酒（扣减）",
            "Refresh List": "刷新列表",
            "Recent Beverage Transactions:": "近期酒水交易：",
            "Action": "操作",
            "Balance": "余额",
            "Staff": "员工",
            "Stock Control": "库存管理",
            "Tracked Stock Items Only": "仅显示追踪库存商品",
            "Reorder At": "补货警戒",
            "Stock Adjustment": "库存调整",
            "Selected Item: -": "已选商品：-",
            "Add": "增加",
            "Deduct": "扣减",
            "Apply": "应用",
            "Refresh": "刷新",
            "In Stock": "有库存",
            "Out of Stock": "缺货",
            "Low Stock": "低库存",

            # Notebook/tab texts across system
            "📋 Items": "📋 商品",
            "📦 Stock Control": "📦 库存管理",
            "📜 Sales History": "📜 销售历史",
            "📅 Daily Summary": "📅 每日汇总",
            "🏆 Top Items": "🏆 热销商品",
            "📦 Stock Status": "📦 库存状态",
            "📦 Daily Sales": "📦 每日销售",
            "📋 Audit Log": "📋 审计日志",
            "🌅 Opening / Shift": "🌅 开盘 / 班次",
            "💸 Expenses": "💸 支出",
            "💵 Other Income": "💵 其他收入",
            "🌙 End of Day": "🌙 日结",
            "🏪 Rent": "🏪 租金",

            # Common report tab labels (without emoji variants)
            "Sales History": "销售历史",
            "Daily Summary": "每日汇总",
            "Top Items": "热销商品",
            "Stock Status": "库存状态",
            "Daily Sales": "每日销售",
            "Audit Log": "审计日志",
            "Opening / Shift": "开盘 / 班次",
            "Other Income": "其他收入",
            "End of Day": "日结",

            # Shared list, dialog and management vocabulary
            "ID": "编号",
            "Staff ID": "员工编号",
            "Name": "名称",
            "Role": "角色",
            "Active": "启用",
            "Table": "桌位",
            "Sort": "排序",
            "Type": "类型",
            "Icon": "图标",
            "Color": "颜色",
            "Order": "顺序",
            "Export": "导出",
            "Date": "日期",
            "DateTime": "日期时间",
            "Price": "价格",
            "Category": "类别",
            "Stock Qty": "库存数量",
            "Stock Mode": "库存模式",
            "Barcode": "条码",
            "Stock": "库存",
            "Source / Category": "来源／类别",
            "Shop / Company": "店家／公司",
            "Notes": "备注",
            "Refund ID": "退款编号",
            "Order ID": "订单号",
            "Reason": "原因",
            "Tenant": "租户",
            "Stall": "档口",
            "Rate": "费率",
            "Method": "方式",
            "By": "处理人",
            "Updated": "更新时间",
            "Start Date": "开始日期",
            "Total Days": "累计天数",
            "Paid Until": "已付至",
            "Trans": "交易数",
            "Cash": "现金",

            "Staff Management": "员工管理",
            "👥 Staff Management": "👥 员工管理",
            "Add New Staff": "新增员工",
            "Staff ID:": "员工编号：",
            "PIN:": "PIN：",
            "➕ Add Staff": "➕ 新增员工",
            "🔑 Change PIN": "🔑 更改 PIN",
            "🔄 Change Role": "🔄 更改角色",
            "🔄 Toggle Active": "🔄 启用／停用",
            "🗑️ Remove": "🗑️ 删除",
            "💾 Save": "💾 保存",
            "💾 Update": "💾 更新",

            "🪑 Table Maintenance": "🪑 桌位维护",
            "➕ Add Table": "➕ 新增桌位",
            "Add Table": "新增桌位",
            "Table code (e.g. A7, B4):": "桌位编号（例如 A7、B4）：",

            "💸 Expense/Income Categories": "💸 收支类别",
            "💸 Expense/Income Category Management": "💸 收支类别管理",
            "📂 Category Management": "📂 商品类别管理",
            "Add Category": "新增类别",
            "Add New Category": "新增类别",
            "Edit Category": "编辑类别",
            "➕ Add Category": "➕ 新增类别",
            "✏️ Edit Category": "✏️ 编辑类别",
            "🗑 Delete Selected": "🗑 删除所选",
            "Icon (Emoji):": "图标（Emoji）：",
            "Color (Hex):": "颜色（Hex）：",
            "Sort Order (Position):": "排序位置：",
            "Exclude from Export:": "不包括在导出中：",
            "🎨 Pick": "🎨 选择",

            "Item Management": "商品管理",
            "⚙️ Item Management": "⚙️ 商品管理",
            "📦 Stock Control": "📦 库存管理",
            "ACTIONS": "操作",
            "Add New Item": "新增商品",
            "Edit Item": "编辑商品",
            "Track Stock:": "追踪库存：",
            "Track Quantity": "追踪数量",
            "Low Stock Threshold:": "低库存警戒：",
            "Barcode (optional):": "条码（可选）：",
            "Barcode:": "条码：",
            "💾 Save Item": "💾 保存商品",
            "💾 Update Item": "💾 更新商品",
            "➕ Add Item": "➕ 新增商品",
            "📝 Edit Item": "📝 编辑商品",
            "🗑️ Delete": "🗑️ 删除",
            "🌡️ Variants": "🌡️ 选项",
            "📦 Sets": "📦 数量组合",
            "No Tracking": "不追踪",
            "Track Stock": "追踪库存",
            "Note:": "备注：",
            "💾 Apply": "💾 应用",

            "💰 POS Cashflow Summary": "💰 POS 现金流汇总",
            "POS-recorded movement only · Not final accounting profit": "仅限 POS 已记录资金流 · 不是最终会计利润",
            "From:": "开始：",
            "To:": "结束：",
            "Type:": "类型：",
            "Shop:": "店家：",
            "Quick range:": "快捷范围：",
            "This Month": "本月",
            "Last Month": "上个月",
            "Last 30 Days": "最近30天",
            "✓ Apply Dates": "✓ 应用日期",
            "💵 Recorded Cash In": "💵 已记录流入",
            "🧾 Recorded Cash Out": "🧾 已记录流出",
            "📈 Recorded Balance": "📈 已记录差额",
            "➕ Add New": "➕ 新增",
            "✏️ Edit Selected": "✏️ 编辑所选",

            "↩️ Refund Management": "↩️ 退款管理",
            "Select Order to Refund:": "选择要退款的订单：",
            "Refund Records:": "退款记录：",
            "Select items to refund:": "选择要退款的商品：",
            "Reason:": "原因：",
            "Return sellable items to stock": "可再销售的商品退回库存",
            "Refund Amount: RM 0.00": "退款金额：RM 0.00",
            "✅ Confirm Refund": "✅ 确认退款",
            "↩️ Process Refund": "↩️ 处理退款",

            "🏪 Stall Management": "🏪 档口管理",
            "➕  Add New Tenant": "➕  新增租户",
            "🏪  All Tenants": "🏪  所有租户",
            "📅  Today's Rent Payments": "📅  今日租金付款",
            "📜  Payment History": "📜  付款历史",
            "Stall No:": "档口编号：",
            "Freq:": "周期：",
            "Rate (RM):": "费率 (RM)：",
            "Start Date:": "开始日期：",
            "➕  Add Tenant": "➕  新增租户",
            "🔄 Sync All Dates": "🔄 同步所有日期",
            "✏ Edit": "✏ 编辑",
            "🗑 Void Selected Payment": "🗑 作废所选付款",

            "Reports & Analytics": "报表与分析",
            "📊 Reports & Analytics": "📊 报表与分析",
            "Payment Reconciliation": "付款对账",
            "💰 Payment Reconciliation": "💰 付款对账",
            "Select Date:": "选择日期：",
            "Discrepancies:": "差异记录：",
            "🔄 Reconcile Again": "🔄 重新对账",
            "Previous Day": "前一天",
            "Last 7 Days": "最近7天",
            "Custom": "自定义",
            "All": "全部",
            "Filter:": "筛选：",
            "Period:": "周期：",

            "✖ Close": "✖ 关闭",
            "❌ Close": "❌ 关闭",
            "✖ Cancel": "✖ 取消",
            "🔍 Search": "🔍 查询",
            "✖ Clear": "✖ 清除",
            "🔄 Refresh": "🔄 刷新",
            "⬇ Export": "⬇ 导出",
            "Export Success": "导出成功",
            "Export Failed": "导出失败",
            "No Data": "没有资料",
            "Confirm Delete": "确认删除",
            "Invalid Date": "日期无效",
            "Invalid Date Range": "日期范围无效",
            "Missing": "缺少资料",
            "Invalid": "输入无效",
            "Saved": "已保存",
        }

        # Keep popup windows consistently centered when opened.
        self._setup_auto_center_toplevels()
        self._setup_auto_sortable_treeviews()

        report_startup(65, "Building interface / 正在建立操作界面…")
        self.build_ui()
        report_startup(82, "Loading items / 正在读取商品…")
        self.load_items()
        self.update_time()
        self.update_order_number()
        report_startup(94, "Applying shortcuts and language / 正在套用语言与快捷键…")
        self.setup_keyboard_shortcuts()
        self._start_auto_localizer()
        self._notify_startup_data_repair()
        report_startup(100, "Ready / 准备完成")

    def load_categories(self):
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT name, icon, color FROM product_categories WHERE active=1 ORDER BY sort_order")
            rows = c.fetchall()
            self.category_list_prices = category_list_prices(conn)
        
        self.categories = [r['name'] for r in rows]
        self.category_icons = {r['name']: r['icon'] for r in rows}
        self.category_colors = {r['name']: r['color'] for r in rows}
        
        if hasattr(self, '_cat_btn_frame'):
            self._refresh_category_pills()

    def _notify_startup_data_repair(self):
        repaired = int(getattr(self, '_repaired_orphan_cashflow_count', 0) or 0)
        if repaired <= 0:
            return

        def _show_notice():
            messagebox.showinfo(
                "Data Repair Summary",
                f"Startup maintenance cleaned {repaired} orphan cashflow row(s)."
            )

        # Defer popup slightly so main UI has time to render first.
        self.root.after(350, _show_notice)

    def _make_cat_btn(self, parent, label, key, active_bg, cmd):
        btn = RoundedButton(parent, text=label, command=cmd,
                            bg="#F8FAFC", fg="#475569",
                            font=("Arial", 9, "bold"),
                            padx=12, pady=5, radius=14)
        btn.pack(side="left", padx=3)
        self._cat_btns[key]      = btn
        self._cat_active_bg[key] = active_bg
        return btn

    def tr(self, key, **kwargs):
        table = self.i18n.get(self.lang, self.i18n['en'])
        text = table.get(key, self.i18n['en'].get(key, key))
        if kwargs:
            try:
                return text.format(**kwargs)
            except Exception:
                return text
        return text

    def l(self, en_text, zh_text):
        return zh_text if self.lang == "zh" else en_text

    def _translate_ui_text(self, text):
        if not isinstance(text, str):
            return text
        if self.lang != "zh":
            return text

        # Preserve leading/trailing spaces used by some labels.
        lead = len(text) - len(text.lstrip(" "))
        tail = len(text) - len(text.rstrip(" "))
        core = text.strip(" ")

        mapped = self._ui_text_map_zh.get(core)
        if mapped is None:
            mapped = self._ui_text_map_zh.get(text)
        if mapped is None:
            return text
        return (" " * lead) + mapped + (" " * tail)

    def _localize_widget_tree(self, root_widget):
        try:
            widgets = [root_widget]
            widgets.extend(root_widget.winfo_children())
        except Exception:
            return

        while widgets:
            w = widgets.pop(0)

            # Window titles
            try:
                if isinstance(w, (tk.Toplevel, tk.Tk)):
                    if not hasattr(w, "_en_title_base"):
                        setattr(w, "_en_title_base", w.title())
                    base_title = getattr(w, "_en_title_base", w.title())
                    zh_title = self._translate_ui_text(base_title)
                    if self.lang == "zh":
                        if zh_title != base_title and w.title() != zh_title:
                            w.title(zh_title)
                            setattr(w, "_ui_title_localized", True)
                    else:
                        if getattr(w, "_ui_title_localized", False) and w.title() != base_title:
                            w.title(base_title)
                        setattr(w, "_ui_title_localized", False)
            except Exception:
                pass

            # Standard text-bearing widgets
            try:
                if not getattr(w, "_skip_auto_localize_text", False):
                    text_now = w.cget("text")
                    if isinstance(text_now, str):
                        if not hasattr(w, "_en_text_base"):
                            setattr(w, "_en_text_base", text_now)
                        base_text = getattr(w, "_en_text_base", text_now)
                        zh_text = self._translate_ui_text(base_text)
                        if self.lang == "zh":
                            if zh_text != base_text and text_now != zh_text:
                                w.configure(text=zh_text)
                                setattr(w, "_ui_text_localized", True)
                        else:
                            if getattr(w, "_ui_text_localized", False) and text_now != base_text:
                                w.configure(text=base_text)
                            setattr(w, "_ui_text_localized", False)
            except Exception:
                pass

            # Notebook tabs
            try:
                if isinstance(w, ttk.Notebook):
                    if not hasattr(w, "_en_tab_text_base"):
                        tab_map = {}
                        for tab_id in w.tabs():
                            tab_map[tab_id] = w.tab(tab_id, "text")
                        setattr(w, "_en_tab_text_base", tab_map)
                    tab_map = getattr(w, "_en_tab_text_base", {})
                    for tab_id in w.tabs():
                        base_text = tab_map.get(tab_id, w.tab(tab_id, "text"))
                        zh_text = self._translate_ui_text(base_text)
                        if self.lang == "zh":
                            if zh_text != base_text and w.tab(tab_id, "text") != zh_text:
                                w.tab(tab_id, text=zh_text)
                        else:
                            if w.tab(tab_id, "text") != base_text:
                                w.tab(tab_id, text=base_text)
            except Exception:
                pass

            # Treeview column headings
            try:
                if isinstance(w, ttk.Treeview):
                    self._enable_treeview_sorting(w)
                    if not hasattr(w, "_en_heading_text_base"):
                        heading_map = {}
                        for col in w["columns"]:
                            heading_map[col] = w.heading(col, "text")
                        setattr(w, "_en_heading_text_base", heading_map)
                    heading_map = getattr(w, "_en_heading_text_base", {})
                    for col in w["columns"]:
                        base_text = heading_map.get(col, w.heading(col, "text"))
                        target_text = self._translate_ui_text(base_text) if self.lang == "zh" else base_text
                        if w.heading(col, "text") != target_text:
                            w.heading(col, text=target_text)
            except Exception:
                pass

            try:
                widgets.extend(w.winfo_children())
            except Exception:
                pass

    def _run_auto_localizer_tick(self):
        try:
            self._localize_widget_tree(self.root)
            for child in self.root.winfo_children():
                if isinstance(child, tk.Toplevel):
                    self._localize_widget_tree(child)
        except Exception:
            pass
        self._auto_localize_job = self.root.after(30000, self._run_auto_localizer_tick)

    def _start_auto_localizer(self):
        if self._auto_localize_job:
            try:
                self.root.after_cancel(self._auto_localize_job)
            except Exception:
                pass
        self._run_auto_localizer_tick()

    def toggle_language(self):
        self.lang = "zh" if self.lang == "en" else "en"
        self.apply_language()

    def apply_language(self):
        if hasattr(self, '_brand_subtitle_label'):
            self._brand_subtitle_label.config(text=self.tr('brand_subtitle'))
        if hasattr(self, '_hub_btn'):
            self._hub_btn.configure(text=self.tr('hub'))
        if hasattr(self, '_admin_btn'):
            self._admin_btn.configure(text=self.tr('admin_on') if self.admin_mode else self.tr('admin'))
        if hasattr(self, '_lang_btn'):
            self._lang_btn.configure(text=self.tr('lang_btn'))
        if hasattr(self, 'float_table_home_btn'):
            self.float_table_home_btn.configure(text=self.tr('main_tables'))
        if hasattr(self, 'current_order_title_label'):
            self.current_order_title_label.config(text=self.tr('current_order'))
        if hasattr(self, 'pay_indicator_label'):
            self.pay_indicator_label.config(text=self.tr('table_flow_hint'))
        if hasattr(self, '_table_home_title_label'):
            self._table_home_title_label.config(text=self.tr('table_layout'))
        if hasattr(self, '_table_home_hint_label'):
            self._table_home_hint_label.config(text=self.tr('table_layout_hint'))
        self._refresh_order_mode_ui()
        if hasattr(self, '_render_cashier_strip'):
            self._render_cashier_strip()
        if hasattr(self, '_render_table_home_strip'):
            self._render_table_home_strip()
        if hasattr(self, 'table_home_cards_frame'):
            self.refresh_table_home_cards()
        if hasattr(self, 'subtotal_label') and hasattr(self, 'tax_label') and hasattr(self, 'total_label'):
            subtotal, tax, total = self._calc_order_totals(self.order)
            self._render_order_total_labels(subtotal, tax, total)
        self._localize_widget_tree(self.root)

    # ================= STAFF LOGIN =================
    def show_login(self):
        self._staff_login_popup(callback=None)

    def _staff_login_popup(self, callback=None, title=None, parent=None, admin_only=False):
        if title is None:
            title = self.l("👤 Who is serving?", "👤 由谁服务？")
        
        login = tk.Toplevel(parent or self.root)
        login.title(title)
        login.configure(bg=self.colors['bg'])
        login.transient(parent or self.root)
        login.grab_set()
        login.resizable(False, False)
        self._center_window(login, 420, 480, parent or self.root)

        # 1. Header
        hdr = tk.Frame(login, bg=self.colors['header'], height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=title, font=("Arial", 15, "bold"),
                 bg=self.colors['header'], fg="white").pack(pady=18)

        # 2. Body container
        body = tk.Frame(login, bg=self.colors['bg'], padx=40, pady=30)
        body.pack(fill="both", expand=True)

        tk.Label(body, text=self.l("Select Staff Member", "选择员工"), 
                 font=("Arial", 11, "bold"), bg=self.colors['bg'], 
                 fg=self.colors['text_sub']).pack(anchor="w", pady=(0, 5))

        with self.db.get_connection() as conn:
            c = conn.cursor()
            if admin_only:
                c.execute("SELECT staff_id, name FROM staff WHERE active=1 AND role='admin' ORDER BY name")
            else:
                c.execute("SELECT staff_id, name FROM staff WHERE active=1 ORDER BY name")
            staff_list = [(r['staff_id'], r['name']) for r in c.fetchall()]

        staff_display = [f"{s[1]}  ({s[0]})" for s in staff_list]
        staff_var = tk.StringVar()
        cb = ttk.Combobox(body, textvariable=staff_var, values=staff_display,
                          font=("Arial", 13), width=25, state="readonly")
        
        # Default to last logged-in staff if available
        default_index = 0
        if self._last_staff_login_id:
            for i, s in enumerate(staff_list):
                if s[0] == self._last_staff_login_id:
                    default_index = i
                    break
        
        if staff_display:
            cb.set(staff_display[default_index])
        cb.pack(pady=(0, 20))

        tk.Label(body, text=self.l("Enter Secret PIN", "输入密码"), 
                 font=("Arial", 11, "bold"), bg=self.colors['bg'], 
                 fg=self.colors['text_sub']).pack(anchor="w", pady=(0, 5))
        
        pin_wrap = tk.Frame(body, bg=self.colors['border'], padx=1, pady=1)
        pin_wrap.pack(pady=(0, 10))
        
        pin_var = tk.StringVar()
        pin_entry = tk.Entry(pin_wrap, textvariable=pin_var, font=("Arial", 28, "bold"),
                             show="●", width=8, justify="center",
                             bg='white', fg='#1E293B', relief='flat', bd=0)
        pin_entry.pack(padx=1, pady=1)
        pin_entry.focus_set()

        msg_label = tk.Label(body, text="", font=("Arial", 10, "bold"),
                             bg=self.colors['bg'], fg=self.colors['danger'])
        msg_label.pack(pady=5)

        def confirm(event=None):
            sel = staff_var.get()
            pin = pin_var.get().strip()
            if not sel:
                msg_label.config(text=self.l("❌ Please select a staff member", "❌ 请选择员工"))
                return
            
            try:
                idx = staff_display.index(sel)
                sid = staff_list[idx][0]
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    if admin_only:
                        c.execute("SELECT * FROM staff WHERE staff_id=? AND pin=? AND active=1 AND role='admin'", (sid, pin))
                    else:
                        c.execute("SELECT * FROM staff WHERE staff_id=? AND pin=? AND active=1", (sid, pin))
                    staff = c.fetchone()
                
                if staff:
                    sd = {'staff_id': staff['staff_id'], 'name': staff['name'], 'role': staff['role']}
                    self.current_staff = sd
                    self._last_staff_login_id = staff['staff_id']
                    self.staff_label.config(text=f"👤 {staff['name']}  ({staff['staff_id']})",
                                            fg='#10B981') # Lighter green
                    login.destroy()
                    if callback:
                        callback(sd)
                else:
                    msg_label.config(text=self.l("❌ Incorrect PIN. Please try again.", "❌ 密码错误，请重试"))
                    pin_var.set("")
                    pin_entry.focus_set()
            except Exception:
                pass

        pin_entry.bind('<Return>', confirm)
        
        action_f = tk.Frame(body, bg=self.colors['bg'])
        action_f.pack(fill="x", pady=10)
        
        RoundedButton(action_f, text=self.l("Cancel", "取消"), command=login.destroy,
                      bg="#64748B", fg="white",
                      font=("Arial", 12, "bold"), padx=20, pady=10, radius=10
                      ).pack(side="left", expand=True, padx=5)

        RoundedButton(action_f, text=self.l("Login Now", "即刻登录"), command=confirm,
                      bg=self.colors['primary'], fg="white",
                      font=("Arial", 12, "bold"), padx=20, pady=10, radius=10
                      ).pack(side="left", expand=True, padx=5)

    # ================= STAFF MANAGEMENT =================
    # ── CHANGE 1: Admin mode toggle ──────────────────────────────────────────
    def toggle_admin_mode(self):
        if self.admin_mode:
            self.admin_mode = False
            self._admin_btn.configure(text=self.tr('admin'))
            new_bg = self.colors['neutral']
            self._admin_btn._bg_orig  = new_bg
            self._admin_btn._bg       = new_bg
            self._admin_btn._hover_bg = RoundedButton._darken(new_bg, 22)
            self._admin_btn._render()
            self._render_cashier_strip()
            self._render_table_home_strip()
            # self.nav_strip.pack_forget()  # Hidden by user request for single row
            # Stats (stats_wrap) now stay visible always based on user feedback
        else:
            self._staff_login_popup(callback=lambda staff: self._enter_admin_mode(), title=self.l("🔑 Admin PIN Required", "🔑 需要管理员登录"), admin_only=True)

    def _enter_admin_mode(self):
        self.admin_mode = True
        self._admin_btn.configure(text=self.tr('admin_on'))
        new_bg = self.colors['warning']
        self._admin_btn._bg_orig  = new_bg
        self._admin_btn._bg       = new_bg
        self._admin_btn._hover_bg = RoundedButton._darken(new_bg, 22)
        self._admin_btn._render()
        # Keep stats consistently in the top bar
        self.stats_wrap.pack(side="left", padx=16, after=self._brand_frame)
        self._render_cashier_strip()
        self._render_table_home_strip()

    def manual_open_cash_drawer(self):
        """Standalone manual drawer kick with security check."""
        def _do_open(staff_or_none=None):
            drawer_error = open_cash_drawer()
            if drawer_error:
                messagebox.showwarning(self.l("Drawer error", "钱箱错误"), drawer_error, parent=self.root)
                return
            self.record_action("MANUAL_DRAWER", "Manual drawer kick", user=(staff_or_none['name'] if staff_or_none else 'Admin'))
            
        if self.admin_mode:
            _do_open()
        else:
            self._staff_login_popup(callback=_do_open, title=self.l("👤 Drawer Access", "👤钱箱访问"))

    def open_cash_center(self):
        self.show_cash_closing(cashier_mode=not self.admin_mode)

    def open_rent_center(self):
        # Main nav Rent should always open the simple cashier-style rent flow.
        self.show_cashier_rent()

    def _get_action_groups(self):
        C = self.colors
        L = self.l
        return [
            (L("Daily Actions", "日常操作"), [
                {"label": L("🧾 Recent Orders", "🧾 近期订单"), "hint": L("View, edit or reprint recent sales", "查看、编辑或重打近期订单"), "aliases": "orders sales receipt history 订单 销售 小票", "command": self.show_recent_orders, "admin": False, "bg": C['primary']},
                {"label": L("📥 Held Orders", "📥 未结订单"), "hint": L("Resume unpaid held orders", "继续处理未付款订单"), "aliases": "hold unpaid pending 挂单 未付款", "command": self.show_held_orders, "admin": False, "bg": C['neutral']},
                {"label": L("💰 Cash Operations", "💰 现金操作"), "hint": L("Opening cash, expenses, income and cashier cash", "开盘现金、支出、收入及收银现金"), "aliases": "cash drawer expense income 现金 钱箱 支出 收入", "command": self.open_cash_center, "admin": False, "bg": C['warning']},
                {"label": L("🏪 Collect Rent", "🏪 收取租金"), "hint": L("Record a tenant rental payment", "记录租户租金付款"), "aliases": "rent rental tenant stall 租金 租户 档口", "command": self.open_rent_center, "admin": False, "bg": C['success']},
                {"label": L("📦 Stock Adjustment", "📦 库存调整"), "hint": L("Adjust tracked stock only", "调整追踪库存的商品"), "aliases": "stock inventory quantity 库存 数量", "command": self.show_cashier_stock_control, "admin": False, "bg": C['warning']},
                {"label": L("🗄️ Open Cash Drawer", "🗄️ 打开钱箱"), "hint": L("Manually open the cash drawer", "手动打开钱箱"), "aliases": "drawer till kick 钱箱", "command": self.manual_open_cash_drawer, "admin": False, "bg": C['neutral']},
            ]),
            (L("Manager Tools", "管理工具"), [
                {"label": L("⚙️ Items & Inventory Settings", "⚙️ 商品与库存设置"), "hint": L("Menu items, prices, variants and stock settings", "商品、价格、选项及库存设置"), "aliases": "item product menu stock 商品 菜单 库存", "command": self.manage_items, "admin": True, "bg": C['primary']},
                {"label": L("📂 Product Categories", "📂 商品分类"), "hint": L("Manage product categories and display order", "管理商品分类及显示顺序"), "aliases": "category group 分类 类别", "command": self.manage_categories, "admin": True, "bg": C['info']},
                {"label": L("🪑 Table Settings", "🪑 桌位设置"), "hint": L("Add and manage dining tables", "新增及管理桌位"), "aliases": "table maintenance seating 桌位 桌号", "command": self.manage_tables, "admin": True, "bg": C['card']},
                {"label": L("🏪 Tenant & Stall Management", "🏪 租户与档口管理"), "hint": L("Manage tenants, stalls and rental settings", "管理租户、档口及租金设置"), "aliases": "tenant stall rent rental 租户 档口 租金", "command": self.manage_stalls, "admin": True, "bg": C['success']},
                {"label": L("👥 Staff Management", "👥 员工管理"), "hint": L("Manage accounts, roles, PINs and access", "管理员工账号、角色、PIN及权限"), "aliases": "staff employee role pin 员工 权限", "command": self.manage_staff, "admin": True, "bg": C['neutral']},
                {"label": L("↩️ Refund Management", "↩️ 退款管理"), "hint": L("Review orders and process controlled refunds", "查看订单及处理受控退款"), "aliases": "refund return void 退款 退货", "command": self.show_refunds, "admin": True, "bg": C['danger']},
                {"label": L("📊 Finance & Reports", "📊 财务与报表"), "hint": L("Sales records, cash closing, analytics, profit and exports", "销售记录、现金日结、营业分析、利润及导出"), "aliases": "reports sales history cash closing profit accountant export 报表 销售记录 现金日结 利润 会计 导出", "command": self.show_finance_reports_center, "admin": True, "bg": C['primary']},
                {"label": L("💸 Income & Expense Categories", "💸 收支类别"), "hint": L("Manage reusable cashflow categories", "管理可重复使用的收支类别"), "aliases": "expense income category 支出 收入 类别", "command": self.manage_cashflow_categories, "admin": True, "bg": C['warning']},
            ]),
        ]

    def _run_action(self, action, parent=None):
        def _go():
            self.record_action("OPEN_MODULE", f"{action.get('group', 'Unknown')} -> {action['label']}")
            self._run_with_loading(action['command'], parent=parent)

        if action.get('admin') and not self.admin_mode:
            self._staff_login_popup(callback=lambda staff: _go(), title=self.l("🔑 Admin PIN Required", "🔑 需要管理员登录"), parent=parent, admin_only=True)
        else:
            _go()

    def show_action_hub(self):
        C = self.colors
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("Action Hub", "功能中心"))
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        self._center_window(win, 760, 560, self.root)

        hdr = tk.Frame(win, bg=C['header'], height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=L("⚡ Action Hub", "⚡ 功能中心"), font=("Arial", 16, "bold"),
                 bg=C['header'], fg="white").pack(side="left", padx=16, pady=12)
        tk.Label(hdr, text="F7", font=("Arial", 9, "bold"),
                 bg=C['header2'], fg="#E2E8F0", padx=6, pady=2).pack(side="right", padx=14, pady=14)

        search_row = tk.Frame(win, bg=C['bg'])
        search_row.pack(fill="x", padx=14, pady=(12, 8))
        tk.Label(search_row, text=L("Search:", "搜索："), font=("Arial", 11, "bold"),
                 bg=C['bg'], fg=C['text']).pack(side="left")
        q_var = tk.StringVar()
        q_entry = tk.Entry(search_row, textvariable=q_var, font=("Arial", 12),
                           bg='white', fg=C['text'], relief='solid', bd=1)
        q_entry.pack(side="left", fill="x", expand=True, padx=8)
        q_entry.focus_set()

        list_frame = tk.Frame(win, bg=C['bg'])
        list_frame.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        lb = tk.Listbox(list_frame, font=("Arial", 11), activestyle='none')
        sb = tk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        hint_lbl = tk.Label(win, text="", font=("Arial", 10),
                            bg=C['bg'], fg=C['text_muted'], anchor='w', justify='left')
        hint_lbl.pack(fill="x", padx=14, pady=(0, 8))

        action_rows = []
        for group_name, actions in self._get_action_groups():
            for a in actions:
                row = dict(a)
                row['group'] = group_name
                action_rows.append(row)

        current = []

        def refresh_list(*_):
            key = q_var.get().strip().casefold()
            lb.delete(0, tk.END)
            current.clear()
            if not key:
                hint_lbl.config(text=L(
                    "Type a function name or keyword to search.",
                    "请输入功能名称或关键词进行搜索。"
                ))
                return
            for a in action_rows:
                searchable = " ".join((
                    a['label'], a['hint'], a['group'], a.get('aliases', '')
                )).casefold()
                if key not in searchable:
                    continue
                admin_tag = L(" [ADMIN]", " [管理员]") if a.get('admin') else ""
                lb.insert(tk.END, f"{a['group']}  •  {a['label']}{admin_tag}")
                current.append(a)
            if current:
                lb.selection_set(0)
                show_hint()
            else:
                hint_lbl.config(text=L("No action found.", "没有找到对应功能。"))

        def show_hint(_=None):
            sel = lb.curselection()
            if not sel or sel[0] >= len(current):
                hint_lbl.config(text="")
                return
            a = current[sel[0]]
            lock = L(" • Requires manager access", " • 需要管理员权限") if a.get('admin') else ""
            hint_lbl.config(text=f"{a['hint']}{lock}")

        def run_selected(_=None):
            sel = lb.curselection()
            if not sel or sel[0] >= len(current):
                return
            action = current[sel[0]]
            win.destroy()
            self._run_action(action, parent=self.root)

        q_var.trace_add("write", refresh_list)
        lb.bind("<<ListboxSelect>>", show_hint)
        lb.bind("<Double-Button-1>", run_selected)
        lb.bind("<Return>", run_selected)

        btns = tk.Frame(win, bg=C['bg'])
        btns.pack(fill="x", padx=14, pady=(0, 12))
        RoundedButton(btns, text=L("Open", "打开"), command=run_selected,
                      bg=C['success'], fg='white', font=("Arial", 11, "bold"),
                      padx=18, pady=8, radius=8).pack(side="left")
        RoundedButton(btns, text=L("Close", "关闭"), command=win.destroy,
                      bg=C['neutral'], fg='white', font=("Arial", 11, "bold"),
                      padx=18, pady=8, radius=8).pack(side="right")

        win.bind("<Escape>", lambda e: win.destroy())
        refresh_list()

    def show_finance_reports_center(self):
        """Canonical manager entry for records, cash closing and reporting tools."""
        C = self.colors
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("Finance & Reports", "财务与报表"))
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()
        win.resizable(True, True)
        center_width, center_height = self._center_window(win, 980, 640, self.root)
        win.minsize(min(760, center_width), min(520, center_height))

        header = tk.Frame(win, bg=C['header'], height=68)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header, text=L("📊  Finance & Reports", "📊  财务与报表"),
            font=("Arial", 18, "bold"), bg=C['header'], fg="white"
        ).pack(side="left", padx=20, pady=15)
        tk.Label(
            header,
            text=L(
                "One place for records, closing, analytics, cashflow and exports",
                "集中管理销售记录、日结、分析、现金流及导出"
            ),
            font=("Arial", 10), bg=C['header'], fg="#94A3B8"
        ).pack(side="left", pady=20)

        intro = tk.Frame(win, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        intro.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(
            intro,
            text=L(
                "Choose a tab, then open the required workspace.",
                "请选择标签，再打开所需工作区。"
            ),
            font=("Arial", 10), bg=C['panel'], fg=C['text_sub']
        ).pack(anchor="w", padx=14, pady=9)

        style = ttk.Style(win)
        style.configure(
            "FinanceCenter.TNotebook.Tab",
            font=("Arial", 10, "bold"), padding=[14, 9]
        )
        notebook = ttk.Notebook(win, style="FinanceCenter.TNotebook")
        notebook.pack(fill="both", expand=True, padx=16, pady=6)

        def launch(command):
            try:
                win.grab_release()
            except tk.TclError:
                pass
            win.destroy()
            self._run_with_loading(command)

        def add_tool_tab(tab_text, icon, title, description, features,
                         button_text, command, accent):
            tab = tk.Frame(notebook, bg=C['bg'])
            notebook.add(tab, text=tab_text)

            card = tk.Frame(
                tab, bg=C['panel'], highlightbackground=C['border'],
                highlightthickness=1
            )
            card.pack(fill="both", expand=True, padx=48, pady=34)
            tk.Frame(card, bg=accent, height=6).pack(fill="x")

            title_row = tk.Frame(card, bg=C['panel'])
            title_row.pack(fill="x", padx=28, pady=(24, 8))
            tk.Label(
                title_row, text=icon, font=("Arial", 30),
                bg=C['panel'], fg=accent
            ).pack(side="left", padx=(0, 16))
            title_wrap = tk.Frame(title_row, bg=C['panel'])
            title_wrap.pack(side="left", fill="x", expand=True)
            tk.Label(
                title_wrap, text=title, font=("Arial", 17, "bold"),
                bg=C['panel'], fg=C['text'], anchor="w"
            ).pack(fill="x")
            tk.Label(
                title_wrap, text=description, font=("Arial", 10),
                bg=C['panel'], fg=C['text_muted'], anchor="w", justify="left",
                wraplength=680
            ).pack(fill="x", pady=(4, 0))

            feature_box = tk.Frame(card, bg="#F8FAFC")
            feature_box.pack(fill="both", expand=True, padx=28, pady=12)
            for feature in features:
                tk.Label(
                    feature_box, text=f"✓  {feature}",
                    font=("Arial", 10), bg="#F8FAFC", fg=C['text_sub'],
                    anchor="w"
                ).pack(fill="x", padx=16, pady=5)

            action_row = tk.Frame(card, bg=C['panel'])
            action_row.pack(fill="x", padx=28, pady=(4, 24))
            RoundedButton(
                action_row, text=button_text,
                command=lambda: launch(command),
                bg=accent, fg="white", font=("Arial", 11, "bold"),
                padx=24, pady=10, radius=8
            ).pack(side="right")

        add_tool_tab(
            L("📜 Sales Records", "📜 销售记录"), "📜",
            L("Sales Records", "销售记录"),
            L("Find and work with individual completed orders.", "查找及处理个别已完成订单。"),
            [
                L("Search by order, item or cashier", "按订单、商品或收银员搜索"),
                L("View details, edit and reprint", "查看详情、修改及重打"),
                L("Open payment reconciliation", "打开付款对账"),
            ],
            L("Open Sales Records", "打开销售记录"), self.show_history, C['primary']
        )
        add_tool_tab(
            L("🌙 Cash Closing", "🌙 现金日结"), "🌙",
            L("Cash Closing", "现金日结"),
            L("Manage the complete business-day cash workflow.", "管理完整营业日现金流程。"),
            [
                L("Opening cash and cashier cash", "开盘现金及收银现金"),
                L("Cash Out and Kuih daily settlement", "现金转出及糕点每日结算"),
                L("End-of-day physical cash check", "日终实点现金核对"),
            ],
            L("Open Cash Closing", "打开现金日结"),
            lambda: self.show_cash_closing(cashier_mode=False), C['warning']
        )
        add_tool_tab(
            L("📊 Business Analytics", "📊 营业分析"), "📊",
            L("Business Analytics", "营业分析"),
            L("Review daily totals, item performance and stock reports.", "查看每日总额、商品表现及库存报表。"),
            [
                L("Daily and monthly summaries", "每日及每月汇总"),
                L("Top items and stock status", "热销商品及库存状态"),
                L("Item sales breakdown and audit log", "商品销售明细及审计日志"),
            ],
            L("Open Business Analytics", "打开营业分析"), self.show_reports, C['info']
        )
        add_tool_tab(
            L("📈 POS Cashflow Summary", "📈 POS 现金流汇总"), "📈",
            L("POS Cashflow", "POS 现金流"),
            L("Review money recorded inside POS; this is not final accounting profit.", "查看 POS 内记录的资金流；这不是最终会计利润。"),
            [
                L("Custom date ranges", "自定义日期范围"),
                L("Income and expense categories", "收入及支出类别"),
                L("Recorded cashflow summary", "已记录现金流汇总"),
            ],
            L("Open Cashflow Summary", "打开现金流汇总"), self.show_profit_center, C['success']
        )
        export_frame = tk.Frame(notebook, bg=C['bg'])
        notebook.add(export_frame, text=L("📊 Excel Exports", "📊 Excel 导出"))
        self._populate_excel_exports(export_frame, win)

        footer = tk.Frame(win, bg=C['bg'])
        footer.pack(fill="x", padx=16, pady=(0, 12))
        RoundedButton(
            footer, text=L("Close", "关闭"), command=win.destroy,
            bg=C['neutral'], fg="white", font=("Arial", 10, "bold"),
            padx=18, pady=8, radius=8
        ).pack(side="right")
        win.bind("<Escape>", lambda _event: win.destroy())
        notebook.enable_traversal()

    def _populate_excel_exports(self, export_frame, parent_win):
        L = self.l
        canvas = tk.Canvas(export_frame, bg=self.colors['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(export_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=scrollbar.set)
        exp_f = tk.Frame(canvas, bg=self.colors['bg'], padx=18, pady=12)
        content = canvas.create_window((0, 0), window=exp_f, anchor="nw")
        exp_f.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(content, width=event.width))
        tk.Label(exp_f, text=L("📊 Business Data Export", "📊 营业数据导出"), font=("Arial", 18, "bold"),
                 bg=self.colors['bg'], fg=self.colors['primary']).pack(anchor="w", pady=(0, 12))

        dg = tk.LabelFrame(exp_f, text=L(" Daily Reports ", " 日报表 "), font=("Arial", 11, "bold"),
                           bg=self.colors['bg'], padx=12, pady=10)
        dg.pack(fill="x", pady=6)
        RoundedButton(dg, text=L("📊 Export Daily FULL", "📊 导出完整日报"),
                      command=lambda: self._export_daily_excel(datetime.today().strftime('%Y-%m-%d'), parent_win, apply_filter=False),
                      bg='#0F766E', fg="white", font=("Arial", 11, "bold"),
                      padx=24, pady=12, radius=8).pack(side="left", padx=12)
        RoundedButton(dg, text=L("📊 Export Daily FILTERED", "📊 导出过滤日报"),
                      command=lambda: self._export_daily_excel(datetime.today().strftime('%Y-%m-%d'), parent_win, apply_filter=True),
                      bg='#14B8A6', fg="white", font=("Arial", 11, "bold"),
                      padx=24, pady=12, radius=8).pack(side="left", padx=12)

        mg = tk.LabelFrame(exp_f, text=L(" Monthly Reports ", " 月报表 "), font=("Arial", 11, "bold"),
                           bg=self.colors['bg'], padx=12, pady=10)
        mg.pack(fill="x", pady=6)
        RoundedButton(mg, text=L("📅 Export Monthly FULL", "📅 导出完整月报"),
                      command=lambda: self._export_monthly_income_expense_excel(parent=parent_win, apply_filter=False),
                      bg='#1D4ED8', fg="white", font=("Arial", 11, "bold"),
                      padx=24, pady=12, radius=8).pack(side="left", padx=12)
        RoundedButton(mg, text=L("📅 Export Monthly FILTERED", "📅 导出过滤月报"),
                      command=lambda: self._export_monthly_income_expense_excel(parent=parent_win, apply_filter=True),
                      bg='#3B82F6', fg="white", font=("Arial", 11, "bold"),
                      padx=24, pady=12, radius=8).pack(side="left", padx=12)

        ag = tk.LabelFrame(
            exp_f, text=L(" Accountant Export ", " 会计专用导出 "),
            font=("Arial", 11, "bold"), bg=self.colors['bg'], padx=12, pady=10
        )
        ag.pack(fill="x", pady=6)
        tk.Label(
            ag,
            text=L(
                "Generate the dedicated 919 Restaurant workbook for accounting.",
                "生成 919 Restaurant 会计专用工作簿。"
            ),
            font=("Arial", 10), bg=self.colors['bg'], fg=self.colors['text_muted'], wraplength=300, justify="left"
        ).pack(side="left", padx=(4, 18))
        RoundedButton(
            ag, text=L("📘 Export Accountant Workbook", "📘 导出会计工作簿"),
            command=lambda: self._export_accountant_sales_exclusive(parent_win),
            bg='#7C3AED', fg="white", font=("Arial", 11, "bold"),
            padx=24, pady=12, radius=8
        ).pack(side="left", padx=12)

    def _rename_cashflow_category(self, old_name, new_name, category_type="Expense",
                                  actor="Admin"):
        """Rename one report group and all matching references atomically."""
        old_name = str(old_name or "").strip()
        new_name = str(new_name or "").strip()
        category_type = str(category_type or "").strip().title()
        if category_type not in ("Expense", "Income"):
            raise ValueError("Category type must be Expense or Income.")
        if not old_name or not new_name:
            raise ValueError("Old and new category names are required.")
        if old_name == new_name:
            raise ValueError("Enter a different category name.")

        protected = {
            ("expense", "stock purchase"),
            ("expense", "staff advance"),
            ("expense", "cash transfer"),
            ("expense", "kuih purchase"),
            ("income", "kuih sales"),
            ("income", "stall rental"),
        }
        if (category_type.casefold(), old_name.casefold()) in protected:
            raise ValueError("This system report group cannot be renamed.")

        with self.db.get_connection() as conn:
            row = conn.execute(
                """SELECT name FROM cashflow_categories
                   WHERE type=? AND LOWER(TRIM(name))=LOWER(TRIM(?))
                   ORDER BY id LIMIT 1""",
                (category_type, old_name),
            ).fetchone()
            if not row:
                raise ValueError("Report group not found. Save the expense option first to add it.")
            old_name = str(row['name'])

            duplicate = conn.execute(
                """SELECT 1 FROM cashflow_categories
                   WHERE type=? AND LOWER(TRIM(name))=LOWER(TRIM(?))
                     AND name<>? LIMIT 1""",
                (category_type, new_name, old_name),
            ).fetchone()
            if duplicate:
                raise ValueError("A report group with this name already exists.")

            category_cursor = conn.execute(
                "UPDATE cashflow_categories SET name=? WHERE type=? AND name=?",
                (new_name, category_type, old_name),
            )
            option_count = 0
            transaction_count = 0
            if category_type == "Expense":
                option_count = conn.execute(
                    "UPDATE expense_options SET category=?,updated_at=? WHERE category=?",
                    (new_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), old_name),
                ).rowcount
                transaction_count = conn.execute(
                    """UPDATE cash_transactions SET category=?
                       WHERE category=?
                         AND LOWER(type) IN ('expense','stock_purchase')""",
                    (new_name, old_name),
                ).rowcount
            else:
                transaction_count = conn.execute(
                    """UPDATE cash_transactions SET category=?
                       WHERE category=? AND LOWER(type) IN ('income','kuih_sales')""",
                    (new_name, old_name),
                ).rowcount

            cashflow_count = conn.execute(
                "UPDATE cashflow SET source=? WHERE type=? AND source=?",
                (new_name, category_type, old_name),
            ).rowcount
            self.record_action(
                "CASHFLOW_CATEGORY_RENAME",
                (
                    f"{category_type}: {old_name} -> {new_name} | "
                    f"groups {category_cursor.rowcount}, options {option_count}, "
                    f"cash transactions {transaction_count}, cashflow {cashflow_count}"
                ),
                user=actor,
                conn=conn,
            )

        return {
            "categories": category_cursor.rowcount,
            "expense_options": option_count,
            "cash_transactions": transaction_count,
            "cashflow": cashflow_count,
        }

    def manage_expense_options(self):
        """Manage ordinary Cash Out choices without changing protected workflows."""
        C = self.colors
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("Expense Management", "开销选项管理"))
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()
        win.resizable(True, True)
        actual_w, actual_h = self._center_window(win, 1180, 720, self.root)
        win.minsize(min(880, actual_w), min(580, actual_h))

        header = tk.Frame(win, bg=C['header'], height=66)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header, text=L("💸  Expense Management", "💸  开销选项管理"),
            font=("Arial", 18, "bold"), bg=C['header'], fg="white"
        ).pack(side="left", padx=20, pady=15)
        tk.Label(
            header,
            text=L(
                "Controls the expense categories shown in Cash Out",
                "管理现金转出画面显示的开销类别"
            ),
            font=("Arial", 10), bg=C['header'], fg="#CBD5E1"
        ).pack(side="left", pady=18)

        notice = tk.Label(
            win,
            text=L(
                "Only active choices created here appear in Cash Out. Link a tracked item only when saving this expense must also add stock.",
                "现金转出只显示这里已启用的选项；只有需要在保存开销时同步加库存，才关联追踪库存商品。"
            ),
            font=("Arial", 9), bg="#EFF6FF", fg="#1E3A8A",
            anchor="w", justify="left", padx=14, pady=8
        )
        notice.pack(fill="x", padx=16, pady=(10, 4))

        body = tk.Frame(win, bg=C['bg'])
        body.pack(fill="both", expand=True, padx=16, pady=(4, 12))
        body.grid_columnconfigure(0, weight=3, uniform="expense_manage")
        body.grid_columnconfigure(1, weight=2, uniform="expense_manage")
        body.grid_rowconfigure(0, weight=1)

        list_panel = tk.Frame(
            body, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1
        )
        list_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        tk.Label(
            list_panel, text=L("Expense Options", "现有开销选项"),
            font=("Arial", 13, "bold"), bg=C['panel'], fg=C['text']
        ).pack(anchor="w", padx=16, pady=(12, 4))

        search_var = tk.StringVar()
        result_var = tk.StringVar()
        query_host = tk.Frame(list_panel, bg=C['panel'])
        query_host.pack(fill="x")

        tree_wrap = tk.Frame(list_panel, bg=C['panel'])
        tree_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        columns = ("ID", "Option", "Category", "Stock Link", "Quantity", "Order", "Status")
        tree = ttk.Treeview(
            tree_wrap, columns=columns, show="headings", height=14,
            style="Manage.Treeview"
        )
        headings = {
            "ID": "ID", "Option": L("Cash Out Category", "现金转出类别"),
            "Category": L("Report Group", "报表分组"),
            "Stock Link": L("Stock Link", "库存关联"),
            "Quantity": L("Quantity", "数量"),
            "Order": L("Order", "排序"), "Status": L("Status", "状态"),
        }
        widths = {
            "ID": 48, "Option": 175, "Category": 135, "Stock Link": 150,
            "Quantity": 82, "Order": 55, "Status": 75,
        }
        for column in columns:
            tree.heading(column, text=headings[column])
            tree.column(
                column, width=widths[column], minwidth=45,
                anchor="w" if column in ("Option", "Category", "Stock Link") else "center"
            )
        tree.tag_configure("inactive", foreground=C['text_muted'])
        scrollbar = tk.Scrollbar(tree_wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        list_buttons = tk.Frame(list_panel, bg=C['panel'])
        list_buttons.pack(fill="x", padx=12, pady=(0, 12))

        editor = tk.Frame(
            body, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1
        )
        editor.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        editor.grid_columnconfigure(0, weight=1, uniform="expense_fields")
        editor.grid_columnconfigure(1, weight=1, uniform="expense_fields")
        tk.Label(
            editor, text=L("Add / Edit Option", "新增／修改选项"),
            font=("Arial", 13, "bold"), bg=C['panel'], fg=C['text']
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 8))

        option_id = [None]
        name_en_var = tk.StringVar()
        name_zh_var = tk.StringVar()
        category_var = tk.StringVar(value="Ingredients")
        quantity_mode_var = tk.StringVar()
        stock_item_var = tk.StringVar()
        stock_choice_map = {}
        sort_var = tk.StringVar(value="100")
        active_var = tk.BooleanVar(value=True)
        form_status_var = tk.StringVar()

        quantity_labels = {
            "required": L("Required", "必须填写"),
            "optional": L("Optional", "选填"),
            "none": L("Not used", "不使用"),
        }
        quantity_from_label = {value: key for key, value in quantity_labels.items()}
        no_stock_label = L("No stock change", "不影响库存")

        def field_label(text, row, column, columnspan=1):
            tk.Label(
                editor, text=text, font=("Arial", 9, "bold"),
                bg=C['panel'], fg=C['text']
            ).grid(
                row=row, column=column, columnspan=columnspan,
                sticky="w", padx=16, pady=(4, 2)
            )

        field_label(L("Category Name (English):", "类别名称（英文）："), 1, 0)
        field_label(L("Category Name (Chinese, optional):", "类别名称（中文，选填）："), 1, 1)
        name_en_entry = tk.Entry(editor, textvariable=name_en_var, font=("Arial", 10))
        name_en_entry.grid(row=2, column=0, sticky="ew", padx=(16, 6), pady=(0, 5))
        tk.Entry(editor, textvariable=name_zh_var, font=("Arial", 10)).grid(
            row=2, column=1, sticky="ew", padx=(6, 16), pady=(0, 5)
        )

        field_label(L("Report Group:", "报表分组："), 3, 0, 2)
        report_group_host = tk.Frame(editor, bg=C['panel'])
        report_group_host.grid(
            row=4, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 5)
        )
        category_combo = ttk.Combobox(
            report_group_host, textvariable=category_var,
            state="normal", font=("Arial", 10)
        )
        category_combo.pack(side="left", fill="x", expand=True)
        edit_report_group_button = tk.Button(
            report_group_host, text="✎",
            command=lambda: edit_report_group(),
            font=("Segoe UI Symbol", 12, "bold"), bg=C['neutral'], fg="white",
            activebackground=C['header2'], activeforeground="white",
            relief="flat", bd=0, padx=9, pady=1, cursor="hand2"
        )
        edit_report_group_button.pack(side="left", padx=(5, 0))

        field_label(L("Linked Stock Item (optional):", "关联库存商品（选填）："), 5, 0)
        field_label(L("Quantity:", "数量："), 5, 1)
        stock_item_combo = ttk.Combobox(
            editor, textvariable=stock_item_var,
            state="readonly", font=("Arial", 10)
        )
        stock_item_combo.grid(row=6, column=0, sticky="ew", padx=(16, 6), pady=(0, 5))
        quantity_combo = ttk.Combobox(
            editor, textvariable=quantity_mode_var,
            values=list(quantity_labels.values()), state="readonly", font=("Arial", 10)
        )
        quantity_combo.grid(row=6, column=1, sticky="ew", padx=(6, 16), pady=(0, 5))

        stock_note_var = tk.StringVar()
        tk.Label(
            editor, textvariable=stock_note_var, font=("Arial", 8),
            bg=C['panel'], fg=C['text_sub'], anchor="w", justify="left", wraplength=390
        ).grid(row=7, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 4))

        field_label(L("Display Order:", "显示排序："), 8, 0)
        tk.Entry(editor, textvariable=sort_var, font=("Arial", 10)).grid(
            row=9, column=0, sticky="ew", padx=(16, 6), pady=(0, 5)
        )
        tk.Checkbutton(
            editor, text=L("Active in Cash Out", "在现金转出中启用"),
            variable=active_var, bg=C['panel'], activebackground=C['panel'],
            font=("Arial", 9, "bold")
        ).grid(row=9, column=1, sticky="w", padx=(6, 16), pady=(0, 5))

        field_label(L("Staff Hint (English, optional):", "员工提示（英文，选填）："), 10, 0, 2)
        help_en_text = tk.Text(editor, height=3, font=("Arial", 9), wrap="word")
        help_en_text.grid(row=11, column=0, columnspan=2, sticky="nsew", padx=16, pady=(0, 5))
        field_label(L("Staff Hint (Chinese, optional):", "员工提示（中文，选填）："), 12, 0, 2)
        help_zh_text = tk.Text(editor, height=3, font=("Arial", 9), wrap="word")
        help_zh_text.grid(row=13, column=0, columnspan=2, sticky="nsew", padx=16, pady=(0, 5))
        editor.grid_rowconfigure(11, weight=1)
        editor.grid_rowconfigure(13, weight=1)

        editor_buttons = tk.Frame(editor, bg=C['panel'])
        editor_buttons.grid(row=14, column=0, columnspan=2, sticky="ew", padx=12, pady=(8, 4))
        tk.Label(
            editor, textvariable=form_status_var, font=("Arial", 9, "bold"),
            bg=C['panel'], fg=C['success'], anchor="w"
        ).grid(row=15, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 10))

        def load_category_choices():
            protected = {"Stock Purchase", "Staff Advance", "Cash Transfer", "Kuih Purchase"}
            with self.db.get_connection() as conn:
                values = [
                    str(row['name']) for row in conn.execute(
                        "SELECT DISTINCT name FROM cashflow_categories "
                        "WHERE type='Expense' ORDER BY name"
                    ).fetchall()
                    if str(row['name']) not in protected
                ]
            category_combo.configure(values=values)

        def edit_report_group():
            current_name = category_var.get().strip()
            if not current_name:
                messagebox.showwarning(
                    L("No Report Group", "未选择报表分组"),
                    L("Choose a report group first.", "请先选择一个报表分组。"),
                    parent=win,
                )
                return

            with self.db.get_connection() as conn:
                existing = conn.execute(
                    """SELECT name FROM cashflow_categories
                       WHERE type='Expense' AND LOWER(TRIM(name))=LOWER(TRIM(?))
                       ORDER BY id LIMIT 1""",
                    (current_name,),
                ).fetchone()
            if not existing:
                messagebox.showinfo(
                    L("New Report Group", "新报表分组"),
                    L(
                        "This is a new report-group name. Save the expense option first; then it can be renamed with the pencil button.",
                        "这是新的报表分组名称。请先保存开销选项，之后即可用铅笔按钮重命名。"
                    ),
                    parent=win,
                )
                return

            current_name = str(existing['name'])
            new_name = simpledialog.askstring(
                L("Rename Report Group", "重命名报表分组"),
                L(
                    "New report-group name:\n\nExisting expense options and historical reports using this group will be updated together.",
                    "新的报表分组名称：\n\n使用此分组的开销选项和历史报表会一起更新。"
                ),
                initialvalue=current_name,
                parent=win,
            )
            if new_name is None:
                return
            new_name = new_name.strip()
            if not new_name or new_name == current_name:
                return
            if not messagebox.askyesno(
                L("Confirm Rename", "确认重命名"),
                L(
                    f"Rename '{current_name}' to '{new_name}' everywhere?",
                    f"确定把“{current_name}”在所有记录中改为“{new_name}”？"
                ),
                parent=win,
            ):
                return

            try:
                counts = self._rename_cashflow_category(
                    current_name, new_name, "Expense", actor="Admin"
                )
            except Exception as exc:
                messagebox.showerror(
                    L("Unable to Rename", "无法重命名"), str(exc), parent=win
                )
                return

            category_var.set(new_name)
            load_category_choices()
            selected_id = option_id[0]
            load_options(selected_id)
            messagebox.showinfo(
                L("Report Group Updated", "报表分组已更新"),
                L(
                    f"Renamed to '{new_name}'. Updated {counts['expense_options']} expense option(s) and {counts['cashflow']} historical report record(s).",
                    f"已改为“{new_name}”。同步更新 {counts['expense_options']} 个开销选项及 {counts['cashflow']} 笔历史报表记录。"
                ),
                parent=win,
            )

        def stock_display_for_id(stock_item_id):
            for display, mapped_id in stock_choice_map.items():
                if mapped_id == stock_item_id:
                    return display
            return no_stock_label

        def update_stock_link_fields(_event=None):
            stock_item_id = stock_choice_map.get(stock_item_var.get())
            if stock_item_id is not None:
                quantity_mode_var.set(quantity_labels["required"])
                quantity_combo.configure(state="disabled")
                stock_note_var.set(L(
                    "Quantity is required. Saving this Cash Out also increases the linked item's stock.",
                    "必须填写数量；保存现金转出时会同步增加所关联商品的库存。"
                ))
            else:
                quantity_combo.configure(state="readonly")
                stock_note_var.set(L(
                    "Leave unlinked for furniture, utensils and other expenses that should not change stock.",
                    "家具、餐具及其他不应改变库存的开销，请保持“不影响库存”。"
                ))

        def load_stock_choices(preselect_id=None):
            stock_choice_map.clear()
            stock_choice_map[no_stock_label] = None
            with self.db.get_connection() as conn:
                rows = conn.execute(
                    """SELECT id,name,category,stock,active FROM items
                       WHERE track_stock=1
                       ORDER BY active DESC,category,name"""
                ).fetchall()
            for row in rows:
                status = "" if int(row['active'] or 0) else L(" · inactive", " · 已停用")
                display = L(
                    f"{row['name']} · {row['category']} (stock {int(row['stock'] or 0)}){status}",
                    f"{row['name']} · {row['category']}（库存 {int(row['stock'] or 0)}）{status}"
                )
                stock_choice_map[display] = int(row['id'])
            stock_item_combo.configure(values=list(stock_choice_map))
            stock_item_var.set(stock_display_for_id(preselect_id))
            update_stock_link_fields()

        def load_options(preselect_id=None):
            tree.delete(*tree.get_children())
            keyword = search_var.get().strip()
            with self.db.get_connection() as conn:
                if keyword:
                    like_value = f"%{keyword}%"
                    rows = conn.execute(
                        """SELECT eo.*,i.name AS stock_item_name
                           FROM expense_options eo
                           LEFT JOIN items i ON i.id=eo.stock_item_id
                           WHERE eo.name_en LIKE ? OR COALESCE(eo.name_zh,'') LIKE ?
                              OR eo.category LIKE ? OR COALESCE(i.name,'') LIKE ?
                           ORDER BY eo.sort_order,eo.id""",
                        (like_value, like_value, like_value, like_value),
                    ).fetchall()
                else:
                    rows = conn.execute(
                        """SELECT eo.*,i.name AS stock_item_name
                           FROM expense_options eo
                           LEFT JOIN items i ON i.id=eo.stock_item_id
                           ORDER BY eo.sort_order,eo.id"""
                    ).fetchall()
            for row in rows:
                option_name = str(row['name_en'])
                if str(row['name_zh'] or '').strip():
                    option_name += f" / {row['name_zh']}"
                stock_text = str(row['stock_item_name'] or '').strip() or L("None", "无")
                quantity_text = quantity_labels.get(
                    str(row['quantity_mode']), str(row['quantity_mode'])
                )
                is_active = int(row['active'] or 0) == 1
                tree.insert(
                    "", "end", iid=str(row['id']),
                    values=(
                        row['id'], option_name, row['category'], stock_text,
                        quantity_text, row['sort_order'],
                        L("Active", "启用") if is_active else L("Inactive", "停用"),
                    ),
                    tags=() if is_active else ("inactive",),
                )
            result_var.set(L(f"{len(rows)} option(s)", f"{len(rows)} 个选项"))
            if preselect_id is not None and tree.exists(str(preselect_id)):
                tree.selection_set(str(preselect_id))
                tree.focus(str(preselect_id))
                tree.see(str(preselect_id))
                populate_form()

        def clear_form():
            option_id[0] = None
            name_en_var.set("")
            name_zh_var.set("")
            category_var.set("Ingredients")
            quantity_mode_var.set(quantity_labels["none"])
            stock_item_var.set(no_stock_label)
            update_stock_link_fields()
            active_var.set(True)
            with self.db.get_connection() as conn:
                next_order = conn.execute(
                    "SELECT COALESCE(MAX(sort_order),0)+10 FROM expense_options"
                ).fetchone()[0]
            sort_var.set(str(int(next_order or 10)))
            help_en_text.delete("1.0", "end")
            help_zh_text.delete("1.0", "end")
            form_status_var.set(L("New option", "正在新增选项"))
            tree.selection_remove(tree.selection())
            name_en_entry.focus_set()

        def populate_form(_event=None):
            selection = tree.selection()
            if not selection:
                return
            selected_id = int(selection[0])
            with self.db.get_connection() as conn:
                row = conn.execute(
                    "SELECT * FROM expense_options WHERE id=?", (selected_id,)
                ).fetchone()
            if not row:
                return
            option_id[0] = selected_id
            name_en_var.set(str(row['name_en'] or ''))
            name_zh_var.set(str(row['name_zh'] or ''))
            category_var.set(str(row['category'] or ''))
            quantity_mode_var.set(
                quantity_labels.get(str(row['quantity_mode']), quantity_labels['none'])
            )
            stock_item_var.set(stock_display_for_id(row['stock_item_id']))
            update_stock_link_fields()
            sort_var.set(str(int(row['sort_order'] or 0)))
            active_var.set(int(row['active'] or 0) == 1)
            help_en_text.delete("1.0", "end")
            help_en_text.insert("1.0", str(row['help_en'] or ''))
            help_zh_text.delete("1.0", "end")
            help_zh_text.insert("1.0", str(row['help_zh'] or ''))
            form_status_var.set(L(f"Editing option #{selected_id}", f"正在修改选项 #{selected_id}"))

        def save_option():
            name_en = name_en_var.get().strip()
            name_zh = name_zh_var.get().strip()
            category = category_var.get().strip()
            if not name_en or not category:
                messagebox.showerror(
                    L("Missing Information", "资料不完整"),
                    L("English name and report category are required.", "必须填写英文名称与报表类别。"),
                    parent=win,
                )
                return
            try:
                sort_order = int(sort_var.get().strip())
            except ValueError:
                messagebox.showerror(
                    L("Invalid Order", "排序无效"),
                    L("Display order must be a whole number.", "显示排序必须是整数。"),
                    parent=win,
                )
                return
            quantity_mode = quantity_from_label.get(quantity_mode_var.get())
            stock_item_id = stock_choice_map.get(stock_item_var.get())
            if stock_item_id is not None:
                quantity_mode = "required"
            if not quantity_mode:
                messagebox.showerror(
                    L("Missing Settings", "设置不完整"),
                    L("Choose the quantity behaviour.", "请选择数量的填写方式。"),
                    parent=win,
                )
                return

            help_en = help_en_text.get("1.0", "end").strip()
            help_zh = help_zh_text.get("1.0", "end").strip()
            selected_id = option_id[0]
            now_text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with self.db.get_connection() as conn:
                duplicate = conn.execute(
                    """SELECT id FROM expense_options
                       WHERE LOWER(TRIM(name_en))=LOWER(TRIM(?)) AND id<>?""",
                    (name_en, int(selected_id or 0)),
                ).fetchone()
                if duplicate:
                    messagebox.showerror(
                        L("Duplicate Option", "选项重复"),
                        L(
                            "An option with this English name already exists. Edit or reactivate it instead.",
                            "相同英文名称的选项已经存在，请修改或重新启用原有选项。"
                        ),
                        parent=win,
                    )
                    return
                conn.execute(
                    """INSERT INTO cashflow_categories (name,type)
                       SELECT ?,'Expense' WHERE NOT EXISTS (
                           SELECT 1 FROM cashflow_categories WHERE name=? AND type='Expense'
                       )""",
                    (category, category),
                )
                if selected_id is None:
                    cursor = conn.execute(
                        """INSERT INTO expense_options
                           (option_key,name_en,name_zh,category,detail_mode,quantity_mode,
                            stock_item_id,help_en,help_zh,sort_order,active,created_at,updated_at)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            f"custom:{uuid.uuid4().hex}", name_en, name_zh, category,
                            "fixed", quantity_mode, stock_item_id, help_en, help_zh, sort_order,
                            1 if active_var.get() else 0, now_text, now_text,
                        ),
                    )
                    selected_id = int(cursor.lastrowid)
                    action = "EXPENSE_OPTION_ADD"
                else:
                    conn.execute(
                        """UPDATE expense_options
                           SET name_en=?,name_zh=?,category=?,detail_mode='fixed',quantity_mode=?,
                               stock_item_id=?,help_en=?,help_zh=?,sort_order=?,active=?,updated_at=?
                           WHERE id=?""",
                        (
                            name_en, name_zh, category, quantity_mode, stock_item_id,
                            help_en, help_zh, sort_order,
                            1 if active_var.get() else 0, now_text, int(selected_id),
                        ),
                    )
                    action = "EXPENSE_OPTION_UPDATE"
            self.record_action(action, f"{name_en} | {category}")
            load_category_choices()
            load_options(selected_id)
            form_status_var.set(L("Saved successfully", "保存成功"))

        def toggle_selected():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(
                    L("No Selection", "未选择"),
                    L("Select an expense option first.", "请先选择一个开销选项。"),
                    parent=win,
                )
                return
            selected_id = int(selection[0])
            with self.db.get_connection() as conn:
                row = conn.execute(
                    "SELECT name_en,active FROM expense_options WHERE id=?", (selected_id,)
                ).fetchone()
                if not row:
                    return
                new_active = 0 if int(row['active'] or 0) else 1
                conn.execute(
                    "UPDATE expense_options SET active=?,updated_at=? WHERE id=?",
                    (new_active, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), selected_id),
                )
            self.record_action(
                "EXPENSE_OPTION_STATUS",
                f"{row['name_en']} -> {'Active' if new_active else 'Inactive'}",
            )
            load_options(selected_id)

        self._build_query_toolbar(
            query_host, search_var, load_options,
            on_refresh=lambda: (load_category_choices(), load_stock_choices(), load_options()),
            on_export=lambda: self._export_treeview_csv(
                tree, win, "expense_options.csv"
            ),
            result_var=result_var,
            search_label=L("Search options:", "搜索选项："),
        )
        RoundedButton(
            list_buttons, text="＋ " + L("New Option", "新增选项"), command=clear_form,
            bg=C['primary'], fg="white", font=("Arial", 10, "bold"),
            padx=13, pady=7, radius=8
        ).pack(side="left", padx=4)
        RoundedButton(
            list_buttons, text=L("Enable / Disable", "启用／停用"), command=toggle_selected,
            bg=C['neutral'], fg="white", font=("Arial", 10, "bold"),
            padx=13, pady=7, radius=8
        ).pack(side="left", padx=4)
        RoundedButton(
            list_buttons, text=L("Close", "关闭"), command=win.destroy,
            bg=C['neutral'], fg="white", font=("Arial", 10, "bold"),
            padx=13, pady=7, radius=8
        ).pack(side="right", padx=4)
        RoundedButton(
            editor_buttons, text="💾 " + L("Save Option", "保存选项"), command=save_option,
            bg=C['success'], fg="white", font=("Arial", 10, "bold"),
            padx=15, pady=8, radius=8
        ).pack(side="right", padx=4)

        tree.bind("<<TreeviewSelect>>", populate_form)
        tree.bind("<Double-1>", populate_form)
        stock_item_combo.bind("<<ComboboxSelected>>", update_stock_link_fields)
        load_category_choices()
        load_stock_choices()
        load_options()
        clear_form()

    def show_admin_panel(self):
        C = self.colors
        L = self.l
        self.record_action("OPEN_ADMIN_PANEL", "Opened Admin Control Center")
        win = tk.Toplevel(self.root)
        win.title(L("Admin Panel", "管理面板"))
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        self._center_window(win, 980, 640, self.root)

        hdr = tk.Frame(win, bg=C['header'], height=62)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=L("🛠️  Admin Control Center", "🛠️  管理控制中心"), font=("Arial", 18, "bold"),
                 bg=C['header'], fg="white").pack(side="left", padx=18, pady=14)
        tk.Label(hdr, text=L("Organized access to all manager functions", "集中进入所有管理功能"),
                 font=("Arial", 10), bg=C['header'], fg="#94A3B8").pack(side="left", pady=16)

        body = tk.Frame(win, bg=C['bg'])
        body.pack(fill="both", expand=True, padx=16, pady=10)

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM items WHERE active=1")
            item_count = int(c.fetchone()[0] or 0)
            c.execute("SELECT COUNT(*) FROM staff WHERE active=1")
            staff_count = int(c.fetchone()[0] or 0)
        held_count = self._count_active_unpaid_orders()

        stats = tk.Frame(body, bg=C['bg'])
        stats.pack(fill="x", pady=(0, 12))
        cards = [
            (L("📦 Active Items", "📦 启用商品"), str(item_count), C['primary']),
            (L("👥 Active Staff", "👥 在职员工"), str(staff_count), C['neutral']),
            (L("🧾 Unpaid Holds", "🧾 未付款订单"), str(held_count), C['danger']),
        ]
        for title, val, color in cards:
            card = tk.Frame(stats, bg=color, padx=14, pady=10)
            card.pack(side="left", padx=6, fill="x", expand=True)
            tk.Label(card, text=title, font=("Arial", 9, "bold"), bg=color, fg="#E2E8F0").pack(anchor="w")
            tk.Label(card, text=val, font=("Arial", 15, "bold"), bg=color, fg="white").pack(anchor="w", pady=(4, 0))

        groups_wrap = tk.Frame(body, bg=C['bg'])
        groups_wrap.pack(fill="both", expand=True)
        groups_wrap.grid_columnconfigure(0, weight=1, uniform="admin_columns")
        groups_wrap.grid_columnconfigure(1, weight=1, uniform="admin_columns")
        groups_wrap.grid_rowconfigure(0, weight=1)

        def group_card(parent, title, subtitle):
            f = tk.Frame(parent, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
            tk.Label(f, text=title, font=("Arial", 13, "bold"), bg=C['panel'], fg=C['text']).pack(anchor="w", padx=14, pady=(12, 2))
            tk.Label(f, text=subtitle, font=("Arial", 9), bg=C['panel'], fg=C['text_muted']).pack(anchor="w", padx=14, pady=(0, 10))
            return f

        left_col = tk.Frame(groups_wrap, bg=C['bg'])
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        right_col = tk.Frame(groups_wrap, bg=C['bg'])
        right_col.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ops = group_card(left_col, L("Operations", "营运管理"), L("Manage menu, team and order corrections", "管理菜单、员工及订单修正"))
        ops.pack(fill="both", expand=True, pady=(0, 8))

        control = group_card(
            right_col,
            L("Management & Control", "管理与控制"),
            L("Team access, corrections, finance and system checks", "员工权限、订单修正、财务及系统检查")
        )
        control.pack(fill="both", expand=True, pady=(0, 8))

        def admin_action_row(parent, icon, title, subtitle, command, accent, close_panel=True):
            row = tk.Frame(
                parent, bg="#F8FAFC", height=48, cursor="hand2",
                highlightbackground=C['border'], highlightthickness=1
            )
            row.pack(fill="x", padx=14, pady=2)
            row.pack_propagate(False)
            tk.Frame(row, bg=accent, width=5).pack(side="left", fill="y")

            icon_label = tk.Label(
                row, text=icon, font=("Arial", 15), width=3,
                bg="#F8FAFC", fg=accent, cursor="hand2"
            )
            icon_label.pack(side="left", padx=(5, 1))

            text_wrap = tk.Frame(row, bg="#F8FAFC", cursor="hand2")
            text_wrap.pack(side="left", fill="both", expand=True, pady=6)
            title_label = tk.Label(
                text_wrap, text=title, font=("Arial", 10, "bold"),
                bg="#F8FAFC", fg="#1E293B", anchor="w", cursor="hand2"
            )
            title_label.pack(fill="x")
            subtitle_label = tk.Label(
                text_wrap, text=subtitle, font=("Arial", 8),
                bg="#F8FAFC", fg="#64748B", anchor="w", cursor="hand2"
            )
            subtitle_label.pack(fill="x", pady=(1, 0))

            def open_action(_event=None):
                if close_panel:
                    win.destroy()
                self._run_with_loading(
                    command,
                    parent=(self.root if close_panel else win)
                )

            open_button = tk.Button(
                row, text=L("Open ›", "打开 ›"), command=open_action,
                font=("Arial", 8, "bold"), bg=accent, fg="white",
                activebackground=accent, activeforeground="white",
                relief="flat", bd=0, padx=8, pady=4, cursor="hand2"
            )
            open_button.pack(side="right", padx=8)
            for clickable in (row, icon_label, text_wrap, title_label, subtitle_label):
                clickable.bind("<Button-1>", open_action)

        section_style = {
            "font": ("Arial", 8, "bold"),
            "bg": C['panel'],
            "fg": C['text_muted'],
            "anchor": "w",
        }
        tk.Label(ops, text=L("SETUP & OPERATIONS", "设置与营运"), **section_style).pack(
            fill="x", padx=14, pady=(1, 1)
        )
        admin_action_row(
            ops, "⚙️", L("Items & Inventory Settings", "商品与库存设置"),
            L("Menu items, prices, variants and stock settings", "商品、价格、选项及库存设置"),
            self.manage_items, C['primary']
        )
        admin_action_row(
            ops, "📂", L("Product Categories", "商品分类"),
            L("Category names, colours, icons and display order", "类别名称、颜色、图标及显示顺序"),
            self.manage_categories, C['info']
        )
        admin_action_row(
            ops, "💸", L("Expense Management", "开销选项管理"),
            L("Add, edit, order or disable Cash Out choices", "新增、修改、排序或停用现金转出选项"),
            self.manage_expense_options, C['warning']
        )
        admin_action_row(
            ops, "🪑", L("Table Settings", "桌位设置"),
            L("Create and maintain dining table codes", "新增及维护用餐桌位编号"),
            self.manage_tables, C['primary']
        )
        admin_action_row(
            ops, "🏪", L("Tenant & Stall Management", "租户与档口管理"),
            L("Tenants, stall details and rental settings", "租户、档口资料及租金设置"),
            self.manage_stalls, C['success']
        )

        tk.Label(control, text=L("MANAGER ACCESS", "经理管理"), **section_style).pack(
            fill="x", padx=14, pady=(1, 1)
        )
        admin_action_row(
            control, "👥", L("Staff Management", "员工管理"),
            L("Staff accounts, roles, PINs and access", "员工账号、角色、PIN及权限"),
            self.manage_staff, C['neutral']
        )
        admin_action_row(
            control, "↩️", L("Refund Management", "退款管理"),
            L("Review orders and process controlled refunds", "查看订单及处理受控退款"),
            self.show_refunds, C['danger']
        )

        tk.Label(control, text=L("FINANCE & SYSTEM", "财务与系统"), **section_style).pack(
            fill="x", padx=14, pady=(6, 1)
        )
        admin_action_row(
            control, "📊", L("Finance & Reports", "财务与报表"),
            L("Records, cash closing, analytics, profit and exports", "销售记录、现金日结、分析、利润及导出"),
            self.show_finance_reports_center, C['primary']
        )
        admin_action_row(
            control, "🔍", L("System Data Audit", "系统资料检查"),
            L("Check today's sales summary for inconsistencies", "检查今日销售汇总是否有不一致"),
            self.run_consistency_audit, C['danger'], close_panel=False
        )

        foot = tk.Frame(win, bg=C['bg'])
        foot.pack(fill="x", padx=16, pady=(0, 12))
        RoundedButton(foot, text=L("Close", "关闭"), command=win.destroy,
                      bg=C['neutral'], fg="white", font=("Arial", 10, "bold"),
                      padx=16, pady=7, radius=8).pack(side="right")

    def manage_staff(self):
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("👥 Staff Management", "👥 员工管理"))
        win.geometry("900x750")
        win.configure(bg=self.colors['bg'])
        win.transient(self.root)
        win.grab_set()

        header = tk.Frame(win, bg=self.colors['header'], height=52)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=L("👥 Staff Management", "👥 员工管理"), font=("Arial", 16, "bold"),
                 bg=self.colors['header'], fg='white').pack(pady=10)

        btn_f = tk.Frame(win, bg=self.colors['bg'])
        btn_f.pack(fill="x", side="bottom", padx=20, pady=10)

        staff_search_var = tk.StringVar()
        staff_result_var = tk.StringVar()
        query_host = tk.Frame(win, bg=self.colors['bg'])
        query_host.pack(fill="x", padx=4, pady=(8, 0))

        tree_frame = tk.Frame(win, bg=self.colors['bg'])
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("ID", "Staff ID", "Name", "Role", "Active")
        stree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=12)
        cw = {"ID": 50, "Staff ID": 100, "Name": 160, "Role": 100, "Active": 70}
        headings = {
            "ID": "ID", "Staff ID": L("Staff ID", "员工编号"),
            "Name": L("Name", "姓名"), "Role": L("Role", "角色"),
            "Active": L("Active", "启用"),
        }
        for col in cols:
            stree.heading(col, text=headings[col])
            stree.column(col, width=cw.get(col, 100), anchor="center")
        sb = tk.Scrollbar(tree_frame, orient="vertical", command=stree.yview)
        stree.configure(yscrollcommand=sb.set)
        stree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def load_staff():
            stree.delete(*stree.get_children())
            keyword = staff_search_var.get().strip().casefold()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT id, staff_id, name, role, active FROM staff ORDER BY name")
                rows = c.fetchall()
                if keyword:
                    rows = [
                        r for r in rows
                        if keyword in " ".join(str(r[key] or "") for key in (
                            'id', 'staff_id', 'name', 'role', 'active'
                        )).casefold()
                    ]
                for r in rows:
                    stree.insert("", "end", values=(r['id'], r['staff_id'], r['name'],
                                                     r['role'], "✅" if r['active'] else "❌"))
            staff_result_var.set(
                L(f"{len(rows)} staff member(s)", f"{len(rows)} 位员工")
            )

        def add_staff_dialog():
            d = tk.Toplevel(win)
            d.title("Add Staff")
            d.geometry("500x460")
            d.configure(bg=self.colors['bg'])
            d.transient(win)
            d.grab_set()

            tk.Label(d, text="Add New Staff", font=("Arial", 14, "bold"),
                     bg=self.colors['bg']).pack(pady=16)

            fields = {}
            for label, key in [("Staff ID:", "sid"), ("Name:", "name"), ("PIN:", "pin")]:
                f = tk.Frame(d, bg=self.colors['bg'])
                f.pack(fill="x", padx=40, pady=8)
                tk.Label(f, text=label, font=("Arial", 12), bg=self.colors['bg'],
                         width=12, anchor="w").pack(side="left")
                fields[key] = tk.Entry(f, font=("Arial", 12), width=18,
                                       show="●" if key == "pin" else "")
                fields[key].pack(side="left")

            role_frame = tk.Frame(d, bg=self.colors['bg'])
            role_frame.pack(fill="x", padx=40, pady=8)
            tk.Label(role_frame, text="Role:", font=("Arial", 12), bg=self.colors['bg'],
                     width=12, anchor="w").pack(side="left")
            role_var = tk.StringVar(value="cashier")
            ttk.Combobox(role_frame, textvariable=role_var,
                         values=["cashier", "admin"], font=("Arial", 12),
                         width=16, state="readonly").pack(side="left")

            def save():
                sid = fields['sid'].get().strip().upper()
                name = fields['name'].get().strip()
                pin = fields['pin'].get().strip()
                if not sid or not name or not pin:
                    self._msg_error("Error", "错误", "All fields required", "所有栏位都必须填写。", d)
                    return
                try:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("INSERT INTO staff (staff_id, name, pin, role, created_at) VALUES (?,?,?,?,?)",
                                  (sid, name, pin, role_var.get(),
                                   datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    self._msg_info("Success", "成功", f"Staff '{name}' added!", f"已新增员工“{name}”。", d)
                    d.destroy()
                    load_staff()
                except Exception as e:
                    messagebox.showerror(L("Error", "错误"), str(e), parent=d)

            RoundedButton(d, text="💾 Save", command=save,
                          bg=self.colors['success'], fg="white",
                          font=("Arial", 12, "bold"), padx=20, pady=8, radius=8
                          ).pack(pady=14)
            fields['sid'].focus_set()

        def change_pin():
            sel = stree.selection()
            if not sel:
                self._msg_warning("No Selection", "未选择", "Select a staff member", "请选择一位员工。", win)
                return
            row = stree.item(sel[0])['values']
            db_id = row[0]
            new_pin = simpledialog.askstring(
                L("Change PIN", "更改 PIN"),
                L(f"New PIN for {row[2]}:", f"请输入 {row[2]} 的新 PIN："),
                parent=win, show="●"
            )
            if new_pin:
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("UPDATE staff SET pin=? WHERE id=?", (new_pin.strip(), db_id))
                self._msg_info("Done", "完成", "PIN updated!", "PIN 已更新。", win)

        def toggle_active():
            sel = stree.selection()
            if not sel:
                self._msg_warning("No Selection", "未选择", "Select a staff member", "请选择一位员工。", win)
                return
            row = stree.item(sel[0])['values']
            db_id = row[0]
            name = row[2]
            role = row[3]
            current_active = row[4] == "✅"
            # Warn before deactivating an admin-role staff
            if current_active and str(role).lower() == 'admin':
                if not messagebox.askyesno(
                    L("Confirm Deactivate Admin", "确认停用管理员"),
                    L(
                        f"'{name}' has the Admin role.\nAre you sure you want to deactivate this account?",
                        f"“{name}”拥有管理员角色。\n确定要停用此账号吗？"
                    ),
                    parent=win
                ):
                    return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("UPDATE staff SET active = CASE WHEN active=1 THEN 0 ELSE 1 END WHERE id=?", (db_id,))
            load_staff()

        def delete_staff():
            sel = stree.selection()
            if not sel:
                self._msg_warning("No Selection", "未选择", "Select a staff member", "请选择一位员工。", win)
                return
            row = stree.item(sel[0])['values']
            db_id = row[0]
            name = row[2]
            
            # Safety: don't delete the primary Admin (ID 1)
            if str(db_id) == "1":
                self._msg_error("Error", "错误", "Primary Admin account cannot be deleted.", "主要管理员账号不能删除。", win)
                return
            
            if not messagebox.askyesno(
                L("Confirm Delete", "确认删除"),
                L(
                    f"Permanently delete staff '{name}'?\nThis action cannot be undone.",
                    f"永久删除员工“{name}”？\n此操作无法复原。"
                ), parent=win
            ):
                return
            
            def _proceed_delete(admin_staff):
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("DELETE FROM staff WHERE id=?", (db_id,))
                load_staff()
                messagebox.showinfo(L("Success", "成功"), L(f"Staff '{name}' removed.", f"已删除员工“{name}”。"), parent=win)

            self._staff_login_popup(callback=_proceed_delete, title=L("👤 Admin PIN Required", "👤 需要管理员 PIN"), parent=win)

        def change_role():
            sel = stree.selection()
            if not sel:
                self._msg_warning("No Selection", "未选择", "Select a staff member", "请选择一位员工。", win)
                return
            row = stree.item(sel[0])['values']
            db_id = row[0]
            name = row[2]
            current_role = row[3]

            # Safety: primary Admin (ID 1) must remain admin
            if str(db_id) == "1":
                self._msg_error("Error", "错误", "Primary Admin role cannot be changed.", "主要管理员的角色不能更改。", win)
                return

            new_role = "cashier" if str(current_role).lower() == "admin" else "admin"

            if not messagebox.askyesno(
                L("Confirm Change Role", "确认更改角色"),
                L(
                    f"Change role of '{name}' from '{current_role}' to '{new_role}'?",
                    f"将“{name}”的角色从“{current_role}”改为“{new_role}”？"
                ),
                parent=win
            ):
                return

            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("UPDATE staff SET role=? WHERE id=?", (new_role, db_id))
            messagebox.showinfo(L("Success", "成功"), L(f"Role updated for '{name}' to {new_role}!", f"“{name}”的角色已更新为 {new_role}。"), parent=win)
            load_staff()

        # Buttons were moved up to pack side="bottom" before tree_frame expands.

        self._build_query_toolbar(
            query_host, staff_search_var, load_staff,
            on_refresh=load_staff,
            on_export=lambda: self._export_treeview_csv(
                stree, win, "staff_list.csv"
            ),
            result_var=staff_result_var,
            search_label=L("Search staff:", "搜索员工：")
        )

        for text, cmd, color in [
            (L("➕ Add Staff", "➕ 新增员工"), add_staff_dialog, self.colors['success']),
            (L("🔑 Change PIN", "🔑 更改 PIN"), change_pin, self.colors['card']),
            (L("🔄 Change Role", "🔄 更改角色"), change_role, self.colors['card']),
            (L("🔄 Toggle Active", "🔄 启用/停用"), toggle_active, self.colors['card']),
            (L("🗑️ Remove", "🗑️ 删除"), delete_staff, self.colors['danger']),
        ]:
            RoundedButton(btn_f, text=text, command=cmd, font=("Arial", 11, "bold"),
                          bg=color, fg="white", padx=14, pady=8, radius=8
                          ).pack(side="left", padx=4)

        load_staff()

    def manage_tables(self):
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("🪑 Table Maintenance", "🪑 桌位维护"))
        win.geometry("640x600")
        win.configure(bg=self.colors['bg'])
        win.transient(self.root)
        win.grab_set()

        hdr = tk.Frame(win, bg=self.colors['header'], height=54)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=L("🪑 Table Maintenance", "🪑 桌位维护"), font=("Arial", 16, "bold"),
                 bg=self.colors['header'], fg="white").pack(side="left", padx=14, pady=10)

        btn_f = tk.Frame(win, bg=self.colors['bg'])
        btn_f.pack(fill="x", side="bottom", padx=14, pady=12)

        table_search_var = tk.StringVar()
        table_result_var = tk.StringVar()
        query_host = tk.Frame(win, bg=self.colors['bg'])
        query_host.pack(fill="x", padx=2, pady=(8, 0))

        tree_wrap = tk.Frame(win, bg=self.colors['bg'])
        tree_wrap.pack(fill="both", expand=True, padx=14, pady=12)
        cols = ("ID", "Table", "Status", "Sort")
        ttree = ttk.Treeview(tree_wrap, columns=cols, show="headings", height=12)
        cw = {"ID": 60, "Table": 160, "Status": 120, "Sort": 100}
        headings = {
            "ID": "ID", "Table": L("Table", "桌位"),
            "Status": L("Status", "状态"), "Sort": L("Sort", "排序"),
        }
        for col in cols:
            ttree.heading(col, text=headings[col])
            ttree.column(col, width=cw.get(col, 120), anchor="center")
        sb = tk.Scrollbar(tree_wrap, orient="vertical", command=ttree.yview)
        ttree.configure(yscrollcommand=sb.set)
        ttree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def load_tables():
            ttree.delete(*ttree.get_children())
            keyword = table_search_var.get().strip().casefold()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("""SELECT id, table_code, active, sort_order
                             FROM dining_tables
                             ORDER BY sort_order ASC, table_code ASC""")
                rows = c.fetchall()
                if keyword:
                    rows = [
                        r for r in rows
                        if keyword in " ".join(str(r[key] or "") for key in (
                            'id', 'table_code', 'active', 'sort_order'
                        )).casefold()
                    ]
                for r in rows:
                    ttree.insert("", "end", values=(
                        r['id'],
                        r['table_code'],
                        "ACTIVE" if int(r['active'] or 0) == 1 else "INACTIVE",
                        int(r['sort_order'] or 0)
                    ))
            table_result_var.set(L(f"{len(rows)} table(s)", f"{len(rows)} 个桌位"))

        def add_table():
            code = simpledialog.askstring(L("Add Table", "新增桌位"), L("Table code (e.g. A7, B4):", "桌位编号（例如 A7、B4）："), parent=win)
            if not code:
                return
            table_code = code.strip().upper()
            if not table_code:
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT COALESCE(MAX(sort_order),0) FROM dining_tables")
                next_sort = int(c.fetchone()[0] or 0) + 1
                try:
                    c.execute(
                        "INSERT INTO dining_tables (table_code, sort_order, active, created_at) VALUES (?,?,1,?)",
                        (table_code, next_sort, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    )
                except Exception:
                    messagebox.showerror(L("Error", "错误"), L(f"Table '{table_code}' already exists.", f"桌位“{table_code}”已经存在。"), parent=win)
                    return
            self.record_action("TABLE_ADD", f"Added table {table_code}")
            load_tables()
            self.refresh_table_cards()
            self.refresh_table_home_cards()

        def toggle_table_active():
            sel = ttree.selection()
            if not sel:
                self._msg_warning("No Selection", "未选择", "Select a table", "请选择一个桌位。", win)
                return
            row = ttree.item(sel[0])['values']
            table_id = int(row[0])
            table_code = str(row[1])
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("UPDATE dining_tables SET active = CASE WHEN active=1 THEN 0 ELSE 1 END WHERE id=?", (table_id,))
            self.record_action("TABLE_TOGGLE", f"Toggled table {table_code}")
            load_tables()
            self.refresh_table_cards()
            self.refresh_table_home_cards()

        self._build_query_toolbar(
            query_host, table_search_var, load_tables,
            on_refresh=load_tables,
            on_export=lambda: self._export_treeview_csv(
                ttree, win, "table_list.csv"
            ),
            result_var=table_result_var,
            search_label=L("Search tables:", "搜索桌位：")
        )

        for txt, cmd, col in [
            (L("➕ Add Table", "➕ 新增桌位"), add_table, self.colors['success']),
            (L("🔄 Toggle Active", "🔄 启用/停用"), toggle_table_active, self.colors['warning']),
        ]:
            RoundedButton(btn_f, text=txt, command=cmd,
                          bg=col, fg="white", font=("Arial", 11, "bold"),
                          padx=12, pady=8, radius=8).pack(side="left", padx=4)

        RoundedButton(btn_f, text=L("✖ Close", "✖ 关闭"), command=win.destroy,
                      bg=self.colors['neutral'], fg="white", font=("Arial", 11, "bold"),
                      padx=12, pady=8, radius=8).pack(side="right", padx=4)

        load_tables()

    def manage_cashflow_categories(self):
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("💸 Expense/Income Categories", "💸 收支类别"))
        win.geometry("700x550")
        win.configure(bg=self.colors['bg'])
        win.transient(self.root)
        win.grab_set()

        header = tk.Frame(win, bg=self.colors['danger'], height=52)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=L("💸 Expense/Income Category Management", "💸 收支类别管理"), font=("Arial", 16, "bold"),
                 bg=self.colors['danger'], fg='white').pack(pady=10)

        category_search_var = tk.StringVar()
        category_result_var = tk.StringVar()
        query_host = tk.Frame(win, bg=self.colors['bg'])
        query_host.pack(fill="x", padx=4, pady=(8, 0))

        tree_frame = tk.Frame(win, bg=self.colors['bg'])
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("ID", "Name", "Type")
        ctree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=12)
        cw = {"ID": 60, "Name": 300, "Type": 150}
        headings = {
            "ID": "ID", "Name": L("Name", "名称"), "Type": L("Type", "类型")
        }
        for col in cols:
            ctree.heading(col, text=headings[col])
            ctree.column(col, width=cw.get(col, 100), anchor="center")
        sb = tk.Scrollbar(tree_frame, orient="vertical", command=ctree.yview)
        ctree.configure(yscrollcommand=sb.set)
        ctree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def load_cats():
            ctree.delete(*ctree.get_children())
            keyword = category_search_var.get().strip().casefold()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT id, name, type FROM cashflow_categories ORDER BY type, name")
                rows = c.fetchall()
                if keyword:
                    rows = [
                        r for r in rows
                        if keyword in " ".join(
                            str(r[key] or "") for key in ('id', 'name', 'type')
                        ).casefold()
                    ]
                for r in rows:
                    ctree.insert("", "end", values=(r['id'], r['name'], r['type']))
            category_result_var.set(
                L(f"{len(rows)} category(s)", f"{len(rows)} 个类别")
            )

        def add_cat_dialog():
            d = tk.Toplevel(win)
            d.title("Add Category")
            d.geometry("400x300")
            d.configure(bg=self.colors['bg'])
            d.transient(win)
            d.grab_set()

            tk.Label(d, text="Add New Category", font=("Arial", 14, "bold"), bg=self.colors['bg']).pack(pady=20)
            
            f1 = tk.Frame(d, bg=self.colors['bg'])
            f1.pack(fill="x", padx=30, pady=8)
            tk.Label(f1, text="Name:", font=("Arial", 11, "bold"), bg=self.colors['bg'], width=10, anchor="w").pack(side="left")
            name_ent = tk.Entry(f1, font=("Arial", 11), width=20)
            name_ent.pack(side="left")

            f2 = tk.Frame(d, bg=self.colors['bg'])
            f2.pack(fill="x", padx=30, pady=8)
            tk.Label(f2, text="Type:", font=("Arial", 11, "bold"), bg=self.colors['bg'], width=10, anchor="w").pack(side="left")
            type_var = tk.StringVar(value="Expense")
            ttk.Combobox(f2, textvariable=type_var, values=["Expense", "Income"], state="readonly", width=18).pack(side="left")

            def save():
                name = name_ent.get().strip()
                ctype = type_var.get()
                if not name:
                    self._msg_error("Error", "错误", "Name is required", "请输入类别名称。", d); return
                try:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("INSERT INTO cashflow_categories (name, type) VALUES (?,?)", (name, ctype))
                    self._msg_info("Success", "成功", f"Category '{name}' added!", f"已新增类别“{name}”。", d)
                    d.destroy(); load_cats()
                except Exception as e:
                    messagebox.showerror(L("Error", "错误"), str(e), parent=d)

            RoundedButton(d, text="💾 Save", command=save, bg=self.colors['success'], fg="white",
                          font=("Arial", 12, "bold"), padx=20, pady=8, radius=8).pack(pady=20)

        def delete_cat():
            sel = ctree.selection()
            if not sel: return
            row = ctree.item(sel[0])['values']
            if not messagebox.askyesno(L("Confirm", "确认"), L(f"Delete category '{row[1]}'?", f"删除类别“{row[1]}”？"), parent=win): return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM cashflow_categories WHERE id=?", (row[0],))
            load_cats()

        self._build_query_toolbar(
            query_host, category_search_var, load_cats,
            on_refresh=load_cats,
            on_export=lambda: self._export_treeview_csv(
                ctree, win, "cashflow_categories.csv"
            ),
            result_var=category_result_var,
            search_label=L("Search categories:", "搜索类别：")
        )

        btn_f = tk.Frame(win, bg=self.colors['bg'])
        btn_f.pack(fill="x", padx=20, pady=10)
        RoundedButton(btn_f, text=L("➕ Add Category", "➕ 新增类别"), command=add_cat_dialog, bg=self.colors['success'], fg="white").pack(side="left", padx=5)
        RoundedButton(btn_f, text=L("🗑 Delete Selected", "🗑 删除所选"), command=delete_cat, bg=self.colors['neutral'], fg="white").pack(side="left", padx=5)
        RoundedButton(btn_f, text=L("❌ Close", "❌ 关闭"), command=win.destroy, bg=self.colors['header'], fg="white").pack(side="right", padx=5)

        load_cats()

    def manage_categories(self):
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("📂 Category Management", "📂 商品类别管理"))
        win.geometry("900x640")
        win.configure(bg=self.colors['bg'])
        win.transient(self.root)
        win.grab_set()

        header = tk.Frame(win, bg=self.colors['header'], height=52)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=L("📂 Category Management", "📂 商品类别管理"), font=("Arial", 16, "bold"),
                 bg=self.colors['header'], fg='white').pack(pady=10)

        product_category_search_var = tk.StringVar()
        product_category_result_var = tk.StringVar()
        query_host = tk.Frame(win, bg=self.colors['bg'])
        query_host.pack(fill="x", padx=4, pady=(8, 0))

        tree_frame = tk.Frame(win, bg=self.colors['bg'])
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("ID", "Name", "Icon", "Color", "Active", "Order", "Export", "List Price")
        ctree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=12)
        cw = {"ID": 50, "Name": 160, "Icon": 80, "Color": 120, "Active": 80, "Order": 80, "Export": 80}
        headings = {
            "ID": "ID", "Name": L("Name", "名称"), "Icon": L("Icon", "图标"),
            "Color": L("Color", "颜色"), "Active": L("Active", "启用"),
            "Order": L("Order", "顺序"), "Export": L("Export", "导出"),
            "List Price": L("List Price", "标价"),
        }
        for col in cols:
            ctree.heading(col, text=headings[col])
            ctree.column(col, width=cw.get(col, 100), anchor="center")
        sb = tk.Scrollbar(tree_frame, orient="vertical", command=ctree.yview)
        ctree.configure(yscrollcommand=sb.set)
        ctree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def load_all_categories():
            ctree.delete(*ctree.get_children())
            keyword = product_category_search_var.get().strip().casefold()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT id, name, icon, color, active, sort_order, exclude_export, show_list_price FROM product_categories ORDER BY sort_order")
                rows = c.fetchall()
                if keyword:
                    rows = [
                        r for r in rows
                        if keyword in " ".join(str(r[key] or "") for key in (
                            'id', 'name', 'icon', 'color', 'active',
                            'sort_order', 'exclude_export'
                        )).casefold()
                    ]
                for r in rows:
                    export_status = "❌ Hidden" if r['exclude_export'] else "✅ Shown"
                    ctree.insert("", "end", values=(r['id'], r['name'], r['icon'],
                                                      r['color'], "✅" if r['active'] else "❌", r['sort_order'],
                                                      export_status, f"RM {LIST_PRICE:.2f}" if r['show_list_price'] else "—"))
            product_category_result_var.set(
                L(f"{len(rows)} category(s)", f"{len(rows)} 个类别")
            )

        def list_price_checkbox(dialog, enabled=0):
            setting = tk.IntVar(master=dialog, value=enabled)
            tk.Checkbutton(
                dialog, variable=setting, bg=self.colors['bg'], activebackground=self.colors['bg'],
                text=L(f"Show RM {LIST_PRICE:.2f} per item as list price", f"显示每件 RM {LIST_PRICE:.2f} 标价"),
                font=("Arial", 11),
            ).pack(anchor="w", padx=40, pady=(8, 0))
            tk.Label(
                dialog, bg=self.colors['bg'], fg=self.colors['text_muted'],
                text=L("Checkout uses prices from Item Management.", "结账按商品管理的价格收费。"),
            ).pack(anchor="w", padx=62)
            return setting

        def add_category_dialog():
            d = tk.Toplevel(win)
            d.title("Add Category")
            d.geometry("500x500")
            d.configure(bg=self.colors['bg'])
            d.transient(win)
            d.grab_set()

            tk.Label(d, text="Add New Category", font=("Arial", 14, "bold"),
                     bg=self.colors['bg']).pack(pady=16)

            fields = {}
            for label, key in [("Name:", "name"), ("Icon (Emoji):", "icon"), ("Color (Hex):", "color"), ("Sort Order (Position):", "order")]:
                f = tk.Frame(d, bg=self.colors['bg'])
                f.pack(fill="x", padx=40, pady=6)
                tk.Label(f, text=label, font=("Arial", 11, "bold"), bg=self.colors['bg'],
                         width=18, anchor="w").pack(side="left")
                fields[key] = tk.Entry(f, font=("Arial", 11), width=16)
                fields[key].pack(side="left")

                if key == 'color':
                    def pick_color_add():
                        c = colorchooser.askcolor(initialcolor=fields['color'].get(), parent=d)[1]
                        if c:
                            fields['color'].delete(0, tk.END)
                            fields['color'].insert(0, c)
                    tk.Button(f, text="🎨 Pick", command=pick_color_add, font=("Arial", 9),
                              padx=4, pady=2).pack(side="left", padx=5)

                if key == 'order':
                    tk.Label(d, text="smaller numbers (e.g. 1) appear first", font=("Arial", 8, "italic"),
                             bg=self.colors['bg'], fg=self.colors['text_muted']).pack(anchor="w", padx=205)

            fields['color'].insert(0, "#475569")
            fields['order'].insert(0, "10")

            exclude_var = tk.IntVar(value=0)
            ck_f = tk.Frame(d, bg=self.colors['bg'])
            ck_f.pack(fill="x", padx=40, pady=6)
            tk.Label(ck_f, text="Exclude from Export:", font=("Arial", 11, "bold"), bg=self.colors['bg'], width=18, anchor="w").pack(side="left")
            tk.Checkbutton(ck_f, variable=exclude_var, bg=self.colors['bg'], activebackground=self.colors['bg']).pack(side="left")
            list_price_var = list_price_checkbox(d)

            def save():
                name = fields['name'].get().strip()
                icon = fields['icon'].get().strip() or "📦"
                color = fields['color'].get().strip() or "#475569"
                order = fields['order'].get().strip() or "10"
                exclude = exclude_var.get()
                if not name:
                    messagebox.showerror(L("Error", "错误"), L("Name is required", "请输入名称。"), parent=d)
                    return
                try:
                    order_int = int(order)
                except ValueError:
                    messagebox.showerror(L("Error", "错误"), L("Sort Order must be a whole number (e.g. 1, 10, 100)", "排序必须是整数（例如 1、10、100）。"), parent=d)
                    return

                try:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("INSERT INTO product_categories (name, icon, color, sort_order, exclude_export, show_list_price) VALUES (?,?,?,?,?,?)",
                                  (name, icon, color, order_int, exclude, list_price_var.get()))
                    messagebox.showinfo(L("Success", "成功"), L(f"Category '{name}' added!", f"已新增类别“{name}”。"), parent=d)
                    d.destroy()
                    load_all_categories()
                    self.load_categories()
                    self.load_items()
                except Exception as e:
                    messagebox.showerror(L("Error", "错误"), str(e), parent=d)

            RoundedButton(d, text="💾 Save", command=save,
                          bg=self.colors['success'], fg="white",
                          font=("Arial", 12, "bold"), padx=20, pady=8, radius=8
                          ).pack(pady=14)
            fields['name'].focus_set()

        def edit_category_dialog():
            sel = ctree.selection()
            if not sel:
                messagebox.showwarning(L("No Selection", "未选择"), L("Select a category", "请选择一个类别。"), parent=win)
                return
            row = ctree.item(sel[0])['values']
            cid = row[0]

            d = tk.Toplevel(win)
            d.title("Edit Category")
            d.geometry("500x500")
            d.configure(bg=self.colors['bg'])
            d.transient(win)
            d.grab_set()

            tk.Label(d, text="Edit Category", font=("Arial", 14, "bold"),
                     bg=self.colors['bg']).pack(pady=16)

            fields = {}
            for label, key, val in [("Name:", "name", row[1]), ("Icon (Emoji):", "icon", row[2]), ("Color (Hex):", "color", row[3]), ("Sort Order (Position):", "order", row[5])]:
                f = tk.Frame(d, bg=self.colors['bg'])
                f.pack(fill="x", padx=40, pady=6)
                tk.Label(f, text=label, font=("Arial", 11, "bold"), bg=self.colors['bg'],
                         width=18, anchor="w").pack(side="left")
                fields[key] = tk.Entry(f, font=("Arial", 11), width=16)
                fields[key].insert(0, str(val))
                fields[key].pack(side="left")

                if key == 'color':
                    def pick_color_edit():
                        c = colorchooser.askcolor(initialcolor=fields['color'].get(), parent=d)[1]
                        if c:
                            fields['color'].delete(0, tk.END)
                            fields['color'].insert(0, c)
                    tk.Button(f, text="🎨 Pick", command=pick_color_edit, font=("Arial", 9),
                              padx=4, pady=2).pack(side="left", padx=5)

                if key == 'order':
                    tk.Label(d, text="smaller numbers (e.g. 1) appear first", font=("Arial", 8, "italic"),
                             bg=self.colors['bg'], fg=self.colors['text_muted']).pack(anchor="w", padx=205)

            # Get current exclude_export value
            curr_exclude = 0
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT exclude_export, show_list_price FROM product_categories WHERE id=?", (cid,))
                row_db = c.fetchone()
                if row_db: curr_exclude = row_db['exclude_export']

            exclude_var = tk.IntVar(value=curr_exclude)
            ck_f = tk.Frame(d, bg=self.colors['bg'])
            ck_f.pack(fill="x", padx=40, pady=6)
            tk.Label(ck_f, text="Exclude from Export:", font=("Arial", 11, "bold"), bg=self.colors['bg'], width=18, anchor="w").pack(side="left")
            tk.Checkbutton(ck_f, variable=exclude_var, bg=self.colors['bg'], activebackground=self.colors['bg']).pack(side="left")
            list_price_var = list_price_checkbox(d, row_db['show_list_price'] if row_db else 0)

            def update():
                name = fields['name'].get().strip()
                icon = fields['icon'].get().strip()
                color = fields['color'].get().strip()
                order = fields['order'].get().strip()
                exclude = exclude_var.get()
                if not name:
                    messagebox.showerror(L("Error", "错误"), L("Name is required", "请输入名称。"), parent=d)
                    return
                try:
                    order_int = int(order)
                except ValueError:
                    messagebox.showerror(L("Error", "错误"), L("Sort Order must be a whole number (e.g. 1, 10, 100)", "排序必须是整数（例如 1、10、100）。"), parent=d)
                    return

                try:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("UPDATE product_categories SET name=?, icon=?, color=?, sort_order=?, exclude_export=?, show_list_price=? WHERE id=?",
                                  (name, icon, color, order_int, exclude, list_price_var.get(), cid))
                        if name != row[1]:
                            c.execute("UPDATE items SET category=? WHERE category=?", (name, row[1]))
                    messagebox.showinfo(L("Success", "成功"), L("Category updated!", "类别已更新。"), parent=d)
                    d.destroy()
                    load_all_categories()
                    self.load_categories()
                    self.load_items()
                except Exception as e:
                    messagebox.showerror(L("Error", "错误"), str(e), parent=d)

            RoundedButton(d, text="💾 Update", command=update,
                          bg=self.colors['success'], fg="white",
                          font=("Arial", 12, "bold"), padx=20, pady=8, radius=8
                          ).pack(pady=14)

        def toggle_active():
            sel = ctree.selection()
            if not sel:
                messagebox.showwarning(L("No Selection", "未选择"), L("Select a category", "请选择一个类别。"), parent=win)
                return
            row = ctree.item(sel[0])['values']
            cid = row[0]
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("UPDATE product_categories SET active = CASE WHEN active=1 THEN 0 ELSE 1 END WHERE id=?", (cid,))
            load_all_categories()
            self.load_categories()
            self.load_items()

        self._build_query_toolbar(
            query_host, product_category_search_var, load_all_categories,
            on_refresh=load_all_categories,
            on_export=lambda: self._export_treeview_csv(
                ctree, win, "product_categories.csv"
            ),
            result_var=product_category_result_var,
            search_label=L("Search categories:", "搜索类别：")
        )

        btn_f = tk.Frame(win, bg=self.colors['bg'])
        btn_f.pack(fill="x", padx=20, pady=10)
        for text, cmd, color in [
            (L("➕ Add Category", "➕ 新增类别"), add_category_dialog, self.colors['success']),
            (L("✏️ Edit Category", "✏️ 编辑类别"), edit_category_dialog, self.colors['warning']),
            (L("🔄 Toggle Active", "🔄 启用/停用"), toggle_active, self.colors['card']),
        ]:
            RoundedButton(btn_f, text=text, command=cmd, font=("Arial", 11, "bold"),
                          bg=color, fg="white", padx=14, pady=8, radius=8
                          ).pack(side="left", padx=4)

        load_all_categories()

    def setup_keyboard_shortcuts(self):
        self.root.bind('<F1>', lambda e: self.quick_checkout())
        self.root.bind('<F2>', lambda e: self.clear_order())
        self.root.bind('<F3>', lambda e: self.focus_search())
        self.root.bind('<F4>', lambda e: self.show_held_orders())
        self.root.bind('<F5>', lambda e: self.load_items())
        self.root.bind('<F6>', lambda e: self.hold_current_order())
        self.root.bind('<F7>', lambda e: self.show_action_hub())
        self.root.bind('<Control-z>', lambda e: self.undo_last_remove())

    def _on_canvas_resize(self, event):
        self.items_canvas.itemconfig(self._btn_window, width=event.width)

    def build_ui(self):
        C = self.colors

        top_bar = tk.Frame(self.root, bg=C['header'], height=68)
        top_bar.pack(fill="x")
        top_bar.pack_propagate(False)

        brand_frame = tk.Frame(top_bar, bg=C['header'])
        brand_frame.pack(side="left", padx=(18, 0), pady=10)
        self._brand_frame = brand_frame
        tk.Label(brand_frame, text="☕", font=("Arial", 22),
                 bg=C['header'], fg="#F59E0B").pack(side="left", padx=(0, 8))
        brand_text = tk.Frame(brand_frame, bg=C['header'])
        brand_text.pack(side="left")
        tk.Label(brand_text, text="919 Restaurant", font=("Arial", 15, "bold"),
                 bg=C['header'], fg="#FFFFFF").pack(anchor="w")
        self._brand_subtitle_label = tk.Label(brand_text, text=self.tr('brand_subtitle'), font=("Arial", 9),
                              bg=C['header'], fg="#94A3B8")
        self._brand_subtitle_label.pack(anchor="w")

        def stat_card(parent, title, var_name, value_fg="#FFFFFF"):
            card = tk.Frame(parent, bg=C['header2'], padx=12, pady=6)
            card.pack(side="left", padx=6, pady=10)
            tk.Label(card, text=title, font=("Arial", 8),
                     bg=C['header2'], fg="#94A3B8").pack(anchor="w")
            lbl_val = tk.Label(card, text="—", font=("Arial", 13, "bold"),
                               bg=C['header2'], fg=value_fg)
            lbl_val.pack(anchor="w")
            return lbl_val

        stats_wrap = tk.Frame(top_bar, bg=C['header'])
        # Pack stats_wrap to make it visible for everyone
        stats_wrap.pack(side="left", padx=10)
        self.stats_wrap = stats_wrap
        self.today_sales_label = stat_card(stats_wrap, "TODAY'S SALES", "sales", "#4ADE80")
        self.today_trans_label = stat_card(stats_wrap, "TRANSACTIONS",  "trans", "#93C5FD")
        self.today_cash_label  = stat_card(stats_wrap, "💵 CASH",  "cash",  "#FCD34D")
        self.today_qr_label    = stat_card(stats_wrap, "📱 QR",    "qr",    "#6EE7B7")

        self.time_label = tk.Label(top_bar, text="", font=("Arial", 12, "bold"),
                                   bg=C['header'], fg="#CBD5E1")
        self.time_label._skip_auto_localize_text = True
        self.time_label.pack(side="right", padx=24)

        self._admin_btn = RoundedButton(top_bar, text=self.tr('admin'),
                                        command=self.toggle_admin_mode,
                                        bg=C['neutral'], fg="white",
                                        font=("Arial", 10, "bold"),
                                        padx=12, pady=8, radius=8)
        self._admin_btn.pack(side="right", padx=8, pady=10)

        self._lang_btn = RoundedButton(top_bar, text=self.tr('lang_btn'),
                           command=self.toggle_language,
                           bg=C['secondary'], fg="white",
                           font=("Arial", 10, "bold"),
                           padx=10, pady=8, radius=8)
        self._lang_btn.pack(side="right", padx=4, pady=10)

        self.staff_label = tk.Label(top_bar, text="", font=("Arial", 11, "bold"),
                                    bg=C['header'], fg=C['warning'])

        self.body_container = tk.Frame(self.root, bg=C['bg'])
        self.body_container.pack(fill="both", expand=True)

        self.table_home_frame = tk.Frame(self.body_container, bg=C['bg'])

        self.main_container = tk.Frame(self.body_container, bg=C['bg'])
        main_container = self.main_container
        main_container.pack(fill="both", expand=True)

        left_panel = tk.Frame(main_container, bg=C['bg'])
        left_panel.pack(side="left", fill="both", expand=True)

        search_wrap = tk.Frame(left_panel, bg=C['bg'])
        search_wrap.pack(fill="x", padx=14, pady=(12, 0))
        search_card = tk.Frame(search_wrap, bg=C['panel'],
                               highlightbackground=C['border2'], highlightthickness=1)
        search_card.pack(fill="x")
        tk.Label(search_card, text="🔍", font=("Arial", 14),
                 bg=C['panel'], fg=C['text_muted']).pack(side="left", padx=(10, 4), pady=8)
        self.search_entry = tk.Entry(search_card, textvariable=self.search_var,
                                     font=("Arial", 13), bg=C['panel'],
                                     fg=C['text'], relief="flat",
                                     insertbackground=C['text'])
        self.search_entry.pack(side="left", fill="x", expand=True, pady=8)
        tk.Label(search_card, text="F3", font=("Arial", 9),
                 bg=C['border'], fg=C['text_muted'],
                 padx=4, pady=2).pack(side="right", padx=8, pady=8)

        cat_bar = tk.Frame(left_panel, bg=C['bg'])
        cat_bar.pack(fill="x", padx=14, pady=8)
        self.category_var = tk.StringVar(value="ALL")

        self._make_cat_btn(cat_bar, "⊞  ALL", "ALL",
                     C['neutral'], lambda: self.filter_category("ALL"))

        # New frame to hold dynamic category buttons for easier refreshing
        self._cat_btn_frame = tk.Frame(cat_bar, bg=C['bg'])
        self._cat_btn_frame.pack(side="left")
        self._refresh_category_pills()

        items_container = tk.Frame(left_panel, bg=C['bg'])
        items_container.pack(fill="both", expand=True, padx=14, pady=(0, 6))
        self.items_canvas = tk.Canvas(items_container, bg=C['bg'], highlightthickness=0)
        scrollbar = tk.Scrollbar(items_container, orient="vertical",
                                 command=self.items_canvas.yview)
        self.buttons_frame = tk.Frame(self.items_canvas, bg=C['bg'])
        self.buttons_frame.bind(
            "<Configure>",
            lambda e: self.items_canvas.configure(
                scrollregion=self.items_canvas.bbox("all")))
        self._btn_window = self.items_canvas.create_window(
            (0, 0), window=self.buttons_frame, anchor="nw")
        self.items_canvas.configure(yscrollcommand=scrollbar.set)
        self.items_canvas.bind("<Configure>", self._on_canvas_resize)
        self.items_canvas.bind("<MouseWheel>", self._on_mousewheel)
        
        # New: Drag-to-scroll (Swipe) bindings
        self.items_canvas.bind("<Button-1>", self._on_scroll_start, add="+")
        self.items_canvas.bind("<B1-Motion>", lambda e: self._on_scroll_drag(e, self.items_canvas), add="+")
        self.buttons_frame.bind("<Button-1>", self._on_scroll_start, add="+")
        self.buttons_frame.bind("<B1-Motion>", lambda e: self._on_scroll_drag(e, self.items_canvas), add="+")
        self.buttons_frame.bind("<MouseWheel>", self._on_mousewheel)

        self.items_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.NUM_COLS = 7
        for i in range(self.NUM_COLS):
            self.buttons_frame.grid_columnconfigure(i, weight=1, uniform="col")

        self.float_table_home_btn = RoundedButton(
            left_panel,
            text=self.tr('main_tables'),
            command=self.show_table_home,
            bg="#B803FF", fg="white",
            font=("Arial", 11, "bold"),
            padx=10, pady=10, radius=10,
            width=100, height=100
        )
        self.float_table_home_btn.place(relx=0.985, rely=0.01, anchor="ne")

        # ── Always-visible cashier strip (no admin needed) ──────────────
        self.cashier_strip = tk.Frame(left_panel, bg=C['panel'], height=52)
        self.cashier_strip.pack(fill="x", side="bottom")
        self.cashier_strip.pack_propagate(False)
        self._render_cashier_strip()

        tk.Frame(left_panel, bg=C['border2'], height=1).pack(fill="x", side="bottom")
        nav_strip = tk.Frame(left_panel, bg=C['panel'], height=56)
        nav_strip.pack_propagate(False)
        # ── CHANGE 1: nav_strip hidden by default, shown only in admin mode ──
        # nav_strip NOT packed here
        self.nav_strip = nav_strip
        nav_items = [
            ("🛠️", "Admin",   self.show_admin_panel),
        ]
        for icon, label, cmd in nav_items:
            nav_btn = tk.Frame(nav_strip, bg=C['panel'], cursor="hand2", width=120, height=42)
            nav_btn.pack(side="left", padx=8, pady=6)
            nav_btn.pack_propagate(False)
            tk.Label(nav_btn, text=icon, font=("Arial", 16),
                     bg=C['panel'], fg=C['header']).pack(side="left", padx=(10, 4))
            tk.Label(nav_btn, text=label, font=("Arial", 10, "bold"),
                     bg=C['panel'], fg=C['text_sub']).pack(side="left", padx=(2, 8))
            _cmd = cmd
            _label = label
            for w in [nav_btn] + nav_btn.winfo_children():
                w.bind("<Button-1>", lambda e, c=_cmd, l=_label: self._run_ui_action(l, c))
            def _enter(e, f=nav_btn):
                f['bg'] = C['neutral_lt']
                for w in f.winfo_children(): w['bg'] = C['neutral_lt']
            def _leave(e, f=nav_btn):
                f['bg'] = C['panel']
                for w in f.winfo_children(): w['bg'] = C['panel']
            nav_btn.bind("<Enter>", _enter)
            nav_btn.bind("<Leave>", _leave)

        def _rebind_nav():
            for child in nav_strip.winfo_children():
                idx = nav_strip.winfo_children().index(child)
                if idx < len(nav_items):
                    _cmd = nav_items[idx][2]
                    _label = nav_items[idx][1]
                    for w in [child] + child.winfo_children():
                        w.bind("<Button-1>", lambda e, c=_cmd, l=_label: self._run_ui_action(l, c))
        self.root.after(50, _rebind_nav)

        right_panel = tk.Frame(main_container, bg=C['panel'], width=480)
        right_panel.pack(side="right", fill="both")
        right_panel.pack_propagate(False)

        order_hdr = tk.Frame(right_panel, bg=C['success'], height=56)
        order_hdr.pack(fill="x")
        order_hdr.pack_propagate(False)
        hdr_inner = tk.Frame(order_hdr, bg=C['success'])
        hdr_inner.pack(expand=True)
        tk.Label(hdr_inner, text="🧾", font=("Arial", 18),
                 bg=C['success'], fg="white").pack(side="left")
        self.current_order_title_label = tk.Label(hdr_inner, text=self.tr('current_order'), font=("Arial", 15, "bold"),
                              bg=C['success'], fg="white")
        self.current_order_title_label.pack(side="left")
        self.order_num_label = tk.Label(right_panel, text="Order  #----  (assign at checkout)",
                        font=("Arial", 10, "bold"),
                        bg=C['panel2'], fg=C['text_sub'], pady=4)
        self.order_table_label = tk.Label(right_panel, text=self.tr('table_label', table='—'),
                          font=("Arial", 9, "bold"),
                          bg=C['panel2'], fg=C['text_muted'], pady=2)
        self.order_table_label.pack(fill="x", padx=0)
        self.order_mode_label = tk.Label(right_panel, text=self.tr('new_order'),
                         font=("Arial", 9, "bold"),
                         bg=C['panel2'], fg=C['text_muted'], pady=2)
        self.order_mode_label.pack(fill="x", padx=0)

        list_container = tk.Frame(right_panel, bg=C['panel'])
        list_container.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        self.order_header_canvas = tk.Canvas(list_container, bg=C['neutral'],
                                              height=28, highlightthickness=0)
        self.order_header_canvas.pack(fill="x")

        def _draw_order_header(event=None):
            hc = self.order_header_canvas
            hc.delete("all")
            w = hc.winfo_width() or 460
            x_name  = 8
            x_qty   = w - 110
            x_price = w - 8
            hc.create_text(x_name,  14, text="ITEM",  anchor="w",
                           font=("Arial", 9, "bold"), fill="white")
            hc.create_text(x_qty,   14, text="QTY",   anchor="center",
                           font=("Arial", 9, "bold"), fill="white")
            hc.create_text(x_price, 14, text="PRICE", anchor="e",
                           font=("Arial", 9, "bold"), fill="white")

        self.order_header_canvas.bind("<Configure>", _draw_order_header)
        self.root.after(30, _draw_order_header)

        list_frame = tk.Frame(list_container, bg=C['panel'])
        list_frame.pack(fill="both", expand=True)
        sb2 = tk.Scrollbar(list_frame)
        sb2.pack(side="right", fill="y")

        self.order_canvas = tk.Canvas(list_frame, bg=C['panel'],
                                      highlightthickness=0, yscrollcommand=sb2.set)
        self.order_canvas.pack(side="left", fill="both", expand=True)
        sb2.config(command=self.order_canvas.yview)

        self.order_list = tk.Listbox(list_frame, selectmode="single")

        self._order_selected = None
        self._order_row_h = 36
        self._order_items_drawn = []

        def _canvas_click(event):
            row = event.y // self._order_row_h
            if row < len(self.order):
                self._order_selected = row
                _draw_order_rows()

        self.order_canvas.bind("<Button-1>", _canvas_click, add="+")
        self.order_canvas.bind("<Button-1>", self._on_scroll_start, add="+")
        self.order_canvas.bind("<B1-Motion>", lambda e: self._on_scroll_drag(e, self.order_canvas), add="+")
        self.order_canvas.bind("<MouseWheel>", self._on_order_mousewheel)

        # Right-click context menu
        def _canvas_right_click(event):
            row = event.y // self._order_row_h
            if row >= len(self.order):
                return
            name, price, qty, cat = self.order[row]
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="➖ Decrease Qty",
                             command=lambda: self.remove_item())
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        self.order_canvas.bind("<Button-3>", _canvas_right_click)

        def _draw_order_rows():
            self.order_canvas.delete("all")
            self._order_items_drawn = []
            w = self.order_canvas.winfo_width() or 460
            rh = self._order_row_h
            x_name  = 8
            x_qty   = w - 110
            x_price = w - 8

            for i, (name, price, qty, cat) in enumerate(self.order):
                y0 = i * rh
                y1 = y0 + rh
                price_dec, qty_int = self._coerce_order_numbers(price, qty)
                display_unit = self._display_unit_price(name, price_dec, cat, self.category_list_prices)
                line_total = display_unit * qty_int

                if i == self._order_selected:
                    self.order_canvas.create_rectangle(0, y0, w, y1,
                        fill=C['success_lt'], outline="")
                    fg  = C['success']
                    fg2 = C['success']
                elif i % 2 == 1:
                    self.order_canvas.create_rectangle(0, y0, w, y1,
                        fill="#F8FAFC", outline="")
                    fg  = C['text']
                    fg2 = C['text_muted']
                else:
                    fg  = C['text']
                    fg2 = C['text_muted']

                # ── Show variant items in two lines so nothing is cut off ──
                if "(" in name and name.endswith(")"):
                    paren_idx    = name.index("(")
                    base_part    = name[:paren_idx].strip()
                    variant_part = name[paren_idx:]          # e.g. "(Hot)"
                    self.order_canvas.create_text(
                        x_name, y0 + rh//2 - 7,
                        text=base_part[:24], anchor="w",
                        font=("Arial", 11, "bold"), fill=fg)
                    self.order_canvas.create_text(
                        x_name + 4, y0 + rh//2 + 8,
                        text=variant_part[:24], anchor="w",
                        font=("Arial", 9), fill=fg2)
                else:
                    self.order_canvas.create_text(
                        x_name, y0 + rh//2,
                        text=name[:26], anchor="w",
                        font=("Arial", 11), fill=fg)

                self.order_canvas.create_text(x_qty, y0 + rh//2,
                    text=f"x{qty_int}", anchor="center",
                    font=("Arial", 11, "bold"), fill=fg)
                show_unit = cat in self.category_list_prices and qty_int > 1 and not name.startswith("[FOC] ")
                self.order_canvas.create_text(x_price, y0 + rh//2 - (7 if show_unit else 0),
                    text=f"RM{float(line_total):.2f}", anchor="e",
                    font=("Arial", 11), fill=fg)
                if show_unit:
                    self.order_canvas.create_text(x_price, y0 + rh//2 + 8,
                        text=self.l(f"RM{display_unit:.2f} each", f"每件 RM{display_unit:.2f}"), anchor="e",
                        font=("Arial", 8), fill=fg2)

                self.order_canvas.create_line(0, y1, w, y1, fill=C['border'], width=1)

            total_h = max(len(self.order) * rh, 1)
            self.order_canvas.configure(scrollregion=(0, 0, w, total_h))
            # Reset view to top if list is small
            if total_h <= (self.order_canvas.winfo_height() or 500):
                self.order_canvas.yview_moveto(0)

        self._draw_order_rows = _draw_order_rows
        self.order_canvas.bind("<Configure>", lambda e: _draw_order_rows())

        action_frame = tk.Frame(right_panel, bg=C['panel'])
        action_frame.pack(fill="x", padx=8, pady=(4, 0))
        for btn_txt, btn_cmd, btn_color in [
            ("➖  Remove", self.remove_item,  C['danger']),
            ("🗑  Clear",  self.clear_order,  C['neutral']),
        ]:
            RoundedButton(action_frame, text=btn_txt, command=btn_cmd,
                          bg=btn_color, fg="white",
                          font=("Arial", 10, "bold"),
                          padx=12, pady=7, radius=8
                          ).pack(side="left", padx=3, pady=4)

        tk.Frame(right_panel, bg=C['border'], height=1).pack(fill="x", padx=8)

        calc_frame = tk.Frame(right_panel, bg=C['panel'])
        calc_frame.pack(fill="x", padx=12, pady=6)

        def calc_row(text, fg_color=None):
            row = tk.Frame(calc_frame, bg=C['panel'])
            row.pack(fill="x", pady=1)
            lbl = tk.Label(row, text=text, font=("Arial", 11),
                           bg=C['panel'], fg=fg_color or C['text_sub'], anchor="e")
            lbl.pack(fill="x")
            return lbl

        self.subtotal_label = calc_row("Subtotal:  RM 0.00")
        self.tax_label      = calc_row("Tax (0%):  —", fg_color=C['text_muted'])

        total_banner = tk.Frame(right_panel, bg=C['success'], height=72)
        total_banner.pack(fill="x", padx=8, pady=6)
        total_banner.pack_propagate(False)
        self.total_label = tk.Label(total_banner, text="TOTAL  RM 0.00",
                                    font=("Arial", 26, "bold"),
                                    bg=C['success'], fg="white")
        self.total_label.pack(expand=True)

        # These labels are dynamic and updated by order math; avoid auto-localizer
        # restoring their initial "...0.00" text snapshots.
        self.subtotal_label._skip_auto_localize_text = True
        self.tax_label._skip_auto_localize_text = True
        self.total_label._skip_auto_localize_text = True

        # Flow: table-based hold / checkout
        pay_indicator = tk.Frame(right_panel, bg=C['panel2'],
                             highlightbackground=C['border'], highlightthickness=1)
        pay_indicator.pack(fill="x", padx=8, pady=(0, 4))
        self.pay_indicator_label = tk.Label(pay_indicator,
                            text=self.tr('table_flow_hint'),
                            font=("Arial", 10, "bold"),
                            bg=C['panel2'], fg=C['text_muted'])
        self.pay_indicator_label.pack(pady=6)

        co_frame = tk.Frame(right_panel, bg=C['panel'])
        co_frame.pack(fill="x", padx=8, pady=(0, 10))
        self.hold_action_btn = RoundedButton(co_frame,
                              text=self.tr('hold_f6'),
                              command=self.hold_current_order,
                              bg=C['success'], fg="white",
                              font=("Arial", 13, "bold"),
                              padx=18, pady=12, radius=12)
        self.hold_action_btn.pack(fill="x")
        self.checkout_action_btn = RoundedButton(co_frame,
                             text=self.tr('checkout'),
                             command=self.checkout_current_order,
                             bg=C['primary'], fg="white",
                             font=("Arial", 16, "bold"),
                             padx=22, pady=20, radius=12)
        self.checkout_action_btn.pack(fill="x", pady=(8, 0))
        self._refresh_order_mode_ui()

        self._build_table_home_ui()
        self.apply_language()
        self.show_table_home()

    def _get_cashier_quick_actions(self):
        """Single source of truth for the always-visible cashier quick bar."""
        actions = [
            ("🧾", self.tr('recent_orders'), self.show_recent_orders),
            ("💰", self.tr('cash'), self.open_cash_center),
            ("🏪", self.tr('rent'), self.open_rent_center),
            ("📦", self.tr('stock'), self.show_cashier_stock_control),
        ]
        if self.admin_mode:
            actions.append(("🛠️", self.tr('admin_short'), self.show_admin_panel))
        return actions

    def _render_cashier_strip(self):
        if not hasattr(self, 'cashier_strip'):
            return
        C = self.colors
        for w in self.cashier_strip.winfo_children():
            w.destroy()

        cashier_strip_items = self._get_cashier_quick_actions()
        for index, (icon, label, cmd) in enumerate(cashier_strip_items):
            csb = tk.Frame(self.cashier_strip, bg=C['panel'], cursor="hand2")
            csb.pack(
                side="left", padx=((10 if index == 0 else 4), 4), pady=4
            )
            tk.Label(csb, text=icon, font=("Arial", 15),
                     bg=C['panel'], fg=C['primary']).pack(side="left")
            tk.Label(csb, text=f" {label}", font=("Arial", 9, "bold"),
                     bg=C['panel'], fg=C['text_sub']).pack(side="left")
            _cmd = cmd
            _label = label
            for w in [csb] + csb.winfo_children():
                w.bind("<Button-1>", lambda e, c=_cmd, l=_label: self._run_ui_action(l, c))

            def _cs_enter(e, f=csb):
                f['bg'] = C['neutral_lt']
                for w in f.winfo_children():
                    w['bg'] = C['neutral_lt']

            def _cs_leave(e, f=csb):
                f['bg'] = C['panel']
                for w in f.winfo_children():
                    w['bg'] = C['panel']

            csb.bind("<Enter>", _cs_enter)
            csb.bind("<Leave>", _cs_leave)

    # ================= ITEMS =================
    def load_items(self):
        for widget in self.buttons_frame.winfo_children():
            widget.destroy()

        category = self.category_var.get()
        search = self.search_var.get().strip().lower()

        with self.db.get_connection() as conn:
            c = conn.cursor()
            if category == "ALL":
                c.execute("""
                    SELECT i.name, i.price, i.category, i.stock, i.low_stock_threshold, i.track_stock 
                    FROM items i
                    LEFT JOIN product_categories pc ON i.category = pc.name
                    WHERE i.active=1 
                    ORDER BY pc.sort_order, i.category, i.name
                """)
            else:
                c.execute("SELECT name, price, category, stock, low_stock_threshold, track_stock FROM items WHERE active=1 AND category=? ORDER BY name", (category,))
            items = c.fetchall()
            c.execute("SELECT DISTINCT i.name FROM item_variants iv JOIN items i ON iv.item_id=i.id")
            items_with_variants = {row[0] for row in c.fetchall()}

        if search:
            items = [item for item in items if search in item[0].lower()]

        if not items:
            tk.Label(self.buttons_frame, text="No items found\n\nAdd items in Item Management",
                     font=("Arial", 14), bg=self.colors['bg'],
                     fg=self.colors['text_muted'], justify="center").grid(row=0, column=0, columnspan=4, pady=100)
            return

        row, col = 0, 0
        canvas_w = self.items_canvas.winfo_width()
        if canvas_w < 10:
            canvas_w = 860
        cell_w = max(80, (canvas_w // self.NUM_COLS) - 16)

        current_cat = None
        for name, price, cat, stock, threshold, track_stock in items:
            if cat != current_cat:
                current_cat = cat
                if col != 0:
                    row += 1
                    col = 0
                sep = tk.Frame(self.buttons_frame, bg="#94A3B8", height=2)
                sep.grid(row=row, column=0, columnspan=self.NUM_COLS,
                         padx=8, pady=(10, 4), sticky="ew")
                # New: Drag support on separator
                sep.bind("<Button-1>", self._on_scroll_start, add="+")
                sep.bind("<B1-Motion>", lambda e: self._on_scroll_drag(e, self.items_canvas), add="+")
                sep.bind("<MouseWheel>", self._on_mousewheel)
                row += 1

            show_stock = (cat in ('Beer', 'Cans')) or (track_stock == 1)
            if not show_stock:
                color      = self.category_colors.get(cat, self.colors['info'])
                stock_text = ""
            elif stock <= 0:
                color      = self.colors['danger']
                stock_text = "OUT OF STOCK"
            elif stock <= threshold:
                color      = self.colors['warning']
                stock_text = f"LOW  {stock}"
            else:
                color      = self.category_colors.get(cat, self.colors['info'])
                stock_text = f"x {stock}"

            icon = self.category_icons.get(cat, '📦')
            has_variants = name in items_with_variants
            btn_lines = [f"{icon}  {name}"]
            list_price = self.category_list_prices.get(cat)
            if list_price is not None:
                btn_lines.append(f"{self.currency} {list_price:.2f}")
            elif not has_variants:
                btn_lines.append(f"{self.currency} {float(price):.2f}")
            if stock_text:
                btn_lines.append(stock_text)
            btn_text = "\n".join(btn_lines)
            btn = RoundedButton(
                self.buttons_frame,
                text=btn_text,
                command=lambda n=name, p=price, c=cat, s=stock, ts=track_stock: self.add_item(n, p, c, s, ts),
                bg=color, fg=self.colors['white'],
                font=("Arial", 11, "bold"),
                padx=10, pady=10, radius=12,
                height=100
            )
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
            # Bind mousewheel to the button itself so scrolling works anywhere
            # Swipe & Wheel support on the button itself
            btn.bind("<MouseWheel>", self._on_mousewheel)
            btn.bind("<Button-1>", self._on_scroll_start, add="+")
            btn.bind("<B1-Motion>", lambda e: self._on_scroll_drag(e, self.items_canvas), add="+")
            col += 1
            if col >= self.NUM_COLS:
                col = 0
                row += 1

        for r in range(row + 1):
            self.buttons_frame.grid_rowconfigure(r, weight=1)

        self.update_today_stats()

    def search_items(self, *args):
        self.load_items()

    def _on_scroll_start(self, event):
        """Record the start position of a mouse drag for scrolling."""
        self._last_scroll_y = event.y_root

    def _on_scroll_drag(self, event, canvas):
        """Scroll the canvas based on mouse drag distance (swipe behavior)."""
        if not hasattr(self, '_last_scroll_y'): return
        
        # Restriction: Only scroll if content exceeds view
        y_range = canvas.yview()
        if y_range[0] <= 0 and y_range[1] >= 1.0:
            return

        # Calculate how much the mouse has moved
        delta_y = self._last_scroll_y - event.y_root
        if abs(delta_y) < 2: return # Threshold for sensitivity
        
        # Scroll the canvas proportionally
        # Move roughly 1 unit per 5 pixels of drag
        scroll_units = delta_y // 5
        if scroll_units != 0:
            canvas.yview_scroll(int(scroll_units), "units")
            self._last_scroll_y = event.y_root

    def _on_mousewheel(self, event):
        """Standard smooth scrolling for items grid."""
        if not hasattr(self, 'items_canvas'): return
        canv = self.items_canvas

        # Restriction: Only scroll if content exceeds view
        y_range = canv.yview()
        if y_range[0] <= 0 and y_range[1] >= 1.0:
            return

        # Standard delta is 120. Normalize to move ~3 units per click.
        units = -1 * (event.delta // 40)
        if units == 0 and event.delta != 0:
            units = -1 if event.delta > 0 else 1
        canv.yview_scroll(units, "units")

    def _on_order_mousewheel(self, event):
        """Standard smooth scrolling for order list."""
        if not hasattr(self, 'order_canvas'): return
        canv = self.order_canvas

        # Restriction: Only scroll if content exceeds view
        y_range = canv.yview()
        if y_range[0] <= 0 and y_range[1] >= 1.0:
            return

        units = -1 * (event.delta // 40)
        if units == 0 and event.delta != 0:
            units = -1 if event.delta > 0 else 1
        canv.yview_scroll(units, "units")

    def _update_cat_pills(self, active_key):
        for key, btn in self._cat_btns.items():
            if key == active_key:
                btn._bg_orig  = self._cat_active_bg.get(key, "#1E293B")
                btn._bg       = btn._bg_orig
                btn._hover_bg = btn._darken(btn._bg_orig, 18)
                btn._fg       = "#FFFFFF"
                btn._fg_orig  = "#FFFFFF"
                btn._render()
            else:
                btn._bg_orig  = "#F8FAFC"
                btn._bg       = "#F8FAFC"
                btn._hover_bg = "#E2E8F0"
                btn._fg       = "#475569"
                btn._fg_orig  = "#475569"
                btn._render()

    def _refresh_category_pills(self):
        if not hasattr(self, '_cat_btn_frame'): return
        for child in self._cat_btn_frame.winfo_children():
            child.destroy()

        C = self.colors
        for cat in self.categories:
            icon = self.category_icons.get(cat, "📦")
            color = self.category_colors.get(cat, C['neutral'])
            _k = cat
            self._make_cat_btn(self._cat_btn_frame, f"{icon} {cat}", _k, color,
                         lambda c=_k: self.filter_category(c))

        self._update_cat_pills(self.category_var.get())

    def filter_category(self, category):
        self.category_var.set(category)
        self._update_cat_pills(category)
        self.load_items()

    def _qty_numpad_popup(self, item_name, item_price, callback, parent=None):
        """Show variant→qty numpad popup. callback(qty) when confirmed."""
        C   = self.colors
        ref = parent or self.root
        win = tk.Toplevel(ref)
        win.title("Quantity")
        win.configure(bg="#0F172A")
        win.transient(ref)
        win.grab_set()
        win.resizable(False, False)

        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)

        card = tk.Frame(win, bg="#1E293B", highlightbackground="#334155", highlightthickness=1)
        card.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        card.grid_rowconfigure(0, weight=0)
        card.grid_rowconfigure(1, weight=0)
        card.grid_rowconfigure(2, weight=1)
        card.grid_rowconfigure(3, weight=0)
        card.grid_columnconfigure(0, weight=1)

        hdr = tk.Frame(card, bg="#1E293B")
        hdr.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 4))
        tk.Label(hdr, text="QUANTITY", font=("Arial", 9, "bold"),
                 bg="#1E293B", fg="#94A3B8").pack()
        tk.Label(hdr, text=item_name, font=("Arial", 14, "bold"),
                 bg="#1E293B", fg="white", wraplength=270,
                 justify="center").pack(pady=(4, 0))
        tk.Label(hdr, text=f"{self.currency} {float(item_price):.2f}",
                 font=("Arial", 12, "bold"), bg="#1E293B", fg="#4ADE80").pack(pady=(2, 0))

        # ── Display ──────────────────────────────────────────────────────
        qty_var = tk.StringVar(value="1")
        disp = tk.Label(card, textvariable=qty_var,
                        font=("Arial", 38, "bold"), bg="#0B1324", fg="white",
                        anchor="e", relief="flat")
        disp.grid(row=1, column=0, sticky="ew", padx=14, pady=(8, 6), ipady=10)

        def press(val):
            cur = qty_var.get()
            if val == "⌫":
                qty_var.set(cur[:-1] or "")
            elif val == "C":
                qty_var.set("")
            else:
                if cur == "1" and val.isdigit():
                    raw = str(val)
                else:
                    raw = cur + str(val)
                raw = raw.lstrip("0") or raw
                if len(raw) <= 3:
                    qty_var.set(raw)

        def confirm():
            try:
                q = int(qty_var.get())
                if q <= 0: raise ValueError
            except ValueError:
                q = 1
            win.destroy()
            callback(q)

        # ── Numpad grid ───────────────────────────────────────────────────
        pad = tk.Frame(card, bg="#1E293B")
        pad.grid(row=2, column=0, padx=14, pady=4)
        for r, row_vals in enumerate([["7","8","9"],["4","5","6"],["1","2","3"]]):
            for c2, v in enumerate(row_vals):
                tk.Button(pad, text=v, font=("Arial",18,"bold"),
                          width=4, relief="flat", bd=0,
                          bg="#334155", fg="white", activebackground="#475569", activeforeground="white",
                          command=lambda v=v: press(v)
                          ).grid(row=r, column=c2, padx=3, pady=3, ipadx=8, ipady=11, sticky="nsew")
        for c2, (v, bg) in enumerate([("0","#334155"),("⌫","#475569"),("C","#991B1B")]):
            tk.Button(pad, text=v, font=("Arial",18,"bold"),
                      width=4, relief="flat", bd=0,
                      bg=bg, fg="white", activebackground="#475569", activeforeground="white",
                      command=lambda v=v: press(v)
                      ).grid(row=3, column=c2, padx=3, pady=3, ipadx=8, ipady=11, sticky="nsew")

        # ── Confirm button — always pinned at bottom ──────────────────────
        tk.Button(card, text="✅  Confirm", font=("Arial",14,"bold"),
                  bg="#16A34A", fg="white", relief="flat",
                  activebackground="#15803D", activeforeground="white", bd=0,
                  command=confirm
                  ).grid(row=3, column=0, sticky="ew", padx=14, pady=(8, 14), ipady=12)

        win.bind("<Return>",    lambda e: confirm())
        win.bind("<Escape>",    lambda e: win.destroy())
        win.bind("<BackSpace>", lambda e: press("⌫"))
        for k in "0123456789":
            win.bind(k, lambda e, v=k: press(v))

        # Size after widgets are created
        win.update_idletasks()
        w = win.winfo_reqwidth() + 20
        h = win.winfo_reqheight() + 20
        self._center_window(win, w, h, ref)
        win.focus_set()

    def add_item(self, name, price, category, stock, track_stock=0):
        if track_stock != 0 and stock <= 0:
            messagebox.showwarning("Out of Stock", f"{name} is currently out of stock!")
            return

        eff_stock = 999 if track_stock == 0 else stock
        base_name = name.split(" (")[0] if " (" in name else name
        
        # Check for Sets
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("""SELECT iqs.quantity, iqs.set_price 
                         FROM item_quantity_sets iqs
                         JOIN items i ON iqs.item_id = i.id
                         WHERE i.name = ? ORDER BY iqs.quantity DESC""", (base_name,))
            sets = c.fetchall()

        # Check for Variants
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("""SELECT iv.id, iv.variant_name, iv.price_adjustment
                         FROM item_variants iv
                         JOIN items i ON iv.item_id = i.id
                         WHERE i.name = ? ORDER BY iv.id""", (base_name,))
            variants = c.fetchall()

        def start_set_selection(final_name, final_price):
            if sets:
                def after_set(display_name, set_price):
                    def after_qty(qty):
                        self._do_add_item(display_name, set_price, category,
                                          eff_stock, track_stock=track_stock, qty=qty)
                    self._qty_numpad_popup(display_name, set_price, after_qty)
                self._show_set_selection_popup(final_name, final_price, category, eff_stock, sets,
                                               track_stock=track_stock, on_choose=after_set)
            else:
                def after_qty(qty):
                    self._do_add_item(final_name, final_price, category,
                                      eff_stock, track_stock=track_stock, qty=qty)
                self._qty_numpad_popup(final_name, final_price, after_qty)

        if variants:
            def after_variant(v_display_name, v_final_price):
                start_set_selection(v_display_name, v_final_price)
            self._show_variant_popup(name, price, category, eff_stock, variants,
                                     track_stock=track_stock, on_choose=after_variant)
        else:
            start_set_selection(name, price)

    def _show_set_selection_popup(self, name, base_price, category, stock, sets,
                                   track_stock=1, on_choose=None):
        popup = tk.Toplevel(self.root)
        popup.title(f"Select Set — {name}")
        popup.configure(bg=self.colors['neutral'])
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        tk.Label(popup, text=f"📦 {name} - Buy Set?", font=("Arial", 16, "bold"),
                 bg=self.colors['neutral'], fg=self.colors['white']).pack(pady=(18, 6), padx=30)
        tk.Label(popup, text="Select a set option:", font=("Arial", 11),
                 bg=self.colors['neutral'], fg=self.colors['light']).pack(pady=(0, 12))

        btn_frame = tk.Frame(popup, bg=self.colors['neutral'])
        btn_frame.pack(padx=24, pady=4)

        def choose(display_name, set_price):
            popup.destroy()
            if on_choose:
                on_choose(display_name, set_price)

        # "Single" Option
        cat_color = self.category_colors.get(category, self.colors['info'])
        RoundedButton(btn_frame, text=f"Single\n{self.currency} {float(base_price):.2f}",
                      command=lambda: choose(name, base_price),
                      bg=cat_color, fg="white",
                      font=("Arial", 13, "bold"), padx=20, pady=14, radius=10
                      ).pack(side="left", padx=6, pady=6)

        # Set Options
        for s in sets:
            s_qty, s_price = s['quantity'], s['set_price']
            btn_text = f"Set ({s_qty})\n{self.currency} {float(s_price):.2f}"
            display_name = f"{name} Set({s_qty})"
            RoundedButton(btn_frame, text=btn_text,
                          command=lambda dn=display_name, sp=s_price: choose(dn, sp),
                          bg=self.colors['primary'], fg="white",
                          font=("Arial", 13, "bold"), padx=20, pady=14, radius=10
                          ).pack(side="left", padx=6, pady=6)

        RoundedButton(popup, text="✖ Cancel", command=popup.destroy,
                      bg=self.colors['danger'], fg="white",
                      font=("Arial", 10), padx=16, pady=6, radius=8
                      ).pack(pady=(6, 18))

        popup.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - popup.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - popup.winfo_height()) // 2
        popup.geometry(f"+{x}+{y}")

    def _show_variant_popup(self, name, base_price, category, stock, variants,
                             track_stock=1, qty=1, on_choose=None):
        """on_choose(display_name, final_price) if provided, else legacy direct-add."""
        popup = tk.Toplevel(self.root)
        popup.title(f"Select option — {name}")
        popup.configure(bg=self.colors['neutral'])
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        tk.Label(popup, text=f"☕ {name}", font=("Arial", 16, "bold"),
                 bg=self.colors['neutral'], fg=self.colors['white']).pack(pady=(18, 6), padx=30)
        tk.Label(popup, text="Select an option:", font=("Arial", 11),
                 bg=self.colors['neutral'], fg=self.colors['light']).pack(pady=(0, 12))

        btn_frame = tk.Frame(popup, bg=self.colors['neutral'])
        btn_frame.pack(padx=24, pady=4)

        def choose(v_name, v_price):
            popup.destroy()
            display_name = f"{name} ({v_name})"
            if on_choose:
                on_choose(display_name, v_price)
            else:
                self._do_add_item(display_name, v_price, category, stock,
                                  track_stock=track_stock, qty=qty)

        for v in variants:
            final_price = float(base_price) + float(v['price_adjustment'])
            btn_text = f"{v['variant_name']}\n{self.currency} {final_price:.2f}"
            cat_color = self.category_colors.get(category, self.colors['info'])
            RoundedButton(btn_frame, text=btn_text,
                          command=lambda vn=v['variant_name'], vp=final_price: choose(vn, vp),
                          bg=cat_color, fg="white",
                          font=("Arial", 13, "bold"), padx=20, pady=14, radius=10
                          ).pack(side="left", padx=6, pady=6)

        RoundedButton(popup, text="✖ Cancel", command=popup.destroy,
                      bg=self.colors['danger'], fg="white",
                      font=("Arial", 10), padx=16, pady=6, radius=8
                      ).pack(pady=(6, 18))

        popup.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - popup.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - popup.winfo_height()) // 2
        popup.geometry(f"+{x}+{y}")

    def _get_item_multiplier(self, name):
        if " Set(" in name and name.endswith(")"):
            try:
                idx = name.rfind(" Set(")
                s_val = name[idx+5:-1]
                return int(s_val)
            except:
                pass
        return 1

    def _do_add_item(self, name, price, category, stock, track_stock=1, qty=1):
        if not self.current_table_code:
            messagebox.showwarning("Select Table", "Please select table first.")
            self.open_table_selector()
            return

        added_multiplier = self._get_item_multiplier(name)
        total_being_added = qty * added_multiplier

        # Calculate total currently in order for the SAME base item/variant
        # We need to resolve base name for accurate stock tracking check
        base_name_to_track = name
        if " Set(" in name:
            base_name_to_track = name.split(" Set(")[0].strip()

        current_in_order_units = 0
        if track_stock != 0:
            for n, p, q, c in self.order:
                m = self._get_item_multiplier(n)
                # If it's a set, the base name before " Set(" should match
                # if it's a single, the name should match
                check_base = n.split(" Set(")[0].strip() if " Set(" in n else n
                if check_base == base_name_to_track:
                    current_in_order_units += (q * m)

        if track_stock != 0 and (current_in_order_units + total_being_added) > stock:
            messagebox.showwarning("Insufficient Stock", 
                                   f"Insufficient stock for {base_name_to_track}.\n"
                                   f"Available: {stock}\n"
                                   f"In Order: {current_in_order_units}\n"
                                   f"Adding: {total_being_added}")
            return

        price = Decimal(str(price))
        for i, (n, p, q, c) in enumerate(self.order):
            if n == name:
                self.order[i] = (n, p, q + qty, c)
                self.update_order_display()
                return
        self.order.append((name, price, qty, category))
        self.update_order_display()

    def _coerce_order_numbers(self, price, qty):
        try:
            price_dec = Decimal(str(price))
        except Exception:
            price_dec = Decimal('0.00')
        try:
            qty_int = int(qty)
        except Exception:
            qty_int = 0
        return price_dec, max(0, qty_int)


    def _calc_order_totals(self, order_items=None):
        items = self.order if order_items is None else order_items
        subtotal = Decimal('0.00')
        for name, price, qty, _ in items:
            price_dec, qty_int = self._coerce_order_numbers(price, qty)
            subtotal += price_dec * qty_int
        tax = (subtotal * self.tax_rate).quantize(Decimal('0.01'), ROUND_HALF_UP)
        total = subtotal + tax
        return subtotal, tax, total

    def update_order_display(self):
        if hasattr(self, '_draw_order_rows'):
            self._draw_order_rows()
        subtotal, tax, self.total = self._calc_order_totals(self.order)
        self._render_order_total_labels(subtotal, tax, self.total)
        self._refresh_order_mode_ui()

    def _render_order_total_labels(self, subtotal, tax, total):
        tax_pct = float(self.tax_rate * 100)
        adjustment = self._display_price_adjustment(self.order, self.category_list_prices)
        displayed_subtotal = subtotal - adjustment
        displayed_total = total - adjustment
        if self.lang == "zh":
            self.subtotal_label.config(text=f"小计：{self.currency} {displayed_subtotal:.2f}")
            self.tax_label.config(text=f"税 ({tax_pct:.0f}%)：{self.currency} {float(tax):.2f}")
            self.total_label.config(text=f"总计  {self.currency} {displayed_total:.2f}")
        else:
            self.subtotal_label.config(text=f"Subtotal: {self.currency} {displayed_subtotal:.2f}")
            self.tax_label.config(text=f"Tax ({tax_pct:.0f}%): {self.currency} {float(tax):.2f}")
            self.total_label.config(text=f"TOTAL  {self.currency} {displayed_total:.2f}")

    def _refresh_order_mode_ui(self):
        if not hasattr(self, 'hold_action_btn'):
            return

        active_hold_id = int(self._active_held_order_id) if self._active_held_order_id else None

        if active_hold_id:
            if hasattr(self, 'order_mode_label'):
                self.order_mode_label.config(text=self.tr('editing_hold', id=active_hold_id), fg=self.colors['warning'])
            self.hold_action_btn.configure(text=self.tr('update_hold_f6'))
        else:
            if hasattr(self, 'order_mode_label'):
                self.order_mode_label.config(text=self.tr('new_order'), fg=self.colors['text_muted'])
            self.hold_action_btn.configure(text=self.tr('hold_f6'))
        if hasattr(self, 'checkout_action_btn') and not self.checkout_action_btn.winfo_manager():
            self.checkout_action_btn.pack(fill="x", pady=(8, 0))
        if hasattr(self, 'checkout_action_btn'):
            self.checkout_action_btn.configure(text=self.tr('checkout'))

        if hasattr(self, 'order_table_label'):
            tcode = self.current_table_code or "—"
            self.order_table_label.config(text=self.tr('table_label', table=tcode))

    def _get_active_tables(self):
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("""SELECT table_code FROM dining_tables
                         WHERE active=1
                         ORDER BY sort_order ASC, table_code ASC""")
            return [str(r['table_code']) for r in c.fetchall()]

    def select_table(self, table_code):
        new_table = str(table_code)
        if self.current_table_code == new_table:
            if self._force_table_reload_once:
                return self._load_active_hold_for_current_table()
            # Same table can be tapped again from home screen; reload held orders
            # when no active cart is currently loaded.
            if (not self.order) and (not self._active_held_order_id):
                return self._load_active_hold_for_current_table()
            return True

        prev_table = self.current_table_code
        prev_order = list(self.order)
        prev_hold_id = self._active_held_order_id

        had_prev_order = bool(prev_table and prev_order)
        self.order = []
        self._active_held_order_id = None

        self.current_table_code = new_table
        if not self._load_active_hold_for_current_table():
            self.current_table_code = prev_table
            self.order = prev_order
            self._active_held_order_id = prev_hold_id
            self.update_order_display()
            self._refresh_order_mode_ui()
            self.refresh_table_cards()
            self.refresh_table_home_cards()
            return False

        if had_prev_order:
            if not self._save_order_snapshot_as_hold(prev_table, prev_order, prev_hold_id):
                self.current_table_code = prev_table
                self.order = prev_order
                self._active_held_order_id = prev_hold_id
                self.update_order_display()
                self._refresh_order_mode_ui()
                self.refresh_table_cards()
                self.refresh_table_home_cards()
                return False

        self._refresh_order_mode_ui()
        self.refresh_table_cards()
        self.refresh_table_home_cards()
        return True

    def _save_order_snapshot_as_hold(self, table_code, order_items, hold_id=None):
        if not table_code or not order_items:
            return True

        now = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        subtotal, tax, total = self._calc_order_totals(order_items)

        items_json = json.dumps([(n, float(p), q, c) for n, p, q, c in order_items])
        cashier_name = (self.current_staff or {}).get('name', 'Unknown')

        try:
            with self.db.get_connection() as conn:
                c = conn.cursor()
                if hold_id:
                    c.execute(
                        """UPDATE held_orders
                           SET updated_at=?, cashier=?, items=?, subtotal=?, tax=?, total=?, status='HELD', table_code=?
                           WHERE id=? AND status='HELD' AND table_code=?""",
                        (now, cashier_name, items_json, float(subtotal), float(tax), float(total), table_code, int(hold_id), table_code)
                    )
                    if c.rowcount == 0:
                        raise ValueError("This unpaid order has been paid, cancelled or moved. Reload the table.")
                else:
                    c.execute(
                        """INSERT INTO held_orders
                           (created_at, updated_at, cashier, items, subtotal, tax, total, status, table_code)
                           VALUES (?,?,?,?,?,?,?,'HELD',?)""",
                        (now, now, cashier_name, items_json, float(subtotal), float(tax), float(total), table_code)
                    )
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save previous table order: {str(e)}", parent=self.root)
            return False

    def _load_active_hold_for_current_table(self):
        table_code = self.current_table_code
        if not table_code:
            self.order = []
            self._active_held_order_id = None
            self.update_order_display()
            return True

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute(
                """SELECT id, items, updated_at, total
                   FROM held_orders
                   WHERE status='HELD' AND table_code=?
                   ORDER BY id DESC
                """,
                (table_code,)
            )
            rows = c.fetchall()

        if not rows:
            self.order = []
            self._active_held_order_id = None
            self.update_order_display()
            return True

        popup = tk.Toplevel(self.root)
        popup.title(self.tr('select_order_title', table=table_code))
        popup.configure(bg=self.colors['neutral'])
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        pw, ph = 500, 460
        sx = self.root.winfo_x() + (self.root.winfo_width() - pw) // 2
        sy = self.root.winfo_y() + (self.root.winfo_height() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        tk.Label(popup, text=self.tr('table_has_orders', table=table_code, count=len(rows)), font=("Arial", 14, "bold"),
                 bg=self.colors['neutral'], fg="white").pack(pady=(14, 8))
        tk.Label(popup, text=self.tr('select_order_hint'), font=("Arial", 10),
                 bg=self.colors['neutral'], fg=self.colors['light']).pack(pady=(0, 8))

        list_wrap = tk.Frame(popup, bg=self.colors['neutral'])
        list_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        canvas = tk.Canvas(list_wrap, bg=self.colors['neutral'], highlightthickness=0)
        sb = tk.Scrollbar(list_wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=self.colors['neutral'])
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cwin = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cwin, width=e.width))
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        selected = {"ok": False}

        def open_hold(hid):
            selected["ok"] = self._load_held_order_into_current(int(hid), parent=popup)
            popup.destroy()

        for row in rows:
            hid = int(row['id'])
            updated = row['updated_at'] or ""
            total = float(row['total'] or 0)
            txt = f"#{hid}   {updated[:16]}   {self.currency}{total:.2f}"
            RoundedButton(inner, text=txt,
                          command=lambda x=hid: open_hold(x),
                          bg=self.colors['primary'], fg="white",
                          font=("Arial", 10, "bold"),
                          padx=10, pady=8, radius=8).pack(fill="x", pady=4)

        btns = tk.Frame(popup, bg=self.colors['neutral'])
        btns.pack(fill="x", padx=14, pady=(0, 12))

        def new_order():
            self.order = []
            self._active_held_order_id = None
            self.update_order_display()
            selected["ok"] = True
            popup.destroy()

        RoundedButton(btns, text=self.tr('new_order_btn'),
                      command=new_order,
                      bg=self.colors['success'], fg="white",
                      font=("Arial", 11, "bold"), padx=12, pady=8, radius=8).pack(side="left", padx=4)
        RoundedButton(btns, text=self.tr('cancel'),
                      command=popup.destroy,
                      bg=self.colors['warning'], fg="white",
                      font=("Arial", 11, "bold"), padx=12, pady=8, radius=8).pack(side="right", padx=4)

        self.root.wait_window(popup)
        return bool(selected["ok"])

    def open_table_selector(self):
        tables = self._get_active_tables()
        if not tables:
            messagebox.showwarning(
                self.l("No Tables", "没有桌位"),
                self.l("No active tables found. Please add table in maintenance.", "没有可用桌位，请先在桌位维护中添加。"),
                parent=self.root
            )
            return

        popup = tk.Toplevel(self.root)
        popup.title(self.l("Select Table", "选择桌位"))
        popup.configure(bg=self.colors['neutral'])
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        pw, ph = 920, 340
        sx = self.root.winfo_x() + (self.root.winfo_width() - pw) // 2
        sy = self.root.winfo_y() + (self.root.winfo_height() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        tk.Label(popup, text=self.l("🪑 Select Table", "🪑 选择桌位"), font=("Arial", 18, "bold"),
                 bg=self.colors['neutral'], fg="white").pack(pady=(14, 10))

        grid = tk.Frame(popup, bg=self.colors['neutral'])
        grid.pack(fill="both", expand=True, padx=14, pady=8)

        cols = 10
        for idx, code in enumerate(tables):
            r = idx // cols
            cidx = idx % cols
            btn = RoundedButton(
                grid,
                text=code,
                command=lambda t=code: (popup.destroy(), self.select_table(t)),
                bg=self.colors['success'], fg="white",
                font=("Arial", 11, "bold"),
                padx=12, pady=8, radius=8,
                width=78, height=44
            )
            btn.grid(row=r, column=cidx, padx=4, pady=4, sticky="nsew")

        for i in range(cols):
            grid.grid_columnconfigure(i, weight=1)

        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def _build_table_home_ui(self):
        C = self.colors
        for w in self.table_home_frame.winfo_children():
            w.destroy()

        title_card = tk.Frame(self.table_home_frame, bg=C['panel'],
                              highlightbackground=C['border2'], highlightthickness=1)
        title_card.pack(fill="x", padx=14, pady=(14, 8))
        self._table_home_title_label = tk.Label(title_card, text=self.tr('table_layout'), font=("Arial", 18, "bold"),
                            bg=C['panel'], fg=C['header'])
        self._table_home_title_label.pack(padx=14, pady=(12, 2), anchor="w")
        self._table_home_hint_label = tk.Label(title_card, text=self.tr('table_layout_hint'), font=("Arial", 10),
                               bg=C['panel'], fg=C['text_muted'])
        self._table_home_hint_label.pack(padx=14, pady=(0, 10), anchor="w")

        self.table_home_cards_frame = tk.Frame(self.table_home_frame, bg=C['bg'])
        self.table_home_cards_frame.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        self.table_home_cards_frame.bind("<Configure>", lambda e: self.refresh_table_home_cards())

        self.table_home_strip = tk.Frame(self.table_home_frame, bg=C['panel'], height=52)
        self.table_home_strip.pack(fill="x", side="bottom")
        self.table_home_strip.pack_propagate(False)
        self._render_table_home_strip()

    def _render_table_home_strip(self):
        if not hasattr(self, 'table_home_strip'):
            return
        C = self.colors
        for w in self.table_home_strip.winfo_children():
            w.destroy()

        strip_items = self._get_cashier_quick_actions()

        for index, (icon, label, cmd) in enumerate(strip_items):
            csb = tk.Frame(self.table_home_strip, bg=C['panel'], cursor="hand2")
            csb.pack(
                side="left", padx=((10 if index == 0 else 4), 4), pady=4
            )
            tk.Label(csb, text=icon, font=("Arial", 15),
                     bg=C['panel'], fg=C['primary']).pack(side="left")
            tk.Label(csb, text=f" {label}", font=("Arial", 9, "bold"),
                     bg=C['panel'], fg=C['text_sub']).pack(side="left")
            _cmd = cmd
            _label = label
            for w in [csb] + csb.winfo_children():
                w.bind("<Button-1>", lambda e, c=_cmd, l=_label: self._run_ui_action(l, c))

    def show_table_home(self):
        # Persist current table cart before leaving so re-entry is always loaded from DB.
        if self.current_table_code and self.order:
            saved = self._save_order_snapshot_as_hold(
                self.current_table_code,
                list(self.order),
                self._active_held_order_id
            )
            if saved:
                self.order = []
                self._active_held_order_id = None
                self.update_order_display()
            else:
                return False

        if hasattr(self, 'main_container'):
            self.main_container.pack_forget()
        if hasattr(self, 'table_home_frame'):
            self.table_home_frame.pack(fill="both", expand=True)
            self.refresh_table_home_cards()
        self.refresh_recent_orders_bar()
        self.refresh_table_cards()

    def _enter_table_from_home(self, table_code):
        prev = self.current_table_code
        self._force_table_reload_once = True
        ok = self.select_table(table_code)
        self._force_table_reload_once = False
        if ok and self.current_table_code == str(table_code):
            if hasattr(self, 'table_home_frame'):
                self.table_home_frame.pack_forget()
            if hasattr(self, 'main_container'):
                self.main_container.pack(fill="both", expand=True)
        else:
            self.current_table_code = prev

    def refresh_table_cards(self):
        if not hasattr(self, 'table_cards_frame'):
            return

        for w in self.table_cards_frame.winfo_children():
            w.destroy()

        tables = self._get_active_tables()
        if not tables:
            tk.Label(self.table_cards_frame, text="No active tables",
                     font=("Arial", 10, "bold"),
                     bg=self.colors['panel'], fg=self.colors['text_muted']).pack(pady=8)
            self.current_table_code = None
            self._refresh_order_mode_ui()
            return

        if self.current_table_code and self.current_table_code not in tables:
            self.current_table_code = None

        cols = 10
        for idx, code in enumerate(tables):
            r = idx // cols
            cidx = idx % cols
            active = (code == self.current_table_code)
            bg = self.colors['success'] if active else self.colors['neutral']
            fg = "white"
            btn = RoundedButton(self.table_cards_frame,
                                text=code,
                                command=lambda t=code: self.select_table(t),
                                bg=bg, fg=fg,
                                font=("Arial", 10, "bold"),
                                padx=10, pady=6, radius=8,
                                width=82, height=42)
            btn.grid(row=r, column=cidx, padx=4, pady=4, sticky="nsew")

        for i in range(cols):
            self.table_cards_frame.grid_columnconfigure(i, weight=1)

        self._refresh_order_mode_ui()

    def _get_table_hold_map(self):
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("""
                SELECT table_code, COUNT(*) AS cnt, SUM(total) AS total
                FROM held_orders
                WHERE status='HELD' AND table_code IS NOT NULL AND TRIM(table_code) != ''
                GROUP BY table_code
            """)
            return {str(r['table_code']): (int(r['cnt'] or 0), float(r['total'] or 0)) for r in c.fetchall()}

    def _get_active_unpaid_rows(self, limit=None):
        """Return each independent unpaid bill, including bills at the same table."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            base_sql = """
                SELECT h.id, h.updated_at, h.total, h.items, h.cashier, h.table_code
                FROM held_orders h
                WHERE h.status='HELD'
                ORDER BY h.id DESC
            """
            if limit is not None:
                c.execute(base_sql + " LIMIT ?", (int(limit),))
            else:
                c.execute(base_sql)
            return c.fetchall()

    def _count_active_unpaid_orders(self):
        rows = self._get_active_unpaid_rows(limit=None)
        return len(rows)

    def refresh_table_home_cards(self):
        if not hasattr(self, 'table_home_cards_frame'):
            return

        for w in self.table_home_cards_frame.winfo_children():
            w.destroy()

        tables = self._get_active_tables()
        hold_map = self._get_table_hold_map()

        if not tables:
            tk.Label(self.table_home_cards_frame, text=self.tr('no_active_tables'),
                     font=("Arial", 12, "bold"),
                     bg=self.colors['bg'], fg=self.colors['text_muted']).pack(pady=20)
            return

        grouped = {}
        for code in tables:
            prefix = ""
            suffix = ""
            for ch in str(code):
                if ch.isalpha() and not suffix:
                    prefix += ch
                elif ch.isdigit():
                    suffix += ch
            row_key = (prefix or "?").upper()
            try:
                num_key = int(suffix) if suffix else 9999
            except Exception:
                num_key = 9999
            grouped.setdefault(row_key, []).append((num_key, str(code)))

        row_keys = sorted(grouped.keys())
        max_display_cols = 10
        max_cols = min(max_display_cols, max(len(grouped[k]) for k in row_keys))

        frame_w = int(self.table_home_cards_frame.winfo_width() or 0)
        if frame_w <= 0:
            frame_w = max(900, int(self.root.winfo_width() * 0.75))
        pad_x = 12
        usable_w = max(360, frame_w - (pad_x * 2 * max_cols))
        cell_w = max(84, usable_w // max(1, max_cols))

        current_grid_row = 0
        for key in row_keys:
            row_items = sorted(grouped[key], key=lambda t: (t[0], t[1]))
            for idx, (_n, code) in enumerate(row_items):
                r = current_grid_row + (idx // max_display_cols)
                cidx = idx % max_display_cols
                hold_data = hold_map.get(code, (0, 0.0))
                hold_count, hold_total = hold_data
                has_hold = hold_count > 0
                if has_hold:
                    bg = "#FCA5A5"
                    fg = "#7F1D1D"
                else:
                    bg = "#E5E7EB"
                    fg = "#334155"

                if not has_hold:
                    label = code
                else:
                    count_label = self.l(f"{hold_count} orders", f"{hold_count} 单")
                    label = f"{code} · {count_label}\n{self.currency} {hold_total:.2f}"
                btn = RoundedButton(
                    self.table_home_cards_frame,
                    text=label,
                    command=lambda t=code: self._enter_table_from_home(t),
                    bg=bg, fg=fg,
                    font=("Arial", 11, "bold"),
                    padx=10, pady=8, radius=6,
                    width=cell_w, height=96
                )
                btn.grid(row=r, column=cidx, padx=8, pady=8, sticky="nsew")
            
            # Advance the grid row by how many display rows were used for this group
            current_grid_row += (len(row_items) + max_display_cols - 1) // max_display_cols

        for i in range(max_cols):
            self.table_home_cards_frame.grid_columnconfigure(i, weight=1)
        for i in range(current_grid_row):
            self.table_home_cards_frame.grid_rowconfigure(i, weight=1)

    def remove_item(self):
        idx = self._order_selected
        if idx is None or idx >= len(self.order):
            messagebox.showwarning("No Selection", "Please select an item to remove")
            return
        self.last_removed = self.order[idx]
        name, price, qty, cat = self.order[idx]
        if qty > 1:
            self.order[idx] = (name, price, qty - 1, cat)
        else:
            self.order.pop(idx)
            if self._order_selected >= len(self.order):
                self._order_selected = len(self.order) - 1 if self.order else None
        self.update_order_display()

    def undo_last_remove(self):
        if hasattr(self, 'last_removed'):
            self.order.append(self.last_removed)
            self.update_order_display()
            delattr(self, 'last_removed')

    def clear_order(self):
        hold_id = self._active_held_order_id
        if not self.order and not hold_id:
            return

        if hold_id:
            confirm_title = self.l("Confirm Clear", "确认清空")
            confirm_msg = self.l(
                "Clear current order and cancel this unpaid order?",
                "确认清空当前订单并取消这笔未结订单吗？"
            )
        else:
            confirm_title = self.l("Confirm Clear", "确认清空")
            confirm_msg = self.l(
                "Clear current order?",
                "确认清空当前订单吗？"
            )

        if not messagebox.askyesno(confirm_title, confirm_msg, parent=self.root):
            return

        if hold_id:
            now = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            with self.db.get_connection() as conn:
                conn.execute("UPDATE held_orders SET status='CANCELLED', updated_at=? WHERE id=? AND status='HELD'",
                             (now, hold_id))
            self.record_action("HOLD_CANCEL", "Unpaid order cancelled by clear order")

        self.order = []
        self._active_held_order_id = None
        self.update_order_display()
        self.refresh_recent_orders_bar()
        self.refresh_table_cards()
        self.refresh_table_home_cards()

    def hold_current_order(self):
        if not self.order:
            messagebox.showwarning("Empty Order", "There is no order to hold")
            return
        if not self.current_table_code:
            messagebox.showwarning("Select Table", "Please select table first.")
            self.open_table_selector()
            return

        staff = dict(self.current_staff or {})
        staff.setdefault('name', 'Unknown')
        staff.setdefault('staff_id', '')
        self._print_and_hold_order(staff)

    def _load_held_order_into_current(self, hold_id, parent=None):
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, items, table_code FROM held_orders WHERE id=? AND status='HELD'", (hold_id,))
            row = c.fetchone()

        if row is None:
            messagebox.showerror("Not Found", "Unpaid order is not available", parent=parent)
            self.refresh_recent_orders_bar()
            return False

        if self.order and self._active_held_order_id != hold_id:
            if not messagebox.askyesno(
                "Replace Current Order",
                "Current order is not empty. Replace it with this unpaid order?",
                parent=parent
            ):
                return False

        try:
            raw_items = json.loads(row['items'])
            loaded = [(n, Decimal(str(p)), int(q), c) for n, p, q, c in raw_items]
        except Exception:
            messagebox.showerror("Corrupt Data", "Unable to load held order items", parent=parent)
            return False

        self.order = loaded
        self._active_held_order_id = int(hold_id)
        if row['table_code']:
            self.current_table_code = str(row['table_code'])
        self.update_order_display()
        self.refresh_table_cards()
        self.record_action("HOLD_RESUME", "Loaded unpaid order")
        return True

    def _pay_held_order_from_card(self, hold_id):
        if self._load_held_order_into_current(hold_id, parent=self.root):
            self.checkout_current_order()

    def refresh_recent_orders_bar(self):
        if not hasattr(self, 'recent_hold_cards_frame'):
            return

        for w in self.recent_hold_cards_frame.winfo_children():
            w.destroy()

        rows = self._get_active_unpaid_rows(limit=18)

        self.recent_hold_count_label.config(text=f"{len(rows)} unpaid")

        if not rows:
            tk.Label(self.recent_hold_cards_frame,
                     text="No unpaid\norders",
                     justify="center",
                     font=("Arial", 10, "bold"),
                     bg=self.colors['panel2'], fg=self.colors['text_muted']).grid(
                         row=0, column=0, padx=8, pady=10, sticky="nsew"
                     )
            return

        cols = 2
        for idx, row in enumerate(rows):
            hid = int(row['id'])
            r = idx // cols
            cidx = idx % cols
            self.recent_hold_cards_frame.grid_columnconfigure(cidx, minsize=114, weight=0)
            self.recent_hold_cards_frame.grid_rowconfigure(r, minsize=116, weight=0)

            try:
                items_list = json.loads(row['items'])
                item_count = sum(int(i[2]) for i in items_list)
            except Exception:
                item_count = 0

            updated = row['updated_at'] or ""
            updated_short = updated[11:16] if len(updated) >= 16 else "--:--"

            card = tk.Frame(self.recent_hold_cards_frame,
                            bg="#B91C1C",
                            width=102, height=104,
                            highlightbackground="#FCA5A5",
                            highlightthickness=1,
                            cursor="hand2")
            card.grid(row=r, column=cidx, padx=6, pady=6, sticky="nsew")
            card.grid_propagate(False)

            tk.Label(card, text=self.tr('unpaid_order'),
                     font=("Arial", 10, "bold"),
                     bg="#B91C1C", fg="white").pack(pady=(8, 0))
            tk.Label(card, text=f"{updated_short}",
                     font=("Arial", 8),
                     bg="#B91C1C", fg="#FECACA").pack(pady=(0, 2))
            tk.Label(card, text=f"{item_count} item(s)",
                     font=("Arial", 8, "bold"),
                     bg="#B91C1C", fg="white").pack()
            tk.Label(card, text=f"{self.currency}{float(row['total']):.2f}",
                     font=("Arial", 9, "bold"),
                     bg="#B91C1C", fg="white").pack(pady=(2, 0))

            def _enter(_, frame=card):
                frame['bg'] = "#991B1B"
                for ch in frame.winfo_children():
                    ch['bg'] = "#991B1B"

            def _leave(_, frame=card):
                frame['bg'] = "#B91C1C"
                for ch in frame.winfo_children():
                    ch['bg'] = "#B91C1C"

            card.bind("<Enter>", _enter)
            card.bind("<Leave>", _leave)

            for w in [card] + card.winfo_children():
                w.bind("<Enter>", _enter)
                w.bind("<Leave>", _leave)
                w.bind("<Button-1>", lambda e, hold=hid: self._pay_held_order_from_card(hold))

    def show_held_orders(self):
        win = tk.Toplevel(self.root)
        win.title(self.l("📥 Held Orders", "📥 未结订单"))
        win.geometry("760x500")
        win.configure(bg=self.colors['bg'])
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text=self.l("📥 Held Orders", "📥 未结订单"), font=("Arial", 16, "bold"),
                 bg=self.colors['bg'], fg=self.colors['header']).pack(pady=(14, 6))

        held_search_var = tk.StringVar()
        held_result_var = tk.StringVar()
        held_query_host = tk.Frame(win, bg=self.colors['bg'])
        held_query_host.pack(fill="x", padx=0, pady=(0, 2))

        tree_wrap = tk.Frame(win, bg=self.colors['bg'])
        tree_wrap.pack(fill="both", expand=True, padx=14, pady=8)

        cols = (
            self.l("Updated", "更新时间"),
            self.l("Cashier", "收银员"),
            self.l("Items", "数量"),
            self.l("Total", "总额"),
        )
        tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", height=14)
        widths = {
            self.l("Updated", "更新时间"): 220,
            self.l("Cashier", "收银员"): 160,
            self.l("Items", "数量"): 90,
            self.l("Total", "总额"): 120,
        }
        for col in cols:
            tree.heading(col, text=col)
            anchor = "e" if col == "Total" else "center"
            tree.column(col, width=widths.get(col, 120), anchor=anchor)

        sb = tk.Scrollbar(tree_wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        hold_map = {}

        def load_held_orders():
            tree.delete(*tree.get_children())
            hold_map.clear()
            rows = self._get_active_unpaid_rows(limit=None)
            keyword = held_search_var.get().strip().casefold()
            if keyword:
                rows = [
                    row for row in rows
                    if keyword in " ".join((
                        str(row['id'] or ''), str(row['updated_at'] or ''),
                        str(row['cashier'] or ''), str(row['table_code'] or ''),
                        str(row['items'] or ''),
                        str(row['note'] or '') if 'note' in row.keys() else '',
                    )).casefold()
                ]
            for row in rows:
                try:
                    items_list = json.loads(row['items'])
                    item_count = sum(int(i[2]) for i in items_list)
                except Exception:
                    item_count = 0
                hid = int(row['id'])
                hold_map[str(hid)] = row
                tree.insert("", "end", iid=str(hid), values=(
                    row['updated_at'],
                    row['cashier'] or "-",
                    item_count,
                    f"{self.currency} {float(row['total']):.2f}"
                ))
            held_result_var.set(
                self.l(f"{len(rows)} unpaid order(s)", f"{len(rows)} 笔未结订单")
            )

        def resume_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning(self.l("No Selection", "未选择"), self.l("Select an unpaid order", "请选择未结订单"), parent=win)
                return

            hid = int(sel[0])
            if not self._load_held_order_into_current(hid, parent=win):
                return
            win.destroy()
            messagebox.showinfo(self.l("Loaded", "已载入"), self.l("Loaded unpaid order", "已载入未结订单"))

        def delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning(self.l("No Selection", "未选择"), self.l("Select an unpaid order", "请选择未结订单"), parent=win)
                return
            hid = int(sel[0])
            if not messagebox.askyesno(self.l("Delete Unpaid Order", "删除未结订单"), self.l("Delete this unpaid order?", "确认删除这笔未结订单？"), parent=win):
                return
            with self.db.get_connection() as conn:
                conn.execute("DELETE FROM held_orders WHERE id=?", (hid,))
            load_held_orders()
            self.refresh_recent_orders_bar()
            self.record_action("HOLD_DELETE", "Deleted unpaid order")

        self._build_query_toolbar(
            held_query_host, held_search_var, load_held_orders,
            on_refresh=load_held_orders,
            on_export=lambda: self._export_treeview_csv(
                tree, win, "held_orders.csv"
            ),
            result_var=held_result_var,
            search_label=self.l("Order / item / cashier:", "订单／商品／收银员：")
        )

        btn_row = tk.Frame(win, bg=self.colors['bg'])
        btn_row.pack(fill="x", padx=14, pady=(0, 14))
        RoundedButton(btn_row, text=self.l("▶ Resume", "▶ 继续"), command=resume_selected,
                      bg=self.colors['success'], fg="white",
                      font=("Arial", 11, "bold"), padx=14, pady=8, radius=8).pack(side="left", padx=4)
        RoundedButton(btn_row, text=self.l("🗑 Delete", "🗑 删除"), command=delete_selected,
                      bg=self.colors['danger'], fg="white",
                      font=("Arial", 11, "bold"), padx=14, pady=8, radius=8).pack(side="left", padx=4)
        RoundedButton(btn_row, text=self.l("✖ Close", "✖ 关闭"), command=win.destroy,
                      bg=self.colors['primary'], fg="white",
                      font=("Arial", 11, "bold"), padx=14, pady=8, radius=8).pack(side="right", padx=4)

        load_held_orders()

    def print_receipt(self):
        if not self.order:
            messagebox.showwarning(self.l("Empty Order", "空订单"), self.l("Please add items to the order", "请先添加商品"))
            return
        self._staff_login_popup(callback=self._print_and_hold_order,
                                title=self.l("👤 Print & Hold by", "👤 打印挂单员工"))

    def _display_unit_price(self, name, actual_price, category, list_prices):
        if category in list_prices and not name.startswith("[FOC] "):
            return Decimal(str(list_prices[category]))
        return actual_price

    def _display_item_prices(self, items, list_prices):
        for item in items:
            if not isinstance(item, (list, tuple)) or len(item) < 3:
                continue
            name, price, qty = item[:3]
            category = item[3] if len(item) >= 4 else None
            unit_price, quantity = self._coerce_order_numbers(price, qty)
            shown_price = self._display_unit_price(name, unit_price, category, list_prices)
            yield name, quantity, unit_price, shown_price

    def _display_price_adjustment(self, items, list_prices):
        if not list_prices:
            return Decimal("0.00")
        return sum(
            ((actual - shown) * qty for _, qty, actual, shown in self._display_item_prices(items, list_prices)),
            Decimal("0.00"),
        ).quantize(Decimal("0.01"), ROUND_HALF_UP)

    def _display_sale_total(self, sale):
        total = Decimal(str(sale['total'] or 0))
        list_prices = json.loads(sale['category_list_prices'])
        if not list_prices:
            return total, ""
        adjustment = self._display_price_adjustment(json.loads(sale['items']), list_prices)
        return total - adjustment, "*" if adjustment else ""

    def _receipt_item_lines(self, items, list_prices):
        lines = []
        for name, quantity, unit_price, shown_price in self._display_item_prices(items, list_prices):
            name_fixed = self.cjk_pad(self.cjk_truncate(name, 16), 16)
            amount = self.cjk_pad(f"{shown_price * quantity:.2f}", 8, "right")
            lines.append(f"{name_fixed}x{quantity:>2} {self.currency} {amount}")
            if quantity > 1 and shown_price != unit_price:
                lines.append(f"  {self.currency} {shown_price:.2f} each")
        return "\n".join(lines)

    def _receipt_adjustment_line(self, adjustment):
        return f"Price adjustment: {self.currency} {adjustment:+.2f}\n" if adjustment else ""

    def _receipt_total_lines(self, actual_total, adjustment, *, paid):
        text = f"TOTAL:            {self.currency} {actual_total - adjustment:>10.2f}\n"
        if adjustment:
            label = "Amount paid" if paid else "Amount due"
            text += self._receipt_adjustment_line(adjustment)
            text += f"{label + ':':18}{self.currency} {actual_total:>10.2f}\n"
        return text

    def _saved_receipt_totals(self, sale):
        total = Decimal(str(sale['total'] or 0))
        tax = Decimal(str(sale['tax'] or 0))
        discount = Decimal(str(sale['discount'] or 0))
        discount_line = f"Discount:      -{self.currency} {discount:>10.2f}\n" if discount else ""
        return total + discount - tax, tax, total, discount_line

    def checkout_current_order(self):
        if not self.order:
            messagebox.showwarning(self.l("Empty Order", "空订单"), self.l("Please add items to the order", "请先添加商品"))
            return
        if not self.current_table_code:
            messagebox.showwarning(
                self.l("Select Table", "选择桌位"),
                self.l("Please select table first.", "请先选择桌位。"),
                parent=self.root,
            )
            self.open_table_selector()
            return
        self._staff_login_popup(
            callback=lambda s: self._show_payment_selection(s, ask_print=True),
            title=self.l("👤 Checkout by", "👤 结账员工")
        )

    def _print_and_hold_order(self, staff):
        if not self.current_table_code:
            messagebox.showwarning(self.l("Select Table", "选择桌位"), self.l("Please select table first.", "请先选择桌位。"))
            self.open_table_selector()
            return

        now = datetime.now()
        date_str = now.strftime("%d-%m-%Y %H:%M:%S")

        subtotal, tax, final_total = self._calc_order_totals(self.order)

        items_json = json.dumps([(n, float(p), q, c) for n, p, q, c in self.order])

        with self.db.get_connection() as conn:
            c = conn.cursor()
            if self._active_held_order_id:
                c.execute(
                    """UPDATE held_orders
                       SET updated_at=?, cashier=?, items=?, subtotal=?, tax=?, total=?, note=?, status='HELD', table_code=?
                       WHERE id=? AND status='HELD' AND table_code=?""",
                    (date_str, staff['name'], items_json,
                     float(subtotal), float(tax), float(final_total), 'Printed awaiting payment', self.current_table_code,
                     int(self._active_held_order_id), self.current_table_code)
                )
                if c.rowcount == 0:
                    messagebox.showerror(
                        self.l("Order Changed", "订单已变更"),
                        self.l("This unpaid order has been paid, cancelled or moved. Reload the table.",
                               "此订单已结账、取消或转桌，请重新载入桌位。"), parent=self.root)
                    return
                hold_id = int(self._active_held_order_id)
            else:
                c.execute(
                    """INSERT INTO held_orders
                       (created_at, updated_at, cashier, items, subtotal, tax, total, note, status, table_code)
                       VALUES (?,?,?,?,?,?,?,?, 'HELD',?)""",
                    (date_str, date_str, staff['name'], items_json,
                     float(subtotal), float(tax), float(final_total), 'Printed awaiting payment', self.current_table_code)
                )
                hold_id = c.lastrowid

            list_prices = list_prices_for_sale(conn, (cat for _, _, _, cat in self.order))
        receipt = self._receipt_item_lines(self.order, list_prices)
        adjustment = self._display_price_adjustment(self.order, list_prices)
        total_lines = self._receipt_total_lines(final_total, adjustment, paid=False)

        text = f"""
{'================================'}
         919 Restaurant
        (CT0159652-K)
  NO.1 G2 BLOK D, JALAN INDAH
    12/2 TAMAN INDAH 84000
        MUAR JOHOR
{'================================'}
Date: {date_str}
Table: {self.current_table_code}
Cashier: {staff['name']}  ({staff['staff_id']})
{'================================'}
{receipt}
{'--------------------------------'}
Subtotal:         {self.currency} {subtotal - adjustment:>10.2f}
Tax ({float(self.tax_rate * 100):.0f}%):         {self.currency} {float(tax):>10.2f}
{'================================'}
{total_lines}{'================================'}
PAYMENT: PENDING
Complete payment from this table checkout.
{'================================'}
"""
        text = "\n\n" + text + "\n\n\n\x1dV\x00"

        do_print = self._confirm_print_popup(
            title=self.l("Confirm Print", "确认打印"),
            message=self.l("Print receipt for unpaid order?", "是否打印未结订单小票？")
        )

        printed_ok = False
        print_error = None
        if do_print:
            try:
                print_to_thermal(PRINTER_NAME, text)
                printed_ok = True
            except Exception as error:
                print_error = str(error)

        self.order = []
        self._active_held_order_id = None
        if print_error:
            messagebox.showwarning(self.l("Print failed", "打印失败"),
                                   self.l("Order saved. Check the printer queue, then reopen the table to print again.",
                                          "订单已挂单。请检查打印队列，然后重新打开桌位打印。")
                                   + f"\n\n{print_error}", parent=self.root)
        self.update_order_display()
        self.refresh_recent_orders_bar()
        self.refresh_table_home_cards()
        action_text = "Printed and held" if printed_ok else "Held without print"
        self.record_action("PRINT_HOLD", f"{action_text} unpaid order ({self.currency} {float(final_total):.2f})", user=staff['name'])
        self.show_table_home()

    # ── IMPROVED: Big payment method selection popup ──────────────────────
    def _show_payment_selection(self, staff, ask_print=True):
        """Full-screen style payment selector — hard to press wrong button."""
        _, _, actual_total = self._calc_order_totals()
        adjustment = self._display_price_adjustment(self.order, self.category_list_prices)
        displayed_total = actual_total - adjustment
        popup = tk.Toplevel(self.root)
        popup.title(self.l("Select Payment", "选择付款方式"))
        popup.configure(bg="#1E293B")
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        pw, ph = 560, 420
        sx = self.root.winfo_x() + (self.root.winfo_width()  - pw) // 2
        sy = self.root.winfo_y() + (self.root.winfo_height() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        # Total display
        tk.Label(popup,
             text=self.l(f"TOTAL  {self.currency} {displayed_total:.2f}",
                         f"总计  {self.currency} {displayed_total:.2f}"),
             font=("Arial", 28, "bold"),
             bg="#1E293B", fg="#4ADE80").pack(pady=(28, 6))

        tk.Label(popup,
               text=self.l("Choose Payment Method", "请选择付款方式"),
             font=("Arial", 13),
             bg="#1E293B", fg="#94A3B8").pack(pady=(0, 20))

        btn_row = tk.Frame(popup, bg="#1E293B")
        btn_row.pack(pady=8)

        def choose(method):
            cash_received = None
            if method == "CASH":
                cash_received = self._cash_balance_popup(actual_total, displayed_total=displayed_total)
                if cash_received is None:
                    return
            
            popup.destroy()
            # We always pass print_receipt=True so that _do_checkout handles the 2nd step (Confirm Print)
            self._do_checkout(staff, print_receipt=True, payment_method=method, cash_received=cash_received)

        # Two big buttons
        pay_options = [
            ("CASH (1)", "CASH", "#15803D"),
            ("QR PAY (2)", "QR",   "#0369A1"),
        ]
        for label, val, color in pay_options:
            RoundedButton(btn_row, text=label, command=lambda v=val: choose(v),
                          bg=color, fg="white",
                          font=("Arial", 18, "bold"),
                          padx=24, pady=28, radius=14,
                          width=150, height=140
                          ).pack(side="left", padx=12)

        # Keyboard shortcuts
        popup.bind("1", lambda e: choose("CASH"))
        popup.bind("2", lambda e: choose("QR"))
        popup.bind("<Escape>", lambda e: popup.destroy())

        shortcut_lbl = tk.Label(popup,
                                text=self.l("Shortcut: 1=Cash  2=QR  Esc=Cancel", 
                                            "快捷键：1=现金  2=二维码  Esc=取消"),
                                font=("Arial", 10, "bold"),
                                bg="#1E293B", fg="#94A3B8")
        shortcut_lbl.pack(pady=(16, 0))

    def _show_cash_qr_payment_choice(self, parent, amount, title, context, callback):
        """Show Cash and QR as the final payment choice for non-sale income."""
        popup = tk.Toplevel(parent)
        popup.title(self.l("Select Payment", "选择付款方式"))
        popup.configure(bg="#1E293B")
        popup.transient(parent)
        popup.grab_set()
        popup.resizable(False, False)

        pw, ph = 560, 410
        parent.update_idletasks()
        sx = parent.winfo_rootx() + max(0, (parent.winfo_width() - pw) // 2)
        sy = parent.winfo_rooty() + max(0, (parent.winfo_height() - ph) // 2)
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        tk.Label(
            popup, text=title, font=("Arial", 15, "bold"),
            bg="#1E293B", fg="white"
        ).pack(pady=(24, 4))
        tk.Label(
            popup, text=f"RM {float(amount):.2f}", font=("Arial", 30, "bold"),
            bg="#1E293B", fg="#4ADE80"
        ).pack(pady=(0, 4))
        tk.Label(
            popup, text=context, font=("Arial", 11),
            bg="#1E293B", fg="#CBD5E1", wraplength=500
        ).pack(pady=(0, 6))
        tk.Label(
            popup,
            text=self.l("Choose Payment Method", "请选择付款方式"),
            font=("Arial", 13, "bold"), bg="#1E293B", fg="#94A3B8"
        ).pack(pady=(4, 14))

        def restore_parent_grab():
            try:
                if parent.winfo_exists():
                    parent.grab_set()
            except Exception:
                pass

        def close_popup():
            popup.destroy()
            restore_parent_grab()

        def choose(method):
            popup.destroy()
            restore_parent_grab()
            callback(method)

        btn_row = tk.Frame(popup, bg="#1E293B")
        btn_row.pack(pady=4)
        for label, method, color in [
            (self.l("CASH (1)", "现金 (1)"), "CASH", "#15803D"),
            (self.l("QR PAY (2)", "二维码 (2)"), "QR", "#0369A1"),
        ]:
            RoundedButton(
                btn_row, text=label, command=lambda value=method: choose(value),
                bg=color, fg="white", font=("Arial", 18, "bold"),
                padx=24, pady=24, radius=14, width=150, height=120
            ).pack(side="left", padx=12)

        tk.Label(
            popup,
            text=self.l("1=Cash   2=QR   Esc=Cancel", "1=现金   2=二维码   Esc=取消"),
            font=("Arial", 10, "bold"), bg="#1E293B", fg="#94A3B8"
        ).pack(pady=(14, 0))

        popup.bind("1", lambda _event: choose("CASH"))
        popup.bind("2", lambda _event: choose("QR"))
        popup.bind("<Escape>", lambda _event: close_popup())
        popup.protocol("WM_DELETE_WINDOW", close_popup)
        popup.focus_force()

    def _cash_balance_popup(self, total_amount, *, displayed_total):
        popup = tk.Toplevel(self.root)
        popup.title(self.l("Cash Balance", "找零计算"))
        popup.configure(bg=self.colors['neutral'])
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        pw, ph = 420, 320
        sx = self.root.winfo_x() + (self.root.winfo_width() - pw) // 2
        sy = self.root.winfo_y() + (self.root.winfo_height() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        # Top Section: Title and Total
        header_f = tk.Frame(popup, bg=self.colors['neutral'])
        header_f.pack(fill="x", pady=(20, 10))

        tk.Label(header_f, text=self.l("💵 Cash Payment", "💵 现金付款"), font=("Arial", 16, "bold"),
                 bg=self.colors['neutral'], fg="white").pack()

        tk.Label(header_f, text=self.l(f"Total {self.currency} {displayed_total:.2f}",
                                      f"总额 {self.currency} {displayed_total:.2f}"),
                 font=("Arial", 20, "bold"),
                 bg=self.colors['neutral'], fg="#4ADE80").pack(pady=4)

        form = tk.Frame(popup, bg=self.colors['neutral'], padx=30)
        form.pack(fill="x", pady=10)

        # Open cash drawer when balance popup appears
        drawer_error = open_cash_drawer()
        if drawer_error:
            messagebox.showwarning(
                self.l("Drawer error", "钱箱错误"),
                self.l("Windows could not send the drawer command. Check the receipt printer queue.",
                       "Windows 无法发送开钱箱指令。请检查小票打印队列。") + f"\n\n{drawer_error}",
                parent=popup,
            )

        tk.Label(form, text=self.l("Cash Received:", "实收现金："), font=("Arial", 13),
                 bg=self.colors['neutral'], fg=self.colors['light']).pack(side="left")

        received_var = tk.StringVar(value=f"{total_amount:.2f}")
        entry = tk.Entry(form, textvariable=received_var, width=12, font=("Arial", 18, "bold"),
                         justify="right", relief="flat", bd=0, highlightthickness=1)
        entry.pack(side="right")

        res_frame = tk.Frame(popup, bg=self.colors['neutral'])
        res_frame.pack(fill="x", pady=15)

        balance_lbl = tk.Label(res_frame, text=self.l("Balance: RM 0.00", "找零：RM 0.00"),
                               font=("Arial", 18, "bold"),
                               bg=self.colors['neutral'], fg="#4ADE80")
        balance_lbl.pack()

        result = {"cash": None}

        def _parse_cash_value():
            raw = (received_var.get() or "").strip().replace(self.currency, "").replace("RM", "").strip()
            try:
                return Decimal(raw)
            except Exception:
                return None

        def _refresh_balance(*_):
            recv = _parse_cash_value()
            if recv is None:
                balance_lbl.config(text=self.l("Balance: Invalid amount", "找零：金额无效"), fg="#F87171")
                return
            balance = (recv - total_amount).quantize(Decimal('0.01'), ROUND_HALF_UP)
            if balance >= 0:
                balance_lbl.config(text=self.l(f"Balance: {self.currency} {float(balance):.2f}", f"找零：{self.currency} {float(balance):.2f}"), fg="#4ADE80")
            else:
                balance_lbl.config(text=self.l(f"Short: {self.currency} {float(abs(balance)):.2f}", f"差额：{self.currency} {float(abs(balance)):.2f}"), fg="#F87171")

        def _confirm():
            recv = _parse_cash_value()
            if recv is None:
                messagebox.showwarning(self.l("Invalid Amount", "金额无效"), self.l("Enter valid cash received amount.", "请输入有效实收金额。"), parent=popup)
                return
            if recv < total_amount:
                messagebox.showwarning(self.l("Insufficient Cash", "现金不足"), self.l(
                    f"Enter at least {self.currency} {total_amount:.2f}.",
                    f"请输入至少 {self.currency} {total_amount:.2f}。",
                ), parent=popup)
                return
            result["cash"] = float(recv)
            popup.destroy()

        def _cancel():
            popup.destroy()

        btn_row = tk.Frame(popup, bg=self.colors['neutral'])
        btn_row.pack(pady=(10, 20))
        RoundedButton(btn_row, text=self.l("✅ Confirm", "✅ 确认"),
                      command=_confirm,
                      bg=self.colors['success'], fg="white",
                      font=("Arial", 14, "bold"),
                      padx=28, pady=10, radius=10).pack(side="left", padx=10)
        RoundedButton(btn_row, text=self.l("✖ Cancel", "✖ 取消"),
                      command=_cancel,
                      bg=self.colors['warning'], fg="white",
                      font=("Arial", 13, "bold"),
                      padx=24, pady=10, radius=10).pack(side="left", padx=10)

        received_var.trace_add("write", _refresh_balance)
        popup.bind("<Return>", lambda e: _confirm())
        popup.bind("<Escape>", lambda e: _cancel())
        popup.protocol("WM_DELETE_WINDOW", _cancel)
        entry.focus_set()
        entry.selection_range(0, "end")
        _refresh_balance()
        self.root.wait_window(popup)
        return result["cash"]


    def _confirm_print_popup(self, title=None, message=None):
        if title is None:
            title = self.l("Confirm Print", "确认打印")
        if message is None:
            message = self.l("Print receipt now?", "现在打印小票吗？")
        popup = tk.Toplevel(self.root)
        popup.title(title)
        popup.configure(bg=self.colors['neutral'])
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        pw, ph = 360, 210
        sx = self.root.winfo_x() + (self.root.winfo_width() - pw) // 2
        sy = self.root.winfo_y() + (self.root.winfo_height() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        tk.Label(popup, text="🖨️", font=("Arial", 24),
                 bg=self.colors['neutral'], fg="white").pack(pady=(16, 0))
        tk.Label(popup, text=self.l("Print Receipt?", "打印小票？"), font=("Arial", 14, "bold"),
                 bg=self.colors['neutral'], fg="white").pack(pady=(4, 2))
        tk.Label(popup, text=message, font=("Arial", 10),
                 bg=self.colors['neutral'], fg=self.colors['light']).pack(pady=(0, 14))

        btn_row = tk.Frame(popup, bg=self.colors['neutral'])
        btn_row.pack()

        choice = {"print": False}

        def choose(do_print):
            choice["print"] = bool(do_print)
            popup.destroy()

        RoundedButton(btn_row, text=self.l("Print (1)", "打印 (1)"),
                      command=lambda: choose(True),
                      bg=self.colors['success'], fg="white",
                      font=("Arial", 12, "bold"),
                      padx=22, pady=10, radius=10).pack(side="left", padx=8)

        RoundedButton(btn_row, text=self.l("Not Print (2)", "不打印 (2)"),
                      command=lambda: choose(False),
                      bg=self.colors['warning'], fg="white",
                      font=("Arial", 12, "bold"),
                      padx=22, pady=10, radius=10).pack(side="left", padx=8)

        popup.bind("1", lambda e: choose(True))
        popup.bind("2", lambda e: choose(False))
        popup.bind("<Return>", lambda e: choose(True))
        popup.bind("<Escape>", lambda e: choose(False))
        popup.protocol("WM_DELETE_WINDOW", lambda: choose(False))
        self.root.wait_window(popup)
        return choice["print"]

    def _adjust_tracked_stock_for_items(self, conn, items, direction,
                                        source_type="SYSTEM", reference_id=None,
                                        recorded_by=None, notes=None):
        """Apply stock changes for tracked items.

        `items` accepts tuples like (name, price, qty) or (name, price, qty, category).
        `direction` is +1 to add back stock, -1 to deduct stock.
        """
        if direction not in (-1, 1):
            raise ValueError("direction must be -1 or 1")

        qty_by_name = {}
        for row in items or []:
            if not row or len(row) < 3:
                continue
            name = str(row[0]).strip()
            if not name:
                continue
            try:
                qty = int(row[2] or 0)
            except Exception:
                qty = 0
            if qty <= 0:
                continue
            qty_by_name[name] = qty_by_name.get(name, 0) + qty

        if not qty_by_name:
            return

        def _resolve_item_name_for_stock(cursor, raw_name):
            # Prefer exact item name. If not found, fallback to base name for variant labels
            # stored in orders like "Coffee (Hot)" or "Anchor Set(3)".
            # Web FOC replacements use a zero-price "[FOC] " display prefix;
            # strip it so desktop checkout still deducts the replacement stock.
            if raw_name.startswith("[FOC] "):
                raw_name = raw_name[6:].strip()
            
            # 1. Check for exact match first
            cursor.execute("SELECT id, name, stock, track_stock FROM items WHERE name=?", (raw_name,))
            row = cursor.fetchone()
            if row:
                return int(row['id']), str(row['name']), int(row['stock'] or 0), int(row['track_stock'] or 0), 1

            multiplier = 1
            search_name = raw_name
            
            # 2. Check for " Set(N)" suffix and extract multiplier
            if " Set(" in raw_name and raw_name.endswith(")"):
                idx = raw_name.rfind(" Set(")
                s_val = raw_name[idx+5:-1]
                if s_val.isdigit():
                    multiplier = int(s_val)
                search_name = raw_name[:idx].strip()
            
            # 3. Check modified search_name (base name)
            cursor.execute("SELECT id, name, stock, track_stock FROM items WHERE name=?", (search_name,))
            row = cursor.fetchone()
            if row:
                return int(row['id']), str(row['name']), int(row['stock'] or 0), int(row['track_stock'] or 0), multiplier

            # 4. Check for parentheses (Variant handling) e.g. "Coffee (Hot)" -> "Coffee"
            if "(" in search_name and search_name.endswith(")"):
                base_name = search_name[:search_name.rfind("(")].strip()
                if base_name:
                    cursor.execute("SELECT id, name, stock, track_stock FROM items WHERE name=?", (base_name,))
                    row = cursor.fetchone()
                    if row:
                        return int(row['id']), str(row['name']), int(row['stock'] or 0), int(row['track_stock'] or 0), multiplier
            
            return None, None, 0, 0, 1

        c = conn.cursor()
        resolved_qty = {}
        for name, qty in qty_by_name.items():
            item_id, resolved_name, stock_before, track_flag, multiplier = _resolve_item_name_for_stock(c, name)
            if not resolved_name or track_flag == 0:
                continue
            effective_qty = qty * multiplier
            key = (item_id, resolved_name)
            resolved_qty[key] = resolved_qty.get(key, 0) + effective_qty

        movement_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        for (item_id, resolved_name), qty in resolved_qty.items():
            c.execute("SELECT stock FROM items WHERE id=?", (item_id,))
            current = c.fetchone()
            if not current:
                raise RuntimeError(f"Tracked item disappeared during stock update: {resolved_name}")
            before = int(current['stock'] or 0)
            delta = direction * qty
            after = before + delta
            if after < 0:
                raise ValueError(f"Insufficient stock for {resolved_name}: {before}, requested {qty}")
            c.execute("UPDATE items SET stock=? WHERE id=?", (after, item_id))
            c.execute("""INSERT INTO stock_movements
                         (datetime,item_id,item_name,quantity_delta,stock_before,stock_after,
                          source_type,reference_id,recorded_by,notes)
                         VALUES (?,?,?,?,?,?,?,?,?,?)""",
                      (movement_time, item_id, resolved_name, delta, before, after,
                       source_type, reference_id, recorded_by or self._audit_actor("System"), notes))

    def _apply_stock_adjustment_batch(self, lines, source_type, staff, notes="",
                                      request_token=None):
        """Apply one multi-item stock document exactly once and all-or-nothing."""
        directions = {
            "STOCK_IN": 1,
            "INTERNAL_USE": -1,
            "WASTE": -1,
            "CORRECTION_ADD": 1,
            "CORRECTION_DEDUCT": -1,
        }
        source_type = str(source_type or "").strip().upper()
        if source_type not in directions:
            raise ValueError("Unknown stock adjustment action.")

        normalized = []
        seen_item_ids = set()
        for line in lines or []:
            if isinstance(line, dict):
                item_raw = line.get("item_id")
                quantity_raw = line.get("quantity")
            else:
                try:
                    item_raw, quantity_raw = line[0], line[1]
                except (TypeError, IndexError):
                    raise ValueError("Each stock row needs an item and quantity.")
            try:
                item_id = int(item_raw)
                quantity_number = float(quantity_raw)
            except (TypeError, ValueError):
                raise ValueError("Each stock row needs a valid item and quantity.")
            if item_id <= 0 or quantity_number <= 0 or not quantity_number.is_integer():
                raise ValueError("Stock quantities must be positive whole numbers.")
            if item_id in seen_item_ids:
                raise ValueError("The same item appears more than once in this stock document.")
            seen_item_ids.add(item_id)
            normalized.append((item_id, int(quantity_number)))

        if not normalized:
            raise ValueError("Add at least one stock item before saving.")

        actor = str(self._staff_value(staff, "name", "Admin") or "Admin")
        notes = str(notes or "").strip()
        token = str(request_token or uuid.uuid4().hex).strip()
        if not token:
            token = uuid.uuid4().hex
        direction = directions[source_type]
        movement_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

        with self.db.get_connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            c = conn.cursor()
            existing = c.execute(
                """SELECT id,item_count,total_quantity FROM stock_adjustment_batches
                   WHERE request_token=?""", (token,)
            ).fetchone()
            if existing:
                return {
                    "batch_id": int(existing["id"]),
                    "item_count": int(existing["item_count"]),
                    "total_quantity": int(existing["total_quantity"]),
                    "duplicate": True,
                    "lines": [],
                }

            prepared = []
            for item_id, quantity in normalized:
                row = c.execute(
                    """SELECT id,name,stock FROM items
                       WHERE id=? AND active=1 AND track_stock=1""", (item_id,)
                ).fetchone()
                if not row:
                    raise ValueError(f"Tracked stock item #{item_id} was not found.")
                before = int(row["stock"] or 0)
                delta = quantity * direction
                after = before + delta
                if after < 0:
                    raise ValueError(
                        f"Insufficient stock for {row['name']}: current {before}, requested {quantity}."
                    )
                prepared.append({
                    "item_id": int(row["id"]), "item_name": str(row["name"]),
                    "quantity": quantity, "delta": delta,
                    "before": before, "after": after,
                })

            total_quantity = sum(row["quantity"] for row in prepared)
            c.execute(
                """INSERT INTO stock_adjustment_batches
                   (datetime,source_type,item_count,total_quantity,recorded_by,notes,request_token)
                   VALUES (?,?,?,?,?,?,?)""",
                (movement_time, source_type, len(prepared), total_quantity,
                 actor, notes or None, token)
            )
            batch_id = int(c.lastrowid)

            for row in prepared:
                c.execute(
                    "UPDATE items SET stock=? WHERE id=?",
                    (row["after"], row["item_id"])
                )
                c.execute(
                    """INSERT INTO stock_movements
                       (datetime,item_id,item_name,quantity_delta,stock_before,stock_after,
                        source_type,reference_id,recorded_by,notes)
                       VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (movement_time, row["item_id"], row["item_name"], row["delta"],
                     row["before"], row["after"], source_type, batch_id, actor,
                     notes or None)
                )

            summary = ", ".join(
                f"{row['item_name']} {row['delta']:+d}" for row in prepared
            )
            self.record_action(
                "STOCK_BATCH",
                f"Batch #{batch_id} | {source_type} | {summary}" +
                (f" | Note: {notes}" if notes else ""),
                user=actor, conn=conn
            )

        return {
            "batch_id": batch_id,
            "item_count": len(prepared),
            "total_quantity": total_quantity,
            "duplicate": False,
            "lines": prepared,
        }

    def _do_checkout(self, staff, print_receipt=True, payment_method=None, cash_received=None):
        if payment_method is None:
            payment_method = self.payment.get()

        now = datetime.now()
        date_str = now.strftime("%d-%m-%Y %H:%M:%S")
        biz_date = get_business_date(now)

        with self.db.get_connection() as conn:
            c = conn.cursor()
            s, e = business_day_range(biz_date)
            # Bug fix: use < not <= for end boundary to avoid double-counting
            c.execute("SELECT COUNT(*) FROM sales WHERE to_iso(datetime) >= ? AND to_iso(datetime) < ?",
                      (s.strftime("%Y-%m-%d %H:%M:%S"), e.strftime("%Y-%m-%d %H:%M:%S")))
            try:
                order_num = c.fetchone()[0] + 1
            except Exception:
                order_num = 1

            subtotal, tax, final_total = self._calc_order_totals(self.order)

            list_prices = list_prices_for_sale(conn, (cat for _, _, _, cat in self.order))
            receipt = self._receipt_item_lines(self.order, list_prices)
            adjustment = self._display_price_adjustment(self.order, list_prices)
            total_lines = self._receipt_total_lines(final_total, adjustment, paid=True)
            # Bug fix: renamed inner variable from 'c' to 'cat' — 'c' was shadowing the cursor
            items_json = json.dumps([(n, float(p), q, cat) for n, p, q, cat in self.order])

            text = f"""
{'================================'}
         919 Restaurant
        (CT0159652-K)
  NO.1 G2 BLOK D, JALAN INDAH
    12/2 TAMAN INDAH 84000
        MUAR JOHOR
{'================================'}
Date: {date_str}
Order: #{order_num:04d}
Table: {self.current_table_code or '-'}
Cashier: {staff['name']}  ({staff['staff_id']})
{'================================'}
{receipt}
{'--------------------------------'}
Subtotal:         {self.currency} {subtotal - adjustment:>10.2f}
Tax ({float(self.tax_rate * 100):.0f}%):         {self.currency} {float(tax):>10.2f}
{'================================'}
{total_lines}{'================================'}
Payment: {payment_method}
{'================================'}
    Terima Kasih! Thank You!
      Please Come Again!
{'================================'}
"""
            text = "\n\n" + text + "\n\n\n\x1dV\x00"

            try:
                c.execute("""INSERT INTO sales (datetime, total, payment, items, discount, tax, cashier, table_code, category_list_prices)
                            VALUES (?,?,?,?,?,?,?,?,?)""",
                          (date_str, float(final_total), payment_method, items_json,
                           0, float(tax), staff['name'], self.current_table_code, json.dumps(list_prices)))
                sale_id = c.lastrowid
                self._adjust_tracked_stock_for_items(
                    conn, self.order, -1, source_type="SALE", reference_id=sale_id,
                    recorded_by=staff['name'], notes=f"Checkout table {self.current_table_code or '-'}"
                )

                if self._active_held_order_id:
                    c.execute("UPDATE held_orders SET status='PAID', updated_at=? WHERE id=?",
                              (date_str, self._active_held_order_id))

                self._record_cashflow(conn, "Income", "Sales", sale_id,
                                      float(final_total), payment_method, staff['name'],
                                      dt=now)

                pm_lower = str(payment_method or '').lower().strip()
                # Bug fix: use safe column mapping instead of f-string SQL interpolation
                _PM_COL = {"cash": "cash_total", "qr": "qr_total", "card": "card_total"}
                pm_col  = _PM_COL.get(pm_lower, "cash_total")
                c.execute("SELECT * FROM daily_summary WHERE date=?", (biz_date,))
                if c.fetchone():
                    c.execute(f"""UPDATE daily_summary
                                 SET total_sales = total_sales + ?,
                                     transaction_count = transaction_count + 1,
                                     {pm_col} = {pm_col} + ?
                                 WHERE date=?""",
                              (float(final_total), float(final_total), biz_date))
                else:
                    cash = float(final_total) if pm_lower == "cash" else 0
                    qr   = float(final_total) if pm_lower == "qr"   else 0
                    card = float(final_total) if pm_lower == "card"  else 0
                    c.execute("""INSERT INTO daily_summary
                                 (date, total_sales, transaction_count, cash_total, qr_total, card_total, total_discount)
                                 VALUES (?,?,?,?,?,?,?)""",
                                        (biz_date, float(final_total), 1, cash, qr, card, 0))

                c.execute("""INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)""",
                          (date_str, 'SALE',
                           f'Order #{order_num:04d} - Total: {self.currency}{float(final_total):.2f} - {payment_method}',
                           staff['name']))

                conn.commit()

                if print_receipt:
                    do_print = self._confirm_print_popup(
                        title="Confirm Print",
                        message="Checkout saved. Print receipt now?"
                    )
                    if do_print:
                        try:
                            print_to_thermal(PRINTER_NAME, text)
                        except Exception as error:
                            messagebox.showwarning(
                                self.l("Receipt print failed", "小票打印失败"),
                                self.l(f"Sale #{sale_id} is saved. Check the printer queue, then reprint from Recent Orders.",
                                       f"订单 #{sale_id} 已保存。请检查打印队列，然后从近期订单重新打印。")
                                + f"\n\n{error}", parent=self.root,
                            )

                self.order = []
                self._active_held_order_id = None
                self.update_order_display()
                self.refresh_recent_orders_bar()
                self.refresh_table_home_cards()
                self.load_items()
                self.update_order_number()
                self.show_table_home()

            except Exception as e:
                conn.rollback()
                messagebox.showerror("Error", f"Transaction failed: {str(e)}")

    def _record_cashflow(self, conn, cf_type, source, ref_id, amount, payment_method, recorded_by, notes="", dt=None, supplier="", biz_date=None):
        """Insert a row into cashflow. Call inside an existing db connection."""
        now = dt or datetime.now()
        conn.execute(
            """INSERT INTO cashflow
               (date, datetime, type, source, reference_id, amount, payment_method, recorded_by, notes, supplier)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (biz_date or get_business_date(now), now.strftime("%d-%m-%Y %H:%M:%S"),
             cf_type, source, ref_id, float(amount), str(payment_method or '').upper(),
             recorded_by, notes, supplier))

    @staticmethod
    def _staff_value(staff, key, default=""):
        """Read authentication fields from either a dict or sqlite Row."""
        try:
            return staff[key] or default
        except (KeyError, TypeError, IndexError):
            return default

    def _get_kuih_batch(self, biz_date=None):
        biz_date = biz_date or today_biz_date()
        with self.db.get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM kuih_batches WHERE date=? LIMIT 1", (biz_date,)
            ).fetchone()
        return dict(row) if row else None

    def _create_kuih_batch(self, cost_amount, full_sale_value, notes, staff, biz_date=None):
        """Record the morning kuih purchase as one cash-only daily batch."""
        cost = float(cost_amount or 0)
        target = float(full_sale_value or 0)
        if cost <= 0:
            raise ValueError("Kuih cost must be more than zero.")
        if target < 0:
            raise ValueError("Full-sale value cannot be negative.")

        biz_date = biz_date or today_biz_date()
        actor = str(self._staff_value(staff, 'name', 'Staff'))
        now = datetime.now()
        with self.db.get_connection() as conn:
            c = conn.cursor()
            if c.execute("SELECT 1 FROM kuih_batches WHERE date=?", (biz_date,)).fetchone():
                raise ValueError("A kuih batch has already been recorded for this business day.")

            c.execute(
                """INSERT INTO kuih_batches
                   (date, opened_at, opened_by, cost_amount, full_sale_value,
                    status, leftover_action, purchase_notes)
                   VALUES (?,?,?,?,?,'OPEN','INTERNAL_CONSUMPTION',?)""",
                (biz_date, now.strftime("%Y-%m-%d %H:%M:%S"), actor,
                 cost, target, str(notes or '').strip())
            )
            batch_id = c.lastrowid
            c.execute(
                """INSERT INTO cash_transactions
                   (datetime,type,category,amount,description,recorded_by,date,
                    payment_method,supplier,quantity)
                   VALUES (?,?,?,?,?,?,?,'CASH','',NULL)""",
                (now.strftime("%Y-%m-%d %H:%M:%S"), 'kuih_cost', 'Kuih Purchase',
                 cost, str(notes or '').strip(), actor, biz_date)
            )
            tx_id = c.lastrowid
            self._record_cashflow(
                conn, "Expense", "Kuih Purchase", tx_id, cost, "CASH", actor,
                str(notes or '').strip(), dt=now, biz_date=biz_date
            )
            c.execute(
                "UPDATE kuih_batches SET cost_transaction_id=? WHERE id=?",
                (tx_id, batch_id)
            )
            self.record_action(
                "KUIH_OPEN",
                f"Batch #{batch_id} | Cost RM {cost:.2f} | Full-sale RM {target:.2f}",
                user=actor, conn=conn
            )
        return self._get_kuih_batch(biz_date)

    def _settle_kuih_batch(self, actual_sales, notes, staff, biz_date=None):
        """Close today's kuih batch using only the cash actually collected."""
        sales = float(actual_sales or 0)
        if sales < 0:
            raise ValueError("Actual kuih sales cannot be negative.")

        biz_date = biz_date or today_biz_date()
        actor = str(self._staff_value(staff, 'name', 'Staff'))
        now = datetime.now()
        with self.db.get_connection() as conn:
            c = conn.cursor()
            batch = c.execute(
                "SELECT * FROM kuih_batches WHERE date=?", (biz_date,)
            ).fetchone()
            if not batch:
                raise ValueError("Record the morning kuih cost before settlement.")
            if str(batch['status']).upper() != 'OPEN':
                raise ValueError("This kuih batch has already been settled.")

            sales_tx_id = None
            if sales > 0:
                c.execute(
                    """INSERT INTO cash_transactions
                       (datetime,type,category,amount,description,recorded_by,date,
                        payment_method,supplier,quantity)
                       VALUES (?,?,?,?,?,?,?,'CASH','',NULL)""",
                    (now.strftime("%Y-%m-%d %H:%M:%S"), 'kuih_sales', 'Kuih Sales',
                     sales, str(notes or '').strip(), actor, biz_date)
                )
                sales_tx_id = c.lastrowid
                self._record_cashflow(
                    conn, "Income", "Kuih Sales", sales_tx_id, sales, "CASH", actor,
                    str(notes or '').strip(), dt=now, biz_date=biz_date
                )

            c.execute(
                """UPDATE kuih_batches
                   SET actual_sales=?, status='CLOSED',
                       leftover_action='INTERNAL_CONSUMPTION', settlement_notes=?,
                       settled_at=?, settled_by=?, sales_transaction_id=?
                   WHERE id=?""",
                (sales, str(notes or '').strip(), now.strftime("%Y-%m-%d %H:%M:%S"),
                 actor, sales_tx_id, batch['id'])
            )
            result = calculate_kuih_batch_summary(
                batch['cost_amount'], sales, batch['full_sale_value']
            )
            self.record_action(
                "KUIH_SETTLE",
                f"Batch #{batch['id']} | Sales RM {sales:.2f} | Margin RM {result['actual_margin']:.2f}",
                user=actor, conn=conn
            )
        return self._get_kuih_batch(biz_date)

    def _void_kuih_batch(self, batch_id, staff):
        """Void a mistaken kuih batch and both linked cash movements."""
        actor = str(self._staff_value(staff, 'name', 'Admin'))
        with self.db.get_connection() as conn:
            c = conn.cursor()
            batch = c.execute("SELECT * FROM kuih_batches WHERE id=?", (int(batch_id),)).fetchone()
            if not batch:
                raise ValueError("Kuih batch not found.")
            for tx_id, source, cf_type in (
                (batch['cost_transaction_id'], 'Kuih Purchase', 'Expense'),
                (batch['sales_transaction_id'], 'Kuih Sales', 'Income'),
            ):
                if tx_id:
                    c.execute(
                        "DELETE FROM cashflow WHERE reference_id=? AND source=? AND type=?",
                        (tx_id, source, cf_type)
                    )
                    c.execute("DELETE FROM cash_transactions WHERE id=?", (tx_id,))
            c.execute("DELETE FROM kuih_batches WHERE id=?", (batch['id'],))
            self.record_action(
                "KUIH_VOID", f"Void batch #{batch['id']} for {batch['date']}",
                user=actor, conn=conn
            )

    def _create_cash_out(self, purpose, detail, amount, staff, supplier="", notes="",
                         quantity=None, item_id=None, deduction_month="", item_name=""):
        """Create one structured cash-out record and any related stock/staff row."""
        purpose = str(purpose or '').strip()
        allowed = {'expense', 'stock_purchase', 'staff_advance', 'cash_transfer'}
        if purpose not in allowed:
            raise ValueError("Unknown cash-out purpose.")
        amount = float(amount or 0)
        if amount <= 0:
            raise ValueError("Cash-out amount must be more than zero.")

        actor = str(self._staff_value(staff, 'name', 'Staff'))
        biz_date = today_biz_date()
        now = datetime.now()
        detail = str(detail or '').strip()
        selected_item = str(item_name or '').strip()
        supplier = str(supplier or '').strip()
        notes = str(notes or '').strip()
        if not detail:
            raise ValueError("Choose a detail/category.")

        cashflow_note = notes
        qty_value = None
        if quantity not in (None, ""):
            try:
                qty_value = float(quantity)
            except (TypeError, ValueError):
                qty_value = 0
            if qty_value <= 0:
                raise ValueError("Quantity must be more than zero.")
            if qty_value.is_integer():
                qty_value = int(qty_value)

        with self.db.get_connection() as conn:
            c = conn.cursor()
            stock_row = None
            if purpose == 'stock_purchase':
                if not qty_value or item_id is None:
                    raise ValueError("Choose a stock item and enter a positive quantity.")
                if not float(qty_value).is_integer():
                    raise ValueError("Tracked stock quantity must be a positive whole number.")
                qty_value = int(qty_value)
                stock_row = c.execute(
                    "SELECT id,name,stock FROM items WHERE id=? AND active=1 AND track_stock=1",
                    (int(item_id),)
                ).fetchone()
                if not stock_row:
                    raise ValueError("Tracked stock item not found.")
                selected_item = str(stock_row['name'])
                cashflow_note = f"{selected_item} × {qty_value}" + (f" | {notes}" if notes else "")
            elif purpose == 'staff_advance':
                staff_id = str(self._staff_value(staff, 'advance_staff_id', ''))
                advance_name = str(self._staff_value(staff, 'advance_staff_name', detail))
                detail = 'Staff Advance'
                selected_item = f"Staff Advance — {advance_name}"
                if not deduction_month:
                    deduction_month = now.strftime("%Y-%m")
                cashflow_note = f"{advance_name} | Deduct {deduction_month}" + (f" | {notes}" if notes else "")
            elif purpose == 'cash_transfer':
                selected_item = str(item_name or detail).strip()
                detail = 'Cash Transfer'
                cashflow_note = selected_item + (f" | {notes}" if notes else "")

            flow_source = detail

            c.execute(
                """INSERT INTO cash_transactions
                   (datetime,type,category,amount,description,recorded_by,date,
                    payment_method,supplier,quantity,item_name)
                   VALUES (?,?,?,?,?,?,?,'CASH',?,?,?)""",
                (now.strftime("%Y-%m-%d %H:%M:%S"), purpose, detail, amount,
                 notes, actor, biz_date, supplier, qty_value, selected_item)
            )
            tx_id = c.lastrowid
            self._record_cashflow(
                conn, "Expense", flow_source, tx_id, amount, "CASH", actor,
                cashflow_note, dt=now, supplier=supplier, biz_date=biz_date
            )

            if purpose == 'stock_purchase':
                before = int(stock_row['stock'] or 0)
                after = before + qty_value
                c.execute("UPDATE items SET stock=? WHERE id=?", (after, stock_row['id']))
                c.execute(
                    """INSERT INTO stock_movements
                       (datetime,item_id,item_name,quantity_delta,stock_before,stock_after,
                        source_type,reference_id,recorded_by,notes)
                       VALUES (?,?,?,?,?,?,'STOCK_PURCHASE',?,?,?)""",
                    (now.strftime("%d-%m-%Y %H:%M:%S"), stock_row['id'], selected_item,
                     qty_value, before, after, tx_id, actor, notes or None)
                )
            elif purpose == 'staff_advance':
                c.execute(
                    """INSERT INTO staff_advances
                       (date,datetime,staff_id,staff_name,amount,deduction_month,status,
                        notes,recorded_by,cash_transaction_id)
                       VALUES (?,?,?,?,?,?,'OUTSTANDING',?,?,?)""",
                    (biz_date, now.strftime("%Y-%m-%d %H:%M:%S"), staff_id,
                     advance_name, amount, deduction_month, notes, actor, tx_id)
                )

            self.record_action(
                "CASH_OUT", f"{selected_item} | {detail} | RM {amount:.2f}",
                user=actor, conn=conn
            )
        return tx_id

    def _void_cash_out(self, tx_id, staff):
        """Void one structured cash-out, safely reversing stock when required."""
        actor = str(self._staff_value(staff, 'name', 'Admin'))
        with self.db.get_connection() as conn:
            c = conn.cursor()
            tx = c.execute(
                """SELECT * FROM cash_transactions
                   WHERE id=? AND type IN ('expense','stock_purchase','staff_advance','cash_transfer')""",
                (int(tx_id),)
            ).fetchone()
            if not tx:
                raise ValueError("Cash-out record not found.")

            if tx['type'] == 'stock_purchase':
                movement = c.execute(
                    """SELECT * FROM stock_movements
                       WHERE reference_id=? AND source_type='STOCK_PURCHASE'
                       ORDER BY id DESC LIMIT 1""", (tx['id'],)
                ).fetchone()
                if movement:
                    item = c.execute(
                        "SELECT id,name,stock FROM items WHERE id=?", (movement['item_id'],)
                    ).fetchone()
                    qty = int(movement['quantity_delta'] or 0)
                    current = int(item['stock'] or 0) if item else 0
                    if not item or qty <= 0 or current < qty:
                        raise ValueError(
                            "Cannot void this stock purchase because some of its stock has already been used. "
                            "Use Stock Control to make a correction instead."
                        )
                    after = current - qty
                    c.execute("UPDATE items SET stock=? WHERE id=?", (after, item['id']))
                    c.execute(
                        """INSERT INTO stock_movements
                           (datetime,item_id,item_name,quantity_delta,stock_before,stock_after,
                            source_type,reference_id,recorded_by,notes)
                           VALUES (?,?,?,?,?,?,'VOID_STOCK_PURCHASE',?,?,?)""",
                        (datetime.now().strftime("%d-%m-%Y %H:%M:%S"), item['id'], item['name'],
                         -qty, current, after, tx['id'], actor, 'Voided cash-out record')
                    )
            elif tx['type'] == 'staff_advance':
                advance = c.execute(
                    "SELECT status FROM staff_advances WHERE cash_transaction_id=?", (tx['id'],)
                ).fetchone()
                if advance and str(advance['status']).upper() != 'OUTSTANDING':
                    raise ValueError("This staff advance has already been processed and cannot be voided.")
                c.execute("DELETE FROM staff_advances WHERE cash_transaction_id=?", (tx['id'],))

            c.execute(
                "DELETE FROM cashflow WHERE type='Expense' AND reference_id=? AND source<>'Refund'",
                (tx['id'],)
            )
            c.execute("DELETE FROM cash_transactions WHERE id=?", (tx['id'],))
            self.record_action(
                "VOID_CASH_OUT",
                f"#{tx['id']} | {tx['type']} | {tx['category']} | RM {float(tx['amount']):.2f}",
                user=actor, conn=conn
            )

    def _get_all_suppliers(self):
        """Fetch all unique suppliers from both cash_transactions and cashflow tables."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            # UNION automatically handles duplicates and returns a unique set.
            # COLLATE NOCASE ensures case-insensitive sorting.
            query = """
                SELECT supplier FROM cash_transactions WHERE supplier IS NOT NULL AND supplier != ''
                UNION
                SELECT supplier FROM cashflow WHERE supplier IS NOT NULL AND supplier != ''
                ORDER BY supplier COLLATE NOCASE
            """
            c.execute(query)
            return [r[0] for r in c.fetchall()]

    def check_data_consistency(self, biz_date=None):
        """Compare totals in sales, cashflow, and daily_summary for a specific business date."""
        if biz_date is None:
            biz_date = today_biz_date()
        
        with self.db.get_connection() as conn:
            c = conn.cursor()
            total_sales, cash_sales, qr_sales, card_sales = compute_sales_payment_breakdown(conn, biz_date)
            
            c.execute("SELECT total_sales, transaction_count, cash_total, qr_total, card_total FROM daily_summary WHERE date=?", (biz_date,))
            ds = c.fetchone()
            
            if not ds:
                return False, f"No daily summary entry found for {biz_date}."
            
            ds_total = float(ds['total_sales'] or 0)
            ds_cash  = float(ds['cash_total'] or 0)
            ds_qr    = float(ds['qr_total'] or 0)
            
            diff_total = abs(total_sales - ds_total)
            diff_cash  = abs(cash_sales - ds_cash)
            diff_qr    = abs(qr_sales - ds_qr)
            
            if diff_total < 0.01 and diff_cash < 0.01 and diff_qr < 0.01:
                return True, f"Data is consistent for {biz_date}.\nTotal: RM {total_sales:.2f}\nCash: RM {cash_sales:.2f}\nQR: RM {qr_sales:.2f}"
            else:
                msg = f"Inconsistency detected for {biz_date}!\n\n"
                msg += f"Source of Truth (Sales/Cashflow):\nTotal: {total_sales:.2f}, Cash: {cash_sales:.2f}, QR: {qr_sales:.2f}\n\n"
                msg += f"Daily Summary Table:\nTotal: {ds_total:.2f}, Cash: {ds_cash:.2f}, QR: {ds_qr:.2f}"
                return False, msg

    def rebuild_daily_summary(self, biz_date):
        """Force re-calculate daily_summary from sales and cashflow for a date."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            total_sales, cash_sales, qr_sales, card_sales = compute_sales_payment_breakdown(conn, biz_date)
            
            # Count transactions in sales table for this biz date
            s, e = business_day_range(biz_date)
            c.execute("SELECT COUNT(*) FROM sales WHERE amended = 0 AND to_iso(datetime) >= ? AND to_iso(datetime) < ?",
                      (s.strftime("%Y-%m-%d %H:%M:%S"), e.strftime("%Y-%m-%d %H:%M:%S")))
            tx_count = c.fetchone()[0]
            
            c.execute("""INSERT OR REPLACE INTO daily_summary 
                         (date, total_sales, transaction_count, cash_total, qr_total, card_total, total_discount)
                         VALUES (?,?,?,?,?,?,0)""",
                      (biz_date, total_sales, tx_count, cash_sales, qr_sales, card_sales))
            conn.commit()
        return True

    def run_consistency_audit(self):
        ok, msg = self.check_data_consistency()
        if ok:
            messagebox.showinfo("System Audit", "✅ " + msg)
        else:
            if messagebox.askyesno("Data Inconsistency", msg + "\n\nWould you like to rebuild the Daily Summary for today?"):
                self.rebuild_daily_summary(today_biz_date())
                messagebox.showinfo("Fixed", "Daily Summary has been rebuilt.")

    def quick_checkout(self):
        if self.order:
            self.checkout_current_order()

    def update_order_number(self):
        # Order number is assigned only at checkout; do not show pre-checkout number.
        return

    def update_time(self):
        now = datetime.now()
        biz = get_business_date(now)
        time_str = now.strftime("%d %b  |  %H:%M:%S")
        if biz != now.strftime("%Y-%m-%d"):
            time_str += f"  (Biz Day {biz})"
        self.time_label.config(text=time_str)
        self.root.after(1000, self.update_time)

    def update_today_stats(self):
        biz = today_biz_date()
        with self.db.get_connection() as conn:
            total_sales, cash_sales, qr_sales, _card_sales = compute_sales_payment_breakdown(conn, biz)
            c = conn.cursor()
            # Count unique transactions for this biz_date that aren't fully refunded
            c.execute("""SELECT COUNT(DISTINCT reference_id) 
                         FROM cashflow 
                         WHERE type='Income' AND source LIKE 'Sales%' AND date=?
                           AND reference_id NOT IN (
                               SELECT DISTINCT reference_id FROM cashflow 
                               WHERE type='Expense' AND source='Refund' AND date=?
                           )""", (biz, biz))
            tx_count = int((c.fetchone() or [0])[0] or 0)

        self.today_sales_label.config(text=f"RM {float(total_sales):.2f}")
        self.today_trans_label.config(text=f"{tx_count}")
        self.today_cash_label.config(text=f"RM {float(cash_sales):.2f}")
        self.today_qr_label.config(text=f"RM {float(qr_sales):.2f}")

    def focus_search(self):
        self.search_entry.focus_set()

    def _audit_actor(self, fallback="System"):
        try:
            if self.current_staff and self.current_staff.get('name'):
                return self.current_staff['name']
        except Exception:
            pass
        return fallback

    def _repair_orphan_cashflow_refs(self):
        """Remove legacy cashflow rows whose cash_transactions reference no longer exists."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute(
                """SELECT cf.id
                   FROM cashflow cf
                   LEFT JOIN cash_transactions ct ON ct.id = cf.reference_id
                   WHERE cf.source NOT LIKE 'Sales%'
                     AND cf.source <> 'Refund'
                     AND cf.source <> 'Rental'
                     AND cf.source <> 'Stall Rental'
                     AND cf.reference_id IS NOT NULL
                     AND cf.type IN ('Income','Expense')
                     AND ct.id IS NULL"""
            )
            ids = [int(r['id']) for r in c.fetchall()]
            if not ids:
                return 0

            c.executemany("DELETE FROM cashflow WHERE id=?", [(i,) for i in ids])
            c.execute(
                "INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                (
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                    "DATA_REPAIR",
                    f"Removed {len(ids)} orphan cashflow row(s) with missing cash_transactions reference",
                    "System",
                ),
            )
            return len(ids)

    def record_action(self, action, details, user=None, conn=None):
        actor = user or self._audit_actor("Admin")
        ts = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        if conn is not None:
            conn.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                         (ts, action, details, actor))
            return
        with self.db.get_connection() as _conn:
            _conn.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                          (ts, action, details, actor))

    def _show_loading_overlay(self, message=None, parent=None):
        """Paint a lightweight wait card before a synchronous screen is built."""
        owner = parent or self.root
        try:
            if not owner.winfo_exists():
                owner = self.root
        except Exception:
            owner = self.root

        overlay = tk.Toplevel(owner)
        overlay._skip_auto_center = True
        overlay.overrideredirect(True)
        overlay.configure(bg="#0F172A", highlightthickness=1, highlightbackground="#334155")
        try:
            overlay.transient(owner)
            overlay.attributes("-topmost", True)
        except tk.TclError:
            pass

        card = tk.Frame(overlay, bg="#0F172A", padx=28, pady=20)
        card.pack(fill="both", expand=True)
        tk.Label(
            card, text="⏳", font=("Arial", 24),
            bg="#0F172A", fg="white"
        ).pack(side="left", padx=(0, 16))
        text_wrap = tk.Frame(card, bg="#0F172A")
        text_wrap.pack(side="left", fill="both", expand=True)
        tk.Label(
            text_wrap,
            text=message or self.l("Loading, please wait…", "加载中，请稍候…"),
            font=("Arial", 12, "bold"), bg="#0F172A", fg="white",
            anchor="w"
        ).pack(fill="x")
        tk.Label(
            text_wrap,
            text=self.l("Preparing the requested screen", "正在准备所选画面"),
            font=("Arial", 9), bg="#0F172A", fg="#94A3B8", anchor="w"
        ).pack(fill="x", pady=(3, 7))
        progress = ttk.Progressbar(text_wrap, mode="indeterminate", length=260)
        progress.pack(fill="x")
        progress.start(12)

        overlay.update_idletasks()
        width = max(410, overlay.winfo_reqwidth())
        height = max(112, overlay.winfo_reqheight())
        try:
            owner.update_idletasks()
            x = owner.winfo_rootx() + max(0, (owner.winfo_width() - width) // 2)
            y = owner.winfo_rooty() + max(0, (owner.winfo_height() - height) // 2)
        except tk.TclError:
            x = max(0, (overlay.winfo_screenwidth() - width) // 2)
            y = max(0, (overlay.winfo_screenheight() - height) // 2)
        overlay.geometry(f"{width}x{height}+{x}+{y}")
        overlay.lift()
        try:
            owner.configure(cursor="watch")
        except tk.TclError:
            pass
        overlay._loading_progress = progress
        overlay._loading_owner = owner
        overlay.update_idletasks()
        return overlay

    def _hide_loading_overlay(self, overlay):
        if overlay is None:
            return
        try:
            overlay._loading_progress.stop()
        except Exception:
            pass
        try:
            overlay._loading_owner.configure(cursor="")
        except Exception:
            pass
        try:
            overlay.destroy()
        except tk.TclError:
            pass

    def _run_with_loading(self, callback, message=None, parent=None, on_complete=None):
        """Run a Tk screen-building action after its loading card is visible."""
        if getattr(self, "_loading_action_pending", False):
            return "break"
        self._loading_action_pending = True
        overlay = self._show_loading_overlay(message=message, parent=parent)

        def execute():
            try:
                callback()
                if on_complete:
                    on_complete()
            except Exception as exc:
                messagebox.showerror(
                    self.l("Unable to Open", "无法打开"), str(exc), parent=self.root
                )
            finally:
                self._hide_loading_overlay(overlay)
                self._loading_action_pending = False

        self.root.after(80, execute)
        return "break"

    def _run_ui_action(self, label, callback):
        self.record_action("UI_ACTION", f"Clicked: {label}")
        return self._run_with_loading(
            callback,
            on_complete=self._auto_refresh_views
        )

    def _auto_refresh_views(self):
        # Keep DB-backed cards/lists in sync after user interactions.
        try:
            self.refresh_recent_orders_bar()
        except Exception:
            pass
        try:
            self.refresh_table_cards()
        except Exception:
            pass
        try:
            self.refresh_table_home_cards()
        except Exception:
            pass

    def _safe_window_dimensions(self, window, width, height):
        """Clamp a requested window size to the current display's usable area."""
        try:
            screen_width = int(window.winfo_screenwidth())
            screen_height = int(window.winfo_screenheight())
        except Exception:
            screen_width, screen_height = 1366, 768
        usable_width = max(480, screen_width - 32)
        usable_height = max(360, screen_height - 72)
        return (
            max(320, min(int(width), usable_width)),
            max(240, min(int(height), usable_height)),
        )

    def _center_window(self, window, width, height, parent=None):
        width, height = self._safe_window_dimensions(window, width, height)
        target = parent or self.root
        window.update_idletasks()
        try:
            target.update_idletasks()
            px, py = target.winfo_x(), target.winfo_y()
            pw, ph = target.winfo_width(), target.winfo_height()
            if pw <= 1 or ph <= 1:
                raise ValueError("Parent size unavailable")
            x = px + max(0, (pw - width) // 2)
            y = py + max(0, (ph - height) // 2)
        except Exception:
            sw = window.winfo_screenwidth()
            sh = window.winfo_screenheight()
            x = max(0, (sw - width) // 2)
            y = max(0, (sh - height) // 2)
        try:
            sw = int(window.winfo_screenwidth())
            sh = int(window.winfo_screenheight())
            x = max(0, min(int(x), max(0, sw - width)))
            y = max(0, min(int(y), max(0, sh - height - 40)))
        except Exception:
            pass
        try:
            min_width, min_height = window.minsize()
            window.minsize(min(int(min_width), width), min(int(min_height), height))
        except Exception:
            pass
        window.geometry(f"{width}x{height}+{x}+{y}")
        try:
            window._auto_centered_once = True
        except Exception:
            pass
        return width, height

    def _setup_auto_center_toplevels(self):
        """Fit and center every new Toplevel once, including on smaller displays."""
        def _on_map(event):
            win = event.widget
            if not isinstance(win, tk.Toplevel):
                return
            if getattr(win, "_skip_auto_center", False):
                return
            if getattr(win, "_auto_centered_once", False):
                return
            try:
                win.update_idletasks()
                w = int(win.winfo_width() or 0)
                h = int(win.winfo_height() or 0)
                if w <= 1:
                    w = int(win.winfo_reqwidth() or 480)
                if h <= 1:
                    h = int(win.winfo_reqheight() or 360)
                self._center_window(win, max(320, w), max(220, h), self.root)
                win.after_idle(lambda current=win: self._localize_widget_tree(current))
            except Exception:
                pass

        # Bind on class so it applies to every future tk.Toplevel.
        self.root.bind_class("Toplevel", "<Map>", _on_map, add="+")

    def _tree_sort_value(self, value):
        text = str(value or "").strip()
        cleaned = (
            text.replace(self.currency, "")
                .replace(",", "")
                .replace("#", "")
                .split("[", 1)[0]
                .strip()
        )
        try:
            return (0, float(cleaned))
        except ValueError:
            pass
        for fmt in (
            "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
            "%d-%m-%Y",
        ):
            try:
                return (1, datetime.strptime(cleaned, fmt))
            except ValueError:
                continue
        return (2, text.casefold())

    def _sort_treeview_column(self, tree, column, descending=False):
        try:
            rows = [
                (self._tree_sort_value(tree.set(item_id, column)), item_id)
                for item_id in tree.get_children("")
            ]
            rows.sort(key=lambda pair: pair[0], reverse=bool(descending))
            for index, (_value, item_id) in enumerate(rows):
                tree.move(item_id, "", index)
            tree.heading(
                column,
                command=lambda: self._sort_treeview_column(
                    tree, column, not bool(descending)
                )
            )
        except Exception:
            pass

    def _enable_treeview_sorting(self, tree):
        if getattr(tree, "_pos_sorting_enabled", False):
            return
        tree._pos_sorting_enabled = True
        try:
            for column in tree["columns"]:
                existing_command = tree.heading(column, "command")
                if existing_command:
                    continue
                tree.heading(
                    column,
                    command=lambda col=column: self._sort_treeview_column(
                        tree, col, False
                    )
                )
        except Exception:
            pass

        # Every table also gets the same keyboard/right-click fallback actions,
        # including small popup lists that do not have room for a full toolbar.
        try:
            menu = tk.Menu(tree, tearoff=False)
            menu.add_command(
                label=self.l("🔍 Find in displayed rows", "🔍 在显示资料中查找"),
                command=lambda: self._find_treeview_text(tree)
            )
            menu.add_command(
                label=self.l("📋 Copy selected rows", "📋 复制所选资料"),
                command=lambda: self._copy_treeview_selection(tree)
            )
            menu.add_separator()
            menu.add_command(
                label=self.l("⬇ Export displayed rows", "⬇ 导出显示资料"),
                command=lambda: self._export_treeview_csv(
                    tree, tree.winfo_toplevel(), "POS_List.csv"
                )
            )
            tree._pos_context_menu = menu

            def show_menu(event):
                row_id = tree.identify_row(event.y)
                if row_id and row_id not in tree.selection():
                    tree.selection_set(row_id)
                    tree.focus(row_id)
                menu.tk_popup(event.x_root, event.y_root)

            tree.bind("<Button-3>", show_menu, add="+")
            tree.bind("<Control-f>", lambda _event: self._find_treeview_text(tree), add="+")
            tree.bind(
                "<Control-e>",
                lambda _event: self._export_treeview_csv(
                    tree, tree.winfo_toplevel(), "POS_List.csv"
                ),
                add="+"
            )
        except Exception:
            pass

    def _find_treeview_text(self, tree):
        query = simpledialog.askstring(
            self.l("Find", "查找"),
            self.l("Enter text to find in displayed rows:", "输入要在显示资料中查找的文字："),
            parent=tree.winfo_toplevel()
        )
        if query is None:
            return
        needle = query.strip().casefold()
        if not needle:
            return
        rows = list(tree.get_children(""))
        current = tree.focus()
        start_at = rows.index(current) + 1 if current in rows else 0
        ordered = rows[start_at:] + rows[:start_at]
        for item_id in ordered:
            values = tree.item(item_id, "values")
            if needle in " ".join(str(value) for value in values).casefold():
                tree.selection_set(item_id)
                tree.focus(item_id)
                tree.see(item_id)
                return
        messagebox.showinfo(
            self.l("Not Found", "找不到"),
            self.l("No displayed row contains that text.", "显示资料中没有包含该文字的记录。"),
            parent=tree.winfo_toplevel()
        )

    def _copy_treeview_selection(self, tree):
        selected = list(tree.selection())
        if not selected:
            messagebox.showwarning(
                self.l("No Selection", "未选择"),
                self.l("Select one or more rows first.", "请先选择一行或多行资料。"),
                parent=tree.winfo_toplevel()
            )
            return
        lines = ["\t".join(str(value) for value in tree.item(item_id, "values")) for item_id in selected]
        tree.clipboard_clear()
        tree.clipboard_append("\n".join(lines))

    def _setup_auto_sortable_treeviews(self):
        """Make every tabular list sortable by clicking a column heading."""
        def _on_tree_map(event):
            tree = event.widget
            if isinstance(tree, ttk.Treeview):
                self._enable_treeview_sorting(tree)
                try:
                    tree.after_idle(lambda current=tree: self._localize_widget_tree(current))
                except Exception:
                    pass

        # ttk class bind tags differ across Tk builds (for example Treeview vs
        # TTreeview), so listen globally and filter by the actual widget type.
        self.root.bind_all("<Map>", _on_tree_map, add="+")

    def _export_treeview_csv(self, tree, parent=None, default_name="POS_List.csv"):
        rows = list(tree.get_children(""))
        if not rows:
            messagebox.showwarning(
                self.l("No Data", "没有资料"),
                self.l("There are no displayed rows to export.", "目前没有可导出的显示资料。"),
                parent=parent or self.root
            )
            return
        columns = list(tree["columns"])
        path = filedialog.asksaveasfilename(
            parent=parent or self.root,
            initialfile=default_name,
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title=self.l("Export Displayed Rows", "导出目前显示资料")
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.writer(handle)
                writer.writerow([tree.heading(col, "text") or col for col in columns])
                for item_id in rows:
                    writer.writerow(tree.item(item_id, "values"))
            messagebox.showinfo(
                self.l("Export Success", "导出成功"),
                self.l(f"Saved to:\n{path}", f"已保存到：\n{path}"),
                parent=parent or self.root
            )
        except Exception as exc:
            messagebox.showerror(
                self.l("Export Failed", "导出失败"),
                str(exc),
                parent=parent or self.root
            )

    def _build_query_toolbar(
        self, parent, search_var, on_search, on_clear=None, on_refresh=None,
        on_export=None, result_var=None, search_label=None
    ):
        """Build the shared Search → Clear → Refresh → Export list toolbar."""
        C = self.colors
        bar = tk.Frame(
            parent, bg=C['panel'],
            highlightbackground=C['border'], highlightthickness=1
        )
        bar.pack(fill="x", padx=16, pady=(4, 8))
        bar.columnconfigure(0, weight=1)
        search_row = tk.Frame(bar, bg=C['panel'])
        action_row = tk.Frame(bar, bg=C['panel'])
        search_row.grid(row=0, column=0, sticky="ew")
        action_row.grid(row=0, column=1, sticky="e")
        tk.Label(
            search_row,
            text=search_label or self.l("Search:", "搜索："),
            font=("Arial", 10, "bold"), bg=C['panel'], fg=C['text']
        ).pack(side="left", padx=(12, 5), pady=8)
        entry = tk.Entry(
            search_row, textvariable=search_var, font=("Arial", 11),
            bg="white", fg=C['text'], relief="solid", bd=1
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=8)
        entry.bind("<Return>", lambda _event: on_search())

        def clear_query():
            search_var.set("")
            if on_clear:
                on_clear()
            else:
                on_search()

        for text_value, command, color in [
            (self.l("🔍 Search", "🔍 查询"), on_search, C['primary']),
            (self.l("✖ Clear", "✖ 清除"), clear_query, C['neutral']),
            (self.l("🔄 Refresh", "🔄 刷新"), on_refresh or on_search, C['info']),
        ]:
            RoundedButton(
                action_row, text=text_value, command=command,
                bg=color, fg="white", font=("Arial", 9, "bold"),
                padx=10, pady=5, radius=7
            ).pack(side="left", padx=3)
        if on_export:
            RoundedButton(
                action_row, text=self.l("⬇ Export", "⬇ 导出"), command=on_export,
                bg=C['success'], fg="white", font=("Arial", 9, "bold"),
                padx=10, pady=5, radius=7
            ).pack(side="left", padx=3)
        if result_var is not None:
            tk.Label(
                search_row, textvariable=result_var, font=("Arial", 9, "bold"),
                bg=C['panel'], fg=C['text_muted']
            ).pack(side="right", padx=10)

        def reflow(event=None):
            try:
                width = int(event.width if event is not None else bar.winfo_width())
                compact = width < 780
                if getattr(bar, "_query_compact", None) == compact:
                    return
                bar._query_compact = compact
                if compact:
                    search_row.grid_configure(row=0, column=0, columnspan=2, sticky="ew")
                    action_row.grid_configure(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 7))
                else:
                    search_row.grid_configure(row=0, column=0, columnspan=1, sticky="ew")
                    action_row.grid_configure(row=0, column=1, columnspan=1, sticky="e", padx=(0, 8), pady=0)
            except Exception:
                pass

        bar.bind("<Configure>", reflow, add="+")
        bar.after_idle(reflow)
        return bar, entry

    def _msg_info(self, en_title, zh_title, en_message, zh_message, parent=None):
        return messagebox.showinfo(
            self.l(en_title, zh_title), self.l(en_message, zh_message),
            parent=parent or self.root
        )

    def _msg_warning(self, en_title, zh_title, en_message, zh_message, parent=None):
        return messagebox.showwarning(
            self.l(en_title, zh_title), self.l(en_message, zh_message),
            parent=parent or self.root
        )

    def _msg_error(self, en_title, zh_title, en_message, zh_message, parent=None):
        return messagebox.showerror(
            self.l(en_title, zh_title), self.l(en_message, zh_message),
            parent=parent or self.root
        )

    def _ask_yes_no(self, en_title, zh_title, en_message, zh_message, parent=None):
        return messagebox.askyesno(
            self.l(en_title, zh_title), self.l(en_message, zh_message),
            parent=parent or self.root
        )

    def show_quantity_sets_dialog(self, item_id, item_name, parent_win):
        win = tk.Toplevel(parent_win)
        win.title(f"📦 Quantity Sets - {item_name}")
        win.geometry("500x550")
        win.configure(bg=self.colors['bg'])
        win.transient(parent_win)
        win.grab_set()

        tk.Label(win, text=f"Quantity Sets: {item_name}", font=("Arial", 14, "bold"),
                 bg=self.colors['bg'], fg=self.colors['header']).pack(pady=14)

        form = tk.Frame(win, bg=self.colors['bg'])
        form.pack(fill="x", padx=30, pady=10)

        tk.Label(form, text="Set Quantity:", bg=self.colors['bg'], font=("Arial", 11)).grid(row=0, column=0, sticky="w", pady=4)
        qty_var = tk.StringVar()
        tk.Entry(form, textvariable=qty_var, font=("Arial", 11), width=10).grid(row=0, column=1, sticky="w", padx=10)

        tk.Label(form, text="Set Price (RM):", bg=self.colors['bg'], font=("Arial", 11)).grid(row=1, column=0, sticky="w", pady=4)
        price_var = tk.StringVar()
        tk.Entry(form, textvariable=price_var, font=("Arial", 11), width=10).grid(row=1, column=1, sticky="w", padx=10)

        def add_set():
            try:
                q = int(qty_var.get())
                p = float(price_var.get())
            except ValueError:
                messagebox.showerror("Error", "Quantity and Price must be numeric", parent=win)
                return
            if q <= 1:
                messagebox.showerror("Error", "Quantity must be greater than 1", parent=win)
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("INSERT INTO item_quantity_sets (item_id, quantity, set_price) VALUES (?,?,?)",
                          (item_id, q, p))
            qty_var.set(""); price_var.set("")
            load_sets()

        RoundedButton(win, text="➕ Add New Quantity Set", command=add_set, bg=self.colors['success'], fg="white",
                      font=("Arial", 11, "bold"), padx=14, pady=8, radius=8
                      ).pack(pady=10)

        tree_f = tk.Frame(win, bg=self.colors['bg'])
        tree_f.pack(fill="both", expand=True, padx=20, pady=10)

        cols = ("ID", "Quantity", "Set Price")
        tree = ttk.Treeview(tree_f, columns=cols, show="headings", height=8)
        tree.heading("ID", text="ID")
        tree.heading("Quantity", text="Qty")
        tree.heading("Set Price", text="Set Price")
        tree.column("ID", width=50)
        tree.column("Quantity", width=100)
        tree.column("Set Price", width=120)

        sb = tk.Scrollbar(tree_f, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def load_sets():
            tree.delete(*tree.get_children())
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT id, quantity, set_price FROM item_quantity_sets WHERE item_id=? ORDER BY quantity DESC", (item_id,))
                for r in c.fetchall():
                    tree.insert("", "end", values=(r[0], r[1], f"{float(r[2]):.2f}"))

        def delete_set():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("No Selection", "Select a set to delete", parent=win); return
            sid = tree.item(sel[0])['values'][0]
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM item_quantity_sets WHERE id=?", (sid,))
            load_sets()

        btn_f = tk.Frame(win, bg=self.colors['bg'])
        btn_f.pack(fill="x", side="bottom", padx=20, pady=10)
        RoundedButton(btn_f, text="🗑️ Delete Set", command=delete_set, bg=self.colors['danger'], fg="white",
                      font=("Arial", 11, "bold"), padx=14, pady=6, radius=8).pack(side="left")
        RoundedButton(btn_f, text="✖ Close", command=win.destroy, bg=self.colors['neutral'], fg="white",
                      font=("Arial", 11), padx=14, pady=6, radius=8).pack(side="right")

        load_sets()

    def require_password(self, correct_password, on_success, parent=None):
        parent_win = parent if parent else self.root
        dialog = tk.Toplevel(parent_win)
        dialog.title("Manager Access")
        dialog.configure(bg=self.colors['bg'])
        dialog.transient(parent_win)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Premium feel: remove standard title bar if we were doing custom, 
        # but for now let's just make the interior stunning.
        self._center_window(dialog, 480, 320, parent_win)

        # Outer padding frame
        outer = tk.Frame(dialog, bg=self.colors['bg'], padx=2, pady=2)
        outer.pack(fill="both", expand=True)

        # 1. Stunning Header
        hdr = tk.Frame(outer, bg=self.colors['header'], height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        
        tk.Label(hdr, text="🔒", font=("Arial", 18), 
                 bg=self.colors['header'], fg="white").pack(side="left", padx=(20, 10))
        tk.Label(hdr, text="Manager Access", font=("Arial", 15, "bold"),
                 bg=self.colors['header'], fg="white").pack(side="left")

        # 2. Main Body Card
        body = tk.Frame(outer, bg=self.colors['bg'], padx=30, pady=25)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Authorization Required", font=("Arial", 14, "bold"),
                 bg=self.colors['bg'], fg=self.colors['text']).pack(pady=(0, 5))
        
        tk.Label(body, text="Please enter the manager password below.",
                 font=("Arial", 10), bg=self.colors['bg'],
                 fg=self.colors['text_sub']).pack()

        # Input Wrapper for a "premium" entry feel
        entry_wrap = tk.Frame(body, bg=self.colors['border'], padx=1, pady=1)
        entry_wrap.pack(pady=20)
        
        password_entry = tk.Entry(entry_wrap, font=("Arial", 20, "bold"), show="●", width=15,
                                  justify="center", bg='white', fg=self.colors['text'],
                                  relief='flat', bd=0)
        password_entry.pack(padx=1, pady=1)
        password_entry.focus()

        msg_lbl = tk.Label(body, text="", font=("Arial", 10, "bold"),
                           bg=self.colors['bg'], fg=self.colors['danger'])
        msg_lbl.pack(pady=(0, 10))

        def verify(event=None):
            if password_entry.get() == correct_password:
                self.record_action("ACCESS_GRANTED", "Manager access granted")
                dialog.destroy()
                on_success()
            else:
                self.record_action("ACCESS_DENIED", "Manager access denied")
                msg_lbl.config(text="❌ Incorrect password. Try again.")
                password_entry.delete(0, tk.END)
                # Visual shake or feedback could go here
                password_entry.focus_set()

        password_entry.bind("<Return>", verify)

        # 3. Action Buttons
        btn_row = tk.Frame(body, bg=self.colors['bg'])
        btn_row.pack(fill="x")

        RoundedButton(btn_row, text="Cancel", command=dialog.destroy,
                      bg="#64748B", fg="white",
                      font=("Arial", 11, "bold"), padx=25, pady=10, radius=10
                      ).pack(side="left", expand=True, padx=5)

        RoundedButton(btn_row, text="Confirm", command=verify,
                      bg=self.colors['success'], fg="white",
                      font=("Arial", 11, "bold"), padx=25, pady=10, radius=10
                      ).pack(side="left", expand=True, padx=5)

        dialog.bind("<Return>", lambda e: verify())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

    # ================= MANAGE ITEMS =================
    def show_cashier_stock_control(self):
        self.manage_items(stock_only=True)

    def manage_items(self, stock_only=False):
        L = self.l
        manage_window = tk.Toplevel(self.root)
        manage_window.title(L("Stock Adjustment", "库存调整") if stock_only else L("Items & Inventory Settings", "商品与库存设置"))
        manage_window.geometry("1300x900")
        manage_window.configure(bg=self.colors['bg'])
        manage_window.transient(self.root)

        header = tk.Frame(manage_window, bg=self.colors['neutral'], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=L("📦 Stock Adjustment", "📦 库存调整") if stock_only else L("⚙️ Items & Inventory Settings", "⚙️ 商品与库存设置"), font=("Arial", 20, "bold"),
                 bg=self.colors['neutral'], fg=self.colors['light']).pack(pady=12)

        nb = ttk.Notebook(manage_window)
        nb.pack(fill="both", expand=True, padx=16, pady=(10, 10))

        # Increase font size for Treeview in this window
        style = ttk.Style(manage_window)
        style.configure("Manage.Treeview", font=("Arial", 12), rowheight=32)
        style.configure("Manage.Treeview.Heading", font=("Arial", 12, "bold"))

        item_tab = tk.Frame(nb, bg=self.colors['bg'])
        stock_tab = tk.Frame(nb, bg=self.colors['bg'])
        stock_document_tab = tk.Frame(nb, bg=self.colors['bg'])
        nb.add(item_tab, text=L("📋 Items", "📋 商品"))
        nb.add(stock_tab, text=L("📦 Stock Status", "📦 库存状态"))
        nb.add(stock_document_tab, text=L("🧾 Stock Document", "🧾 库存单据"))

        if stock_only:
            nb.hide(item_tab)
            nb.select(stock_document_tab)

        def manage_quantity_sets():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select an item first", "请先选择一个商品。"), parent=manage_window)
                return
            item_vals = tree.item(sel[0])['values']
            item_id = item_vals[0]
            item_name = item_vals[1]
            self.show_quantity_sets_dialog(item_id, item_name, manage_window)

        sidebar = tk.Frame(item_tab, bg=self.colors['panel2'], width=200, highlightbackground=self.colors['border'], highlightthickness=1)
        sidebar.pack(side="left", fill="y", padx=(10, 0), pady=10)
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text=L("ACTIONS", "操作"), font=("Arial", 10, "bold"),
                 bg=self.colors['panel2'], fg=self.colors['text_sub']).pack(pady=(16, 10))

        main_item_frame = tk.Frame(item_tab, bg=self.colors['bg'])
        main_item_frame.pack(side="left", fill="both", expand=True)

        item_search_var = tk.StringVar()
        item_result_var = tk.StringVar()
        item_query_host = tk.Frame(main_item_frame, bg=self.colors['bg'])
        item_query_host.pack(fill="x", padx=0, pady=(8, 0))

        tree_frame = tk.Frame(main_item_frame, bg=self.colors['bg'])
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("ID", "Name", "Price", "Category", "Stock Qty", "Reorder At", "Stock Mode", "Barcode")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20, style="Manage.Treeview")
        
        # Configure Low Stock Tag
        tree.tag_configure('low_stock', foreground='#EF4444', font=("Arial", 12, "bold"))

        for col in columns:
            tree.heading(col, text={
                "ID": "ID", "Name": L("Name", "名称"),
                "Price": L("Price", "价格"), "Category": L("Category", "类别"),
                "Stock Qty": L("Stock Qty", "库存数量"),
                "Reorder At": L("Reorder At", "补货警戒"),
                "Stock Mode": L("Stock Mode", "库存模式"),
                "Barcode": L("Barcode", "条码"),
            }[col])
            tree.column(col, width=120, anchor="center")
        tree.column("Name", width=200, anchor="w")

        scrollbar = tk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def load_all_items():
            tree.delete(*tree.get_children())
            keyword = item_search_var.get().strip()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                query = """SELECT id, name, price, category, stock, low_stock_threshold, track_stock, barcode
                           FROM items WHERE active=1"""
                params = []
                if keyword:
                    query += " AND (name LIKE ? OR category LIKE ? OR barcode LIKE ? OR CAST(id AS TEXT) LIKE ?)"
                    params.extend([f"%{keyword}%"] * 4)
                query += " ORDER BY track_stock DESC, category, name"
                c.execute(query, params)
                items = c.fetchall()
            for item in items:
                is_tracked = int(item['track_stock'] or 0) != 0
                stock_mode = L("Track Stock", "追踪库存") if is_tracked else L("No Tracking", "不追踪")
                stock_qty = item['stock'] if is_tracked else "—"
                threshold = item['low_stock_threshold'] if is_tracked else "—"
                
                tags = ()
                if is_tracked and int(item['stock'] or 0) <= int(item['low_stock_threshold'] or 0):
                    tags = ('low_stock',)

                tree.insert("", "end", values=(
                    item['id'], item['name'], item['price'], item['category'],
                    stock_qty, threshold, stock_mode, item['barcode'] or ""
                ), tags=tags)
            item_result_var.set(L(f"{len(items)} item(s)", f"{len(items)} 个商品"))

        def add_item_dialog():
            dialog = tk.Toplevel(manage_window)
            dialog.title("Add New Item")
            dialog.geometry("680x650")
            dialog.configure(bg=self.colors['bg'])
            dialog.transient(manage_window)
            dialog.grab_set()

            tk.Label(dialog, text="Add New Item", font=("Arial", 18, "bold"),
                     bg=self.colors['bg']).pack(pady=20)

            fields = {}
            entries = [
                ("Name:", "name"), ("Price:", "price"),
                ("Low Stock Threshold:", "threshold"), ("Barcode (optional):", "barcode")
            ]
            for label_text, key in entries:
                frame = tk.Frame(dialog, bg=self.colors['bg'])
                frame.pack(fill="x", padx=40, pady=8)
                tk.Label(frame, text=label_text, font=("Arial", 14),
                         bg=self.colors['bg'], width=20, anchor="w").pack(side="left")
                fields[key] = tk.Entry(frame, font=("Arial", 14), width=20)
                fields[key].pack(side="left", padx=10)

            cat_frame = tk.Frame(dialog, bg=self.colors['bg'])
            cat_frame.pack(fill="x", padx=40, pady=8)
            tk.Label(cat_frame, text="Category:", font=("Arial", 14),
                     bg=self.colors['bg'], width=20, anchor="w").pack(side="left")
            fields['category'] = ttk.Combobox(cat_frame, values=self.categories,
                                               font=("Arial", 14), width=18, state="readonly")
            if self.categories:
                fields['category'].set(self.categories[0])
            fields['category'].pack(side="left", padx=10)

            ts_var = tk.IntVar(value=0)
            ts_frame = tk.Frame(dialog, bg=self.colors['bg'])
            ts_frame.pack(fill="x", padx=40, pady=6)
            tk.Label(ts_frame, text="Track Stock:", font=("Arial", 14),
                     bg=self.colors['bg'], width=20, anchor="w").pack(side="left")
            tk.Checkbutton(ts_frame,
                           text="Track Quantity",
                           variable=ts_var, font=("Arial", 12),
                           bg=self.colors['bg'], fg=self.colors['text_sub'],
                           selectcolor="white",
                           activebackground=self.colors['bg']).pack(side="left")

            def save_item():
                try:
                    name_raw = fields['name'].get().strip()
                    price_raw = fields['price'].get().strip()
                    threshold_raw = fields['threshold'].get().strip()
                    
                    if not name_raw:
                        raise ValueError("Name is required")
                    
                    try:
                        price = float(price_raw)
                    except ValueError:
                        raise ValueError("Price must be a number (e.g. 15 or 15.50)")
                        
                    try:
                        threshold = int(threshold_raw)
                    except ValueError:
                        raise ValueError("Threshold must be a whole number (e.g. 5)")

                    category = fields['category'].get()
                    barcode  = fields['barcode'].get().strip() or None
                    track    = ts_var.get()
                    
                    if threshold < 0:
                        raise ValueError("Threshold cannot be negative")
                        
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("""INSERT INTO items (name, price, category, stock, low_stock_threshold, track_stock, barcode, created_at)
                                    VALUES (?,?,?,?,?,?,?,?)""",
                                  (name_raw, price, category, 0, threshold, track, barcode,
                                   datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    messagebox.showinfo(L("Success", "成功"), L(f"Item '{name_raw}' added!", f"已新增商品“{name_raw}”。"), parent=dialog)
                    dialog.destroy()
                    load_all_items()
                    load_stock_items()
                    self.load_items()
                except ValueError as e:
                    messagebox.showerror(L("Error", "错误"), str(e), parent=dialog)
                except sqlite3.IntegrityError:
                    messagebox.showerror(L("Error", "错误"), L("Item name already exists!", "商品名称已经存在。"), parent=dialog)
                except Exception as e:
                    messagebox.showerror(L("Error", "错误"), L(f"Unexpected error: {str(e)}", f"发生未预期错误：{str(e)}"), parent=dialog)

            btn_frame = tk.Frame(dialog, bg=self.colors['bg'])
            btn_frame.pack(pady=20)
            RoundedButton(btn_frame, text="💾 Save Item", command=save_item,
                          bg=self.colors['success'], fg="white",
                          font=("Arial", 13, "bold"), padx=30, pady=10, radius=8
                          ).pack(side="left", padx=5)
            RoundedButton(btn_frame, text="✖ Cancel", command=dialog.destroy,
                          bg=self.colors['danger'], fg="white",
                          font=("Arial", 13, "bold"), padx=30, pady=10, radius=8
                          ).pack(side="left", padx=5)

            fields['name'].focus_set()
            dialog.bind("<Return>", lambda e: save_item())
            dialog.bind("<Escape>", lambda e: dialog.destroy())

        def edit_item():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select an item to edit", "请选择要编辑的商品。"), parent=manage_window)
                return

            item = tree.item(selection[0])
            item_id = item['values'][0]

            dialog = tk.Toplevel(manage_window)
            dialog.title("Edit Item")
            dialog.geometry("680x740")
            dialog.configure(bg=self.colors['bg'])
            dialog.transient(manage_window)
            dialog.grab_set()

            tk.Label(dialog, text="Edit Item", font=("Arial", 18, "bold"),
                     bg=self.colors['bg']).pack(pady=20)

            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM items WHERE id=?", (item_id,))
                current = c.fetchone()

            if current is None:
                messagebox.showerror(L("Missing Item", "找不到商品"), L("Item not found", "找不到该商品。"), parent=manage_window)
                return

            fields = {}
            entries = [
                ("Name:", "name", current['name']),
                ("Price:", "price", current['price']),
                ("Low Stock Threshold:", "threshold", current['low_stock_threshold']),
                ("Barcode:", "barcode", current['barcode'] or "")
            ]
            for label_text, key, default in entries:
                frame = tk.Frame(dialog, bg=self.colors['bg'])
                frame.pack(fill="x", padx=40, pady=8)
                tk.Label(frame, text=label_text, font=("Arial", 14),
                         bg=self.colors['bg'], width=20, anchor="w").pack(side="left")
                fields[key] = tk.Entry(frame, font=("Arial", 14), width=20)
                fields[key].insert(0, str(default))
                fields[key].pack(side="left", padx=10)

            cat_frame = tk.Frame(dialog, bg=self.colors['bg'])
            cat_frame.pack(fill="x", padx=40, pady=8)
            tk.Label(cat_frame, text="Category:", font=("Arial", 14),
                     bg=self.colors['bg'], width=20, anchor="w").pack(side="left")
            fields['category'] = ttk.Combobox(cat_frame, values=self.categories,
                                               font=("Arial", 14), width=18, state="readonly")
            curr_cat = current['category']
            if curr_cat not in self.categories:
                # If current category was deactivated or deleted, add it temporarily to show it
                fields['category']['values'] = list(self.categories) + [curr_cat]
            fields['category'].set(curr_cat)
            fields['category'].pack(side="left", padx=10)

            ts_edit_var = tk.IntVar(value=int(current['track_stock'] or 0))
            ts_edit_frame = tk.Frame(dialog, bg=self.colors['bg'])
            ts_edit_frame.pack(fill="x", padx=40, pady=6)
            tk.Label(ts_edit_frame, text="Track Stock:", font=("Arial", 14),
                     bg=self.colors['bg'], width=20, anchor="w").pack(side="left")
            tk.Checkbutton(ts_edit_frame,
                           text="Track Quantity",
                           variable=ts_edit_var, font=("Arial", 12),
                           bg=self.colors['bg'], fg=self.colors['text_sub'],
                           selectcolor="white",
                           activebackground=self.colors['bg']).pack(side="left")

            def update_item():
                try:
                    name_raw = fields['name'].get().strip()
                    price_raw = fields['price'].get().strip()
                    threshold_raw = fields['threshold'].get().strip()
                    
                    if not name_raw:
                        raise ValueError("Name is required")
                        
                    try:
                        price = float(price_raw)
                    except ValueError:
                        raise ValueError("Price must be a number (e.g. 15 or 15.50)")
                        
                    try:
                        threshold = int(threshold_raw)
                    except ValueError:
                        raise ValueError("Threshold must be a whole number (e.g. 5)")

                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("""UPDATE items SET name=?, price=?, category=?,
                                    low_stock_threshold=?, track_stock=?, barcode=? WHERE id=?""",
                                  (name_raw, price, fields['category'].get(),
                                   threshold, ts_edit_var.get(),
                                   fields['barcode'].get().strip() or None, item_id))
                    messagebox.showinfo(L("Success", "成功"), L("Item updated!", "商品已更新。"), parent=dialog)
                    dialog.destroy()
                    load_all_items()
                    load_stock_items()
                    self.load_items()
                except ValueError as e:
                    messagebox.showerror(L("Error", "错误"), str(e), parent=dialog)
                except Exception as e:
                    messagebox.showerror(L("Error", "错误"), L(f"Unexpected error: {str(e)}", f"发生未预期错误：{str(e)}"), parent=dialog)

            btn_frame = tk.Frame(dialog, bg=self.colors['bg'])
            btn_frame.pack(pady=20)
            RoundedButton(btn_frame, text="💾 Update Item", command=update_item,
                          bg=self.colors['success'], fg="white",
                          font=("Arial", 13, "bold"), padx=30, pady=10, radius=8
                          ).pack(side="left", padx=5)
            RoundedButton(btn_frame, text="✖ Cancel", command=dialog.destroy,
                          bg=self.colors['danger'], fg="white",
                          font=("Arial", 13, "bold"), padx=30, pady=10, radius=8
                          ).pack(side="left", padx=5)

            fields['name'].focus_set()
            dialog.bind("<Return>", lambda e: update_item())
            dialog.bind("<Escape>", lambda e: dialog.destroy())

        def delete_item():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select an item to delete", "请选择要删除的商品。"), parent=manage_window)
                return
            item = tree.item(selection[0])
            item_name = item['values'][1]
            if not messagebox.askyesno(L("Confirm Delete", "确认删除"),
                                       L(f"Are you sure you want to delete '{item_name}'?\n\nThis will deactivate the item.",
                                         f"确定要删除“{item_name}”吗？\n\n商品将被停用。"),
                                       parent=manage_window):
                return
            item_id = item['values'][0]
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("UPDATE items SET active=0 WHERE id=?", (item_id,))
            messagebox.showinfo(L("Success", "成功"), L("Item deactivated!", "商品已停用。"), parent=manage_window)
            load_all_items()
            load_stock_items()
            self.load_items()

        def manage_variants():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select an item first", "请先选择一个商品。"), parent=manage_window)
                return
            item_data = tree.item(selection[0])['values']
            item_id   = item_data[0]
            item_name = item_data[1]
            base_price = float(item_data[2])

            vwin = tk.Toplevel(manage_window)
            vwin.title(f"Variants — {item_name}")
            vwin.geometry("640x640")
            vwin.configure(bg=self.colors['bg'])
            vwin.transient(manage_window)
            vwin.grab_set()

            tk.Label(vwin, text=f"🌡️ Variants for: {item_name}",
                     font=("Arial", 16, "bold"), bg=self.colors['bg']).pack(pady=12)
            tk.Label(vwin, text=f"Base price: {self.currency} {base_price:.2f}  |  Variant price = Base + Adjustment",
                     font=("Arial", 12), bg=self.colors['bg'], fg="gray").pack()

            vtree_frame = tk.Frame(vwin, bg=self.colors['bg'])
            vtree_frame.pack(fill="both", expand=True, padx=20, pady=10)
            vcols = ("ID", "Variant Name", "Price Adjust", "Final Price")
            vtree = ttk.Treeview(vtree_frame, columns=vcols, show="headings", height=8, style="Manage.Treeview")
            for col in vcols:
                vtree.heading(col, text=col)
                vtree.column(col, width=110, anchor="center")
            vtree.column("Variant Name", width=180)
            vtree.pack(fill="both", expand=True)

            def load_variants():
                vtree.delete(*vtree.get_children())
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("SELECT id, variant_name, price_adjustment FROM item_variants WHERE item_id=? ORDER BY id", (item_id,))
                    for v in c.fetchall():
                        final = base_price + float(v['price_adjustment'])
                        sign  = "+" if float(v['price_adjustment']) >= 0 else ""
                        vtree.insert("", "end", values=(
                            v['id'], v['variant_name'],
                            f"{sign}{float(v['price_adjustment']):.2f}",
                            f"{self.currency} {final:.2f}"
                        ))

            add_frame = tk.Frame(vwin, bg=self.colors['bg'])
            add_frame.pack(fill="x", padx=20, pady=6)

            tk.Label(add_frame, text="Name:", font=("Arial", 13), bg=self.colors['bg']).pack(side="left")
            vname_var = tk.StringVar()
            tk.Entry(add_frame, textvariable=vname_var, font=("Arial", 13), width=12).pack(side="left", padx=6)

            tk.Label(add_frame, text="Adjust (RM):", font=("Arial", 13), bg=self.colors['bg']).pack(side="left")
            vadj_var = tk.StringVar(value="0")
            tk.Entry(add_frame, textvariable=vadj_var, font=("Arial", 13), width=7).pack(side="left", padx=6)

            tk.Label(add_frame, text="e.g. HOT=0, COLD=+0.50", font=("Arial", 9),
                     bg=self.colors['bg'], fg="gray").pack(side="left", padx=4)

            def add_variant():
                name = vname_var.get().strip()
                if not name:
                    messagebox.showerror(L("Error", "错误"), L("Variant name required", "请输入选项名称。"), parent=vwin); return
                try:
                    adj = float(vadj_var.get())
                except ValueError:
                    messagebox.showerror(L("Error", "错误"), L("Adjustment must be a number", "价格调整必须是数字。"), parent=vwin); return
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("INSERT INTO item_variants (item_id, variant_name, price_adjustment) VALUES (?,?,?)",
                              (item_id, name, adj))
                vname_var.set(""); vadj_var.set("0")
                load_variants(); load_stock_items(); self.load_items()

            def delete_variant():
                sel = vtree.selection()
                if not sel:
                    messagebox.showwarning(L("No Selection", "未选择"), L("Select a variant to delete", "请选择要删除的选项。"), parent=vwin); return
                vid = vtree.item(sel[0])['values'][0]
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("DELETE FROM item_variants WHERE id=?", (vid,))
                load_variants(); load_stock_items(); self.load_items()

            vbtn_frame = tk.Frame(vwin, bg=self.colors['bg'])
            vbtn_frame.pack(fill="x", padx=20, pady=6)
            RoundedButton(vbtn_frame, text="➕ Add Variant", command=add_variant,
                          bg=self.colors['success'], fg="white",
                          font=("Arial", 11, "bold"), padx=14, pady=6, radius=8
                          ).pack(side="left", padx=4)
            RoundedButton(vbtn_frame, text="🗑️ Delete Selected", command=delete_variant,
                          bg=self.colors['danger'], fg="white",
                          font=("Arial", 11, "bold"), padx=14, pady=6, radius=8
                          ).pack(side="left", padx=4)

            load_variants()

        # Buttons were moved up to pack side="bottom" before tree_frame expands.

        # ----- Stock Status Tab (tracked items only) -----
        stock_top = tk.Frame(stock_tab, bg=self.colors['bg'])
        stock_top.pack(fill="x", padx=20, pady=(14, 8))

        tk.Label(stock_top, text=L("Tracked Stock Items Only", "仅显示追踪库存商品"), font=("Arial", 16, "bold"),
                 bg=self.colors['bg'], fg=self.colors['text']).pack(side="left")

        stock_search_var = tk.StringVar()
        stock_result_var = tk.StringVar()
        stock_query_host = tk.Frame(stock_tab, bg=self.colors['bg'])
        stock_query_host.pack(fill="x", padx=4, pady=(0, 2))

        stock_tree_wrap = tk.Frame(stock_tab, bg=self.colors['bg'])
        stock_tree_wrap.pack(fill="both", expand=True, padx=20, pady=8)

        s_cols = ("ID", "Name", "Category", "Stock", "Reorder At", "Status")
        stock_tree = ttk.Treeview(stock_tree_wrap, columns=s_cols, show="headings", height=14, style="Manage.Treeview")
        s_w = {"ID": 60, "Name": 280, "Category": 130, "Stock": 110, "Reorder At": 110, "Status": 180}
        for col in s_cols:
            stock_tree.heading(col, text={
                "ID": "ID", "Name": L("Name", "名称"),
                "Category": L("Category", "类别"), "Stock": L("Stock", "库存"),
                "Reorder At": L("Reorder At", "补货警戒"),
                "Status": L("Status", "状态"),
            }[col])
            stock_tree.column(col, width=s_w.get(col, 120), anchor="center")
        stock_tree.column("Name", anchor="w")
        stock_tree.tag_configure('low', foreground='#D97706')
        stock_tree.tag_configure('out', foreground='#DC2626')

        s_sb = tk.Scrollbar(stock_tree_wrap, orient="vertical", command=stock_tree.yview)
        stock_tree.configure(yscrollcommand=s_sb.set)
        stock_tree.pack(side="left", fill="both", expand=True)
        s_sb.pack(side="right", fill="y")

        adjust_box = tk.LabelFrame(
            stock_document_tab, text="  " + L("Stock Document", "库存单据") + "  ",
            font=("Arial", 13, "bold"), bg=self.colors['bg'],
            fg=self.colors['text'], padx=10, pady=8
        )
        adjust_box.pack(fill="both", expand=True, padx=20, pady=16)

        stock_actions = {
            L("Stock In", "进货入库"): "STOCK_IN",
            L("Internal Use", "店内自用"): "INTERNAL_USE",
            L("Waste / Damaged", "损耗／损坏"): "WASTE",
            L("Correction Add", "盘点调增"): "CORRECTION_ADD",
            L("Correction Deduct", "盘点调减"): "CORRECTION_DEDUCT",
        }
        stock_action_var = tk.StringVar(value=next(iter(stock_actions)))
        stock_note_var = tk.StringVar()

        adjustment_meta = tk.Frame(adjust_box, bg=self.colors['bg'])
        adjustment_meta.pack(fill="x", pady=(0, 4))
        tk.Label(
            adjustment_meta, text=L("Action:", "动作："),
            font=("Arial", 10, "bold"), bg=self.colors['bg']
        ).pack(side="left")
        ttk.Combobox(
            adjustment_meta, textvariable=stock_action_var, values=list(stock_actions),
            state="readonly", font=("Arial", 10), width=20
        ).pack(side="left", padx=(6, 18))
        tk.Label(
            adjustment_meta, text=L("Invoice / Note:", "进货单号／备注："),
            font=("Arial", 10, "bold"), bg=self.colors['bg']
        ).pack(side="left")
        tk.Entry(
            adjustment_meta, textvariable=stock_note_var, font=("Arial", 10)
        ).pack(side="left", fill="x", expand=True, padx=(6, 0))

        tk.Label(
            adjust_box,
            text=L(
                "One document can contain several tracked items. Each item may appear only once; the whole document saves together.",
                "一张单据可包含多个追踪库存商品；同一商品只能出现一次，整张单据会一次保存。"
            ),
            font=("Arial", 9), bg=self.colors['bg'], fg=self.colors['text_sub'],
            anchor="w"
        ).pack(fill="x", pady=(0, 4))

        stock_line_header = tk.Frame(adjust_box, bg=self.colors['panel2'])
        stock_line_header.pack(fill="x")
        stock_line_header.grid_columnconfigure(1, weight=1)
        tk.Label(stock_line_header, text="#", width=4, font=("Arial", 9, "bold"),
                 bg=self.colors['panel2']).grid(row=0, column=0, padx=3, pady=3)
        tk.Label(stock_line_header, text=L("Tracked Item", "追踪库存商品"),
                 font=("Arial", 9, "bold"), bg=self.colors['panel2'], anchor="w").grid(
                     row=0, column=1, padx=5, pady=3, sticky="ew")
        tk.Label(stock_line_header, text=L("Current", "现有库存"), width=16,
                 font=("Arial", 9, "bold"), bg=self.colors['panel2']).grid(
                     row=0, column=2, padx=5, pady=3)
        tk.Label(stock_line_header, text=L("Quantity", "数量"), width=10,
                 font=("Arial", 9, "bold"), bg=self.colors['panel2']).grid(
                     row=0, column=3, padx=5, pady=3)
        tk.Label(stock_line_header, text="", width=6,
                 bg=self.colors['panel2']).grid(row=0, column=4, padx=3)

        stock_lines_view = tk.Frame(adjust_box, bg=self.colors['bg'])
        stock_lines_view.pack(fill="both", expand=True)
        stock_lines_canvas = tk.Canvas(
            stock_lines_view, height=320, bg=self.colors['bg'],
            highlightthickness=0, bd=0
        )
        stock_lines_scroll = tk.Scrollbar(
            stock_lines_view, orient="vertical", command=stock_lines_canvas.yview
        )
        stock_lines_canvas.configure(yscrollcommand=stock_lines_scroll.set)
        stock_lines_canvas.pack(side="left", fill="both", expand=True)
        stock_lines_scroll.pack(side="right", fill="y")
        stock_lines_host = tk.Frame(stock_lines_canvas, bg=self.colors['bg'])
        stock_lines_window = stock_lines_canvas.create_window(
            (0, 0), window=stock_lines_host, anchor="nw"
        )
        stock_lines_host.bind(
            "<Configure>",
            lambda _event: stock_lines_canvas.configure(
                scrollregion=stock_lines_canvas.bbox("all")
            )
        )
        stock_lines_canvas.bind(
            "<Configure>",
            lambda event: stock_lines_canvas.itemconfigure(
                stock_lines_window, width=event.width
            )
        )

        stock_adjustment_rows = []
        stock_choice_map = {}
        stock_display_by_id = {}
        stock_document_token = [uuid.uuid4().hex]

        def refresh_stock_row_choices():
            used_ids = {
                line["item_id"] for line in stock_adjustment_rows
                if line["item_id"] is not None
            }
            for line in stock_adjustment_rows:
                own_id = line["item_id"]
                values = [
                    display for display, item in stock_choice_map.items()
                    if item["id"] not in used_ids or item["id"] == own_id
                ]
                line["combo"].configure(values=values)
                if own_id in stock_display_by_id:
                    line["item_var"].set(stock_display_by_id[own_id])

        def refresh_stock_choice_map():
            selected_ids = [line["item_id"] for line in stock_adjustment_rows]
            stock_choice_map.clear()
            stock_display_by_id.clear()
            with self.db.get_connection() as conn:
                rows = conn.execute(
                    """SELECT id,name,category,stock FROM items
                       WHERE active=1 AND track_stock=1 ORDER BY category,name"""
                ).fetchall()
            for row in rows:
                display = f"{row['name']}  ·  {row['category']}"
                if display in stock_choice_map:
                    display = f"{display}  [#{row['id']}]"
                item = {
                    "id": int(row['id']), "name": str(row['name']),
                    "category": str(row['category']), "stock": int(row['stock'] or 0),
                }
                stock_choice_map[display] = item
                stock_display_by_id[item["id"]] = display
            for line, selected_id in zip(stock_adjustment_rows, selected_ids):
                line["item_id"] = selected_id if selected_id in stock_display_by_id else None
                if line["item_id"] is None:
                    line["item_var"].set("")
                    line["current_label"].config(text="-")
                else:
                    item = stock_choice_map[stock_display_by_id[selected_id]]
                    line["current_label"].config(text=str(item["stock"]))
            refresh_stock_row_choices()

        def on_adjustment_item_selected(line):
            item = stock_choice_map.get(line["item_var"].get())
            if not item:
                line["item_id"] = None
                line["current_label"].config(text="-")
                return
            if any(
                other is not line and other["item_id"] == item["id"]
                for other in stock_adjustment_rows
            ):
                messagebox.showerror(
                    L("Duplicate Item", "商品重复"),
                    L(
                        "This item is already in the document. Change its quantity on the existing row.",
                        "这个商品已经在单据内，请直接修改原有那一排的数量。"
                    ),
                    parent=manage_window
                )
                line["item_id"] = None
                line["item_var"].set("")
                line["current_label"].config(text="-")
                refresh_stock_row_choices()
                return
            line["item_id"] = item["id"]
            line["current_label"].config(text=str(item["stock"]))
            refresh_stock_row_choices()

        def renumber_stock_rows():
            for index, line in enumerate(stock_adjustment_rows, 1):
                line["number_label"].config(text=str(index))

        def remove_stock_adjustment_row(line):
            if len(stock_adjustment_rows) == 1:
                line["item_id"] = None
                line["item_var"].set("")
                line["qty_var"].set("1")
                line["current_label"].config(text="-")
            else:
                line["frame"].destroy()
                stock_adjustment_rows.remove(line)
            renumber_stock_rows()
            refresh_stock_row_choices()

        def _positive_row_quantity(line):
            try:
                quantity = int(line["qty_var"].get().strip())
                return quantity if quantity > 0 else None
            except (TypeError, ValueError):
                return None

        def add_stock_adjustment_row(require_previous=True):
            if require_previous and stock_adjustment_rows:
                previous = stock_adjustment_rows[-1]
                if previous["item_id"] is None or _positive_row_quantity(previous) is None:
                    messagebox.showwarning(
                        L("Complete Current Row", "请先完成当前一排"),
                        L(
                            "Choose an item and enter a positive whole quantity before adding another row.",
                            "请先选择商品并填写正整数数量，才新增下一排。"
                        ),
                        parent=manage_window
                    )
                    return

            row_frame = tk.Frame(stock_lines_host, bg=self.colors['bg'])
            row_frame.pack(fill="x", pady=2)
            row_frame.grid_columnconfigure(1, weight=1)
            item_var = tk.StringVar()
            qty_var = tk.StringVar(value="1")
            number_label = tk.Label(
                row_frame, text=str(len(stock_adjustment_rows) + 1), width=4,
                font=("Arial", 10, "bold"), bg=self.colors['bg']
            )
            number_label.grid(row=0, column=0, padx=3, pady=2)
            combo = ttk.Combobox(
                row_frame, textvariable=item_var, state="readonly", font=("Arial", 10)
            )
            combo.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
            current_label = tk.Label(
                row_frame, text="-", width=16, font=("Arial", 10),
                bg=self.colors['bg'], fg=self.colors['text_sub']
            )
            current_label.grid(row=0, column=2, padx=5, pady=2)
            qty_entry = tk.Entry(
                row_frame, textvariable=qty_var, width=10,
                font=("Arial", 11), justify="center"
            )
            qty_entry.grid(row=0, column=3, padx=5, pady=2)
            line = {
                "frame": row_frame, "number_label": number_label,
                "item_var": item_var, "qty_var": qty_var, "combo": combo,
                "current_label": current_label, "item_id": None,
            }
            tk.Button(
                row_frame, text="✕", width=4,
                command=lambda current=line: remove_stock_adjustment_row(current),
                bg="#FEE2E2", fg=self.colors['danger'], relief="flat",
                font=("Arial", 10, "bold"), cursor="hand2"
            ).grid(row=0, column=4, padx=5, pady=2)
            combo.bind(
                "<<ComboboxSelected>>",
                lambda _event, current=line: on_adjustment_item_selected(current)
            )
            stock_adjustment_rows.append(line)
            refresh_stock_row_choices()
            stock_lines_host.update_idletasks()
            stock_lines_canvas.yview_moveto(1.0)

        def reset_stock_adjustment_rows():
            for line in list(stock_adjustment_rows):
                line["frame"].destroy()
            stock_adjustment_rows.clear()
            add_stock_adjustment_row(require_previous=False)

        refresh_stock_choice_map()
        add_stock_adjustment_row(require_previous=False)

        def load_stock_items():
            stock_tree.delete(*stock_tree.get_children())
            keyword = stock_search_var.get().strip()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                if keyword:
                    c.execute("""
                        SELECT id, name, category, stock, low_stock_threshold
                        FROM items
                        WHERE active=1 AND track_stock=1 AND name LIKE ?
                        ORDER BY category, name
                    """, (f"%{keyword}%",))
                else:
                    c.execute("""
                        SELECT id, name, category, stock, low_stock_threshold
                        FROM items
                        WHERE active=1 AND track_stock=1
                        ORDER BY category, name
                    """)
                rows = c.fetchall()

            for row in rows:
                tag = ''
                status = L("✅ In Stock", "✅ 有库存")
                if int(row['stock']) <= 0:
                    status = L("❌ Out of Stock", "❌ 缺货"); tag = 'out'
                elif int(row['stock']) <= int(row['low_stock_threshold'] or 0):
                    status = L("⚠️ Low Stock", "⚠️ 低库存"); tag = 'low'
                stock_tree.insert("", "end", iid=str(row['id']), values=(
                    row['id'], row['name'], row['category'], row['stock'],
                    row['low_stock_threshold'], status
                ), tags=(tag,) if tag else ())
            stock_result_var.set(L(f"{len(rows)} item(s)", f"{len(rows)} 个商品"))
            refresh_stock_choice_map()

        def apply_stock_adjustment():
            action = stock_action_var.get()
            source_type = stock_actions.get(action)
            if not source_type:
                messagebox.showerror(
                    L("Missing Action", "缺少动作"),
                    L("Choose a stock action.", "请选择库存动作。"),
                    parent=manage_window
                )
                return

            entries = []
            seen_item_ids = set()
            for index, line in enumerate(stock_adjustment_rows, 1):
                item_id = line["item_id"]
                quantity = _positive_row_quantity(line)
                if item_id is None:
                    messagebox.showwarning(
                        L("Missing Item", "缺少商品"),
                        L(
                            f"Choose a tracked item on row {index}.",
                            f"请选择第 {index} 排的追踪库存商品。"
                        ),
                        parent=manage_window
                    )
                    return
                if quantity is None:
                    messagebox.showerror(
                        L("Invalid Quantity", "数量无效"),
                        L(
                            f"Row {index} quantity must be a positive whole number.",
                            f"第 {index} 排数量必须是正整数。"
                        ),
                        parent=manage_window
                    )
                    return
                if item_id in seen_item_ids:
                    messagebox.showerror(
                        L("Duplicate Item", "商品重复"),
                        L(
                            "The same item cannot appear twice in one document.",
                            "同一商品不能在一张单据中出现两次。"
                        ),
                        parent=manage_window
                    )
                    return
                seen_item_ids.add(item_id)
                entries.append({"item_id": item_id, "quantity": quantity})

            note = stock_note_var.get().strip()
            submission_token = stock_document_token[0]
            captured_entries = [dict(entry) for entry in entries]

            def do_adjustment(staff=None):
                actor = staff or self.current_staff or {"name": "Admin"}
                try:
                    result = self._apply_stock_adjustment_batch(
                        captured_entries, source_type, actor,
                        notes=note, request_token=submission_token
                    )
                except Exception as exc:
                    messagebox.showerror(
                        L("Unable to Save", "无法保存"), str(exc), parent=manage_window
                    )
                    return

                load_stock_items()
                load_all_items()
                self.load_items()
                if result["duplicate"]:
                    messagebox.showwarning(
                        L("Already Saved", "已经保存"),
                        L(
                            f"Stock document #{result['batch_id']} was already saved. Inventory was not added again.",
                            f"库存单据 #{result['batch_id']} 已经保存；系统没有再次增加库存。"
                        ),
                        parent=manage_window
                    )
                    return

                stock_document_token[0] = uuid.uuid4().hex
                stock_note_var.set("")
                reset_stock_adjustment_rows()
                refresh_stock_choice_map()
                detail_lines = "\n".join(
                    f"{row['item_name']}: {row['before']} → {row['after']}"
                    for row in result["lines"][:8]
                )
                if len(result["lines"]) > 8:
                    detail_lines += L("\n…more items", "\n……还有其他商品")
                messagebox.showinfo(
                    L("Stock Document Saved", "库存单据已保存"),
                    L(
                        f"Document #{result['batch_id']}\n{result['item_count']} item(s), total quantity {result['total_quantity']}\n\n{detail_lines}",
                        f"单据 #{result['batch_id']}\n{result['item_count']} 项商品，总数量 {result['total_quantity']}\n\n{detail_lines}"
                    ),
                    parent=manage_window
                )

            if not self.admin_mode:
                self._staff_login_popup(
                    callback=do_adjustment,
                    title=self.l("🔑 Admin PIN Required", "🔑 需要管理员登录"),
                    parent=manage_window, admin_only=True
                )
            else:
                do_adjustment()

        stock_btns = tk.Frame(adjust_box, bg=self.colors['bg'])
        stock_btns.pack(fill="x", padx=0, pady=(6, 0))
        RoundedButton(
            stock_btns, text="＋ " + L("Add Another Item", "增加下一样商品"),
            command=add_stock_adjustment_row,
            bg=self.colors['primary'], fg="white",
            font=("Arial", 11, "bold"), padx=16, pady=8, radius=8
        ).pack(side="left", padx=4)
        RoundedButton(
            stock_btns, text=L("Clear Form", "清空表格"),
            command=reset_stock_adjustment_rows,
            bg=self.colors['neutral'], fg="white",
            font=("Arial", 10, "bold"), padx=14, pady=8, radius=8
        ).pack(side="left", padx=4)
        RoundedButton(stock_btns, text="💾 " + L("Save Stock Document", "保存库存单据"), command=apply_stock_adjustment,
                      bg=self.colors['success'], fg="white",
                      font=("Arial", 11, "bold"), padx=16, pady=8, radius=8
                      ).pack(side="right", padx=4)
        self._build_query_toolbar(
            item_query_host, item_search_var, load_all_items,
            on_refresh=load_all_items,
            on_export=lambda: self._export_treeview_csv(
                tree, manage_window, "items.csv"
            ),
            result_var=item_result_var,
            search_label=L("Search items:", "搜索商品：")
        )
        self._build_query_toolbar(
            stock_query_host, stock_search_var, load_stock_items,
            on_refresh=load_stock_items,
            on_export=lambda: self._export_treeview_csv(
                stock_tree, manage_window, "stock_items.csv"
            ),
            result_var=stock_result_var,
            search_label=L("Search stock:", "搜索库存：")
        )

        load_all_items()
        load_stock_items()

        if not stock_only:
            # Side buttons
            for _t, _c, _bg in [
                ("➕ Add Item",  add_item_dialog, self.colors['success']),
                ("📝 Edit Item", edit_item,       self.colors['warning']),
                ("🗑️ Delete",    lambda: delete_item() if self.admin_mode else self._staff_login_popup(callback=lambda staff: delete_item(), title=self.l("🔑 Admin PIN Required", "🔑 需要管理员登录"), parent=manage_window, admin_only=True), self.colors['danger']),
                ("🌡️ Variants",  manage_variants,  self.colors['primary']),
                ("📦 Sets",      manage_quantity_sets, self.colors['info']),
            ]:
                RoundedButton(sidebar, text=_t, command=_c,
                              bg=_bg, fg="white",
                              font=("Arial", 11, "bold"),
                              padx=20, pady=10, radius=8
                              ).pack(fill="x", padx=14, pady=6)

    # ================= QUICK STORE BEVERAGE (IMPROVED) =================
    # ── Cashier quick: change payment of last order ─────────────────────────
    def show_recent_orders(self):
        """Cashier panel: recent orders with inline payment toggle, details, reprint, edit."""
        C = self.colors
        L = (lambda en, zh: zh if self.lang == "zh" else en)
        win = tk.Toplevel(self.root)
        win.title(L("Recent Orders", "近期订单"))
        win.geometry("1280x760")
        win.minsize(1180, 680)
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()

        # ── Header ────────────────────────────────────────────────────────
        hdr = tk.Frame(win, bg=C['primary'], height=56)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text=L("📋  Recent Orders", "📋  近期订单"),
                 font=("Arial", 17, "bold"), bg=C['primary'], fg="white"
                 ).pack(side="left", padx=20, pady=10)

        filter_var = tk.StringVar(value="today")
        date_filter_var = tk.StringVar(value=today_biz_date())
        item_search_var = tk.StringVar()
        result_summary_var = tk.StringVar()

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT name FROM items ORDER BY name")
            item_search_values = [str(r['name']) for r in c.fetchall() if r['name']]

        filter_panel = tk.Frame(
            win, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1
        )
        filter_panel.pack(fill="x", padx=12, pady=(8, 4))

        period_row = tk.Frame(filter_panel, bg=C['panel'])
        period_row.pack(fill="x", padx=10, pady=(8, 3))
        tk.Label(period_row, text=L("Period:", "时段："), font=("Arial", 10, "bold"),
                 bg=C['panel'], fg=C['text']).pack(side="left", padx=(0, 4))
        for label, val in [
            (L("1 Hour", "1小时"), "1"),
            (L("3 Hours", "3小时"), "3"),
            (L("Today", "今天"), "today"),
            (L("Previous Day", "前一天"), "yesterday"),
            (L("7 Days", "7天"), "7days"),
        ]:
            tk.Radiobutton(
                period_row, text=label, variable=filter_var, value=val,
                font=("Arial", 10), bg=C['panel'], fg=C['text'],
                selectcolor=C['panel'], activebackground=C['panel'],
                command=lambda: load_orders()
            ).pack(side="left", padx=6)

        lookup_row = tk.Frame(filter_panel, bg=C['panel'])
        lookup_row.pack(fill="x", padx=10, pady=(3, 8))
        tk.Label(lookup_row, text=L("Business Date:", "营业日期："),
                 font=("Arial", 10, "bold"), bg=C['panel'], fg=C['text']).pack(side="left")
        DatePicker(
            lookup_row, textvariable=date_filter_var, width=11,
            font=("Arial", 10), language=self.lang
        ).pack(side="left", padx=(5, 4))

        def apply_chosen_date():
            filter_var.set("date")
            load_orders()

        RoundedButton(
            lookup_row, text=L("View Date", "查看日期"), command=apply_chosen_date,
            bg=C['neutral'], fg="white", font=("Arial", 9, "bold"),
            padx=9, pady=4, radius=6
        ).pack(side="left", padx=(0, 14))

        tk.Label(lookup_row, text=L("Item Name:", "商品名称："),
                 font=("Arial", 10, "bold"), bg=C['panel'], fg=C['text']).pack(side="left")
        item_search_box = ttk.Combobox(
            lookup_row, textvariable=item_search_var, values=item_search_values,
            width=27, font=("Arial", 10)
        )
        item_search_box.pack(side="left", padx=5)
        RoundedButton(
            lookup_row, text=L("🔍 Find Orders", "🔍 查找订单"), command=lambda: load_orders(),
            bg=C['primary'], fg="white", font=("Arial", 9, "bold"),
            padx=10, pady=4, radius=6
        ).pack(side="left", padx=3)

        def clear_item_search():
            item_search_var.set("")
            load_orders()

        RoundedButton(
            lookup_row, text=L("Clear Item", "清除商品"), command=clear_item_search,
            bg=C['neutral'], fg="white", font=("Arial", 9, "bold"),
            padx=9, pady=4, radius=6
        ).pack(side="left", padx=3)
        tk.Label(
            lookup_row, textvariable=result_summary_var, font=("Arial", 9, "bold"),
            bg=C['panel'], fg=C['text_muted']
        ).pack(side="right", padx=6)
        item_search_box.bind("<Return>", lambda _event: load_orders())
        item_search_box.bind("<<ComboboxSelected>>", lambda _event: load_orders())

        # ── Column headers ────────────────────────────────────────────────
        col_hdr = tk.Frame(win, bg=C['neutral'], height=30)
        col_hdr.pack(fill="x", padx=12, pady=(8,0))
        col_hdr.pack_propagate(False)
        col_specs = [
            ("id", "#", 60),
            ("time", L("Time", "时间"), 120),
            ("total", L("Total", "总额"), 80),
            ("pay", L("Payment", "付款"), 75),
            ("cashier", L("Cashier", "收银员"), 90),
            ("match", L("Matching Item", "匹配商品"), 170),
            ("status", L("Status", "状态"), 110),
            ("change", L("Change Payment", "更改付款"), 135),
            ("actions", L("Actions", "操作"), 270),
        ]
        for _key, text, w in col_specs:
            hcell = tk.Frame(col_hdr, bg=C['neutral'], width=w, height=28)
            hcell.pack(side="left", padx=2)
            hcell.pack_propagate(False)
            tk.Label(hcell, text=text, font=("Arial",9,"bold"),
                     bg=C['neutral'], fg="white", anchor="w"
                     ).pack(fill="both", padx=4)

        # ── Scrollable list ───────────────────────────────────────────────
        list_outer = tk.Frame(win, bg=C['bg'])
        list_outer.pack(fill="both", expand=True, padx=12, pady=(4,10))
        canvas = tk.Canvas(list_outer, bg=C['bg'], highlightthickness=0)
        sb     = tk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
        rows_frame = tk.Frame(canvas, bg=C['bg'])
        rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        _cwin = canvas.create_window((0,0), window=rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(_cwin, width=e.width))
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1*(e.delta//120),"units"))
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        PAY_COLORS = {"CASH":"#16A34A","QR":"#0369A1","CARD":"#7C3AED"}

        def _do_reprint(s, parent_win):
            """Build receipt text and send to printer; always show result popup."""
            try:
                table_code = s['table_code'] if ('table_code' in s.keys() and s['table_code']) else '-'
                items = json.loads(s['items'])
                list_prices = json.loads(s['category_list_prices'])
                receipt = self._receipt_item_lines(items, list_prices)
                if not receipt:
                    messagebox.showwarning(L("No Items", "无项目"), L("This order has no printable items.", "该订单没有可打印项目。"), parent=parent_win)
                    return

                subtotal, tax, total, discount_line = self._saved_receipt_totals(s)
                adjustment = self._display_price_adjustment(items, list_prices)
                text = (
                    "\n\n"
                    "================================\n"
                    "         919 Restaurant\n"
                    "        (CT0159652-K)\n"
                    "  NO.1 G2 BLOK D, JALAN INDAH\n"
                    "    12/2 TAMAN INDAH 84000\n"
                    "        MUAR JOHOR\n"
                    "================================\n"
                    f"Date: {s['datetime']}\n"
                    f"Order: #{s['id']:04d} (REPRINT)\n"
                    f"Table: {table_code}\n"
                    f"Cashier: {s['cashier']}\n"
                    "================================\n"
                    f"{receipt}\n"
                    "--------------------------------\n"
                    f"Subtotal:      {self.currency} {subtotal - adjustment:>10.2f}\n"
                    f"{discount_line}"
                    f"Tax:           {self.currency} {tax:>10.2f}\n"
                    "================================\n"
                    f"{self._receipt_total_lines(total, adjustment, paid=True)}"
                    "================================\n"
                    f"Payment: {s['payment']}\n"
                    "================================\n"
                    "    Thank You! Please Come Again!\n"
                    "================================\n"
                    "\n\n\n\x1dV\x00"
                )
                print_to_thermal(PRINTER_NAME, text)
                messagebox.showinfo(L("✅ Reprinted", "✅ 已重打"),
                    L(f"Receipt for Order #{s['id']:04d} sent to printer!", f"订单 #{s['id']:04d} 小票已发送到打印机！"), parent=parent_win)
            except Exception as e:
                messagebox.showerror(L("Print Error", "打印错误"),
                    L(f"Could not print receipt:\n{e}", f"无法打印小票：\n{e}"), parent=parent_win)

        def load_orders():
            for w in rows_frame.winfo_children():
                w.destroy()

            fval = filter_var.get()
            now = datetime.now()
            period_label = ""
            if fval == "today":
                biz = today_biz_date()
                s_dt, e_dt = business_day_range(biz)
                period_label = L("Today", "今天")
            elif fval == "yesterday":
                current_biz = datetime.strptime(today_biz_date(), "%Y-%m-%d")
                biz = (current_biz - timedelta(days=1)).strftime("%Y-%m-%d")
                s_dt, e_dt = business_day_range(biz)
                period_label = L(f"Previous Day ({biz})", f"前一天 ({biz})")
            elif fval == "7days":
                current_biz = datetime.strptime(today_biz_date(), "%Y-%m-%d")
                first_biz = (current_biz - timedelta(days=6)).strftime("%Y-%m-%d")
                s_dt, _ = business_day_range(first_biz)
                _, e_dt = business_day_range(today_biz_date())
                period_label = L("Last 7 Business Days", "最近7个营业日")
            elif fval == "date":
                try:
                    selected_biz = normalize_date_text(date_filter_var.get())
                    date_filter_var.set(selected_biz)
                except ValueError:
                    messagebox.showerror(
                        L("Invalid Date", "日期无效"),
                        L("Please choose a valid business date.", "请选择有效的营业日期。"),
                        parent=win
                    )
                    return
                s_dt, e_dt = business_day_range(selected_biz)
                period_label = L(f"Business Date {selected_biz}", f"营业日期 {selected_biz}")
            else:
                hours = int(fval)
                s_dt = now - timedelta(hours=hours)
                e_dt = now
                period_label = L(f"Last {hours} Hour(s)", f"最近 {hours} 小时")

            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute(
                    """SELECT * FROM sales
                       WHERE COALESCE(amended,0) != 1
                         AND to_iso(datetime) >= ?
                         AND to_iso(datetime) <= ?
                       ORDER BY id DESC""",
                    (
                        s_dt.strftime("%Y-%m-%d %H:%M:%S"),
                        e_dt.strftime("%Y-%m-%d %H:%M:%S")
                    )
                )
                all_rows = c.fetchall()

            item_query = item_search_var.get().strip()
            orders = []
            match_summaries = {}
            matched_units = 0
            for row in all_rows:
                matches = find_sale_item_matches(row['items'], item_query)
                if item_query and not matches:
                    continue
                orders.append(row)
                if matches:
                    summary = ", ".join(f"{name} x{qty}" for name, qty in matches)
                    match_summaries[int(row['id'])] = summary
                    matched_units += sum(qty for _name, qty in matches)

            if item_query:
                result_summary_var.set(
                    L(
                        f"{len(orders)} order(s) · {matched_units} matching unit(s) · {period_label}",
                        f"{len(orders)} 笔订单 · {matched_units} 件匹配商品 · {period_label}"
                    )
                )
            else:
                result_summary_var.set(
                    L(f"{len(orders)} order(s) · {period_label}",
                      f"{len(orders)} 笔订单 · {period_label}")
                )

            refund_map = {}
            if orders:
                sale_ids = [int(r['id']) for r in orders if r['id'] is not None]
                if sale_ids:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        for offset in range(0, len(sale_ids), 500):
                            id_batch = sale_ids[offset:offset + 500]
                            qmarks = ",".join(["?"] * len(id_batch))
                            c.execute(
                                f"SELECT sale_id, COALESCE(SUM(refund_amount),0) AS refunded FROM refunds WHERE sale_id IN ({qmarks}) GROUP BY sale_id",
                                tuple(id_batch)
                            )
                            refund_map.update({
                                int(r['sale_id']): float(r['refunded'] or 0)
                                for r in c.fetchall()
                            })

            if not orders:
                empty_text = (
                    L(f"No orders containing '{item_query}' in this period.",
                      f"此时段没有包含“{item_query}”的订单。")
                    if item_query else
                    L("No orders found for this period.", "此时段没有订单。")
                )
                tk.Label(rows_frame, text=empty_text,
                         font=("Arial",13), bg=C['bg'], fg=C['text_muted']).pack(pady=40)
                return

            has_adjusted_display = False
            for i, sale in enumerate(orders):
                if i > 0 and i % 10 == 0:
                    rows_frame.update_idletasks()

                refunded_amt = float(refund_map.get(int(sale['id']), 0.0))
                sale_total = float(sale['total'] or 0)
                is_refunded = refunded_amt > 0.005
                is_full_refund = sale_total > 0 and refunded_amt + 0.005 >= sale_total
                row_bg = "#FECACA" if is_full_refund else ("#FEE2E2" if is_refunded else ("#F8FAFC" if i % 2 == 0 else C['bg']))
                try:    amended      = int(sale['amended'] or 0)
                except: amended      = 0
                try:    amended_from = sale['amended_from']
                except: amended_from = None

                # Hide old/legacy amended versions of orders
                if amended == 1:
                    continue

                row = tk.Frame(rows_frame, bg=row_bg)
                row.pack(fill="x", padx=2, pady=1)

                # ── Order # ──────────────────────────────────────────────
                id_f = tk.Frame(row, bg=row_bg, width=60, height=42)
                id_f.pack(side="left", padx=2); id_f.pack_propagate(False)
                tk.Label(id_f, text=f"#{sale['id']}", font=("Arial",10,"bold"),
                         bg=row_bg, fg=C['text_muted'] if amended else C['text']).pack(anchor="w")
                if is_full_refund:
                    tk.Label(id_f, text="refunded", font=("Arial",7),
                             bg="#EF4444", fg="white", padx=2).pack(anchor="w")
                elif is_refunded:
                    tk.Label(id_f, text="part-ref", font=("Arial",7),
                             bg="#FECACA", fg="#991B1B", padx=2).pack(anchor="w")
                elif amended == 2 or amended_from:
                    tk.Label(id_f, text="amend", font=("Arial",7),
                             bg="#DCFCE7", fg="#166534", padx=2).pack(anchor="w")

                # ── Time ─────────────────────────────────────────────────
                time_f = tk.Frame(row, bg=row_bg, width=120, height=42)
                time_f.pack(side="left", padx=2); time_f.pack_propagate(False)
                tk.Label(time_f, text=sale['datetime'][:16], font=("Arial",10),
                         bg=row_bg, fg=C['text'], anchor="w").pack(fill="both", padx=4)

                # ── Total ────────────────────────────────────────────────
                total_f = tk.Frame(row, bg=row_bg, width=80, height=42)
                total_f.pack(side="left", padx=2); total_f.pack_propagate(False)
                display_total, star = self._display_sale_total(sale)
                has_adjusted_display |= bool(star)
                tk.Label(total_f, text=f"{self.currency} {display_total:.2f}{star}", font=("Arial",10,"bold"),
                         bg=row_bg, fg=C['danger'] if is_refunded else C['success'], anchor="e").pack(fill="both", padx=4)

                # ── Current payment badge ─────────────────────────────────
                pay = sale['payment'] or "CASH"
                pay_wrap = tk.Frame(row, bg=row_bg, width=75, height=42)
                pay_wrap.pack(side="left", padx=2); pay_wrap.pack_propagate(False)
                pf  = tk.Frame(pay_wrap, bg=PAY_COLORS.get(pay,"#475569"), padx=6, pady=2)
                pf.pack(side="left", padx=4, pady=4)
                tk.Label(pf, text=pay, font=("Arial",9,"bold"),
                         bg=PAY_COLORS.get(pay,"#475569"), fg="white").pack()

                # ── Cashier ───────────────────────────────────────────────
                cashier_f = tk.Frame(row, bg=row_bg, width=90, height=42)
                cashier_f.pack(side="left", padx=2); cashier_f.pack_propagate(False)
                tk.Label(cashier_f, text=(sale['cashier'] or "-")[:10], font=("Arial",9),
                         bg=row_bg, fg=C['text_sub'], anchor="w").pack(fill="both", padx=4)

                # ── Item matched by the search ────────────────────────────
                match_f = tk.Frame(row, bg=row_bg, width=170, height=42)
                match_f.pack(side="left", padx=2); match_f.pack_propagate(False)
                match_text = match_summaries.get(int(sale['id']), "—")
                tk.Label(
                    match_f, text=match_text, font=("Arial",8,"bold"),
                    bg=row_bg, fg=C['primary'] if item_query else C['text_muted'],
                    anchor="w", justify="left", wraplength=164
                ).pack(fill="both", padx=4)

                # ── Status ────────────────────────────────────────────────
                status_f = tk.Frame(row, bg=row_bg, width=110, height=42)
                status_f.pack(side="left", padx=2); status_f.pack_propagate(False)
                status_text = ((f"Refunded {self.currency}{refunded_amt:.2f}") if is_full_refund
                               else (f"Part Refund {self.currency}{refunded_amt:.2f}") if is_refunded
                               else ("Amend" if (amended == 2 or amended_from) else ("Edited" if amended == 1 else "Original")))
                status_fg = (C['danger'] if is_refunded
                             else (C['success'] if (amended == 2 or amended_from) else (C['warning'] if amended == 1 else C['text_muted'])))
                tk.Label(status_f, text=status_text,
                         font=("Arial",9), bg=row_bg,
                         fg=status_fg,
                         anchor="w").pack(fill="both", padx=4)

                # ── Inline payment toggle (rounded buttons) ───────────────
                pt = tk.Frame(row, bg=row_bg, width=135, height=42)
                pt.pack(side="left", padx=2); pt.pack_propagate(False)
                if is_refunded:
                    tk.Label(pt, text=L("Refunded (Locked)", "已退款（锁定）"),
                             font=("Arial",8,"bold"), bg=row_bg, fg=C['danger']).pack(side="left", padx=6, pady=8)
                else:
                    for pm, col in [("CASH","#16A34A"),("QR","#0369A1")]:
                        is_cur = (pay == pm)
                        def _cp(m=pm, s2=sale):
                            def _do(staff):
                                with self.db.get_connection() as conn2:
                                    self._change_sale_payment(conn2, s2, m)
                                    conn2.execute(
                                        "INSERT INTO audit_log (timestamp,action,details,user) VALUES (?,?,?,?)",
                                        (datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                                         "CHANGE_PAYMENT",
                                         f"Sale #{s2['id']} payment: {str(s2['payment'] or '').upper() or 'CASH'} -> {m}",
                                         staff['name']))
                                self.update_today_stats()
                                load_orders()
                            self._staff_login_popup(callback=_do, parent=win)
                        RoundedButton(pt, text=pm,
                                      command=_cp,
                                      bg=col if is_cur else "#CBD5E1",
                                      fg="white" if is_cur else "#475569",
                                      font=("Arial",8,"bold"),
                                      padx=6, pady=4, radius=6
                                      ).pack(side="left", padx=2)

                # ── Action buttons (rounded) ──────────────────────────────
                act = tk.Frame(row, bg=row_bg, width=270, height=42)
                act.pack(side="left", padx=2); act.pack_propagate(False)

                def _view(s=sale, active_query=item_query):
                    try:   items_data = json.loads(s['items'])
                    except: items_data = []
                    list_prices = json.loads(s['category_list_prices'])
                    lines = [
                        f"Order #{s['id']}  |  {s['datetime']}",
                        f"Cashier: {s['cashier']}  |  Payment: {s['payment']}",
                        "-"*38,
                    ]
                    for name, qty, _, display_price in self._display_item_prices(items_data, list_prices):
                        marker = "▶" if (
                            active_query and active_query.casefold() in str(name).casefold()
                        ) else " "
                        lines.append(
                            f"{marker} {str(name)[:20]:20s}  x{qty}  {self.currency}{display_price * qty:.2f}"
                        )
                    adjustment = self._display_price_adjustment(items_data, list_prices)
                    lines += ["-"*38]
                    if adjustment:
                        lines.append(self._receipt_adjustment_line(adjustment).rstrip())
                    lines.append(L(f"Amount paid: {self.currency} {s['total']:.2f}",
                                   f"实收金额：{self.currency} {s['total']:.2f}"))
                    messagebox.showinfo(f"Order #{s['id']}", "\n".join(lines), parent=win)

                def _reprint(s=sale): _do_reprint(s, win)

                def _edit(s=sale):
                    def do_edit(staff):
                        self._open_edit_order(s, win,
                            lambda: (self.update_today_stats(), load_orders()), staff)
                    self._staff_login_popup(callback=do_edit, parent=win)

                def _refund(s=sale): self._quick_full_refund(s, win, lambda: (self.update_today_stats(), self.load_items(), load_orders()))

                actions = [
                    (L("Details", "详情"),  _view,    C['neutral']),
                    (L("Reprint", "重打"),  _reprint, C['card']),
                ]
                if not is_refunded:
                    actions.append((L("Refund", "退款"), _refund, C['danger']))
                    actions.append((L("Edit", "编辑"), _edit, C['warning']))

                for btn_txt, btn_cmd, btn_col in actions:
                    RoundedButton(act, text=btn_txt, command=btn_cmd,
                                  bg=btn_col, fg="white",
                                  font=("Arial",9,"bold"),
                                  padx=8, pady=5, radius=7
                                  ).pack(side="left", padx=2)

            if has_adjusted_display:
                result_summary_var.set(result_summary_var.get() + L(
                    " · * Displayed total. See Details for amount paid.",
                    " · * 显示金额，实收金额见详情。"))

        load_orders()

    def _payment_summary_column(self, payment_method):
        pm = str(payment_method or '').strip().upper()
        if pm == 'CASH':
            return 'cash_total'
        if pm == 'QR':
            return 'qr_total'
        if pm == 'CARD':
            return 'card_total'
        return None

    def _change_sale_payment(self, conn, sale_row, new_payment):
        """Change sale payment and reconcile daily_summary/cashflow totals."""
        sale_id = int(sale_row['id'])
        old_payment = str(sale_row['payment'] or '').strip().upper() or 'CASH'
        new_payment = str(new_payment or '').strip().upper()
        if not new_payment or new_payment == old_payment:
            return

        sale_total = float(sale_row['total'] or 0)
        biz_date = biz_date_of_sale(str(sale_row['datetime'] or ''))

        conn.execute("UPDATE sales SET payment=? WHERE id=?", (new_payment, sale_id))

        old_col = self._payment_summary_column(old_payment)
        new_col = self._payment_summary_column(new_payment)
        if old_col or new_col:
            conn.execute(
                """INSERT OR IGNORE INTO daily_summary
                   (date, total_sales, transaction_count, cash_total, qr_total, card_total, total_discount)
                   VALUES (?,0,0,0,0,0,0)""",
                (biz_date,)
            )
            if old_col:
                conn.execute(f"UPDATE daily_summary SET {old_col} = {old_col} - ? WHERE date=?",
                             (sale_total, biz_date))
            if new_col:
                conn.execute(f"UPDATE daily_summary SET {new_col} = {new_col} + ? WHERE date=?",
                             (sale_total, biz_date))

        conn.execute(
            """UPDATE cashflow
               SET payment_method=?
               WHERE type='Income' AND reference_id=? AND source LIKE 'Sales%'""",
            (new_payment, sale_id)
        )

    def quick_change_last_payment(self):
        """Change payment method of the most recent sale — no admin required."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, datetime, total, payment, cashier FROM sales ORDER BY id DESC LIMIT 1")
            last = c.fetchone()
        if not last:
            messagebox.showinfo("No Records", "No sales records found."); return

        C = self.colors
        popup = tk.Toplevel(self.root)
        popup.title("Change Last Order Payment")
        popup.configure(bg="#1E293B")
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)
        pw, ph = 500, 360
        sx = self.root.winfo_x() + (self.root.winfo_width()  - pw) // 2
        sy = self.root.winfo_y() + (self.root.winfo_height() - ph) // 2
        popup.geometry(f"{pw}x{ph}+{sx}+{sy}")

        tk.Label(popup, text="Change Last Order Payment",
                 font=("Arial", 16, "bold"), bg="#1E293B", fg="white").pack(pady=(20, 4))
        tk.Label(popup,
                 text=f"Order #{last['id']}  |  {last['datetime']}  |  RM {last['total']:.2f}",
                 font=("Arial", 10), bg="#1E293B", fg="#94A3B8").pack(pady=(0, 4))
        tk.Label(popup,
                 text=f"Current Payment: {last['payment']}",
                 font=("Arial", 12, "bold"), bg="#1E293B", fg="#FCD34D").pack(pady=(0, 16))

        btn_row = tk.Frame(popup, bg="#1E293B")
        btn_row.pack(pady=6)
        for label, method, color in [
            ("💵 Cash\nCASH",  "CASH", "#16A34A"),
            ("📱 QR\nQR",     "QR",   "#0369A1"),
        ]:
            def _pick(m=method):
                def _do(staff):
                    with self.db.get_connection() as conn2:
                        self._change_sale_payment(conn2, last, m)
                        conn2.execute(
                            "INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                            (datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                             "CHANGE_PAYMENT",
                             f"Sale #{last['id']} payment changed from {str(last['payment'] or '').upper() or 'CASH'} to {m}",
                             staff['name']))
                    popup.destroy()
                    self.update_today_stats()
                    messagebox.showinfo("✅ Updated",
                        f"Order #{last['id']} payment changed to {m}")
                self._staff_login_popup(callback=_do, parent=popup)
            RoundedButton(btn_row, text=label, command=_pick,
                          bg=color, fg="white",
                          font=("Arial", 13, "bold"), padx=22, pady=16, radius=10
                          ).pack(side="left", padx=8)

        RoundedButton(popup, text="✖ Cancel", command=popup.destroy,
                      bg=C['neutral'], fg="white",
                      font=("Arial", 10), padx=16, pady=6, radius=8).pack(pady=(10, 0))

    # ── Cashier quick: edit last order items ────────────────────────────────
    def quick_edit_last_order(self):
        """Open Edit Order for the most recent sale — no admin required."""
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM sales ORDER BY id DESC LIMIT 1")
            last = c.fetchone()
            if not last:
                messagebox.showinfo("No Records", "No sales records found."); return
            c.execute("SELECT COALESCE(SUM(refund_amount),0) AS total_refund FROM refunds WHERE sale_id=?", (last['id'],))
            ref_res = c.fetchone()
            refunded = float(ref_res['total_refund'] if ref_res else 0)
        
        if refunded > 0.005:
            messagebox.showwarning("Cannot Edit", "Cannot edit a refunded order.")
            return

        self._staff_login_popup(
            callback=lambda staff: self._open_edit_order(last, self.root, lambda: self.update_today_stats(), staff))

    # ================= HISTORY =================
    def show_history(self):
        L = (lambda en, zh: zh if self.lang == "zh" else en)
        history_window = tk.Toplevel(self.root)
        history_window.title(L("Sales History", "销售历史"))
        history_window.geometry("1600x900")
        history_window.configure(bg=self.colors['bg'])
        history_window.transient(self.root)

        header = tk.Frame(history_window, bg=self.colors['primary'], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=L("📜 Sales History", "📜 销售历史"), font=("Arial", 20, "bold"),
                 bg=self.colors['header'], fg='white').pack(pady=10)

        # Use one notebook only (no nested report notebook) to avoid duplicated-window feel.
        outer_nb = ttk.Notebook(history_window)
        outer_nb.pack(fill="both", expand=True, padx=6, pady=6)

        history_tab = tk.Frame(outer_nb, bg=self.colors['bg'])
        outer_nb.add(history_tab, text=L("📜 Sales History", "📜 销售历史"))

        # No report tabs here anymore; reports are moved to the dedicated Reports window.

        history_window_ref = history_window

        filter_frame = tk.Frame(history_tab, bg=self.colors['bg'])
        filter_frame.pack(fill="x", padx=20, pady=8)
        row1 = tk.Frame(filter_frame, bg=self.colors['bg'])
        row1.pack(fill="x")
        tk.Label(row1, text=L("Filter:", "筛选："), font=("Arial", 12, "bold"),
                 bg=self.colors['bg']).pack(side="left", padx=5)
        filter_var = tk.StringVar(value="today")
        filters = [
            (L("Today (Biz Day)", "今天（营业日）"), "today"),
            (L("Previous Day", "前一天"), "yesterday"),
            (L("Last 7 Days", "最近7天"), "last7"),
            (L("This Month", "本月"), "month"),
            (L("All", "全部"), "all"),
            (L("Custom", "自定义"), "custom"),
        ]
        for text, val in filters:
            tk.Radiobutton(row1, text=text, variable=filter_var, value=val,
                           font=("Arial", 11), bg=self.colors['bg'], fg=self.colors['text'],
                           command=lambda: toggle_custom()).pack(side="left", padx=5)

        row2 = tk.Frame(filter_frame, bg=self.colors['bg'])
        tk.Label(row2, text=L("From:", "开始："), font=("Arial", 11), bg=self.colors['bg']).pack(side="left", padx=(10,2))
        from_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        dp_h_from = DatePicker(
            row2, textvariable=from_var,
            language=self.lang
        )
        dp_h_from.pack(side="left")

        tk.Label(row2, text=L("To:", "结束："), font=("Arial", 11), bg=self.colors['bg']).pack(side="left", padx=(10,2))
        to_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        dp_h_to = DatePicker(
            row2, textvariable=to_var,
            language=self.lang
        )
        dp_h_to.pack(side="left")

        tk.Label(
            row2,
            text=L("Type or choose a date", "可输入或从日历选择"),
            font=("Arial", 9), bg=self.colors['bg'], fg="gray"
        ).pack(side="left", padx=5)
        RoundedButton(row2, text=L("🔍 Search", "🔍 查询"), command=lambda: load_history(),
                      bg=self.colors['neutral'], fg="white",
                      font=("Arial", 11, "bold"), padx=10, pady=2, radius=6
                      ).pack(side="left", padx=5)

        def toggle_custom():
            if filter_var.get() == "custom":
                row2.pack(fill="x", pady=4)
            else:
                row2.pack_forget()
                load_history()

        history_search_var = tk.StringVar()
        history_result_var = tk.StringVar()
        query_host = tk.Frame(history_tab, bg=self.colors['bg'])
        query_host.pack(fill="x", padx=4, pady=(0, 2))

        tree_frame = tk.Frame(history_tab, bg=self.colors['bg'])
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)

        kpi_row = tk.Frame(history_tab, bg=self.colors['bg'])
        kpi_row.pack(fill="x", padx=20, pady=(0, 6))
        kpi_row.columnconfigure(0, weight=1)
        kpi_row.columnconfigure(1, weight=1)
        kpi_row.columnconfigure(2, weight=1)

        def _history_kpi(parent, title, color):
            card = tk.Frame(parent, bg="#FFFFFF", highlightbackground="#D1D5DB", highlightthickness=1)
            tk.Frame(card, bg=color, height=6).pack(fill="x")
            tk.Label(card, text=title, font=("Arial", 10, "bold"), bg="#FFFFFF", fg="#475569").pack(pady=(8, 2))
            value_lbl = tk.Label(card, text="-", font=("Arial", 15, "bold"), bg="#FFFFFF", fg="#0F172A")
            value_lbl.pack(pady=(0, 10))
            return card, value_lbl

        sales_card, sales_kpi_lbl = _history_kpi(kpi_row, L("Total Sales", "总销售额"), self.colors['success'])
        tx_card, tx_kpi_lbl = _history_kpi(kpi_row, L("Transactions", "交易数"), self.colors['primary'])
        pay_card, pay_kpi_lbl = _history_kpi(kpi_row, L("Cash / QR", "现金 / 二维码"), self.colors['warning'])
        sales_card.grid(row=0, column=0, padx=4, sticky="nsew")
        tx_card.grid(row=0, column=1, padx=4, sticky="nsew")
        pay_card.grid(row=0, column=2, padx=4, sticky="nsew")

        history_current_data = []

        columns = (
            L("ID", "编号"), L("DateTime", "日期时间"), L("Total", "总额"),
            L("Payment", "付款"), L("Cashier", "收银员"), L("Items", "商品数")
        )
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        col_widths = {
            L("ID", "编号"): 50,
            L("DateTime", "日期时间"): 160,
            L("Total", "总额"): 100,
            L("Payment", "付款"): 80,
            L("Cashier", "收银员"): 120,
            L("Items", "商品数"): 80,
        }
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths.get(col, 100), anchor="center")
        tree.column(L("DateTime", "日期时间"), anchor="w")
        tree.column(L("Cashier", "收银员"), anchor="w")
        scrollbar = tk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def load_history():
            tree.delete(*tree.get_children())
            history_current_data.clear()
            filter_type = filter_var.get()
            now = datetime.now()

            start_biz = None
            end_biz = None
            period_label = L("All", "全部")

            if filter_type == "today":
                start_biz = end_biz = today_biz_date()
                period_label = f"{L('Biz Day', '营业日')}: {start_biz}"
            elif filter_type == "yesterday":
                start_biz = end_biz = get_business_date(now - timedelta(days=1))
                period_label = f"{L('Previous Day', '前一天')}: {start_biz}"
            elif filter_type == "last7":
                start_biz = get_business_date(now - timedelta(days=6))
                end_biz = today_biz_date()
                period_label = f"{start_biz} to {end_biz}"
            elif filter_type == "month":
                start_biz = get_business_date(now.replace(day=1))
                end_biz = today_biz_date()
                period_label = f"{start_biz} to {end_biz}"
            elif filter_type == "custom":
                try:
                    normalized_from = normalize_date_text(from_var.get())
                    normalized_to = normalize_date_text(to_var.get())
                    f = datetime.strptime(normalized_from, "%Y-%m-%d")
                    t = datetime.strptime(normalized_to, "%Y-%m-%d")
                    if t < f:
                        raise ValueError
                    from_var.set(normalized_from)
                    to_var.set(normalized_to)
                    start_biz = f.strftime("%Y-%m-%d")
                    end_biz = t.strftime("%Y-%m-%d")
                    period_label = f"{start_biz} to {end_biz}"
                except ValueError:
                    messagebox.showerror(
                        L("Invalid Date", "日期无效"),
                        L(
                            "Choose valid From and To dates; From cannot be after To.",
                            "请选择有效的开始及结束日期，开始日期不可晚于结束日期。"
                        ),
                        parent=history_window_ref
                    )
                    return

            with self.db.get_connection() as conn:
                c = conn.cursor()
                # Bug fix: filter amended=0 to avoid double-counting replaced orders
                c.execute("SELECT * FROM sales WHERE amended=0 ORDER BY id DESC LIMIT 2000")
                all_sales = c.fetchall()

                # Fetch refunds for the same potential period to adjust KPIs
                c.execute("SELECT sale_id, refund_amount FROM refunds ORDER BY id DESC LIMIT 2000")
                all_refunds = c.fetchall()
                refund_map = {} # sale_id -> total_refunded
                for r in all_refunds:
                    sid = r['sale_id']
                    refund_map[sid] = refund_map.get(sid, 0.0) + float(r['refund_amount'] or 0)

            keyword = history_search_var.get().strip().casefold()
            sales = []
            for sale in all_sales:
                if filter_type != "all":
                    sale_biz = biz_date_of_sale(str(sale['datetime'] or ''))
                    if not (start_biz <= sale_biz <= end_biz):
                        continue
                if keyword:
                    searchable = " ".join((
                        str(sale['id'] or ''), str(sale['datetime'] or ''),
                        str(sale['payment'] or ''), str(sale['cashier'] or ''),
                        str(sale['table_code'] or '') if 'table_code' in sale.keys() else '',
                        str(sale['items'] or ''),
                    )).casefold()
                    if keyword not in searchable:
                        continue
                sales.append(sale)

            total_sales = 0.0
            total_cash = 0.0
            total_qr = 0.0
            has_adjusted_display = False
            for sale in sales:
                sid = int(sale['id'])
                items_list  = json.loads(sale['items'])
                item_count  = sum(item[2] for item in items_list)
                sale_total = float(sale['total'] or 0)
                
                # Check for refunds on this specific sale
                refund_amt = refund_map.get(sid, 0.0)
                is_refunded = refund_amt >= (sale_total - 0.005)
                is_partial  = 0 < refund_amt < (sale_total - 0.005)
                
                status_tag = ""
                if is_refunded: status_tag = L(" [REFUNDED]", " [已退款]")
                elif is_partial: status_tag = L(" [PART-REFUNDED]", " [部分退款]")

                payment = str(sale['payment'] or '').strip().upper()
                display_total, star = self._display_sale_total(sale)
                has_adjusted_display |= bool(star)
                tree.insert("", "end", values=(
                    sale['id'], sale['datetime'], f"{self.currency} {display_total:.2f}{star}{status_tag}",
                    payment or '-', sale['cashier'] or "-",
                    L(f"{item_count} items", f"{item_count} 件")
                ))
                
                # Net totals (Sales minus any refunds)
                net_sale = sale_total - refund_amt
                total_sales += net_sale
                if payment == 'CASH':
                    total_cash += net_sale
                elif payment == 'QR':
                    total_qr += net_sale

                history_current_data.append({
                    'ID': sale['id'],
                    'DateTime': sale['datetime'],
                    'Total': f"{sale_total:.2f}",
                    'Payment': payment or '-',
                    'Cashier': sale['cashier'] or '-',
                    'Items': item_count,
                    'BusinessDate': biz_date_of_sale(str(sale['datetime'] or '')),
                })

            tx_count = len(sales)
            sales_kpi_lbl.config(text=f"{self.currency} {total_sales:.2f}")
            tx_kpi_lbl.config(text=f"{tx_count}")
            pay_kpi_lbl.config(text=f"{self.currency} {total_cash:.2f} / {self.currency} {total_qr:.2f}")
            history_result_var.set(
                L(
                    f"{tx_count} order(s) · {period_label}",
                    f"{tx_count} 笔订单 · {period_label}"
                )
            )
            if has_adjusted_display:
                history_result_var.set(history_result_var.get() + L(
                    " · * Displayed total. See Details for amount paid.",
                    " · * 显示金额，实收金额见详情。"))
            # summary_label.config(
            #     text=(f"{L('Period', '周期')}: {period_label} | {L('Total Sales', '总销售额')}: {self.currency} {total_sales:.2f} | "
            #         f"{L('Transactions', '交易数')}: {tx_count}")
            # )

        # Summary label (commented out to avoid redundancy with colored KPI cards)
        # summary_frame = tk.Frame(history_tab, bg=self.colors['card'], height=50)
        # summary_frame.pack(fill="x", pady=5)
        # summary_frame.pack_propagate(False)
        # summary_label = tk.Label(summary_frame, text="", font=("Arial", 14, "bold"),
        #                          bg=self.colors['card'], fg=self.colors['light'])
        # summary_label.pack(pady=10)

        btn_frame = tk.Frame(history_tab, bg=self.colors['bg'])
        btn_frame.pack(fill="x", padx=20, pady=10)

        def view_details():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select a transaction", "请选择一笔交易"), parent=history_window_ref)
                return
            item     = tree.item(selection[0])
            sale_id  = item['values'][0]
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM sales WHERE id=?", (sale_id,))
                sale = c.fetchone()
            items_list = json.loads(sale['items'])
            details = (f"{L('Transaction Details', '交易详情')}\n{'='*40}\n{L('Order ID', '订单号')}: {sale['id']}\n"
                       f"{L('Date & Time', '日期时间')}: {sale['datetime']}\n{L('Cashier', '收银员')}: {sale['cashier']}\n"
                       f"{L('Payment', '付款')}: {sale['payment']}\n\n{L('Items', '商品')}:\n{'-'*40}\n")
            list_prices = json.loads(sale['category_list_prices'])
            adjustment = self._display_price_adjustment(items_list, list_prices)
            details += self._receipt_item_lines(items_list, list_prices) + "\n"
            details += (f"{'-'*40}\n"
                        f"{self._receipt_adjustment_line(adjustment)}"
                        f"{L('Tax', '税')}: {self.currency} {sale['tax']:.2f}\n{'='*40}\n"
                        f"{L('Amount paid', '实收金额')}: {self.currency} {sale['total']:.2f}\n")
            messagebox.showinfo(L("Transaction Details", "交易详情"), details, parent=history_window_ref)

        def edit_order():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select a transaction to edit", "请选择要编辑的交易"), parent=history_window_ref)
                return
            item    = tree.item(selection[0])
            sale_id = item['values'][0]
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM sales WHERE id=?", (sale_id,))
                sale = c.fetchone()
                c.execute("SELECT COALESCE(SUM(refund_amount),0) AS total_refund FROM refunds WHERE sale_id=?", (sale_id,))
                ref_res = c.fetchone()
                refunded = float(ref_res['total_refund'] if ref_res else 0)
                
            if refunded > 0.005:
                messagebox.showwarning(L("Cannot Edit", "无法编辑"), L("Cannot edit a refunded order.", "已退款的订单无法编辑。"), parent=history_window_ref)
                return

            self._staff_login_popup(
                callback=lambda staff: self._open_edit_order(sale, history_window, load_history, staff))

        def reprint_from_history():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select a transaction to reprint", "请选择要重打的交易"), parent=history_window_ref)
                return
            item    = tree.item(selection[0])
            sale_id = item['values'][0]
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM sales WHERE id=?", (sale_id,))
                sale = c.fetchone()
            items_list = json.loads(sale['items'])
            subtotal, tax, total, discount_line = self._saved_receipt_totals(sale)
            list_prices = json.loads(sale['category_list_prices'])
            receipt = self._receipt_item_lines(items_list, list_prices)
            adjustment = self._display_price_adjustment(items_list, list_prices)
            table_code = sale['table_code'] if ('table_code' in sale.keys() and sale['table_code']) else '-'
            text = f"""
{'================================'}
         919 Restaurant
        (CT0159652-K)
  NO.1 G2 BLOK D, JALAN INDAH
    12/2 TAMAN INDAH 84000
        MUAR JOHOR
{'================================'}
Date: {sale['datetime']}
Order: #{sale['id']:04d} (REPRINT)
Table: {table_code}
Cashier: {sale['cashier']}
{'================================'}
{receipt}
{'--------------------------------'}
Subtotal:         {self.currency} {subtotal - adjustment:>10.2f}
{discount_line}Tax:              {self.currency} {tax:>10.2f}
{'================================'}
{self._receipt_total_lines(total, adjustment, paid=True)}{'================================'}
Payment: {sale['payment']}
{'================================'}
    Terima Kasih! Thank You!
      Please Come Again!
{'================================'}
"""
            try:
                print_to_thermal(PRINTER_NAME, "\n\n" + text + "\n\n\n\x1dV\x00")
                messagebox.showinfo(L("Reprinted", "已重打"), L("Receipt Printed", "小票已重打"), parent=history_window_ref)
            except Exception as e:
                messagebox.showerror(L("Print Error", "打印错误"), str(e), parent=history_window_ref)

        def export_history_csv():
            if not history_current_data: return
            import csv
            from tkinter import filedialog
            filepath = filedialog.asksaveasfilename(parent=history_window_ref, defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if not filepath: return
            try:
                with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=['ID', 'DateTime', 'BusinessDate', 'Total', 'Payment', 'Cashier', 'Items'])
                    writer.writeheader(); writer.writerows(history_current_data)
                messagebox.showinfo(L("Export Success", "导出成功"), f"Saved: {filepath}", parent=history_window_ref)
            except Exception as e:
                messagebox.showerror(L("Export Failed", "导出失败"), str(e), parent=history_window_ref)

        def clear_history_filters():
            filter_var.set("today")
            today_text = datetime.now().strftime("%Y-%m-%d")
            from_var.set(today_text)
            to_var.set(today_text)
            row2.pack_forget()
            load_history()

        self._build_query_toolbar(
            query_host, history_search_var, load_history,
            on_clear=clear_history_filters,
            on_refresh=load_history,
            on_export=lambda: self._export_treeview_csv(
                tree, history_window_ref, "sales_history.csv"
            ),
            result_var=history_result_var,
            search_label=L("Order / item / cashier:", "订单／商品／收银员：")
        )

        for _t, _c, _bg in [
            (L("👁️ View Details", "👁️ 查看详情"), view_details,         self.colors['neutral']),
            (L("✏️ Edit Order",   "✏️ 编辑订单"), edit_order,            self.colors['warning']),
            (L("🖨️ Reprint",      "🖨️ 重打小票"), reprint_from_history, self.colors['neutral']),
            (L("💰 Reconcile",    "💰 对账"), lambda: self._payment_reconciliation(history_window), self.colors['neutral']),
        ]:
            RoundedButton(btn_frame, text=_t, command=_c, bg=_bg, fg="white",
                          font=("Arial", 12, "bold"), padx=16, pady=8, radius=8
                          ).pack(side="left", padx=4)
        load_history()

    # ================= EDIT ORDER (IN-PLACE) =================
    def _open_edit_order(self, sale, parent_window, refresh_callback, editor_staff=None):
        edit_window = tk.Toplevel(parent_window)
        edit_window.title(f"Edit Order #{sale['id']}")
        edit_window.geometry("980x920")
        edit_window.configure(bg=self.colors['bg'])
        edit_window.transient(parent_window)
        edit_window.grab_set()

        header = tk.Frame(edit_window, bg=self.colors['neutral'], height=60)
        header.pack(fill="x"); header.pack_propagate(False)
        tk.Label(header, text=f"✏️ Edit Order #{sale['id']}", font=("Arial", 18, "bold"), bg=self.colors['neutral'], fg=self.colors['light']).pack(pady=12)

        info_f = tk.Frame(edit_window, bg=self.colors['bg']); info_f.pack(fill="x", padx=20, pady=10)
        tk.Label(info_f, text=f"Date: {sale['datetime']}", font=("Arial", 11), bg=self.colors['bg']).pack(anchor="w")
        pay_f = tk.Frame(info_f, bg=self.colors['bg']); pay_f.pack(anchor="w", pady=5)
        edit_payment = tk.StringVar(value=sale['payment'])
        for val in ["CASH", "QR"]:
            tk.Radiobutton(pay_f, text=val, variable=edit_payment, value=val, font=("Arial", 11), bg=self.colors['bg']).pack(side="left", padx=10)

        item_f = tk.Frame(edit_window, bg=self.colors['bg']); item_f.pack(fill="both", expand=True, padx=20, pady=5)
        cols = ("Item", "Price", "Qty", "Total")
        itree = ttk.Treeview(item_f, columns=cols, show="headings", height=10)
        for c in cols: itree.heading(c, text=c); itree.column(c, width=120)
        itree.pack(side="left", fill="both", expand=True)

        original_items = json.loads(sale['items'])
        edited_items   = list(original_items)
        tvar = tk.StringVar()

        def refresh_itree():
            itree.delete(*itree.get_children()); tot = Decimal('0')
            for n, p, q, ct in edited_items:
                line = Decimal(str(p)) * q; tot += line
                itree.insert("", "end", values=(n, f"RM {p:.2f}", q, f"RM {line:.2f}"))
            tvar.set(f"Total: RM {tot:.2f}")
        refresh_itree()
        tk.Label(item_f, textvariable=tvar, font=("Arial", 13, "bold"), bg=self.colors['bg']).pack(anchor="e")

        def add_item():
            # Category-separated item picker
            L = self.l
            swin = tk.Toplevel(edit_window)
            swin.title(L("Add Item", "新增商品"))
            swin.configure(bg=self.colors['bg'])
            swin.transient(edit_window)
            swin.grab_set()
            swin.resizable(True, True)
            self._center_window(swin, 780, 620, edit_window)
            swin.minsize(560, 460)

            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT id, name, price, category FROM items WHERE active=1 ORDER BY name")
                all_items = c.fetchall()

            item_map = {it['name']: it for it in all_items}
            categories = item_picker_categories(all_items)
            category_counts = {category: 0 for category in categories}
            for item in all_items:
                item_category = str(item['category'] or "Uncategorized").strip() or "Uncategorized"
                category_counts[item_category] = category_counts.get(item_category, 0) + 1

            # Header
            header = tk.Frame(swin, bg=self.colors['primary'], height=50)
            header.pack(fill="x")
            header.pack_propagate(False)
            tk.Label(
                header, text=L("➕ Add Item", "➕ 新增商品"),
                font=("Arial", 16, "bold"),
                bg=self.colors['primary'], fg="white"
            ).pack(pady=10)

            # Search bar
            search_f = tk.Frame(swin, bg=self.colors['bg'])
            search_f.pack(fill="x", padx=12, pady=(10, 6))
            tk.Label(
                search_f, text=L("Search:", "搜索："),
                font=("Arial", 12, "bold"), bg=self.colors['bg']
            ).pack(side="left")
            search_var = tk.StringVar()
            search_entry = tk.Entry(search_f, textvariable=search_var, font=("Arial", 12))
            search_entry.pack(side="left", fill="x", expand=True, padx=10)
            result_var = tk.StringVar()
            tk.Label(
                search_f, textvariable=result_var, font=("Arial", 9),
                bg=self.colors['bg'], fg=self.colors['text_muted']
            ).pack(side="right", padx=(8, 0))
            tk.Button(
                search_f, text=L("Clear", "清除"),
                command=lambda: search_var.set(""),
                font=("Arial", 9, "bold"), bg="#E2E8F0", fg="#334155",
                activebackground="#CBD5E1", relief="flat", bd=0,
                padx=10, pady=4, cursor="hand2"
            ).pack(side="right")
            search_entry.focus_set()

            # Dynamic category tabs. The horizontal strip stays usable even
            # when the restaurant adds more categories later.
            selected_category = tk.StringVar(value="")
            category_outer = tk.Frame(swin, bg=self.colors['bg'])
            category_outer.pack(fill="x", padx=12, pady=(0, 5))
            tk.Label(
                category_outer, text=L("Category:", "分类："),
                font=("Arial", 10, "bold"), bg=self.colors['bg'],
                fg=self.colors['text_sub']
            ).pack(anchor="w", pady=(0, 3))
            category_canvas = tk.Canvas(
                category_outer, height=42, bg=self.colors['bg'],
                highlightthickness=0
            )
            category_scroll = ttk.Scrollbar(
                category_outer, orient="horizontal", command=category_canvas.xview
            )
            category_canvas.configure(xscrollcommand=category_scroll.set)
            category_canvas.pack(fill="x")
            category_scroll.pack(fill="x", pady=(1, 0))
            category_tabs = tk.Frame(category_canvas, bg=self.colors['bg'])
            category_canvas.create_window((0, 0), window=category_tabs, anchor="nw")
            tab_buttons = {}

            def refresh_category_tab_styles():
                active = selected_category.get()
                for key, button in tab_buttons.items():
                    is_active = key == active
                    button.configure(
                        bg=self.colors['primary'] if is_active else "#E2E8F0",
                        fg="white" if is_active else self.colors['text_sub'],
                        activebackground=self.colors['header2'] if is_active else "#CBD5E1",
                        activeforeground="white" if is_active else self.colors['text']
                    )

            def select_category(category):
                selected_category.set(category)
                refresh_category_tab_styles()
                populate_list()

            category_options = [
                ("", L(f"All ({len(all_items)})", f"全部 ({len(all_items)})"))
            ]
            for category in categories:
                label = L("Uncategorized", "未分类") if category == "Uncategorized" else category
                category_options.append(
                    (category, f"{label} ({category_counts.get(category, 0)})")
                )

            for key, label in category_options:
                button = tk.Button(
                    category_tabs, text=label,
                    command=lambda category=key: select_category(category),
                    font=("Arial", 9, "bold"), relief="flat", bd=0,
                    padx=13, pady=7, cursor="hand2"
                )
                button.pack(side="left", padx=(0, 5))
                tab_buttons[key] = button
                button.bind(
                    "<MouseWheel>",
                    lambda event: category_canvas.xview_scroll(
                        -1 if event.delta > 0 else 1, "units"
                    )
                )

            def sync_category_scroll(_event=None):
                category_canvas.configure(scrollregion=category_canvas.bbox("all"))

            category_tabs.bind("<Configure>", sync_category_scroll)
            category_canvas.bind(
                "<MouseWheel>",
                lambda event: category_canvas.xview_scroll(
                    -1 if event.delta > 0 else 1, "units"
                )
            )
            refresh_category_tab_styles()

            # Treeview
            tv_f = tk.Frame(swin, bg=self.colors['bg'])
            tv_f.pack(fill="both", expand=True, padx=12, pady=5)
            
            cols = ("Name", "Category", "Price")
            tv = ttk.Treeview(tv_f, columns=cols, show="headings", height=15)
            tv.heading("Name", text="Item Name")
            tv.heading("Category", text="Category")
            tv.heading("Price", text="Price")
            tv.column("Name", width=300)
            tv.column("Category", width=150)
            tv.column("Price", width=100, anchor="e")
            
            sb = tk.Scrollbar(tv_f, orient="vertical", command=tv.yview)
            tv.configure(yscrollcommand=sb.set)
            tv.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            def populate_list(*args):
                tv.delete(*tv.get_children())
                matches = filter_item_picker_items(
                    all_items,
                    query=search_var.get(),
                    category=selected_category.get() or None
                )
                for it in matches:
                    display_category = str(it['category'] or L("Uncategorized", "未分类"))
                    tv.insert(
                        "", "end",
                        values=(it['name'], display_category, f"RM {it['price']:.2f}")
                    )
                result_var.set(
                    L(f"{len(matches)} item(s)", f"{len(matches)} 个商品")
                )
                first = tv.get_children()
                if first:
                    tv.selection_set(first[0])
                    tv.focus(first[0])

            populate_list()
            search_var.trace_add("write", populate_list)

            def show_set_picker(it, sets, final_name, final_price):
                vwin = tk.Toplevel(swin)
                vwin.title(f"Select Set — {final_name}")
                vwin.configure(bg=self.colors['neutral'])
                vwin.transient(swin)
                vwin.grab_set()
                vwin.resizable(False, False)

                tk.Label(vwin, text=f"📦 {final_name}", font=("Arial", 16, "bold"),
                         bg=self.colors['neutral'], fg="white").pack(pady=(18, 6), padx=30)
                tk.Label(vwin, text="Select a set option:", font=("Arial", 11),
                         bg=self.colors['neutral'], fg="#E2E8F0").pack(pady=(0, 12))

                btn_frame = tk.Frame(vwin, bg=self.colors['neutral'])
                btn_frame.pack(padx=24, pady=4)

                def choose_s(s_name, s_price):
                    vwin.destroy()
                    swin.destroy()
                    edited_items.append((s_name, float(s_price), 1, it['category']))
                    refresh_itree()

                cat_color = self.category_colors.get(it['category'], self.colors['info'])
                RoundedButton(btn_frame, text=f"Single\n{self.currency} {float(final_price):.2f}",
                              command=lambda: choose_s(final_name, final_price),
                              bg=cat_color, fg="white",
                              font=("Arial", 13, "bold"), padx=20, pady=14, radius=10
                              ).pack(side="left", padx=6, pady=6)

                for s in sets:
                    s_qty, s_price = s['quantity'], s['set_price']
                    btn_text = f"Set ({s_qty})\n{self.currency} {float(s_price):.2f}"
                    display_name = f"{final_name} Set({s_qty})"
                    RoundedButton(btn_frame, text=btn_text,
                                  command=lambda dn=display_name, sp=s_price: choose_s(dn, sp),
                                  bg=self.colors['primary'], fg="white",
                                  font=("Arial", 13, "bold"), padx=20, pady=14, radius=10
                                  ).pack(side="left", padx=6, pady=6)

                RoundedButton(vwin, text="✖ Cancel", command=vwin.destroy,
                              bg=self.colors['danger'], fg="white",
                              font=("Arial", 10), padx=16, pady=6, radius=8
                              ).pack(pady=(6, 18))

                vwin.update_idletasks()
                x = swin.winfo_x() + (swin.winfo_width() - vwin.winfo_width()) // 2
                y = swin.winfo_y() + (swin.winfo_height() - vwin.winfo_height()) // 2
                vwin.geometry(f"+{x}+{y}")

            def show_variant_picker(it, variants, sets):
                vwin = tk.Toplevel(swin)
                vwin.title(f"Select option — {it['name']}")
                vwin.configure(bg=self.colors['neutral'])
                vwin.transient(swin)
                vwin.grab_set()
                vwin.resizable(False, False)

                tk.Label(vwin, text=f"☕ {it['name']}", font=("Arial", 16, "bold"),
                         bg=self.colors['neutral'], fg="white").pack(pady=(18, 6), padx=30)
                tk.Label(vwin, text="Select an option:", font=("Arial", 11),
                         bg=self.colors['neutral'], fg="#E2E8F0").pack(pady=(0, 12))

                btn_frame = tk.Frame(vwin, bg=self.colors['neutral'])
                btn_frame.pack(padx=24, pady=4)

                def choose_v(v_name, v_price):
                    vwin.destroy()
                    display_name = f"{it['name']} ({v_name})"
                    if sets:
                        show_set_picker(it, sets, display_name, v_price)
                    else:
                        swin.destroy()
                        edited_items.append((display_name, float(v_price), 1, it['category']))
                        refresh_itree()

                for v in variants:
                    final_price = float(it['price']) + float(v['price_adjustment'])
                    btn_text = f"{v['variant_name']}\n{self.currency} {final_price:.2f}"
                    cat_color = self.category_colors.get(it['category'], self.colors['info'])
                    RoundedButton(btn_frame, text=btn_text,
                                  command=lambda vn=v['variant_name'], vp=final_price: choose_v(vn, vp),
                                  bg=cat_color, fg="white",
                                  font=("Arial", 13, "bold"), padx=20, pady=14, radius=10
                                  ).pack(side="left", padx=6, pady=6)

                RoundedButton(vwin, text="✖ Cancel", command=vwin.destroy,
                              bg=self.colors['danger'], fg="white",
                              font=("Arial", 10), padx=16, pady=6, radius=8
                              ).pack(pady=(6, 18))

                vwin.update_idletasks()
                x = swin.winfo_x() + (swin.winfo_width() - vwin.winfo_width()) // 2
                y = swin.winfo_y() + (swin.winfo_height() - vwin.winfo_height()) // 2
                vwin.geometry(f"+{x}+{y}")

            def pick(_=None):
                sel = tv.selection()
                if not sel: return
                val = tv.item(sel[0])['values']
                it = item_map[val[0]]
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("SELECT variant_name, price_adjustment FROM item_variants WHERE item_id = ? ORDER BY id", (it['id'],))
                    variants = c.fetchall()
                    c.execute("SELECT quantity, set_price FROM item_quantity_sets WHERE item_id = ? ORDER BY quantity DESC", (it['id'],))
                    sets = c.fetchall()
                if variants:
                    show_variant_picker(it, variants, sets)
                elif sets:
                    show_set_picker(it, sets, it['name'], it['price'])
                else:
                    edited_items.append((it['name'], float(it['price']), 1, it['category']))
                    refresh_itree(); swin.destroy()

            tv.bind("<Double-1>", pick)
            btn_f = tk.Frame(swin, bg=self.colors['bg'])
            btn_f.pack(fill="x", padx=10, pady=10)
            RoundedButton(btn_f, text="✅ Select Item", command=pick, bg=self.colors['success'], fg="white", font=("Arial", 12, "bold"), padx=20, pady=10, radius=8).pack(side="right")

        def remove_item():
            sel = itree.selection()
            if not sel: return
            idx = itree.index(sel[0]); edited_items.pop(idx); refresh_itree()

        btn_f = tk.Frame(edit_window, bg=self.colors['bg']); btn_f.pack(fill="x", padx=20, pady=5)
        tk.Button(btn_f, text="Add Item", command=add_item).pack(side="left", padx=5)
        tk.Button(btn_f, text="Remove", command=remove_item).pack(side="left", padx=5)

        def save_changes():
            sub = sum(Decimal(str(p)) * q for _, p, q, _ in edited_items)
            tx = (sub * self.tax_rate).quantize(Decimal('0.01'), ROUND_HALF_UP)
            ntot = sub + tx; npm = edit_payment.get(); editor = editor_staff['name'] if editor_staff else "Admin"
            try:
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    sid = int(sale['id']); otot = float(sale['total'] or 0); opm = str(sale['payment'] or 'CASH')
                    try: dobj = datetime.strptime(str(sale['datetime']), "%d-%m-%Y %H:%M:%S")
                    except: dobj = datetime.now()
                    biz = get_business_date(dobj)
                    self._adjust_tracked_stock_for_items(
                        conn, original_items, +1, source_type="SALE_EDIT_RESTORE",
                        reference_id=sid, recorded_by=editor,
                        notes="Restore original quantities before applying sale edit"
                    )
                    self._adjust_tracked_stock_for_items(
                        conn, edited_items, -1, source_type="SALE_EDIT_APPLY",
                        reference_id=sid, recorded_by=editor,
                        notes="Apply edited sale quantities"
                    )
                    previous_categories = {item[3] for item in original_items if len(item) >= 4}
                    edited_categories = {item[3] for item in edited_items}
                    list_prices = {
                        cat: price for cat, price in json.loads(sale['category_list_prices']).items()
                        if cat in edited_categories
                    }
                    list_prices.update(list_prices_for_sale(conn, edited_categories - previous_categories))
                    c.execute(
                        """UPDATE sales SET total=?, payment=?, items=?, tax=?, cashier=?,
                                  amended=2, category_list_prices=? WHERE id=?""",
                        (float(ntot), npm, json.dumps([(n, float(p), q, ct) for n, p, q, ct in edited_items]),
                         float(tx), f"[Amended] {editor}", json.dumps(list_prices), sid),
                    )
                    c.execute("UPDATE daily_summary SET total_sales = total_sales - ? + ? WHERE date=?", (otot, float(ntot), biz))
                    opc = self._payment_summary_column(opm); npc = self._payment_summary_column(npm)
                    if opc: c.execute(f"UPDATE daily_summary SET {opc} = {opc} - ? WHERE date=?", (otot, biz))
                    if npc: c.execute(f"UPDATE daily_summary SET {npc} = {npc} + ? WHERE date=?", (float(ntot), biz))
                    c.execute("UPDATE cashflow SET amount=?, payment_method=? WHERE reference_id=? AND source LIKE 'Sales%'", (float(ntot), npm, sid))
                    conn.commit()
                messagebox.showinfo("Success", "Updated"); edit_window.destroy(); refresh_callback(); self.load_items()
            except Exception as e: messagebox.showerror("Error", str(e), parent=edit_window)

        tk.Button(edit_window, text="Save Changes", command=save_changes, bg="green", fg="white", font=("Arial", 12, "bold")).pack(pady=20)


    # ================= PAYMENT RECONCILIATION =================
    def _payment_reconciliation(self, parent_window):
        recon_window = tk.Toplevel(parent_window)
        recon_window.title("💰 Payment Reconciliation")
        recon_window.geometry("920x740")
        recon_window.configure(bg=self.colors['bg'])
        recon_window.transient(parent_window)
        recon_window.grab_set()

        header = tk.Frame(recon_window, bg=self.colors['header'], height=52)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="💰 Payment Reconciliation", font=("Arial", 18, "bold"),
                 bg=self.colors['header'], fg='white').pack(pady=10)

        period_frame = tk.Frame(recon_window, bg=self.colors['bg'])
        period_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(period_frame, text="Select Date:", font=("Arial", 14, "bold"),
                 bg=self.colors['bg']).pack(side="left", padx=5)
        period_var = tk.StringVar(value="today")
        for text, val in [("Today (Biz Day)", "today"), ("This Week", "week"), ("This Month", "month")]:
            tk.Radiobutton(period_frame, text=text, variable=period_var, value=val,
                           font=("Arial", 14), bg=self.colors['bg'],
                           command=lambda: run_reconcile()).pack(side="left", padx=10)

        result_frame = tk.Frame(recon_window, bg=self.colors['white'], relief="sunken", bd=1)
        result_frame.pack(fill="both", expand=True, padx=20, pady=10)
        result_text = tk.Text(result_frame, font=("Courier New", 14), bg=self.colors['white'],
                              relief="flat", padx=15, pady=15)
        result_text.pack(fill="both", expand=True)
        result_text.config(state="disabled")

        tk.Label(recon_window, text="Discrepancies:", font=("Arial", 14, "bold"),
                 bg=self.colors['bg']).pack(anchor="w", padx=20)
        anom_frame = tk.Frame(recon_window, bg=self.colors['bg'])
        anom_frame.pack(fill="x", padx=20, pady=5)
        anom_tree = ttk.Treeview(anom_frame, columns=("ID", "DateTime", "Total", "Payment"),
                                  show="headings", height=6)
        for col, w in [("ID", 60), ("DateTime", 160), ("Total", 100), ("Payment", 100)]:
            anom_tree.heading(col, text=col)
            anom_tree.column(col, width=w)
        anom_tree.pack(fill="x")

        def run_reconcile():
            result_text.config(state="normal")
            result_text.delete("1.0", tk.END)
            anom_tree.delete(*anom_tree.get_children())

            now = datetime.now()
            if period_var.get() == "today":
                biz = today_biz_date()
                s, e = business_day_range(biz)
                s_str = s.strftime("%Y-%m-%d %H:%M:%S")
                e_str = e.strftime("%Y-%m-%d %H:%M:%S")
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("SELECT * FROM sales WHERE to_iso(datetime) >= ? AND to_iso(datetime) <= ? AND amended=0", (s_str, e_str))
                    sales = c.fetchall()
            elif period_var.get() == "week":
                date_filter = (now - timedelta(days=7)).strftime("%Y-%m-%d")
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("SELECT * FROM sales WHERE to_iso(datetime) >= ? AND amended=0", (date_filter,))
                    sales = c.fetchall()
            else:
                date_filter = (now - timedelta(days=30)).strftime("%Y-%m-%d")
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("SELECT * FROM sales WHERE to_iso(datetime) >= ? AND amended=0", (date_filter,))
                    sales = c.fetchall()

            total_from_sales = sum(s['total'] for s in sales)
            cash_total   = sum(s['total'] for s in sales if s['payment'] == 'CASH')
            qr_total     = sum(s['total'] for s in sales if s['payment'] == 'QR')
            payment_sum  = cash_total + qr_total
            diff         = total_from_sales - payment_sum
            status       = "✅ Balanced" if abs(diff) < 0.01 else f"⚠️ Diff: RM {diff:.2f}"

            report = f"""
{'='*40}
     Payment Reconciliation Report
{'='*40}
Period Sales:     RM {total_from_sales:>10.2f}
{'='*40}
Cash (CASH):      RM {cash_total:>10.2f}
QR Pay:         RM {qr_total:>10.2f}
{'─'*40}
Payment Total:    RM {payment_sum:>10.2f}
{'='*40}
Result: {status}
{'='*40}
Transactions: {len(sales)}
"""
            result_text.insert(tk.END, report)
            result_text.config(state="disabled")

            for s in sales:
                if s['payment'] not in ('CASH', 'QR'):
                    anom_tree.insert("", "end", values=(
                        s['id'], s['datetime'], f"RM {s['total']:.2f}", s['payment']
                    ), tags=('anomaly',))
            anom_tree.tag_configure('anomaly', foreground='red')

        run_reconcile()
        RoundedButton(recon_window, text="🔄 Reconcile Again", command=run_reconcile,
                      bg=self.colors['neutral'], fg="white",
                      font=("Arial", 12, "bold"), padx=20, pady=8, radius=8
                      ).pack(pady=10)

    # ================= REPORTS =================
    def _build_reports_in_frame(self, parent_frame):
        notebook = ttk.Notebook(parent_frame)
        notebook.pack(fill="both", expand=True, padx=6, pady=6)
        self._populate_reports_notebook(notebook, parent_frame)

    def show_reports(self):
        report_window = tk.Toplevel(self.root)
        report_window.title("Reports & Analytics")
        report_window.geometry("1600x900")
        report_window.configure(bg=self.colors['bg'])
        report_window.transient(self.root)

        header = tk.Frame(report_window, bg=self.colors['neutral'], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="📊 Reports & Analytics", font=("Arial", 20, "bold"),
                 bg=self.colors['neutral'], fg=self.colors['light']).pack(pady=12)

        # Increase Notebook Tab size for better usability
        style = ttk.Style(report_window)
        style.configure("Large.TNotebook.Tab", font=("Arial", 12, "bold"), padding=[20, 10])
        
        notebook = ttk.Notebook(report_window, style="Large.TNotebook")
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self._populate_reports_notebook(notebook, report_window)

    def _populate_reports_notebook(self, notebook, parent_win):
        L = (lambda en, zh: zh if self.lang == "zh" else en)
        daily_frame = tk.Frame(notebook, bg=self.colors['bg'])
        notebook.add(daily_frame, text=L("📅 Daily Summary", "📅 每日汇总"))

        tk.Label(daily_frame,
                 text=L("📌 Business day = 06:00 today to 05:59 next day", "📌 营业日 = 当天 06:00 到次日 05:59"),
                 font=("Arial", 10), bg=self.colors['bg'], fg=self.colors['warning']
                 ).pack(anchor="w", padx=20, pady=(8,0))

        period_frame = tk.Frame(daily_frame, bg=self.colors['bg'])
        period_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(period_frame, text=L("Period:", "周期："), font=("Arial", 12, "bold"),
                 bg=self.colors['bg']).pack(side="left", padx=5)
        period_var = tk.StringVar(value="7")
        for days in ["7", "30", "90"]:
            tk.Radiobutton(period_frame, text=L(f"Last {days} days", f"最近 {days} 天"), variable=period_var,
                           value=days, font=("Arial", 11), bg=self.colors['bg'],
                           command=lambda: toggle_report_custom()).pack(side="left", padx=5)
        tk.Radiobutton(period_frame, text=L("Custom", "自定义"), variable=period_var, value="custom",
                       font=("Arial", 11), bg=self.colors['bg'],
                       command=lambda: toggle_report_custom()).pack(side="left", padx=5)

        # Open per-day detail in a separate window for clearer viewing.
        RoundedButton(period_frame, text=L("📄 View Details", "📄 查看明细"),
                      command=lambda: open_day_details_popup(),
                      bg=self.colors['primary'], fg="white",
                      font=("Arial", 11, "bold"), padx=12, pady=4, radius=7
                      ).pack(side="right", padx=5)

        # KPI summary panel at top of Daily Summary
        kpi_frame = tk.Frame(daily_frame, bg=self.colors['bg'])
        kpi_frame.pack(fill="x", padx=20, pady=(5, 10))
        
        kpi_vals = {}
        def create_kpi_card(parent, key, title, bg_color):
            card = tk.Frame(parent, bg=bg_color, padx=12, pady=10)
            card.pack(side="left", padx=6, fill="x", expand=True)
            tk.Label(card, text=title, font=("Arial", 9, "bold"), bg=bg_color, fg="#E2E8F0").pack(anchor="w")
            lbl = tk.Label(card, text="—", font=("Arial", 16, "bold"), bg=bg_color, fg="white")
            lbl.pack(anchor="w", pady=(4, 0))
            kpi_vals[key] = lbl

        create_kpi_card(kpi_frame, 'sales', L("Total Sales", "总营业额"), "#1D4ED8")
        create_kpi_card(kpi_frame, 'trans', L("Transactions", "总单数"), "#0369A1")
        create_kpi_card(kpi_frame, 'cash', L("Cash Payments", "现金收入"), "#0F766E")
        create_kpi_card(kpi_frame, 'qr', L("QR Payments", "QR 收入"), "#7C3AED")

        report_custom_frame = tk.Frame(daily_frame, bg=self.colors['bg'])
        tk.Label(report_custom_frame, text=L("From:", "开始："), font=("Arial", 11), bg=self.colors['bg']).pack(side="left", padx=(10,2))
        report_from_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        DatePicker(
            report_custom_frame, textvariable=report_from_var,
            font=("Arial", 11), width=11, language=self.lang
        ).pack(side="left")
        tk.Label(report_custom_frame, text=L("To:", "结束："), font=("Arial", 11), bg=self.colors['bg']).pack(side="left", padx=(10,2))
        report_to_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        DatePicker(
            report_custom_frame, textvariable=report_to_var,
            font=("Arial", 11), width=11, language=self.lang
        ).pack(side="left")
        tk.Label(
            report_custom_frame,
            text=L("Type or choose", "可输入或选择"),
            font=("Arial", 9), bg=self.colors['bg'], fg="gray"
        ).pack(side="left", padx=5)
        RoundedButton(report_custom_frame, text=L("🔍 Search", "🔍 查询"), command=lambda: load_daily_report(),
                      bg=self.colors['neutral'], fg="white",
                      font=("Arial", 11, "bold"), padx=10, pady=2, radius=6
                      ).pack(side="left", padx=5)

        def toggle_report_custom():
            if period_var.get() == "custom":
                report_custom_frame.pack(fill="x", pady=4)
            else:
                report_custom_frame.pack_forget()
                load_daily_report()

        daily_tree_frame = tk.Frame(daily_frame, bg=self.colors['bg'])
        daily_tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        columns = ("No.", "Date", "Sales", "Trans", "Cash", "QR")
        daily_tree = ttk.Treeview(daily_tree_frame, columns=columns, show="headings", height=12)
        daily_tree.heading("No.", text="#")
        daily_tree.column("No.", width=54, anchor="center")
        daily_tree.heading("Date", text="Date")
        daily_tree.column("Date", width=120, anchor="center")
        daily_tree.heading("Sales", text="Sales")
        daily_tree.column("Sales", width=150, anchor="e")
        daily_tree.heading("Trans", text="Trans")
        daily_tree.column("Trans", width=90, anchor="center")
        daily_tree.heading("Cash", text="Cash")
        daily_tree.column("Cash", width=140, anchor="e")
        daily_tree.heading("QR", text="QR")
        daily_tree.column("QR", width=140, anchor="e")
        daily_tree.tag_configure('odd', background='#F8FAFC')
        daily_tree.tag_configure('even', background='#FFFFFF')
        scrollbar = tk.Scrollbar(daily_tree_frame, orient="vertical", command=daily_tree.yview)
        daily_tree.configure(yscrollcommand=scrollbar.set)
        daily_tree.pack(side="left", fill="both", expand=True)

        scrollbar.pack(side="right", fill="y")

        daily_btn_frame = tk.Frame(daily_frame, bg=self.colors['bg'])
        daily_btn_frame.pack(fill="x", padx=20, pady=(10, 8))

        daily_current_rows = []

        def _selected_biz_date():
            sel = daily_tree.selection()
            if not sel:
                return None
            vals = daily_tree.item(sel[0], "values")
            if not vals or len(vals) < 2:
                return None
            return str(vals[1]).strip()

        def load_day_cashflow_details(biz_date=None):
            if not biz_date:
                return "", []

            with self.db.get_connection() as conn:
                total_sales, cash_sales, qr_sales, card_sales = compute_sales_payment_breakdown(conn, biz_date)
                pos_cash, other_income, expenses = compute_cash_summary_components(conn, biz_date)
                expected_cash = float(pos_cash) + float(other_income) - float(expenses)

                s_dt, e_dt = business_day_range(biz_date)
                c = conn.cursor()
                c.execute(
                    "SELECT COUNT(*) FROM sales WHERE to_iso(datetime) >= ? AND to_iso(datetime) <= ?",
                    (s_dt.strftime("%Y-%m-%d %H:%M:%S"), e_dt.strftime("%Y-%m-%d %H:%M:%S"))
                )
                tx_count = int((c.fetchone() or [0])[0] or 0)

                c.execute(
                    """SELECT datetime, type, source, payment_method, amount, recorded_by, notes
                       FROM cashflow
                       WHERE date=?
                       ORDER BY datetime ASC, id ASC""",
                    (biz_date,)
                )
                rows = c.fetchall()
                
                c.execute("SELECT notes FROM cash_closing WHERE closing_date=? ORDER BY rowid DESC LIMIT 1", (biz_date,))
                closing_note = (c.fetchone() or [""])[0] or ""

            summary_text = (
                f"{biz_date} | "
                f"{L('Sales', '销售')}: {self.currency} {float(total_sales):.2f} "
                f"(CASH {self.currency} {float(cash_sales):.2f} | QR {self.currency} {float(qr_sales):.2f}) | "
                f"{L('Trans', '交易')}: {tx_count} | "
                f"{L('Other Income', '其他收入')}: {self.currency} {float(other_income):.2f} | "
                f"{L('Expenses', '支出')}: {self.currency} {float(expenses):.2f} | "
                f"{L('Expected Cash', '应有现金')}: {self.currency} {float(expected_cash):.2f}"
            )
            metrics = {
                'biz_date': biz_date,
                'sales': float(total_sales),
                'cash_sales': float(cash_sales),
                'qr_sales': float(qr_sales),
                'card_sales': float(card_sales),
                'tx_count': int(tx_count),
                'other_income': float(other_income),
                'expenses': float(expenses),
                'expected_cash': float(expected_cash),
                'closing_note': closing_note
            }
            return summary_text, rows, metrics

        def open_day_details_popup():
            biz_date = _selected_biz_date()
            if not biz_date:
                messagebox.showwarning(L("No Selection", "未选择"),
                                       L("Please select a date row first", "请先选择一行日期"),
                                       parent=parent_win)
                return

            summary_text, rows, metrics = load_day_cashflow_details(biz_date)

            win = tk.Toplevel(parent_win)
            win.title(L(f"Day Details - {biz_date}", f"当日明细 - {biz_date}"))
            win.geometry("1600x900")
            win.configure(bg=self.colors['bg'])
            win.transient(parent_win)

            hdr = tk.Frame(win, bg=self.colors['primary'], height=58)
            hdr.pack(fill="x")
            hdr.pack_propagate(False)
            tk.Label(hdr,
                     text=L("Selected Day Cashflow Details", "已选日期现金流明细"),
                     font=("Arial", 16, "bold"),
                     bg=self.colors['primary'], fg="white").pack(side="left", padx=12, pady=10)

            tk.Label(hdr,
                     text=str(metrics.get('biz_date') or ''),
                     font=("Arial", 11, "bold"),
                     bg=self.colors['primary'], fg="#DBEAFE").pack(side="right", padx=14)

            cards_wrap = tk.Frame(win, bg=self.colors['bg'])
            cards_wrap.pack(fill="x", padx=14, pady=(10, 4))

            def _kpi(parent, title, value, bg, fg="white"):
                card = tk.Frame(parent, bg=bg, width=180, height=64)
                card.pack(side="left", padx=4, pady=2, fill="x", expand=True)
                card.pack_propagate(False)
                tk.Label(card, text=title, font=("Arial", 9, "bold"), bg=bg, fg="#E2E8F0").pack(anchor="w", padx=8, pady=(6, 0))
                tk.Label(card, text=value, font=("Arial", 12, "bold"), bg=bg, fg=fg).pack(anchor="w", padx=8)

            _kpi(cards_wrap, "Sales", f"{self.currency} {metrics['sales']:.2f}", "#1D4ED8")
            _kpi(cards_wrap, "Cash / QR", f"{self.currency} {metrics['cash_sales']:.2f} / {self.currency} {metrics['qr_sales']:.2f}", "#0369A1")
            _kpi(cards_wrap, "Other Income", f"{self.currency} {metrics['other_income']:.2f}", "#0F766E")
            _kpi(cards_wrap, "Expenses", f"{self.currency} {metrics['expenses']:.2f}", "#B91C1C")
            _kpi(cards_wrap, "Expected Cash", f"{self.currency} {metrics['expected_cash']:.2f}", "#7C3AED")

            tk.Label(win, text=summary_text, wraplength=1180,
                     font=("Arial", 9, "bold"),
                     bg=self.colors['bg'], fg=self.colors['text_muted'],
                     justify="left", anchor="w").pack(fill="x", padx=14, pady=(0, 8))

            if metrics.get('closing_note'):
                note_f = tk.Frame(win, bg="#FEF3C7", highlightbackground="#F59E0B", highlightthickness=1)
                note_f.pack(fill="x", padx=14, pady=(0, 10))
                tk.Label(note_f, text=f"📝 Closing Remark: {metrics['closing_note']}", 
                         font=("Arial", 11, "italic", "bold"), bg="#FEF3C7", fg="#92400E",
                         wraplength=1400, justify="left").pack(padx=12, pady=10, anchor="w")

            filter_row = tk.Frame(win, bg=self.colors['bg'])
            filter_row.pack(fill="x", padx=14, pady=(0, 8))
            tk.Label(filter_row, text=L("Filter", "筛选"), font=("Arial", 10, "bold"),
                     bg=self.colors['bg'], fg=self.colors['text']).pack(side="left", padx=(0, 8))

            type_var = tk.StringVar(value="All")
            pay_var = tk.StringVar(value="All")
            ttk.Combobox(filter_row, textvariable=type_var, width=12, state="readonly",
                         values=["All", "Income", "Expense"]).pack(side="left", padx=4)
            ttk.Combobox(filter_row, textvariable=pay_var, width=10, state="readonly",
                         values=["All", "CASH", "QR"]).pack(side="left", padx=4)

            row_count_var = tk.StringVar(value="")
            tk.Label(filter_row, textvariable=row_count_var,
                     font=("Arial", 10, "bold"),
                     bg=self.colors['bg'], fg=self.colors['text_muted']).pack(side="right")

            tree_wrap = tk.Frame(win, bg=self.colors['bg'])
            tree_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 10))
            detail_cols = ("Time", "Type", "Source", "Payment", "Amount", "By", "Notes")
            detail_tree = ttk.Treeview(tree_wrap, columns=detail_cols, show="headings", height=18)
            detail_tree.heading("Time", text="Time")
            detail_tree.column("Time", width=150, anchor="center")
            detail_tree.heading("Type", text="Type")
            detail_tree.column("Type", width=90, anchor="center")
            detail_tree.heading("Source", text="Source")
            detail_tree.column("Source", width=160, anchor="w")
            detail_tree.heading("Payment", text="Pay")
            detail_tree.column("Payment", width=90, anchor="center")
            detail_tree.heading("Amount", text="Amount")
            detail_tree.column("Amount", width=130, anchor="e")
            detail_tree.heading("By", text="By")
            detail_tree.column("By", width=130, anchor="center")
            detail_tree.heading("Notes", text="Notes")
            detail_tree.column("Notes", width=390, anchor="w")
            detail_tree.tag_configure('cf_odd', background='#F8FAFC')
            detail_tree.tag_configure('cf_even', background='#FFFFFF')
            detail_tree.tag_configure('cf_expense', foreground='#B91C1C')

            sb = tk.Scrollbar(tree_wrap, orient="vertical", command=detail_tree.yview)
            detail_tree.configure(yscrollcommand=sb.set)
            detail_tree.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            filtered_rows = []

            def _reload_detail_rows(*_):
                detail_tree.delete(*detail_tree.get_children())
                filtered_rows.clear()
                tf = str(type_var.get() or 'All').strip().upper()
                pf = str(pay_var.get() or 'All').strip().upper()
                shown = 0
                for r in rows:
                    r_type = str(r['type'] or '').strip().upper()
                    r_pay = str(r['payment_method'] or '').strip().upper()
                    if tf != 'ALL' and r_type != tf:
                        continue
                    if pf != 'ALL' and r_pay != pf:
                        continue
                    shown += 1
                    filtered_rows.append(r)
                    row_tag = 'cf_even' if shown % 2 == 0 else 'cf_odd'
                    tags = (row_tag, 'cf_expense') if r_type == 'EXPENSE' else (row_tag,)
                    detail_tree.insert(
                        "", "end",
                        values=(
                            str(r['datetime'] or ''),
                            str(r['type'] or ''),
                            str(r['source'] or ''),
                            str(r['payment_method'] or ''),
                            f"{self.currency} {float(r['amount'] or 0):.2f}",
                            str(r['recorded_by'] or '-'),
                            str(r['notes'] or '-'),
                        ),
                        tags=tags,
                    )
                row_count_var.set(f"Rows: {shown}")

            def _export_day_detail_rows(rows_to_export):
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                from openpyxl.utils import get_column_letter
                if not rows_to_export:
                    messagebox.showwarning(
                        L("No Data", "无数据"),
                        L("No detail rows to export", "没有可导出的明细数据"),
                        parent=win,
                    )
                    return

                from tkinter import filedialog as fd
                default_name = f"day_cashflow_{biz_date}_{datetime.now().strftime('%H%M')}"

                if _OPENPYXL_OK:
                    path = fd.asksaveasfilename(
                        parent=win,
                        initialfile=default_name,
                        defaultextension=".xlsx",
                        filetypes=[("Excel Files", "*.xlsx")],
                        title=L("Save Day Details", "保存当日明细"),
                    )
                    if not path:
                        return
                    try:
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.title = "Day Details"
                        ws.append([
                            "Business Date", "Time", "Type", "Source", "Payment", "Amount", "By", "Notes"
                        ])
                        for r in rows_to_export:
                            ws.append([
                                biz_date,
                                str(r['datetime'] or ''),
                                str(r['type'] or ''),
                                str(r['source'] or ''),
                                str(r['payment_method'] or ''),
                                float(r['amount'] or 0),
                                str(r['recorded_by'] or '-'),
                                str(r['notes'] or '-'),
                            ])
                        for col, width in {
                            "A": 14, "B": 20, "C": 10, "D": 18, "E": 10, "F": 12, "G": 14, "H": 42
                        }.items():
                            ws.column_dimensions[col].width = width
                        wb.save(path)
                        messagebox.showinfo(L("Export Success", "导出成功"),
                                            f"{L('Saved to', '已保存到')}:\n{path}", parent=win)
                    except Exception as e:
                        messagebox.showerror(L("Export Failed", "导出失败"), str(e), parent=win)
                    return

                # CSV fallback if openpyxl not available.
                path = fd.asksaveasfilename(
                    parent=win,
                    initialfile=default_name,
                    defaultextension=".csv",
                    filetypes=[("CSV files", "*.csv")],
                    title=L("Save Day Details", "保存当日明细"),
                )
                if not path:
                    return
                try:
                    import csv
                    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                        w = csv.writer(f)
                        w.writerow(["Business Date", "Time", "Type", "Source", "Payment", "Amount", "By", "Notes"])
                        for r in rows_to_export:
                            w.writerow([
                                biz_date,
                                str(r['datetime'] or ''),
                                str(r['type'] or ''),
                                str(r['source'] or ''),
                                str(r['payment_method'] or ''),
                                f"{float(r['amount'] or 0):.2f}",
                                str(r['recorded_by'] or '-'),
                                str(r['notes'] or '-'),
                            ])
                    messagebox.showinfo(L("Export Success", "导出成功"),
                                        f"{L('Saved to', '已保存到')}:\n{path}", parent=win)
                except Exception as e:
                    messagebox.showerror(L("Export Failed", "导出失败"), str(e), parent=win)

            def export_day_detail():
                _export_day_detail_rows(filtered_rows)

            _reload_detail_rows()
            type_var.trace_add("write", _reload_detail_rows)
            pay_var.trace_add("write", _reload_detail_rows)

            btn_row = tk.Frame(win, bg=self.colors['bg'])
            btn_row.pack(fill="x", padx=14, pady=(0, 10))
            RoundedButton(btn_row, text=L("📊 Export Excel", "📊 导出 Excel"), command=export_day_detail,
                          bg=self.colors['primary'], fg="white",
                          font=("Arial", 11, "bold"), padx=16, pady=6, radius=7
                          ).pack(side="left")
            RoundedButton(btn_row, text=L("✖ Close", "✖ 关闭"), command=win.destroy,
                          bg=self.colors['neutral'], fg="white",
                          font=("Arial", 11, "bold"), padx=16, pady=6, radius=7
                          ).pack(side="right")

        daily_tree.bind("<Double-1>", lambda _e: open_day_details_popup())

        def load_daily_report():
            daily_tree.delete(*daily_tree.get_children())
            daily_current_rows.clear()
            end_date = datetime.now()
            if period_var.get() == "custom":
                try:
                    normalized_from = normalize_date_text(report_from_var.get())
                    normalized_to = normalize_date_text(report_to_var.get())
                    start_date = datetime.strptime(normalized_from, "%Y-%m-%d")
                    end_date = datetime.strptime(normalized_to, "%Y-%m-%d")
                    if end_date < start_date:
                        raise ValueError
                    report_from_var.set(normalized_from)
                    report_to_var.set(normalized_to)
                except ValueError:
                    messagebox.showerror(
                        L("Invalid Date", "日期无效"),
                        L(
                            "Choose valid From and To dates; From cannot be after To.",
                            "请选择有效的开始及结束日期，开始日期不可晚于结束日期。"
                        ),
                        parent=parent_win
                    )
                    return
            else:
                days       = int(period_var.get())
                start_date = end_date - timedelta(days=days)

            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM daily_summary WHERE date >= ? AND date <= ? ORDER BY date DESC",
                          (start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")))
                summaries = c.fetchall()

            total_sales = total_trans = total_cash = total_qr = 0
            for idx, summary in enumerate(summaries, 1):
                s_date     = summary['date']
                s_sales    = float(summary['total_sales']        or 0)
                s_trans    = int(summary['transaction_count']    or 0)
                s_cash     = float(summary['cash_total']         or 0)
                s_qr       = float(summary['qr_total']           or 0)
                row_tag = 'even' if idx % 2 == 0 else 'odd'
                daily_tree.insert("", "end", values=(
                    f"{idx}", s_date, f"{self.currency} {s_sales:.2f}", s_trans,
                    f"{self.currency} {s_cash:.2f}", f"{self.currency} {s_qr:.2f}"
                ), tags=(row_tag,))
                daily_current_rows.append({
                    'Date': s_date,
                    'Sales': s_sales,
                    'Transactions': s_trans,
                    'Cash': s_cash,
                    'QR': s_qr
                })
                total_sales += s_sales; total_trans += s_trans
                total_cash  += s_cash;  total_qr    += s_qr

            kpi_vals['sales'].config(text=f"{self.currency} {total_sales:.2f}")
            kpi_vals['trans'].config(text=f"{total_trans}")
            kpi_vals['cash'].config(text=f"{self.currency} {total_cash:.2f}")
            kpi_vals['qr'].config(text=f"{self.currency} {total_qr:.2f}")

            children = daily_tree.get_children()
            if children:
                daily_tree.selection_set(children[0])
                daily_tree.focus(children[0])

        def export_daily_sheet():
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            if not daily_current_rows:
                messagebox.showwarning(L("No Data", "无数据"), L("No daily summary data to export", "没有可导出的每日汇总数据"), parent=parent_win)
                return

            from tkinter import filedialog as fd
            default_name = f"daily_summary_{datetime.now().strftime('%Y%m%d_%H%M')}"

            if _OPENPYXL_OK:
                file_path = fd.asksaveasfilename(
                    parent=parent_win,
                    initialfile=default_name,
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx")],
                    title=L("Save Daily Summary", "保存每日汇总")
                )
                if not file_path:
                    return
                try:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Daily Summary"
                    ws.append(["Date", "Sales", "Transactions", "Cash", "QR"])
                    for row in daily_current_rows:
                        ws.append([
                            row['Date'],
                            float(row['Sales']),
                            int(row['Transactions']),
                            float(row['Cash']),
                            float(row['QR']),
                        ])
                    for col, width in {"A": 14, "B": 14, "C": 14, "D": 14, "E": 14}.items():
                        ws.column_dimensions[col].width = width
                    wb.save(file_path)
                    messagebox.showinfo(L("Export Success", "导出成功"), f"{L('Saved to', '已保存到')}:\n{file_path}", parent=parent_win)
                except Exception as e:
                    messagebox.showerror(L("Export Failed", "导出失败"), str(e), parent=parent_win)
                return

            # Fallback when openpyxl is missing: export CSV so user can still open in spreadsheet apps.
            file_path = fd.asksaveasfilename(
                parent=parent_win,
                initialfile=default_name,
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title=L("Save Daily Summary", "保存每日汇总")
            )
            if not file_path:
                return
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Date", "Sales", "Transactions", "Cash", "QR"])
                    for row in daily_current_rows:
                        writer.writerow([
                            row['Date'],
                            f"{float(row['Sales']):.2f}",
                            int(row['Transactions']),
                            f"{float(row['Cash']):.2f}",
                            f"{float(row['QR']):.2f}",
                        ])
                messagebox.showinfo(L("Export Success", "导出成功"), f"{L('Saved to', '已保存到')}:\n{file_path}", parent=parent_win)
            except Exception as e:
                messagebox.showerror(L("Export Failed", "导出失败"), str(e), parent=parent_win)

        RoundedButton(daily_btn_frame, text=L("🔄 Refresh", "🔄 刷新"), command=load_daily_report,
                  bg=self.colors['neutral'], fg="white",
                  font=("Arial", 11, "bold"), padx=16, pady=7, radius=7
                  ).pack(side="left", padx=5)

        RoundedButton(daily_btn_frame, text=L("📥 Export Table", "📥 导出表格"), command=export_daily_sheet,
                  bg=self.colors['primary'], fg="white",
                  font=("Arial", 11, "bold"), padx=16, pady=7, radius=7
                  ).pack(side="left", padx=5)

        items_frame = tk.Frame(notebook, bg=self.colors['bg'])
        notebook.add(items_frame, text=L("🏆 Top Items", "🏆 热销商品"))
        tk.Label(items_frame, text=L("Best Selling Items", "热销商品排行"), font=("Arial", 16, "bold"),
                 bg=self.colors['bg']).pack(pady=20)

        items_tree_frame = tk.Frame(items_frame, bg=self.colors['bg'])
        items_tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        columns2 = ("Rank", "Item", "Qty Sold", "Revenue", "Category")
        items_tree = ttk.Treeview(items_tree_frame, columns=columns2, show="headings", height=20)
        for col in columns2:
            items_tree.heading(col, text=col)
            items_tree.column(col, width=150)
        items_tree.tag_configure('row_odd', background='#F8FAFC')
        items_tree.tag_configure('row_even', background='#FFFFFF')
        items_tree.tag_configure('out_item', foreground='#DC2626')
        scrollbar2 = tk.Scrollbar(items_tree_frame, orient="vertical", command=items_tree.yview)
        items_tree.configure(yscrollcommand=scrollbar2.set)
        items_tree.pack(side="left", fill="both", expand=True)
        scrollbar2.pack(side="right", fill="y")

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT items FROM sales")
            all_sales = c.fetchall()
            c.execute("SELECT name, stock, track_stock, active FROM items")
            item_stock_map = {
                str(r['name'] or '').strip(): {
                    'stock': int(r['stock'] or 0),
                    'track_stock': int(r['track_stock'] or 0),
                    'active': int(r['active'] or 0),
                }
                for r in c.fetchall()
                if str(r['name'] or '').strip()
            }

        item_stats = {}
        for sale in all_sales:
            items_list = json.loads(sale[0])
            for name, price, qty, cat in items_list:
                if name not in item_stats:
                    item_stats[name] = {'qty': 0, 'revenue': 0, 'category': cat}
                item_stats[name]['qty']     += qty
                item_stats[name]['revenue'] += float(Decimal(str(price)) * int(qty))

        def _resolve_stock_info_for_report(raw_name):
            key = str(raw_name or '').strip()
            if not key:
                return None

            stock_info = item_stock_map.get(key)
            if stock_info:
                return stock_info

            # Variant sales are stored as "Base Item (Variant)", while stock lives
            # on the base item row. Fallback avoids false "missing" red labels.
            if "(" in key and key.endswith(")"):
                base_name = key[:key.rfind("(")].strip()
                if base_name:
                    return item_stock_map.get(base_name)
            return None

        sorted_items = sorted(item_stats.items(), key=lambda x: x[1]['qty'], reverse=True)
        normal_rows = []
        removed_rows = []

        for name, stats in sorted_items:
            stock_info = _resolve_stock_info_for_report(name)
            # "Removed" means no longer in menu or marked inactive.
            is_removed = (not stock_info) or (int(stock_info.get('active', 0)) == 0)
            row = (name, stats, is_removed)
            if is_removed:
                removed_rows.append(row)
            else:
                normal_rows.append(row)

        merged_rows = normal_rows + removed_rows
        for rank, (name, stats, is_removed) in enumerate(merged_rows, 1):
            zebra_tag = 'row_even' if rank % 2 == 0 else 'row_odd'
            row_tags = (zebra_tag, 'out_item') if is_removed else (zebra_tag,)
            items_tree.insert("", "end", values=(
                f"#{rank}", name, stats['qty'],
                f"{self.currency} {stats['revenue']:.2f}", stats['category']
            ), tags=row_tags)

        stock_frame = tk.Frame(notebook, bg=self.colors['bg'])
        notebook.add(stock_frame, text=L("📦 Stock Status", "📦 库存状态"))
        tk.Label(stock_frame, text=L("Current Inventory", "当前库存"), font=("Arial", 16, "bold"),
                 bg=self.colors['bg']).pack(pady=20)

        stock_summary_frame = tk.Frame(stock_frame, bg=self.colors['card'], height=46)
        stock_summary_frame.pack(fill="x", padx=20, pady=(0, 8))
        stock_summary_frame.pack_propagate(False)
        stock_summary_label = tk.Label(stock_summary_frame, text="", font=("Arial", 11, "bold"),
                           bg=self.colors['card'], fg=self.colors['light'])
        stock_summary_label.pack(side="left", padx=14, pady=10)

        def show_stock_movement_history():
            history = tk.Toplevel(win)
            history.title(L("Stock Movement History", "库存变动记录"))
            history.geometry("1180x700")
            history.configure(bg=self.colors['bg'])
            history.transient(win)

            filters = tk.Frame(history, bg=self.colors['bg'])
            filters.pack(fill="x", padx=16, pady=12)
            tk.Label(filters, text=L("Item:", "商品："), font=("Arial", 11, "bold"),
                     bg=self.colors['bg']).pack(side="left")
            item_filter = tk.StringVar()
            filter_entry = tk.Entry(filters, textvariable=item_filter, font=("Arial", 11), width=28)
            filter_entry.pack(side="left", padx=8)
            reconciliation_label = tk.Label(filters, text="", font=("Arial", 10, "bold"),
                                            bg=self.colors['bg'])
            reconciliation_label.pack(side="right", padx=8)

            cols = ("Time", "Item", "Change", "Before", "After", "Source", "Reference", "User", "Notes")
            movement_tree = ttk.Treeview(history, columns=cols, show="headings", height=24)
            widths = {"Time":155, "Item":150, "Change":70, "Before":65, "After":65,
                      "Source":135, "Reference":80, "User":90, "Notes":260}
            for col in cols:
                movement_tree.heading(col, text=col)
                movement_tree.column(col, width=widths[col], anchor="w" if col in ("Item", "Notes") else "center")
            movement_tree.tag_configure("in", foreground=self.colors['success'])
            movement_tree.tag_configure("out", foreground=self.colors['danger'])
            movement_tree.pack(fill="both", expand=True, padx=16, pady=(0, 16))

            def load_movements():
                movement_tree.delete(*movement_tree.get_children())
                keyword = item_filter.get().strip()
                sql = """SELECT datetime,item_name,quantity_delta,stock_before,stock_after,
                                source_type,reference_id,recorded_by,notes
                         FROM stock_movements"""
                params = []
                if keyword:
                    sql += " WHERE item_name LIKE ?"
                    params.append(f"%{keyword}%")
                sql += " ORDER BY id DESC LIMIT 2000"
                with self.db.get_connection() as conn:
                    rows = conn.execute(sql, params).fetchall()
                    mismatches = conn.execute("""
                        SELECT i.name,i.stock,sm.stock_after
                        FROM items i
                        JOIN stock_movements sm ON sm.id=(
                            SELECT id FROM stock_movements
                            WHERE item_id=i.id ORDER BY id DESC LIMIT 1
                        )
                        WHERE i.track_stock=1 AND i.stock<>sm.stock_after
                    """).fetchall()
                if mismatches:
                    reconciliation_label.config(
                        text=L(f"⚠ {len(mismatches)} unexplained balance change(s)",
                               f"⚠ {len(mismatches)} 个无法解释的库存变化"),
                        fg=self.colors['danger'])
                else:
                    reconciliation_label.config(text=L("✅ Ledger reconciled", "✅ 库存账目一致"),
                                                fg=self.colors['success'])
                for r in rows:
                    delta = int(r['quantity_delta'])
                    movement_tree.insert("", "end", values=(
                        r['datetime'], r['item_name'], f"{delta:+d}", r['stock_before'], r['stock_after'],
                        r['source_type'], r['reference_id'] or "-", r['recorded_by'] or "-", r['notes'] or "-"
                    ), tags=("in" if delta > 0 else "out",))

            tk.Button(filters, text=L("Search", "搜索"), command=load_movements,
                      bg=self.colors['primary'], fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=4)
            tk.Button(filters, text=L("Refresh", "刷新"), command=load_movements,
                      bg=self.colors['neutral'], fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=4)
            filter_entry.bind("<Return>", lambda _e: load_movements())
            load_movements()

        tk.Button(stock_summary_frame, text=L("View Movements", "查看库存变动"),
                  command=show_stock_movement_history, bg=self.colors['primary'], fg="white",
                  font=("Arial", 10, "bold")).pack(side="right", padx=12, pady=7)

        stock_tree_frame = tk.Frame(stock_frame, bg=self.colors['bg'])
        stock_tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        columns3 = ("Item", "Category", "Stock", "Status", "Price")
        stock_tree = ttk.Treeview(stock_tree_frame, columns=columns3, show="headings", height=20)
        col_w = {"Item": 200, "Category": 110, "Stock": 100, "Status": 140, "Price": 110}
        for col in columns3:
            stock_tree.heading(col, text=col)
            stock_tree.column(col, width=col_w.get(col, 100), anchor="center")
        stock_tree.tag_configure('row_odd', background='#F8FAFC')
        stock_tree.tag_configure('row_even', background='#FFFFFF')
        scrollbar3 = tk.Scrollbar(stock_tree_frame, orient="vertical", command=stock_tree.yview)
        stock_tree.configure(yscrollcommand=scrollbar3.set)
        stock_tree.pack(side="left", fill="both", expand=True)
        scrollbar3.pack(side="right", fill="y")

        stock_items = []

        def load_stock_status():
            nonlocal stock_items
            stock_tree.delete(*stock_tree.get_children())
            with self.db.get_connection() as conn:
                stock_items = conn.execute("""
                    SELECT name, category, stock, low_stock_threshold, price
                    FROM items
                    WHERE active=1 AND track_stock=1
                    ORDER BY category, name
                """).fetchall()

            for idx, (name, cat, stock, threshold, price) in enumerate(stock_items, 1):
                if stock <= 0:
                    status = "❌ Out of Stock"; tag = 'out'
                elif stock <= threshold:
                    status = "⚠️ Low Stock";    tag = 'low'
                else:
                    status = "✅ In Stock";      tag = ''
                zebra_tag = 'row_even' if idx % 2 == 0 else 'row_odd'
                row_tags = (zebra_tag, tag) if tag else (zebra_tag,)
                stock_tree.insert("", "end", values=(
                    name, cat, stock, status, f"{self.currency} {price:.2f}"
                ), tags=row_tags)

            tracked_count = len(stock_items)
            out_count = sum(1 for _n, _c, s, _t, _p in stock_items if int(s or 0) <= 0)
            low_count = sum(1 for _n, _c, s, t, _p in stock_items if int(s or 0) > 0 and int(s or 0) <= int(t or 0))
            total_units = sum(int(s or 0) for _n, _c, s, _t, _p in stock_items)
            refreshed = datetime.now().strftime("%H:%M:%S")
            stock_summary_label.config(
                text=(f"{L('Tracked Items', '追踪商品')}: {tracked_count} | {L('Total Units', '总库存数')}: {total_units} | "
                      f"{L('Low Stock', '低库存')}: {low_count} | {L('Out of Stock', '缺货')}: {out_count} | "
                      f"{L('Updated', '更新时间')}: {refreshed}")
            )

        load_stock_status()

        def export_stock_status():
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            if not stock_items:
                messagebox.showwarning(L("No Data", "无数据"), L("No stock status data to export", "没有可导出的数据"), parent=parent_win)
                return
            
            from tkinter import filedialog as fd
            default_name = f"stock_status_{datetime.now().strftime('%Y%m%d_%H%M')}"
            
            if _OPENPYXL_OK:
                file_path = fd.asksaveasfilename(
                    parent=parent_win,
                    initialfile=default_name,
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx")],
                    title=L("Save Stock Status", "保存库存状态")
                )
                if not file_path:
                    return
                try:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Stock Status"
                    ws.append([L("Item", "商品"), L("Category", "分类"), L("Stock", "库存"), L("Status", "状态"), L("Price", "单价")])
                    for name, cat, stock, threshold, price in stock_items:
                        if stock <= 0:
                            status_str = L("Out of Stock", "缺货")
                        elif stock <= threshold:
                            status_str = L("Low Stock", "低库存")
                        else:
                            status_str = L("In Stock", "有货")
                        ws.append([name, cat, stock, status_str, float(price)])
                    for col, width in {"A": 22, "B": 14, "C": 12, "D": 14, "E": 12}.items():
                        ws.column_dimensions[col].width = width
                    wb.save(file_path)
                    messagebox.showinfo(L("Export Success", "导出成功"), f"{L('Saved to', '已保存到')}:\n{file_path}", parent=parent_win)
                except Exception as e:
                    messagebox.showerror(L("Export Failed", "导出失败"), str(e), parent=parent_win)
                return

            # CSV fallback
            file_path = fd.asksaveasfilename(
                parent=parent_win,
                initialfile=default_name,
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title=L("Save Stock Status", "保存库存状态")
            )
            if not file_path:
                return
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Item", "Category", "Stock", "Status", "Price"])
                    for name, cat, stock, threshold, price in stock_items:
                        if stock <= 0:
                            status_str = "Out of Stock"
                        elif stock <= threshold:
                            status_str = "Low Stock"
                        else:
                            status_str = "In Stock"
                        writer.writerow([name, cat, stock, status_str, f"{float(price):.2f}"])
                messagebox.showinfo(L("Export Success", "导出成功"), f"{L('Saved to', '已保存到')}:\n{file_path}", parent=parent_win)
            except Exception as e:
                messagebox.showerror(L("Export Failed", "导出失败"), str(e), parent=parent_win)

        stock_btn_frame = tk.Frame(stock_frame, bg=self.colors['bg'])
        stock_btn_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        RoundedButton(stock_btn_frame, text=L("📥 Export Table", "📥 导出表格"), 
                      command=export_stock_status, bg=self.colors['primary'], fg="white",
                      font=("Arial", 11, "bold"), padx=16, pady=7, radius=7
                      ).pack(side="left", padx=5)
        RoundedButton(stock_btn_frame, text=L("🔄 Refresh Now", "🔄 立即刷新"),
                      command=load_stock_status, bg=self.colors['neutral'], fg="white",
                      font=("Arial", 11, "bold"), padx=16, pady=7, radius=7
                      ).pack(side="left", padx=5)

        def refresh_stock_when_tab_open(_event=None):
            try:
                if notebook.nametowidget(notebook.select()) == stock_frame:
                    load_stock_status()
            except Exception:
                pass

        notebook.bind("<<NotebookTabChanged>>", refresh_stock_when_tab_open, add="+")

        daily_items_frame = tk.Frame(notebook, bg=self.colors['bg'])
        notebook.add(daily_items_frame, text=L("📊 Item Sales Breakdown", "📊 商品销售明细"))

        tk.Label(daily_items_frame,
                 text=L("📌 Business Day = 06:00 to 05:59 next day", "📌 营业日 = 06:00 到次日 05:59"),
                 font=("Arial", 10), bg=self.colors['bg'], fg=self.colors['warning']
                 ).pack(anchor="w", padx=20, pady=(6,0))

        di_filter_frame = tk.Frame(daily_items_frame, bg=self.colors['bg'])
        di_filter_frame.pack(fill="x", padx=20, pady=6)
        tk.Label(di_filter_frame, text=L("Date:", "日期："), font=("Arial", 12, "bold"),
                 bg=self.colors['bg']).pack(side="left", padx=5)
        di_period_var = tk.StringVar(value="month")
        for text, val in [
            (L("Today", "今天"), "today"),
            (L("Yesterday", "昨天"), "yesterday"),
            (L("This Week", "本周"), "week"),
            (L("This Month", "本月"), "month"),
            (L("By Date", "按日期"), "date")
        ]:
            tk.Radiobutton(di_filter_frame, text=text, variable=di_period_var, value=val,
                           font=("Arial", 11), bg=self.colors['bg'],
                           command=lambda: toggle_di_date_mode()).pack(side="left", padx=8)

        di_date_frame = tk.Frame(daily_items_frame, bg=self.colors['bg'])
        tk.Label(di_date_frame, text=L("Business Date:", "营业日期："), font=("Arial", 11, "bold"),
                 bg=self.colors['bg']).pack(side="left", padx=(24, 4))
        di_date_var = tk.StringVar(value=today_biz_date())
        DatePicker(
            di_date_frame, textvariable=di_date_var,
            font=("Arial", 11), width=11, language=self.lang
        ).pack(side="left")
        tk.Label(
            di_date_frame, text=L("Type or choose", "可输入或选择"),
            font=("Arial", 9), bg=self.colors['bg'], fg="gray"
        ).pack(side="left", padx=6)
        RoundedButton(di_date_frame, text=L("🔍 View Date", "🔍 查看日期"), command=lambda: load_daily_items(),
                      bg=self.colors['neutral'], fg="white",
                      font=("Arial", 10, "bold"), padx=10, pady=3, radius=6
                      ).pack(side="left", padx=6)

        tk.Label(di_filter_frame, text=L("Category:", "类别："), font=("Arial", 12, "bold"),
                 bg=self.colors['bg']).pack(side="left", padx=(20, 5))
        di_cat_var = tk.StringVar(value="ALL")
        for text, val in [
            (L("All", "全部"),"ALL"),
            (L("🍺 Beer", "🍺 啤酒"),"Beer"),
            (L("🍵 Teh", "🍵 茶"),"Teh"),
            (L("☕ Kopi", "☕ 咖啡"),"Kopi"),
            (L("🍛 Food", "🍛 食物"),"Food"),
            (L("🥤 Cans", "🥤 罐装饮料"),"Cans"),
            (L("🧃 Drinks", "🧃 饮料"),"Drinks")
        ]:
            tk.Radiobutton(di_filter_frame, text=text, variable=di_cat_var, value=val,
                           font=("Arial", 10), bg=self.colors['bg'],
                           command=lambda: load_daily_items()).pack(side="left", padx=4)

        di_tree_frame = tk.Frame(daily_items_frame, bg=self.colors['bg'])
        di_tree_frame.pack(fill="both", expand=True, padx=20, pady=5)
        di_cols = ("Item", "Category", "Sold Qty", "Price", "Revenue")
        di_tree = ttk.Treeview(di_tree_frame, columns=di_cols, show="headings", height=18)
        col_widths_di = {"Item":240,"Category":110,"Sold Qty":100,"Price":100,"Revenue":120}
        for col in di_cols:
            di_tree.heading(col, text=col)
            di_tree.column(col, width=col_widths_di.get(col, 100), anchor="center")
        di_tree.tag_configure('row_odd', background='#F8FAFC')
        di_tree.tag_configure('row_even', background='#FFFFFF')
        di_tree.tag_configure('beer_row', background='#FFF3CD', foreground='#856404')
        di_tree.tag_configure('zero_row', foreground='#999999')
        di_sb = tk.Scrollbar(di_tree_frame, orient="vertical", command=di_tree.yview)
        di_tree.configure(yscrollcommand=di_sb.set)
        di_tree.pack(side="left", fill="both", expand=True)
        di_sb.pack(side="right", fill="y")

        di_summary_frame = tk.Frame(daily_items_frame, bg=self.colors['card'], height=45)
        di_summary_frame.pack(fill="x", pady=5)
        di_summary_frame.pack_propagate(False)
        di_summary_label = tk.Label(di_summary_frame, text="", font=("Arial", 12, "bold"),
                                    bg=self.colors['card'], fg=self.colors['light'])
        di_summary_label.pack(pady=8)

        di_btn_frame = tk.Frame(daily_items_frame, bg=self.colors['bg'])
        di_btn_frame.pack(fill="x", padx=20, pady=8)

        di_current_data = []

        def toggle_di_date_mode():
            if di_period_var.get() == "date":
                di_date_frame.pack(fill="x", padx=20, pady=(0, 4))
            else:
                di_date_frame.pack_forget()
            load_daily_items()

        def load_daily_items():
            di_tree.delete(*di_tree.get_children())
            di_current_data.clear()
            now    = datetime.now()
            period = di_period_var.get()

            def _normalize_daily_cat(raw_cat):
                cat_txt = str(raw_cat or '').strip().lower()
                if cat_txt in ('can drinks', 'cans', 'can', 'tin'):
                    return 'Cans'
                if cat_txt in ('drink', 'drinks', 'beverage', 'cold drinks'):
                    return 'Drinks'
                if cat_txt in ('beer',):
                    return 'Beer'
                if cat_txt in ('teh', 'tea'):
                    return 'Teh'
                if cat_txt in ('kopi', 'coffee'):
                    return 'Kopi'
                if cat_txt in ('food',):
                    return 'Food'
                return str(raw_cat or '').strip() or 'Other'

            start_biz = None
            end_biz = None
            if period == "today":
                biz       = today_biz_date()
                start_biz = end_biz = biz
                label_str = f"{biz} (Biz Day)"
            elif period == "yesterday":
                yesterday_biz = get_business_date(now - timedelta(days=1))
                biz       = yesterday_biz
                start_biz = end_biz = biz
                label_str = f"{biz} (Yesterday Biz Day)"
            elif period == "date":
                try:
                    date_txt = normalize_date_text(di_date_var.get())
                    di_date_var.set(date_txt)
                except ValueError:
                    messagebox.showerror(
                        L("Invalid Date", "日期无效"),
                        L("Please type or choose a valid date.", "请输入或选择有效日期。"),
                        parent=parent_win
                    )
                    return
                biz = date_txt
                start_biz = end_biz = biz
                label_str = f"{biz} (Selected Biz Day)"
            elif period == "week":
                biz_7     = get_business_date(now - timedelta(days=7))
                biz       = None
                start_biz = biz_7
                end_biz = today_biz_date()
                label_str = f"{biz_7} ~ {today_biz_date()}"
            else:
                biz_30    = get_business_date(now - timedelta(days=30))
                biz       = None
                start_biz = biz_30
                end_biz = today_biz_date()
                label_str = f"{biz_30} ~ {today_biz_date()}"

            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT datetime, items FROM sales")
                raw_sales_rows = c.fetchall()
                sales_rows = []
                for r in raw_sales_rows:
                    sale_biz = biz_date_of_sale(str(r['datetime'] or ''))
                    if start_biz <= sale_biz <= end_biz:
                        sales_rows.append(r)

                c.execute("SELECT name, category, price FROM items")
                item_info = {row['name']: {'category': row['category'], 'price': row['price']}
                             for row in c.fetchall()}

            item_sold = {}
            for row in sales_rows:
                try:
                    items_list = json.loads(row['items'])
                except Exception:
                    try:
                        items_list = json.loads(row[1])
                    except Exception:
                        items_list = []
                for name, price, qty, cat in items_list:
                    norm_cat = _normalize_daily_cat(cat)
                    if name not in item_sold:
                        item_sold[name] = {'qty': 0, 'revenue': 0, 'category': norm_cat, 'price': price}
                    item_sold[name]['qty']     += qty
                    item_sold[name]['revenue'] += float(Decimal(str(price)) * int(qty))

            cat_filter  = di_cat_var.get()
            total_qty = total_revenue = 0
            beer_cats = {'Beer'}

            shown_idx = 0
            for name, stats in sorted(item_sold.items(), key=lambda x: x[1]['qty'], reverse=True):
                cat = stats['category']
                if cat_filter != "ALL" and cat != cat_filter: continue
                shown_idx += 1

                price      = item_info.get(name, {}).get('price', stats['price'])

                total_qty     += stats['qty']
                total_revenue += stats['revenue']

                zebra_tag = 'row_even' if shown_idx % 2 == 0 else 'row_odd'
                row_tag = 'beer_row' if cat in beer_cats else ''
                row_tags = (zebra_tag, row_tag) if row_tag else (zebra_tag,)
                di_tree.insert("", "end", values=(
                    name, cat, stats['qty'],
                    f"{self.currency} {price:.2f}",
                    f"{self.currency} {stats['revenue']:.2f}"
                ), tags=row_tags)

                di_current_data.append({
                    'Item': name, 'Category': cat, 'Sold Qty': stats['qty'],
                    'Price': price, 'Revenue': stats['revenue']
                })

            di_summary_label.config(
                text=(f"📅 {label_str}  |  Sold: {total_qty}  |  "
                      f"Revenue: {self.currency} {total_revenue:.2f}"))

            if not di_current_data:
                di_summary_label.config(
                    text=(f"📅 {label_str}  |  {L('No sales data for this filter. Try changing Date or Category.', '当前筛选没有销售数据，请更换日期或类别。')}"))

        def export_csv():
            if not di_current_data:
                messagebox.showwarning(L("No Data", "无数据"), L("No data to export", "没有可导出的数据"), parent=parent_win); return
            import csv
            from tkinter import filedialog
            filepath = filedialog.asksaveasfilename(
                parent=parent_win, defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                initialfile=f"daily_sales_{datetime.now().strftime('%Y%m%d')}.csv")
            if not filepath: return
            try:
                with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=['Item','Category','Sold Qty','Price','Revenue'])
                    writer.writeheader()
                    writer.writerows(di_current_data)
                messagebox.showinfo(L("Export Success", "导出成功"), f"{L('Saved to', '已保存到')}:\n{filepath}", parent=parent_win)
            except Exception as e:
                messagebox.showerror(L("Error", "错误"), f"{L('Export failed', '导出失败')}: {str(e)}", parent=parent_win)


        RoundedButton(di_btn_frame, text=L("🔄 Refresh", "🔄 刷新"), 
                      command=load_daily_items, bg=self.colors['neutral'], fg="white",
                      font=("Arial", 11, "bold"), padx=16, pady=7, radius=7
                      ).pack(side="left", padx=5)

        RoundedButton(di_btn_frame, text=L("📥 Export Table", "📥 导出表格"), 
                      command=export_csv, bg=self.colors['primary'], fg="white",
                      font=("Arial", 11, "bold"), padx=16, pady=7, radius=7
                      ).pack(side="left", padx=5)

        tk.Label(di_btn_frame, text=L("🍺 Beer highlighted yellow", "🍺 啤酒高亮黄色"),
                 font=("Arial", 10), bg=self.colors['bg'], fg=self.colors['secondary']
                 ).pack(side="left", padx=15)

        # ── Audit Log ─────────────────────────────────────────────────────
        audit_frame = tk.Frame(notebook, bg=self.colors['bg'])
        notebook.add(audit_frame, text=L("📋 Audit Log", "📋 审计日志"))

        audit_tree_frame = tk.Frame(audit_frame, bg=self.colors['bg'])
        audit_tree_frame.pack(fill="both", expand=True, padx=16, pady=(10, 4))

        audit_cols = ("ID", "Time", "Action", "Details", "User")
        audit_tree = ttk.Treeview(audit_tree_frame, columns=audit_cols, show="headings", height=20)
        acw = {"ID": 55, "Time": 160, "Action": 90, "Details": 420, "User": 100}
        for col in audit_cols:
            audit_tree.heading(col, text=col)
            audit_tree.column(col, width=acw.get(col, 120), anchor="center")
        audit_tree.column("Details", anchor="w")
        
        audit_tree.tag_configure('row_odd', background='#F8FAFC')
        audit_tree.tag_configure('row_even', background='#FFFFFF')
        audit_tree.tag_configure('SALE',   foreground=self.colors['success'])
        audit_tree.tag_configure('EDIT',   foreground=self.colors['warning'])
        audit_tree.tag_configure('REFUND', foreground=self.colors['danger'])
        
        asb = tk.Scrollbar(audit_tree_frame, orient="vertical", command=audit_tree.yview)
        audit_tree.configure(yscrollcommand=asb.set)
        audit_tree.pack(side="left", fill="both", expand=True)
        asb.pack(side="right", fill="y")

        audit_filter_frame = tk.Frame(audit_frame, bg=self.colors['bg'])
        audit_filter_frame.pack(fill="x", padx=16, pady=(10, 4))
        
        audit_summary_lbl = tk.Label(audit_frame, text="", font=("Arial", 11, "bold"),
                                      bg=self.colors['bg'], fg=self.colors['text_sub'])
        audit_summary_lbl.pack(anchor="w", padx=16, pady=(0, 4))

        audit_period_var = tk.StringVar(value="today")
        audit_action_var = tk.StringVar(value="All")

        def load_audit_log(filter_type="today", action_filter="All"):
            audit_tree.delete(*audit_tree.get_children())
            now = datetime.now()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                sql = "SELECT * FROM audit_log WHERE 1=1"
                params = []
                
                if filter_type == "today":
                    date_filter = now.strftime("%d-%m-%Y")
                    sql += " AND timestamp LIKE ?"
                    params.append(date_filter + "%")
                elif filter_type == "week":
                    week_ago = (now - timedelta(days=7)).strftime("%Y-%m-%d 00:00:00")
                    # Stored timestamps are DD-MM-YYYY, so plain text comparison
                    # gives incorrect results across days/months.
                    sql += " AND to_iso(timestamp) >= ?"
                    params.append(week_ago)
                
                if action_filter and action_filter != "All":
                    sql += " AND action = ?"
                    params.append(action_filter)
                
                sql += " ORDER BY id DESC"
                if filter_type == "all":
                    sql += " LIMIT 2000"
                
                c.execute(sql, tuple(params))
                rows = c.fetchall()
            for idx, r in enumerate(rows, 1):
                action = r['action']
                zebra_tag = 'row_even' if idx % 2 == 0 else 'row_odd'
                audit_tree.insert("", "end", values=(
                    r['id'], r['timestamp'], action,
                    r['details'] or "-", r['user'] or "-"
                ), tags=(zebra_tag, action))
            audit_summary_lbl.config(text=f"{len(rows)} " + L("records", "条记录"))

        def refresh_audit_actions():
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT DISTINCT action FROM audit_log ORDER BY action")
                actions = ["All"] + [r[0] for r in c.fetchall() if r[0]]
                audit_action_cb['values'] = actions

        tk.Label(audit_filter_frame, text=L("Filter:", "筛选："), font=("Arial", 11, "bold"),
                 bg=self.colors['bg']).pack(side="left", padx=(0, 8))
        for _text, _val in [(L("Today", "今天"),"today"), (L("This Week", "本周"),"week"), (L("All", "全部"),"all")]:
            tk.Radiobutton(audit_filter_frame, text=_text, variable=audit_period_var, value=_val,
                           font=("Arial", 11), bg=self.colors['bg'],
                           command=lambda: load_audit_log(audit_period_var.get(), audit_action_var.get())).pack(side="left", padx=6)
        
        tk.Label(audit_filter_frame, text=" |  " + L("Action:", "动作："), font=("Arial", 11),
                 bg=self.colors['bg']).pack(side="left", padx=(10, 4))
        audit_action_cb = ttk.Combobox(audit_filter_frame, textvariable=audit_action_var, state="readonly", width=18)
        audit_action_cb.pack(side="left", padx=4)
        audit_action_cb.bind("<<ComboboxSelected>>", lambda e: load_audit_log(audit_period_var.get(), audit_action_var.get()))
        
        refresh_audit_actions()

        load_audit_log("today")

        load_daily_report()
        load_daily_items()

    # ================= CASHIER CASH (simplified) =================
    def show_cashier_cash(self, start_tab=0):
        """Compatibility route: all cash work now uses one consistent screen."""
        return self.show_cash_closing(cashier_mode=True)

        # Legacy implementation retained below for old source compatibility;
        # it is unreachable because every caller is routed above.
        C  = self.colors
        L = (lambda en, zh: zh if self.lang == "zh" else en)
        win = tk.Toplevel(self.root)
        win.title(L("💰 Cash", "💰 现金"))
        win.geometry("980x700")
        win.minsize(900, 620)
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()

        # ── Shared helpers ────────────────────────────────────────────────
        def get_today():  return today_biz_date()

        def get_opening_cash():
            biz = get_today()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT opening_cash FROM cash_opening WHERE date=? ORDER BY id DESC LIMIT 1", (biz,))
                row = c.fetchone()
                return float(row['opening_cash']) if row else 0.0

        def get_summary_nums():
            """Calculate summary with cashflow + legacy fallback compatibility."""
            biz = get_today()
            with self.db.get_connection() as conn:
                pos_cash, income, expense = compute_cash_summary_components(conn, biz)
            opening  = get_opening_cash()
            expected = opening + pos_cash + income - expense
            return opening, pos_cash, income, expense, expected

        # ── Notebook ──────────────────────────────────────────────────────
        hdr = tk.Frame(win, bg=C['primary'], height=64)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text=L("💰  Cash — ", "💰 现金 — ") + today_biz_date(),
                 font=("Arial", 15, "bold"), bg=C['primary'], fg="white"
             ).pack(side="left", padx=20, pady=(8, 2))
        tk.Label(hdr, text=L("Cashier quick panel", "收银员快捷面板"),
             font=("Arial", 10), bg=C['primary'], fg="#DBEAFE"
             ).pack(side="left", pady=(28, 8), padx=(0, 4))

        style = ttk.Style(win)
        style.configure("CashierTabs.TNotebook", background=C['bg'], borderwidth=0)
        style.configure("CashierTabs.TNotebook.Tab",
                font=("Arial", 11, "bold"),
            padding=(18, 10),
            background="#E5E7EB",
            foreground="#334155")
        style.map("CashierTabs.TNotebook.Tab",
              background=[("selected", "#FFFFFF"), ("active", "#D1D5DB")],
              foreground=[("selected", C['primary'])])

        body = tk.Frame(win, bg=C['bg'])
        body.pack(fill="both", expand=True, padx=14, pady=12)

        nb_wrap = tk.Frame(body, bg="#FFFFFF", highlightbackground="#D1D5DB", highlightthickness=1)
        nb_wrap.pack(fill="both", expand=True)

        nb = ttk.Notebook(nb_wrap, style="CashierTabs.TNotebook")
        nb.pack(fill="both", expand=True, padx=6, pady=6)

        # ════════════════════════════════════════════════════════════════
        # TAB 1 — Today Summary
        # ════════════════════════════════════════════════════════════════
        t1 = tk.Frame(nb, bg=C['bg'])
        nb.add(t1, text=L("  📊 Summary  ", "  📊 汇总  "))

        cards_row = tk.Frame(t1, bg=C['bg'])
        cards_row.pack(fill="x", padx=20, pady=(18, 8))

        card_defs = [
            (L("🪙 Opening", "🪙 开盘"),      C['warning'],  0),
            (L("💵 Cash Sales", "💵 现金销售"),   C['success'],  1),
            (L("➕ Cash Receipts", "➕ 其他现金收入"), "#2563EB",     2),
            (L("💸 Cash Out", "💸 现金转出"),     C['danger'],   3),
        ]
        card_labels = []
        for i, (title, color, _) in enumerate(card_defs):
            cards_row.columnconfigure(i, weight=1)
            # Photo 1 style: Solid vibrant backgrounds, no top bar, white text
            card = tk.Frame(cards_row, bg=color, highlightbackground="#FFFFFF", highlightthickness=0)
            card.grid(row=0, column=i, padx=6, pady=4, sticky="nsew")

            tk.Label(card, text=title, font=("Arial", 11, "bold"),
                     bg=color, fg="white").pack(pady=(12, 2))
            lbl = tk.Label(card, text="RM 0.00", font=("Arial", 20, "bold"),
                           bg=color, fg="white")
            lbl.pack(pady=(0, 14))
            card_labels.append(lbl)

        # Big expected cash banner
        exp_banner = tk.Frame(t1, bg="#334155", height=82)
        exp_banner.pack(fill="x", padx=20, pady=(4,16)); exp_banner.pack_propagate(False)
        exp_lbl = tk.Label(exp_banner, text=L("Expected Cash:  RM 0.00", "应有现金:  RM 0.00"),
                           font=("Arial", 26, "bold"), bg="#334155", fg="white")
        exp_lbl.pack(expand=True)

        # Opening cash quick setup for cashier
        open_row = tk.Frame(t1, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        open_row.pack(fill="x", padx=20, pady=(0, 10))

        tk.Label(open_row, text=L("Opening Cash (Today):", "今日开盘现金："),
                 font=("Arial", 11, "bold"), bg=C['panel']).grid(row=0, column=0, padx=(14, 6), pady=10, sticky="w")

        open_amt_var = tk.StringVar()
        open_amt_entry = tk.Entry(open_row, textvariable=open_amt_var, font=("Arial", 14, "bold"),
                                  width=10, justify="right")
        open_amt_entry.grid(row=0, column=1, padx=6, pady=10)

        open_status_lbl = tk.Label(open_row, text="", font=("Arial", 10, "bold"),
                                   bg=C['panel'], fg=C['text_muted'])
        open_status_lbl.grid(row=0, column=2, padx=(12, 6), pady=10, sticky="w")

        def refresh_opening_status():
            opening_now = get_opening_cash()
            if opening_now > 0:
                open_status_lbl.config(
                    text=L(
                        f"Current: RM {opening_now:.2f} (locked, admin can edit)",
                        f"当前：RM {opening_now:.2f}（已锁定，仅管理员可修改）"
                    ),
                    fg=C['success']
                )
            else:
                open_status_lbl.config(text=L("Not set", "未设置"), fg=C['warning'])

        def save_opening_cash():
            try:
                amt = float(open_amt_var.get())
                if amt < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(L("Error", "错误"), L("Please enter valid opening cash", "请输入有效开盘金额"), parent=win)
                return

            existing = get_opening_cash()
            if existing > 0:
                if not self.admin_mode:
                    messagebox.showwarning(
                        L("Locked", "已锁定"),
                        L(
                            f"Opening cash already set: RM {existing:.2f}\nOnly Admin can change it.",
                            f"开盘现金已设置：RM {existing:.2f}\n仅管理员可以修改。"
                        ),
                        parent=win
                    )
                    return

                if not messagebox.askyesno(
                    L("Replace Opening Cash", "替换开盘现金"),
                    L(f"Opening cash already set: RM {existing:.2f}\nReplace with RM {amt:.2f}?",
                      f"开盘现金已设置：RM {existing:.2f}\n要改为 RM {amt:.2f} 吗？"),
                    parent=win
                ):
                    return

            def _do_save(staff):
                now = datetime.now()
                biz = get_today()
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("DELETE FROM cash_opening WHERE date=?", (biz,))
                    c.execute(
                        "INSERT INTO cash_opening (date, opening_cash, opened_by, opened_at) VALUES (?,?,?,?)",
                        (biz, amt, staff['name'], now.strftime("%Y-%m-%d %H:%M:%S"))
                    )
                    c.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                              (now.strftime("%d-%m-%Y %H:%M:%S"), 'CASH_OPEN',
                               f'Opening Cash RM {amt:.2f}', staff['name']))

                open_amt_var.set("")
                refresh_opening_status()
                refresh_summary()
                messagebox.showinfo(
                    L("Saved", "已保存"),
                    L(f"Opening cash RM {amt:.2f} saved by {staff['name']}",
                      f"开盘现金 RM {amt:.2f} 已保存，记录人：{staff['name']}"),
                    parent=win
                )

            if self.current_staff:
                _do_save(self.current_staff)
            else:
                self._staff_login_popup(callback=_do_save,
                                        title=L("👤 Log in to set opening cash", "👤 登录后设置开盘现金"),
                                        parent=win)

        RoundedButton(open_row, text=L("💾 Set Opening", "💾 设置开盘"),
                      command=save_opening_cash,
                      bg="#F59E0B", fg="white",
                      font=("Arial", 11, "bold"), padx=12, pady=8, radius=8
                      ).grid(row=0, column=3, padx=(6, 14), pady=10)

        open_amt_entry.bind("<Return>", lambda e: save_opening_cash())

        def refresh_summary():
            opening, pos_cash, income, expense, expected = get_summary_nums()
            vals = [opening, pos_cash, income, expense]
            for lbl, v in zip(card_labels, vals):
                lbl.config(text=f"RM {v:.2f}")
            exp_lbl.config(text=L(f"Expected Cash:  RM {expected:.2f}", f"应有现金:  RM {expected:.2f}"))
            refresh_opening_status()

        refresh_summary()

        tk.Label(t1,
             text=L("Tip: Use quick actions below to record income/expense immediately.", "提示：使用下面快捷按钮可立即记录收入/支出。"),
             font=("Arial", 10), bg=C['bg'], fg=C['text_muted']).pack(anchor="w", padx=22, pady=(0, 6))

        # Quick action buttons
        qa_row = tk.Frame(t1, bg=C['bg'])
        qa_row.pack(fill="x", padx=20, pady=4)
        qa_row.columnconfigure(0, weight=1); qa_row.columnconfigure(1, weight=1)

        RoundedButton(qa_row, text=L("➕  Add Expense", "➕  添加支出"),
                      command=lambda: nb.select(1),
                      bg=C['danger'], fg="white",
                      font=("Arial", 14, "bold"), padx=0, pady=20, radius=12
                      ).grid(row=0, column=0, padx=6, sticky="ew")
        RoundedButton(qa_row, text=L("➕  Add Income", "➕  添加收入"),
                      command=lambda: nb.select(1),
                      bg="#2563EB", fg="white",
                      font=("Arial", 14, "bold"), padx=0, pady=20, radius=12
                      ).grid(row=0, column=1, padx=6, sticky="ew")

        # ════════════════════════════════════════════════════════════════
        # TAB 2 — Quick Expense / Income
        # ════════════════════════════════════════════════════════════════
        t2 = tk.Frame(nb, bg=C['bg'])
        nb.add(t2, text=L("  💸 Expense / Income  ", "  💸 支出 / 收入  "))

        def _quick_entry(parent, entry_type, bg_color, icon):
            entry_label = L("Expense", "支出") if entry_type == "Expense" else L("Income", "收入")
            card = tk.Frame(parent, bg=C['panel'],
                            highlightbackground=C['border'], highlightthickness=1)
            card.pack(fill="x", padx=20, pady=10)
            tk.Label(card, text=f"{icon}  {entry_label}",
                     font=("Arial", 13, "bold"), bg=C['panel']).pack(anchor="w", padx=16, pady=(12,4))

            row1 = tk.Frame(card, bg=C['panel'])
            row1.pack(fill="x", padx=16, pady=4)

            tk.Label(row1, text=L("Amount  RM:", "金额  RM："), font=("Arial", 11), bg=C['panel']).pack(side="left")
            amt_var = tk.StringVar()
            amt_e   = tk.Entry(row1, textvariable=amt_var, font=("Arial", 14, "bold"),
                               width=10, justify="right")
            amt_e.pack(side="left", padx=(6,20))

            if entry_type == "Expense":
                cats = [
                    "Ingredients / 食材",
                    "Utilities / 杂费",
                    "Staff Wages / 员工薪水",
                    "Repairs / 维修",
                    "Packaging / 包装",
                    "Transport / 运输",
                    "Beverage Cost / 饮料成本",
                    "Maintenance / 保养",
                    "License / 执照"
                ]
            else:
                cats = ["Rental / 租金", "Sales / 销售", "Other Income / 其他收入"]

            tk.Label(row1, text=L("Category:", "类别："), font=("Arial", 11), bg=C['panel']).pack(side="left")
            cat_var = tk.StringVar(value=cats[0])
            ttk.Combobox(row1, textvariable=cat_var, values=cats,
                         font=("Arial", 11), width=22, state="readonly").pack(side="left", padx=6)

            row2 = tk.Frame(card, bg=C['panel'])
            row2.pack(fill="x", padx=16, pady=4)
            tk.Label(row2, text=L("Notes:", "备注："), font=("Arial", 11), bg=C['panel']).pack(side="left")
            notes_var = tk.StringVar()
            tk.Entry(row2, textvariable=notes_var, font=("Arial", 11), width=38).pack(side="left", padx=6)

            status_lbl = tk.Label(card, text="", font=("Arial", 10),
                                  bg=C['panel'], fg=C['success'])
            status_lbl.pack(anchor="w", padx=16, pady=(0,4))

            def save_entry():
                try:
                    amt = float(amt_var.get())
                    if amt <= 0: raise ValueError
                except ValueError:
                    status_lbl.config(text=L("⚠ Enter a valid amount", "⚠ 请输入有效金额"), fg=C['danger']); return

                def save_with_method(payment_method):
                    payment_method = str(payment_method or "CASH").strip().upper()
                    if payment_method not in ("CASH", "QR"):
                        payment_method = "CASH"

                    def _do_save(staff):
                        self.current_staff = staff
                        biz     = get_today()
                        now     = datetime.now()
                        t       = "expense" if entry_type == "Expense" else "income"
                        cf_type = "Expense" if t == "expense" else "Income"
                        
                        display_src = cat_var.get()
                        cf_src = display_src.split(" / ")[0] if " / " in display_src else display_src
                        
                        with self.db.get_connection() as conn:
                            c = conn.cursor()
                            c.execute("INSERT INTO cash_transactions (datetime,type,category,amount,description,recorded_by,date,payment_method) VALUES (?,?,?,?,?,?,?,?)",
                                      (now.strftime("%Y-%m-%d %H:%M:%S"), t, cf_src,
                                       amt, notes_var.get().strip(), staff['name'], biz,
                                       payment_method))
                            ref_id = c.lastrowid
                            self._record_cashflow(conn, cf_type, cf_src, ref_id,
                                                  amt, payment_method, staff['name'],
                                                  notes_var.get().strip(), dt=now)
                            c.execute("INSERT INTO audit_log (timestamp,action,details,user) VALUES (?,?,?,?)",
                                      (now.strftime("%d-%m-%Y %H:%M:%S"),
                                       t.upper(),
                                       f"{cf_src} RM{amt:.2f} {payment_method} {notes_var.get()}",
                                       staff['name']))

                        amt_var.set(""); notes_var.set("")
                        status_lbl.config(
                            text=L(
                                f"✅ Saved RM{amt:.2f} by {payment_method} — {staff['name']}",
                                f"✅ 已保存 RM{amt:.2f}（{payment_method}）— {staff['name']}"
                            ),
                            fg=C['success']
                        )
                        refresh_summary()
                        if payment_method == "CASH":
                            open_cash_drawer()

                    if self.current_staff:
                        _do_save(self.current_staff)
                    else:
                        self._staff_login_popup(
                            callback=_do_save,
                            title=L("👤 Log in to save entry", "👤 登录后保存记录"),
                            parent=win
                        )

                if entry_type == "Income":
                    display_src = cat_var.get()
                    income_source = (
                        display_src.split(" / ")[0]
                        if " / " in display_src else display_src
                    )
                    self._show_cash_qr_payment_choice(
                        win, amt,
                        L("Other Income Payment", "其他收入付款"),
                        income_source,
                        save_with_method
                    )
                else:
                    save_with_method("CASH")

            button_text = (
                L("💳  Receive Payment", "💳  收款")
                if entry_type == "Income"
                else L("💾  Save Expense", "💾  保存支出")
            )
            RoundedButton(card, text=button_text,
                  command=save_entry, bg=bg_color, fg="white",
                  font=("Arial", 12, "bold"), padx=16, pady=10, radius=9
                  ).pack(anchor="w", padx=16, pady=(4,12))
            return amt_e

        _quick_entry(t2, "Expense", C['danger'],  "💸")
        _quick_entry(t2, "Income",  "#2563EB",    "💵")

        # ════════════════════════════════════════════════════════════════
        # TAB 3 — Close Day
        # ════════════════════════════════════════════════════════════════
        t4 = tk.Frame(nb, bg=C['bg'])
        nb.add(t4, text=L("  🌙 Close Day  ", "  🌙 结账日结  "))

        close_card = tk.Frame(t4, bg=C['panel'],
                              highlightbackground=C['border'], highlightthickness=1)
        close_card.pack(padx=40, pady=30, fill="x")

        tk.Label(close_card, text=L("🌙  End of Day Cash Check", "🌙  日结现金核对"),
                 font=("Arial", 16, "bold"), bg=C['panel']).pack(pady=(20,10), padx=20, anchor="w")

        exp_row = tk.Frame(close_card, bg=C['panel'])
        exp_row.pack(fill="x", padx=20, pady=6)
        tk.Label(exp_row, text=L("Expected Cash:", "应有现金："), font=("Arial", 13), bg=C['panel'], width=18, anchor="w").pack(side="left")
        close_exp_lbl = tk.Label(exp_row, text="RM 0.00",
                                  font=("Arial", 18, "bold"), bg=C['panel'], fg=C['success'])
        close_exp_lbl.pack(side="left")

        act_row2 = tk.Frame(close_card, bg=C['panel'])
        act_row2.pack(fill="x", padx=20, pady=6)
        tk.Label(act_row2, text=L("Actual Cash:", "实点现金："), font=("Arial", 13), bg=C['panel'], width=18, anchor="w").pack(side="left")
        actual_var = tk.StringVar()
        actual_entry = tk.Entry(act_row2, textvariable=actual_var,
                                font=("Arial", 18, "bold"), width=12, justify="right")
        actual_entry.pack(side="left")

        diff_row = tk.Frame(close_card, bg=C['panel'])
        diff_row.pack(fill="x", padx=20, pady=6)
        tk.Label(diff_row, text=L("Difference:", "差额："), font=("Arial", 13), bg=C['panel'], width=18, anchor="w").pack(side="left")
        diff_lbl = tk.Label(diff_row, text="—",
                             font=("Arial", 18, "bold"), bg=C['panel'], fg=C['text_muted'])
        diff_lbl.pack(side="left")

        def on_actual_change(*_):
            try:
                actual = float(actual_var.get())
                _, _, _, _, expected = get_summary_nums()
                diff = actual - expected
                color = C['success'] if abs(diff) < 0.01 else (C['warning'] if abs(diff) < 5 else C['danger'])
                sign  = "+" if diff >= 0 else ""
                diff_lbl.config(text=f"{sign}RM {diff:.2f}", fg=color)
            except ValueError:
                diff_lbl.config(text="—", fg=C['text_muted'])

        actual_var.trace_add("write", on_actual_change)

        def refresh_close_expected():
            _, _, _, _, expected = get_summary_nums()
            close_exp_lbl.config(text=f"RM {expected:.2f}")

        def confirm_close():
            try:
                actual = float(actual_var.get())
            except ValueError:
                messagebox.showerror(L("Error", "错误"), L("Enter actual cash amount", "请输入实点现金金额"), parent=win); return
            opening, pos_cash, income, expense, expected = get_summary_nums()
            diff = actual - expected
            sign = "+" if diff >= 0 else ""
            if not messagebox.askyesno(L("Confirm Close", "确认日结"),
                    L(
                        f"Confirm end of day?\n\nExpected: RM {expected:.2f}\nActual: RM {actual:.2f}\nDifference: {sign}RM {diff:.2f}",
                        f"确认结束营业日？\n\n应有现金：RM {expected:.2f}\n实点现金：RM {actual:.2f}\n差额：{sign}RM {diff:.2f}"
                    ),
                    parent=win):
                return

            def _do_close(staff):
                self.current_staff = staff
                now = datetime.now()
                biz = get_today()
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute(
                        """INSERT INTO cash_closing
                           (closing_date, opening_cash, pos_cash, other_income, expenses,
                            expected_cash, actual_cash, difference, cashier, notes, closed_at)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            biz, float(opening), float(pos_cash), float(income), float(expense),
                            float(expected), float(actual), float(diff), staff['name'],
                            "Cashier quick close", now.strftime("%Y-%m-%d %H:%M:%S")
                        )
                    )
                    c.execute("INSERT INTO audit_log (timestamp,action,details,user) VALUES (?,?,?,?)",
                              (now.strftime("%d-%m-%Y %H:%M:%S"), "DAY_CLOSE",
                               f"Expected RM{expected:.2f}  Actual RM{actual:.2f}  Diff {sign}RM{diff:.2f}",
                               staff['name']))
                messagebox.showinfo(L("✅ Day Closed", "✅ 日结完成"),
                    L(
                        f"Day closed by {staff['name']}\nExpected: RM {expected:.2f}\nActual: RM {actual:.2f}\nDiff: {sign}RM {diff:.2f}",
                        f"日结人：{staff['name']}\n应有现金：RM {expected:.2f}\n实点现金：RM {actual:.2f}\n差额：{sign}RM {diff:.2f}"
                    ),
                    parent=win)
                actual_var.set("")

            if self.current_staff:
                _do_close(self.current_staff)
            else:
                self._staff_login_popup(callback=_do_close,
                                        title=L("👤 Log in to close day", "👤 登录后执行日结"),
                                        parent=win)

        RoundedButton(close_card, text=L("🌙  Confirm Close Day", "🌙  确认结账日结"),
                      command=confirm_close,
                      bg=C['primary'], fg="white",
                      font=("Arial", 14, "bold"), padx=0, pady=16, radius=12
                      ).pack(fill="x", padx=20, pady=(12,24))

        def on_tab_change(event):
            tab = nb.index(nb.select())
            if tab == 0: refresh_summary()
            elif tab == 2: refresh_close_expected()
        nb.bind("<<NotebookTabChanged>>", on_tab_change)

        try:
            nb.select(int(start_tab))
        except Exception:
            pass

    def show_cashier_rent(self):
        C = self.colors
        L = (lambda en, zh: zh if self.lang == "zh" else en)
        win = tk.Toplevel(self.root)
        win.title(L("🏪 Collect Rent", "🏪 收取租金"))
        win.geometry("1100x700")
        win.minsize(1000, 600)
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()

        hdr = tk.Frame(win, bg=C['primary'], height=58)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=L("🏪 Collect Rent", "🏪 收取租金"),
                 font=("Arial", 15, "bold"), bg=C['primary'], fg="white").pack(side="left", padx=16, pady=12)
        tk.Label(hdr, text=today_biz_date(),
                 font=("Arial", 11), bg=C['primary'], fg="#DBEAFE").pack(side="right", padx=16, pady=16)

        body_outer = tk.Frame(win, bg=C['bg'])
        body_outer.pack(fill="both", expand=True, padx=12, pady=10)

        # Left panel (tenant list)
        left_panel = tk.Frame(body_outer, bg=C['bg'])
        left_panel.pack(side="left", fill="both", expand=True)

        rent_canvas = tk.Canvas(left_panel, bg=C['bg'], highlightthickness=0)
        rent_sb = tk.Scrollbar(left_panel, orient="vertical", command=rent_canvas.yview)
        rent_inner = tk.Frame(rent_canvas, bg=C['bg'])
        rent_inner.bind("<Configure>", lambda e: rent_canvas.configure(scrollregion=rent_canvas.bbox("all")))
        _rw = rent_canvas.create_window((0, 0), window=rent_inner, anchor="nw")
        rent_canvas.configure(yscrollcommand=rent_sb.set)
        rent_canvas.bind("<Configure>", lambda e: rent_canvas.itemconfig(_rw, width=e.width))
        rent_canvas.bind("<MouseWheel>", lambda e: rent_canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        rent_canvas.pack(side="left", fill="both", expand=True)
        rent_sb.pack(side="right", fill="y")

        # Right sidebar (today's payments)
        side_panel = tk.Frame(body_outer, bg=C['panel'], width=280, highlightbackground=C['border'], highlightthickness=1)
        side_panel.pack_propagate(False)
        side_panel.pack(side="right", fill="both", padx=(10, 0))

        side_hdr = tk.Frame(side_panel, bg=C['primary'], height=42)
        side_hdr.pack(fill="x"); side_hdr.pack_propagate(False)
        tk.Label(side_hdr, text=L("📋 Today's Payments", "📋 今日已收"), 
                 font=("Arial", 11, "bold"), bg=C['primary'], fg="white").pack(pady=10)

        side_canvas = tk.Canvas(side_panel, bg=C['panel'], highlightthickness=0)
        side_sb = tk.Scrollbar(side_panel, orient="vertical", command=side_canvas.yview)
        side_inner = tk.Frame(side_canvas, bg=C['panel'])
        side_inner.bind("<Configure>", lambda e: side_canvas.configure(scrollregion=side_canvas.bbox("all")))
        _sw = side_canvas.create_window((0, 0), window=side_inner, anchor="nw")
        side_canvas.configure(yscrollcommand=side_sb.set)
        side_canvas.bind("<Configure>", lambda e: side_canvas.itemconfig(_sw, width=e.width))
        side_canvas.pack(side="left", fill="both", expand=True)
        side_sb.pack(side="right", fill="y")

        def _insert_rent_payment_compat(c2, ten, days, amount, biz, now, staff_name, payment_method):
            values_map = {
                "tenant_id": ten['id'], "tenant_name": ten['name'], "stall_no": ten['stall_no'] or "",
                "days": days, "rate": ten['daily_rate'], "amount": amount,
                "pay_date": biz, "datetime": now.strftime("%Y-%m-%d %H:%M:%S"),
                "recorded_by": staff_name, "payment_method": payment_method,
                "notes": f"{ten['name']} +{days:.0f}d"
            }
            c2.execute("PRAGMA table_info(rent_payments)")
            cols = {str(r['name']) for r in c2.fetchall()}
            insert_cols = [k for k in values_map.keys() if k in cols]
            placeholders = ",".join(["?"] * len(insert_cols))
            sql = f"INSERT INTO rent_payments ({','.join(insert_cols)}) VALUES ({placeholders})"
            c2.execute(sql, tuple(values_map[k] for k in insert_cols))

        def load_today_payments():
            for w in side_inner.winfo_children(): w.destroy()
            biz = today_biz_date()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM rent_payments WHERE pay_date=? AND voided=0 ORDER BY datetime DESC", (biz,))
                rows = c.fetchall()
            
            if not rows:
                tk.Label(side_inner, text=L("No payments today.", "今日暂无记录"), 
                         bg=C['panel'], fg=C['text_muted'], font=("Arial", 10)).pack(pady=30)
                return

            for i, r in enumerate(rows):
                entry = tk.Frame(side_inner, bg=C['bg'] if i % 2 == 0 else C['panel'], padx=8, pady=8)
                entry.pack(fill="x")
                
                lbl_f = tk.Frame(entry, bg=entry.cget("bg"))
                lbl_f.pack(side="left", fill="both", expand=True)
                
                tk.Label(lbl_f, text=f"Stall {r['stall_no']}", font=("Arial", 9, "bold"), 
                         bg=entry.cget("bg"), fg=C['primary'], anchor="w").pack(fill="x")
                tk.Label(lbl_f, text=r['tenant_name'], font=("Arial", 10), 
                         bg=entry.cget("bg"), fg=C['text'], anchor="w").pack(fill="x")
                
                amt_f = tk.Frame(entry, bg=entry.cget("bg"))
                amt_f.pack(side="right")
                
                tk.Label(amt_f, text=f"RM {float(r['amount'] or 0):.2f}", font=("Arial", 10, "bold"), 
                         bg=entry.cget("bg"), fg=C['success'], anchor="e").pack()
                payment_method = (
                    str(r['payment_method'] or 'CASH').strip().upper()
                    if 'payment_method' in r.keys() else 'CASH'
                )
                tk.Label(amt_f, text=payment_method, font=("Arial", 8, "bold"),
                         bg=entry.cget("bg"),
                         fg="#0369A1" if payment_method == "QR" else C['text_muted'],
                         anchor="e").pack()
                t_str = r['datetime'][11:16] if r['datetime'] else ""
                tk.Label(amt_f, text=t_str, font=("Arial", 8), 
                         bg=entry.cget("bg"), fg=C['text_muted'], anchor="e").pack()
                
                # Void button in Today's Payments list
                def make_void(pr=r): return lambda: void_today_payment(pr)
                RoundedButton(amt_f, text=L("Void","撤回"), command=make_void(), 
                              bg="#EF4444", fg="white", font=("Arial", 7, "bold"), 
                              padx=4, pady=2, radius=4).pack(side="top", pady=(4,0))

        def void_today_payment(p):
            pid = p['id']
            def do_void(staff):
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("UPDATE rent_payments SET voided=1 WHERE id=?", (pid,))
                    
                    # Recalculate based on non-voided payments for this tenant
                    tenant_id = p['tenant_id']
                    c.execute("SELECT * FROM rent_tenants WHERE id=?", (tenant_id,))
                    tenant = c.fetchone()
                    if not tenant: return
                    
                    c.execute("SELECT SUM(days) FROM rent_payments WHERE tenant_id=? AND COALESCE(voided,0)=0", (tenant_id,))
                    total_days = c.fetchone()[0] or 0
                    
                    start_val = tenant['start_date']
                    if start_val:
                        try:
                            start_dt = datetime.strptime(start_val, "%Y-%m-%d")
                            new_until = (start_dt + timedelta(days=total_days)).strftime("%Y-%m-%d")
                        except: new_until = tenant['paid_until_date']
                    else: new_until = tenant['paid_until_date']

                    c.execute("UPDATE rent_tenants SET total_days_paid=?, paid_until_date=? WHERE id=?", 
                             (total_days, new_until, tenant_id))
                    
                    # Remove from cashflow
                    c.execute("DELETE FROM cashflow WHERE source='Stall Rental' AND reference_id=?", (pid,))
                    
                    # Audit Log
                    c.execute("INSERT INTO audit_log (timestamp,action,details,user) VALUES (?,?,?,?)",
                             (datetime.now().strftime("%d-%m-%Y %H:%M:%S"), 'VOID_RENT_CASHIER',
                              f"Void Rent #{pid} for {tenant['name']} RM{float(p['amount'] or 0):.2f}", staff['name']))
                
                messagebox.showinfo(L("Voided","已退款"), L(f"Payment #{pid} voided successfully.", f"支付记录 #{pid} 已成功撤回。"), parent=win)
                load_rent_rows()
                load_today_payments()
            
            if messagebox.askyesno(L("Void Payment","撤销支付"), L(f"Void payment #{pid} (RM {float(p['amount'] or 0):.2f}) for {p['tenant_name']}?", f"确认撤销 {p['tenant_name']} 的支付记录 #{pid} (RM {float(p['amount'] or 0):.2f})?"), parent=win):
                self._staff_login_popup(callback=do_void, title=L("Admin PIN Required","需要管理员密码"), parent=win)


        def load_rent_rows():
            for w in rent_inner.winfo_children(): w.destroy()
            biz = today_biz_date(); now = datetime.now()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM rent_tenants WHERE active=1 ORDER BY stall_no")
                tenants = c.fetchall()
            
            for ten in tenants:
                row = tk.Frame(rent_inner, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
                row.pack(fill="x", padx=20, pady=8)
                
                info = tk.Frame(row, bg=C['panel'], padx=16, pady=12)
                info.pack(side="left", fill="both", expand=True)
                
                tk.Label(info, text=f"Stall {ten['stall_no'] or '?'}", font=("Arial", 11, "bold"), bg=C['panel'], fg=C['primary']).pack(anchor="w")
                tk.Label(info, text=ten['name'], font=("Arial", 13, "bold"), bg=C['panel'], fg=C['text']).pack(anchor="w")
                if ten['notes']:
                    tk.Label(info, text=f"{ten['notes']}", font=("Arial", 9, "italic"), bg=C['panel'], fg="#64748B").pack(anchor="w", pady=(0, 2))
                
                paid_until = ten['paid_until_date'] or L("Never", "从未支付")
                tk.Label(info, text=f"{L('Paid Until:','支付至:')} {paid_until}", font=("Arial", 10), bg=C['panel'], fg=C['text_muted']).pack(anchor="w", pady=(4,0))

                # Status Calculation
                today = datetime.now().date()
                if not ten['paid_until_date']:
                    status = L("⚠ No payment yet", "⚠ 尚未支付")
                    color = "#DC2626"
                else:
                    try:
                        until = datetime.strptime(ten['paid_until_date'], "%Y-%m-%d").date()
                        diff = (until - today).days
                        if diff < 0:
                            status = L(f"⚠ Unpaid {abs(diff)}d", f"⚠ 欠费 {abs(diff)}天")
                            color = "#DC2626"
                        elif diff == 0:
                            status = L("⚠ Due Today", "⚠ 今日到期")
                            color = "#D97706"
                        else:
                            status = L(f"✅ {diff}d left", f"✅ 剩余 {diff}天")
                            color = "#16A34A"
                    except:
                        status = ""
                        color = C['text_muted']
                
                if status:
                    tk.Label(info, text=status, font=("Arial", 10, "bold"), bg=C['panel'], fg=color).pack(anchor="w", pady=(0,4))
                
                # Payment Section
                act = tk.Frame(row, bg=C['panel'], padx=16)
                act.pack(side="right", fill="y")
                
                is_monthly = (ten['rental_type'] == 'Monthly')
                rate_lbl_text = f"RM {ten['daily_rate']:.2f}/month" if is_monthly else f"RM {ten['daily_rate']:.2f}/day"
                tk.Label(act, text=rate_lbl_text, font=("Arial", 10, "italic"), bg=C['panel']).pack(pady=(12,2))
                
                days_var = tk.StringVar(value="1")
                e = tk.Entry(act, textvariable=days_var, font=("Arial", 13, "bold"), width=4, justify="center")
                e.pack(side="left", padx=10, pady=10)
                
                def collect_rent(t=ten, dv=days_var, is_m=is_monthly):
                    try:
                        units = float(dv.get())
                        if units <= 0: raise ValueError
                    except: messagebox.showerror(L("Error", "错误"), L("Invalid quantity", "数量无效。"), parent=win); return
                    
                    amt = units * t['daily_rate']
                    unit_str = "months" if is_m else "days"

                    def continue_with_payment(payment_method):
                        if not messagebox.askyesno(
                                L("Confirm Rent Payment", "确认租金付款"),
                                L(
                                    f"Receive RM {amt:.2f} by {payment_method} for {units} {unit_str} from {t['name']}?",
                                    f"确认以 {payment_method} 收取 {t['name']} 的租金 RM {amt:.2f}？"
                                ),
                                parent=win):
                            return

                        def do_collect(staff):
                            biz = today_biz_date(); now = datetime.now()
                            with self.db.get_connection() as conn:
                                c = conn.cursor()
                                _insert_rent_payment_compat(
                                    c, t, units, amt, biz, now, staff['name'], payment_method
                                )
                                
                                # Update tenant
                                from datetime import timedelta
                                base = t['paid_until_date']
                                start_date = t['start_date']
                                if base:
                                    current_paid = base
                                elif start_date:
                                    current_paid = start_date
                                else:
                                    current_paid = biz
                                try:
                                    d_obj = datetime.strptime(current_paid, "%Y-%m-%d")
                                    if is_m:
                                        import calendar
                                        from datetime import date
                                        m_count = int(units)
                                        month = d_obj.month - 1 + m_count
                                        year = d_obj.year + month // 12
                                        month = month % 12 + 1
                                        day = min(d_obj.day, calendar.monthrange(year, month)[1])
                                        new_paid = date(year, month, day).strftime("%Y-%m-%d")
                                    else:
                                        new_paid = (d_obj + timedelta(days=units)).strftime("%Y-%m-%d")
                                except: new_paid = biz
                                c.execute("UPDATE rent_tenants SET paid_until_date=?, total_days_paid=total_days_paid+?, last_payment_date=? WHERE id=?", (new_paid, units, biz, t['id']))
                                
                                self._record_cashflow(
                                    conn, "Income", "Stall Rental", c.lastrowid, amt,
                                    payment_method, staff['name'],
                                    f"Rent {t['name']} Stall {t['stall_no']}", dt=now
                                )
                                c.execute(
                                    "INSERT INTO audit_log (timestamp,action,details,user) VALUES (?,?,?,?)",
                                    (now.strftime("%d-%m-%Y %H:%M:%S"), "RENT_COLLECT",
                                     f"RM{amt:.2f} {payment_method} from {t['name']} ({units}{'m' if is_m else 'd'})",
                                     staff['name'])
                                )
                            
                            if payment_method == "CASH":
                                open_cash_drawer()
                            messagebox.showinfo(
                                L("Payment Complete", "付款完成"),
                                L(
                                    f"Received RM {amt:.2f} by {payment_method}\nPaid until: {new_paid}",
                                    f"已通过 {payment_method} 收取 RM {amt:.2f}\n租金支付至：{new_paid}"
                                ),
                                parent=win
                            )
                            load_rent_rows()
                            load_today_payments()

                        self._staff_login_popup(
                            callback=do_collect,
                            title=L("Log in to collect", "登录以收取租金"),
                            parent=win
                        )

                    self._show_cash_qr_payment_choice(
                        win, amt,
                        L("Rent Payment", "租金付款"),
                        L(
                            f"{t['name']} · Stall {t['stall_no'] or '?'} · {units} {unit_str}",
                            f"{t['name']} · 档口 {t['stall_no'] or '?'} · 数量 {units}"
                        ),
                        continue_with_payment
                    )

                RoundedButton(
                    act, text=L("Make Payment", "付款"), command=collect_rent,
                    bg=C['success'], fg="white", font=("Arial", 11, "bold"),
                    padx=16, pady=8, radius=8
                ).pack(side="left", pady=10)

        load_rent_rows()
        load_today_payments()

    def manage_stalls(self):
        C = self.colors
        L = (lambda en, zh: zh if self.lang == "zh" else en)
        win = tk.Toplevel(self.root)
        win.title("🏪 Stall Management")
        win.geometry("1200x800")
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()

        def get_today(): return today_biz_date()

        hdr = tk.Frame(win, bg=C['header'], height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="🏪 Stall Management", font=("Arial", 16, "bold"), bg=C['header'], fg='white').pack(side="left", padx=20, pady=10)
        RoundedButton(hdr, text=L("✖ Close", "✖ 关闭"), command=win.destroy,
                      bg=C['danger'], fg="white", font=("Arial", 11, "bold"),
                      padx=20, pady=10, radius=8).pack(side="right", padx=10, pady=6)

        # RENT TAB — flat layout
        # ═══════════════════════════════════════════════════════════════
        tab_rent = tk.Frame(win, bg=C['bg'])
        tab_rent.pack(fill="both", expand=True)

        rent_outer = tk.Frame(tab_rent, bg=C['bg'])
        rent_outer.pack(fill="both", expand=True)
        rent_canvas = tk.Canvas(rent_outer, bg=C['bg'], highlightthickness=0)
        rent_vsb    = tk.Scrollbar(rent_outer, orient="vertical", command=rent_canvas.yview)
        rent_body   = tk.Frame(rent_canvas, bg=C['bg'])
        rent_body.bind("<Configure>",
                       lambda e: rent_canvas.configure(scrollregion=rent_canvas.bbox("all")))
        _rwin = rent_canvas.create_window((0, 0), window=rent_body, anchor="nw")
        rent_canvas.configure(yscrollcommand=rent_vsb.set)
        rent_canvas.bind("<Configure>",
                         lambda e: rent_canvas.itemconfig(_rwin, width=e.width))
        rent_canvas.bind("<MouseWheel>",
                         lambda e: rent_canvas.yview_scroll(-1*(e.delta//120), "units"))
        rent_canvas.pack(side="left", fill="both", expand=True)
        rent_vsb.pack(side="right", fill="y")

        add_card = tk.Frame(rent_body, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        add_card.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(add_card, text="➕  Add New Tenant", font=("Arial", 13, "bold"), bg=C['panel']).pack(anchor="w", padx=16, pady=(12, 4))
        add_form = tk.Frame(add_card, bg=C['panel'])
        add_form.pack(fill="x", padx=16, pady=(0, 8))

        lbl_kw = dict(font=("Arial", 11), bg=C['panel'])
        ent_kw = dict(font=("Arial", 12))
        tk.Label(add_form, text="Name:", **lbl_kw).grid(row=0, column=0, padx=(0,4), pady=6, sticky="w")
        new_name_var = tk.StringVar()
        tk.Entry(add_form, textvariable=new_name_var, width=18, **ent_kw).grid(row=0, column=1, padx=(0,14), pady=6)
        tk.Label(add_form, text="Stall No:", **lbl_kw).grid(row=0, column=2, padx=(0,4), pady=6, sticky="w")
        new_stall_var = tk.StringVar()
        tk.Entry(add_form, textvariable=new_stall_var, width=8, **ent_kw).grid(row=0, column=3, padx=(0,14), pady=6)
        tk.Label(add_form, text="Freq:", **lbl_kw).grid(row=0, column=4, padx=(0,4), pady=6, sticky="w")
        new_freq_var = tk.StringVar(value="Daily")
        ttk.Combobox(add_form, textvariable=new_freq_var, values=["Daily", "Monthly"], state="readonly", width=8, font=("Arial", 11)).grid(row=0, column=5, padx=(0,14), pady=6)
        tk.Label(add_form, text="Rate (RM):", **lbl_kw).grid(row=0, column=6, padx=(0,4), pady=6, sticky="w")
        new_rate_var = tk.StringVar(value="0.0")
        tk.Entry(add_form, textvariable=new_rate_var, width=8, **ent_kw).grid(row=0, column=7, padx=(0,14), pady=6)
        tk.Label(add_form, text="Start Date:", **lbl_kw).grid(row=1, column=0, padx=(0,4), pady=6, sticky="w")
        new_start_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        DatePicker(
            add_form, textvariable=new_start_var, width=11,
            font=("Arial", 11), language=self.lang, bg=C['panel']
        ).grid(row=1, column=1, padx=(0,10), pady=6, sticky="w")
        tk.Label(
            add_form, text=L("Type or choose", "可输入或选择"),
            font=("Arial", 9), bg=C['panel'], fg=C['text_muted']
        ).grid(row=2, column=1, padx=(0,10), pady=(0,2), sticky="w")
        tk.Label(add_form, text="Notes:", **lbl_kw).grid(row=1, column=2, padx=(0,4), pady=4, sticky="w")
        new_notes_var = tk.StringVar()
        tk.Entry(add_form, textvariable=new_notes_var, width=46, **ent_kw).grid(row=1, column=3, columnspan=5, padx=(0,14), pady=4, sticky="ew")

        add_status_lbl = tk.Label(add_card, text="", font=("Arial", 10), bg=C['panel'], fg=C['success'])
        add_status_lbl.pack(anchor="w", padx=16, pady=(0, 4))

        tenant_search_var = tk.StringVar()
        tenant_result_var = tk.StringVar()

        def do_add_tenant():
            name = new_name_var.get().strip()
            if not name:
                add_status_lbl.config(text="⚠ Name is required", fg=C['danger']); return
            try:
                rate = float(new_rate_var.get())
            except ValueError:
                add_status_lbl.config(text="⚠ Enter a valid rate", fg=C['danger']); return
            try:
                start_date = normalize_date_text(new_start_var.get())
                new_start_var.set(start_date)
            except ValueError:
                add_status_lbl.config(text="⚠ Please type or choose a valid start date", fg=C['danger']); return
            with self.db.get_connection() as conn:
                conn.execute(
                    "INSERT INTO rent_tenants (name, stall_no, rental_type, daily_rate, start_date, notes) VALUES (?,?,?,?,?,?)",
                    (name, new_stall_var.get().strip(), new_freq_var.get(), rate, start_date, new_notes_var.get().strip()))
            new_name_var.set(""); new_stall_var.set("")
            new_rate_var.set("0.0"); new_notes_var.set("")
            new_start_var.set(datetime.now().strftime("%Y-%m-%d"))
            add_status_lbl.config(text=f"✅ '{name}' added", fg=C['success'])
            reload_all_tenants()
            reload_tenant_today()

        RoundedButton(add_form, text="➕  Add Tenant", command=do_add_tenant, bg=C['success'], fg="white", font=("Arial", 11, "bold"), padx=16, pady=8, radius=8).grid(row=1, column=8, padx=8, pady=6)

        # ──────────────────────────────────────────────────────────────────
        # FUNCTIONS FOR RENT TAB
        # ──────────────────────────────────────────────────────────────────
        def reload_all_tenants():
            ten_tree.delete(*ten_tree.get_children())
            keyword = tenant_search_var.get().strip().casefold()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM rent_tenants ORDER BY active DESC, stall_no, name")
                rows = c.fetchall()
                if keyword:
                    rows = [
                        t for t in rows
                        if keyword in " ".join((
                            str(t['id'] or ''), str(t['name'] or ''),
                            str(t['stall_no'] or ''), str(t['rental_type'] or ''),
                            str(t['paid_until_date'] or ''), str(t['notes'] or ''),
                        )).casefold()
                    ]
                for t in rows:
                    until = t['paid_until_date'] or "—"
                    start_date = (t['start_date'] or "").strip()
                    status_text = "✅ Active" if t['active'] else "❌ Off"
                    if not t['active']:
                        tag = "inactive"
                    else:
                        try:
                            ud    = datetime.strptime(until, "%Y-%m-%d").date()
                            left  = (ud - datetime.now().date()).days
                            tag   = "overdue" if left < 0 else ("warn" if left <= 3 else "ok")
                            if left < 0:
                                status_text = f"⚠ Unpaid {abs(left)}d"
                            elif left == 0:
                                status_text = "⚠ Due Today"
                            else:
                                status_text = f"✅ {left}d left"
                        except Exception:
                            try:
                                sdt = datetime.strptime(start_date, "%Y-%m-%d").date() if start_date else datetime.now().date()
                            except Exception:
                                sdt = datetime.now().date()
                            unpaid = max(0, (datetime.now().date() - sdt).days)
                            if unpaid > 0:
                                tag = "overdue"
                                status_text = f"⚠ Unpaid {unpaid}d"
                            else:
                                tag = "ok"
                                status_text = "✅ No due yet"
                    rate_display = f"RM {t['daily_rate']:.2f}/mo" if t['rental_type'] == 'Monthly' else f"RM {t['daily_rate']:.2f}/d"
                    ten_tree.insert("", "end", values=(
                        t['id'], t['name'], t['stall_no'] or "—",
                        rate_display,
                        start_date or "—",
                        f"{t['total_days_paid']:.0f}",
                        until,
                        status_text
                    ), tags=(tag,))
            tenant_result_var.set(L(f"{len(rows)} tenant(s)", f"{len(rows)} 个租户"))

        def toggle_tenant_active():
            sel = ten_tree.selection()
            if not sel: return
            tid = ten_tree.item(sel[0])['values'][0]
            with self.db.get_connection() as conn:
                conn.execute(
                    "UPDATE rent_tenants SET active=CASE WHEN active=1 THEN 0 ELSE 1 END WHERE id=?",
                    (tid,))
            reload_all_tenants()

        def delete_tenant():
            sel = ten_tree.selection()
            if not sel: return
            tid  = ten_tree.item(sel[0])['values'][0]
            name = str(ten_tree.item(sel[0])['values'][1])
            if not messagebox.askyesno(L("Delete Tenant", "删除租户"),
                    L(f"Delete '{name}'?\nPayment history will be kept.", f"删除“{name}”？\n付款历史会保留。"), parent=win):
                return
            with self.db.get_connection() as conn:
                conn.execute("DELETE FROM rent_tenants WHERE id=?", (tid,))
            reload_all_tenants()

        def sync_all_rent_dates():
            if not messagebox.askyesno(L("Sync Dates", "同步日期"),
                                       L("This will recalculate all 'Paid Until' dates based on valid payments and start dates. Continue?",
                                         "系统将根据有效付款及开始日期重新计算所有“已付至”日期。是否继续？"),
                                       parent=win):
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT id, start_date FROM rent_tenants")
                tenants = c.fetchall()
                for t in tenants:
                    tid = t['id']
                    s_date_str = (t['start_date'] or "").strip()
                    c.execute("SELECT SUM(days) as total_days FROM rent_payments WHERE tenant_id=? AND voided=0", (tid,))
                    row = c.fetchone()
                    total_days = float(row['total_days'] or 0)
                    new_until = ""
                    if total_days > 0 and s_date_str:
                        try:
                            s_date = datetime.strptime(s_date_str, "%Y-%m-%d").date()
                            new_until = (s_date + timedelta(days=max(0, total_days - 1))).strftime("%Y-%m-%d")
                        except: pass
                    c.execute("UPDATE rent_tenants SET total_days_paid=?, paid_until_date=? WHERE id=?",
                             (total_days, new_until, tid))
            reload_all_tenants()
            messagebox.showinfo(L("✅ Synced", "✅ 已同步"), L("All rent dates have been corrected.", "所有租金日期已重新计算。"), parent=win)

        def edit_tenant():
            sel = ten_tree.selection()
            if not sel:
                messagebox.showwarning(L("No Selection", "未选择"), L("Select a tenant first", "请先选择一个租户。"), parent=win)
                return
            tid = int(ten_tree.item(sel[0])['values'][0])
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM rent_tenants WHERE id=?", (tid,))
                ten = c.fetchone()
            if not ten:
                messagebox.showerror(L("Missing", "缺少资料"), L("Tenant record not found", "找不到租户记录。"), parent=win)
                return

            dlg = tk.Toplevel(win)
            dlg.title(f"Edit Tenant #{tid}")
            dlg.configure(bg=C['bg'])
            dlg.transient(win)
            dlg.grab_set()
            self._center_window(dlg, 560, 360, win)

            form = tk.Frame(dlg, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
            form.pack(fill="both", expand=True, padx=12, pady=12)

            tk.Label(form, text=f"✏ Edit Tenant  #{tid}", font=("Arial", 14, "bold"),
                     bg=C['panel']).grid(row=0, column=0, columnspan=4, padx=12, pady=(12, 10), sticky="w")

            def _frow(row, label, var_or_factory, width=24):
                tk.Label(form, text=label, font=("Arial", 11), bg=C['panel']).grid(
                    row=row, column=0, padx=(12, 6), pady=6, sticky="w")
                if callable(var_or_factory):
                    var_or_factory(form).grid(row=row, column=1, columnspan=3, padx=(0, 12), pady=6, sticky="ew")
                else:
                    tk.Entry(form, textvariable=var_or_factory, font=("Arial", 11), width=width).grid(
                        row=row, column=1, columnspan=3, padx=(0, 12), pady=6, sticky="ew")

            name_var = tk.StringVar(value=ten['name'] or "")
            stall_var = tk.StringVar(value=ten['stall_no'] or "")
            freq_var = tk.StringVar(value=ten['rental_type'] or "Daily")
            rate_var = tk.StringVar(value=f"{float(ten['daily_rate'] or 0):.2f}")
            start_var = tk.StringVar(value=ten['start_date'] or datetime.now().strftime("%Y-%m-%d"))
            notes_var = tk.StringVar(value=ten['notes'] or "")

            _frow(1, "Name:", name_var); _frow(2, "Stall No:", stall_var)
            _frow(3, "Freq:", lambda f: ttk.Combobox(f, textvariable=freq_var, values=["Daily", "Monthly"], state="readonly", font=("Arial", 11)))
            _frow(4, "Rate (RM):", rate_var)
            _frow(
                5, "Start Date:",
                lambda f: DatePicker(
                    f, textvariable=start_var, font=("Arial", 11),
                    language=self.lang, bg=C['panel']
                )
            )
            _frow(6, "Notes:", notes_var)

            status_lbl = tk.Label(form, text="", font=("Arial", 10), bg=C['panel'], fg=C['danger'])
            status_lbl.grid(row=7, column=0, columnspan=4, padx=12, pady=(2, 8), sticky="w")

            btns = tk.Frame(form, bg=C['panel'])
            btns.grid(row=8, column=0, columnspan=4, padx=12, pady=(0, 12), sticky="w")

            def _save_tenant_edit():
                name = name_var.get().strip()
                if not name:
                    status_lbl.config(text="Name is required", fg=C['danger']); return
                try:
                    rate = float(rate_var.get())
                    if rate < 0: raise ValueError
                except ValueError:
                    status_lbl.config(text="Daily rate must be a valid non-negative number", fg=C['danger']); return
                try:
                    start_date = normalize_date_text(start_var.get())
                    start_var.set(start_date)
                except ValueError:
                    status_lbl.config(text="Please type or choose a valid start date", fg=C['danger']); return

                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("UPDATE rent_tenants SET name=?, stall_no=?, rental_type=?, daily_rate=?, start_date=?, notes=? WHERE id=?",
                              (name, stall_var.get().strip(), freq_var.get(), rate, start_date, notes_var.get().strip(), tid))
                    c.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                              (datetime.now().strftime("%d-%m-%Y %H:%M:%S"), "RENT_EDIT_TENANT",
                               f"Tenant#{tid} updated: {name}, rate RM{rate:.2f}, start {start_date}", self._audit_actor("Admin")))

                messagebox.showinfo(L("Saved", "已保存"), L(f"Tenant '{name}' updated.", f"租户“{name}”已更新。"), parent=win)

            RoundedButton(btns, text="💾 Save", command=_save_tenant_edit, bg=C['success'], fg="white", font=("Arial", 11, "bold"), padx=14, pady=7, radius=7).pack(side="left", padx=(0, 8))
            RoundedButton(btns, text="Cancel", command=dlg.destroy, bg=C['neutral'], fg="white", font=("Arial", 11, "bold"), padx=14, pady=7, radius=7).pack(side="left")

        def reload_tenant_today():
            tr_tree.delete(*tr_tree.get_children())
            biz = get_today()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM rent_payments WHERE pay_date=? ORDER BY id DESC", (biz,))
                rows = c.fetchall()
            for r in rows:
                voided = r['voided'] if 'voided' in r.keys() else 0
                dt_str = r['datetime'] or ""
                tr_tree.insert("", "end", iid=str(r['id']), values=(
                    dt_str[11:16] if len(dt_str) > 11 else "—",
                    r['tenant_name'] or "—", r['stall_no'] or "—",
                    f"{r['days']:.0f}", f"RM{r['rate']:.2f}",
                    f"RM {r['amount']:.2f}",
                    (r['payment_method'] or "CASH").upper(),
                    r['recorded_by'] or "—",
                    "VOID" if voided else "OK"
                ), tags=("voided",) if voided else ())

        def void_rent_payment():
            sel = tr_tree.selection()
            if not sel: return
            pid = int(sel[0])
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM rent_payments WHERE id=?", (pid,))
                row = c.fetchone()
                if not row: return
                if row['voided'] if 'voided' in row.keys() else 0:
                    messagebox.showinfo(L("Wait", "请注意"), L("This payment is already voided", "这笔付款已经作废。"), parent=win); return
                tenant_id = row['tenant_id']
                days = float(row['days'] or 0)
                amount = float(row['amount'] or 0)
                c.execute("SELECT * FROM rent_tenants WHERE id=?", (tenant_id,))
                tenant = c.fetchone()
                if not tenant:
                    messagebox.showerror(L("Error", "错误"), L("Tenant not found", "找不到租户。"), parent=win); return

            def do_void(staff):
                with self.db.get_connection() as conn:
                    c2 = conn.cursor()
                    c2.execute("UPDATE rent_payments SET voided=1 WHERE id=?", (pid,))
                    
                    # Robust recalculation from all non-voided payments
                    c2.execute("SELECT SUM(days) FROM rent_payments WHERE tenant_id=? AND COALESCE(voided,0)=0", (tenant_id,))
                    total_days = c2.fetchone()[0] or 0
                    
                    start_val = tenant['start_date']
                    if start_val:
                        try:
                            start_dt = datetime.strptime(start_val, "%Y-%m-%d")
                            new_until = (start_dt + timedelta(days=total_days)).strftime("%Y-%m-%d")
                        except:
                            new_until = tenant['paid_until_date']
                    else:
                        new_until = tenant['paid_until_date']

                    c2.execute("""UPDATE rent_tenants 
                                 SET total_days_paid = ?,
                                     paid_until_date = ?
                                 WHERE id=?""", (total_days, new_until, tenant_id))

                    c2.execute("DELETE FROM cashflow WHERE source='Stall Rental' AND reference_id=?", (pid,))
                    c2.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                             (datetime.now().strftime("%d-%m-%Y %H:%M:%S"), 'VOID_RENT',
                              f"Void Rent #{pid} for {tenant['name']} (RM{amount:.2f}, {days}d)", staff['name']))
                messagebox.showinfo(L("✅ Voided", "✅ 已作废"), L(f"Payment for {tenant['name']} has been voided.", f"{tenant['name']} 的付款已作废。"), parent=win)

            self._staff_login_popup(callback=do_void, title="👤 Staff PIN — Void Rent Payment", parent=win)

        # ──────────────────────────────────────────────────────────────────
        # SECTION 2: All Tenants list + today's payments side by side
        # ──────────────────────────────────────────────────────────────────
        mid_row = tk.Frame(rent_body, bg=C['bg'])
        mid_row.pack(fill="both", expand=True, padx=12, pady=4)

        # LEFT: All tenants tree
        ten_card = tk.Frame(mid_row, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        ten_card.pack(side="left", fill="both", expand=True, padx=(0, 6))

        ten_hdr = tk.Frame(ten_card, bg=C['panel'])
        ten_hdr.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(ten_hdr, text=L("🏪  All Tenants", "🏪  所有租户"), font=("Arial", 13, "bold"), bg=C['panel']).pack(side="left")

        tenant_query_host = tk.Frame(ten_card, bg=C['panel'])
        tenant_query_host.pack(fill="x", padx=0, pady=(0, 2))

        ten_cols = ("ID", "Name", "Stall", "Rate", "Start Date", "Total Days", "Paid Until", "Status")
        ten_tree_f = tk.Frame(ten_card, bg=C['panel'])
        ten_tree_f.pack(fill="both", expand=True, padx=8, pady=4)
        ten_tree = ttk.Treeview(ten_tree_f, columns=ten_cols, show="headings", height=9)
        ten_w = {"ID":36,"Name":120,"Stall":62,"Rate":78,"Start Date":95,"Total Days":80,"Paid Until":100,"Status":88}
        for col in ten_cols:
            ten_tree.heading(col, text=col)
            ten_tree.column(col, width=ten_w.get(col, 80), anchor="center")
        ten_tree.tag_configure("overdue", foreground="#DC2626")
        ten_tree.tag_configure("ok",      foreground="#16A34A")
        ten_tree.tag_configure("warn",    foreground="#D97706")
        ten_tree.tag_configure("inactive",foreground="#94A3B8")
        ten_sb = tk.Scrollbar(ten_tree_f, orient="vertical", command=ten_tree.yview)
        ten_tree.configure(yscrollcommand=ten_sb.set)
        ten_tree.pack(side="left", fill="both", expand=True)
        ten_sb.pack(side="right", fill="y")

        ten_act = tk.Frame(ten_card, bg=C['panel'])
        ten_act.pack(fill="x", padx=10, pady=(0, 8))

        RoundedButton(ten_act, text="🔄 Sync All Dates", command=sync_all_rent_dates,
                      bg="#6366F1", fg="white", font=("Arial", 9, "bold"),
                      padx=10, pady=4, radius=6).pack(side="right", padx=4)

        RoundedButton(ten_act, text="🔄 Toggle Active", command=toggle_tenant_active,
                      bg=C['warning'], fg="white", font=("Arial", 10, "bold"), padx=10, pady=4, radius=7
                      ).pack(side="left", padx=4)
        RoundedButton(ten_act, text="✏ Edit", command=edit_tenant,
                      bg=C['neutral'], fg="white", font=("Arial", 10, "bold"), padx=10, pady=4, radius=7
                      ).pack(side="left", padx=4)
        RoundedButton(ten_act, text="🗑️ Delete", command=delete_tenant,
                      bg=C['danger'], fg="white", font=("Arial", 10, "bold"), padx=10, pady=4, radius=7
                      ).pack(side="left", padx=4)

        ten_tree.bind("<Double-1>", lambda _e: edit_tenant())
        self._build_query_toolbar(
            tenant_query_host, tenant_search_var, reload_all_tenants,
            on_refresh=reload_all_tenants,
            on_export=lambda: self._export_treeview_csv(
                ten_tree, win, "rent_tenants.csv"
            ),
            result_var=tenant_result_var,
            search_label=L("Search tenants:", "搜索租户：")
        )

        # RIGHT: Today's payments
        today_card = tk.Frame(mid_row, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        today_card.pack(side="left", fill="both", expand=True, padx=(6, 0))

        today_pay_hdr = tk.Frame(today_card, bg=C['panel'])
        today_pay_hdr.pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(today_pay_hdr, text=L("📅  Today's Rent Payments", "📅  今日租金付款"), font=("Arial", 13, "bold"), bg=C['panel']).pack(side="left")
        RoundedButton(
            today_pay_hdr, text=L("⬇ Export", "⬇ 导出"),
            command=lambda: self._export_treeview_csv(
                tr_tree, win, "today_rent_payments.csv"
            ),
            bg=C['success'], fg="white", font=("Arial", 8, "bold"),
            padx=8, pady=3, radius=6
        ).pack(side="right")

        tr_cols = ("Time","Tenant","Stall","Days","Rate","Amount","Method","By","Status")
        tr_tree_f = tk.Frame(today_card, bg=C['panel'])
        tr_tree_f.pack(fill="both", expand=True, padx=8, pady=4)
        tr_tree = ttk.Treeview(tr_tree_f, columns=tr_cols, show="headings", height=9)
        tr_w = {"Time":60,"Tenant":100,"Stall":50,"Days":45,"Rate":60,"Amount":80,
                "Method":58,"By":80,"Status":54}
        for col in tr_cols:
            tr_tree.heading(col, text=col)
            tr_tree.column(col, width=tr_w.get(col, 70), anchor="center")
        tr_tree.tag_configure("voided", foreground="#94A3B8")
        tr_sb2 = tk.Scrollbar(tr_tree_f, orient="vertical", command=tr_tree.yview)
        tr_tree.configure(yscrollcommand=tr_sb2.set)
        tr_tree.pack(side="left", fill="both", expand=True)
        tr_sb2.pack(side="right", fill="y")

        RoundedButton(today_card, text="🗑 Void Selected Payment", command=void_rent_payment,
                      bg=C['danger'], fg="white", font=("Arial", 10, "bold"), padx=12, pady=6, radius=8).pack(pady=4)

        # Initial data load
        reload_all_tenants()
        reload_tenant_today()

        # ──────────────────────────────────────────────────────────────────
        # SECTION 3: Payment History
        # ──────────────────────────────────────────────────────────────────
        hist_card = tk.Frame(rent_body, bg=C['panel'],
                             highlightbackground=C['border'], highlightthickness=1)
        hist_card.pack(fill="x", padx=12, pady=4)

        hist_hdr = tk.Frame(hist_card, bg=C['panel'])
        hist_hdr.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(hist_hdr, text=L("📜  Payment History", "📜  付款历史"),
                 font=("Arial", 13, "bold"), bg=C['panel']).pack(side="left")

        hist_filter_var = tk.StringVar(value="All")
        hist_cb = ttk.Combobox(hist_hdr, textvariable=hist_filter_var,
                               font=("Arial", 10), width=16, state="readonly")
        hist_cb.pack(side="left", padx=16)

        def refresh_hist_cb():
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT name FROM rent_tenants ORDER BY name")
                hist_cb['values'] = ["All"] + [r['name'] for r in c.fetchall()]
        refresh_hist_cb()

        rent_hist_period_var = tk.StringVar(value="all")
        rent_hist_from_var = tk.StringVar(value=datetime.now().replace(day=1).strftime("%Y-%m-%d"))
        rent_hist_to_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        hist_date_row = tk.Frame(hist_card, bg=C['panel'])
        hist_date_row.pack(fill="x", padx=12, pady=(0, 3))
        tk.Label(hist_date_row, text=L("Date:", "日期："), font=("Arial", 9, "bold"), bg=C['panel']).pack(side="left")
        for label, value in [
            (L("Today", "今天"), "today"),
            (L("Previous Day", "前一天"), "yesterday"),
            (L("7 Days", "7天"), "7days"),
            (L("All", "全部"), "all"),
            (L("Custom", "自定义"), "custom"),
        ]:
            tk.Radiobutton(
                hist_date_row, text=label, variable=rent_hist_period_var, value=value,
                command=lambda: load_rent_history(), font=("Arial", 9),
                bg=C['panel'], fg=C['text']
            ).pack(side="left", padx=3)
        tk.Label(hist_date_row, text=L("From:", "开始："), font=("Arial", 9), bg=C['panel']).pack(side="left", padx=(10, 3))
        DatePicker(
            hist_date_row, textvariable=rent_hist_from_var, width=10,
            font=("Arial", 9), language=self.lang, bg=C['panel'],
            callback=lambda: (rent_hist_period_var.set("custom"), load_rent_history())
        ).pack(side="left")
        tk.Label(hist_date_row, text=L("To:", "结束："), font=("Arial", 9), bg=C['panel']).pack(side="left", padx=(7, 3))
        DatePicker(
            hist_date_row, textvariable=rent_hist_to_var, width=10,
            font=("Arial", 9), language=self.lang, bg=C['panel'],
            callback=lambda: (rent_hist_period_var.set("custom"), load_rent_history())
        ).pack(side="left")

        rent_history_search_var = tk.StringVar()
        rent_history_result_var = tk.StringVar()
        rent_history_query_host = tk.Frame(hist_card, bg=C['panel'])
        rent_history_query_host.pack(fill="x", padx=0, pady=(0, 2))

        h_cols = ("Date","Time","Tenant","Stall","Days","Rate","Total","Method","By","Status")
        h_tree_f = tk.Frame(hist_card, bg=C['panel'])
        h_tree_f.pack(fill="both", padx=8, pady=4)
        h_tree = ttk.Treeview(h_tree_f, columns=h_cols, show="headings", height=7)
        h_w = {"Date":85,"Time":52,"Tenant":110,"Stall":55,"Days":48,"Rate":68,
               "Total":85,"Method":58,"By":90,"Status":58}
        for col in h_cols:
            h_tree.heading(col, text=col)
            h_tree.column(col, width=h_w.get(col, 80), anchor="center")
        h_tree.tag_configure("voided", foreground="#94A3B8")
        h_sb = tk.Scrollbar(h_tree_f, orient="vertical", command=h_tree.yview)
        h_tree.configure(yscrollcommand=h_sb.set)
        h_tree.pack(side="left", fill="both", expand=True)
        h_sb.pack(side="right", fill="y")

        def load_rent_history():
            h_tree.delete(*h_tree.get_children())
            tenant_filter = hist_filter_var.get()
            keyword = rent_history_search_var.get().strip().casefold()
            period = rent_hist_period_var.get()
            today = datetime.now()
            start_date = end_date = None
            if period == "today":
                start_date = end_date = today.strftime("%Y-%m-%d")
            elif period == "yesterday":
                start_date = end_date = (today - timedelta(days=1)).strftime("%Y-%m-%d")
            elif period == "7days":
                start_date = (today - timedelta(days=6)).strftime("%Y-%m-%d")
                end_date = today.strftime("%Y-%m-%d")
            elif period == "custom":
                try:
                    start_date = normalize_date_text(rent_hist_from_var.get())
                    end_date = normalize_date_text(rent_hist_to_var.get())
                    if end_date < start_date:
                        raise ValueError
                    rent_hist_from_var.set(start_date)
                    rent_hist_to_var.set(end_date)
                except ValueError:
                    messagebox.showerror(
                        L("Invalid Date Range", "日期范围无效"),
                        L("Choose valid dates; From cannot be after To.", "请选择有效日期，开始日期不可晚于结束日期。"),
                        parent=win
                    )
                    return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                if tenant_filter == "All":
                    c.execute("SELECT * FROM rent_payments ORDER BY id DESC LIMIT 300")
                else:
                    c.execute("SELECT * FROM rent_payments WHERE tenant_name=? ORDER BY id DESC",
                              (tenant_filter,))
                rows = c.fetchall()
                filtered_rows = []
                for r in rows:
                    pay_date = str(r['pay_date'] or '')
                    if start_date and not (start_date <= pay_date <= end_date):
                        continue
                    if keyword:
                        searchable = " ".join((
                            str(r['id'] or ''), str(r['tenant_name'] or ''),
                            str(r['stall_no'] or ''), str(r['payment_method'] or ''),
                            str(r['recorded_by'] or ''), str(r['notes'] or '') if 'notes' in r.keys() else '',
                            pay_date,
                        )).casefold()
                        if keyword not in searchable:
                            continue
                    filtered_rows.append(r)
                for r in filtered_rows:
                    voided = r['voided'] if 'voided' in r.keys() else 0
                    dt_str = r['datetime'] or ""
                    h_tree.insert("", "end", values=(
                        r['pay_date'] or "—",
                        dt_str[11:16] if len(dt_str) > 11 else "—",
                        r['tenant_name'] or "—", r['stall_no'] or "—",
                        f"{r['days']:.0f}", f"RM{r['rate']:.2f}",
                        f"RM {r['amount']:.2f}",
                        (r['payment_method'] or "CASH").upper(),
                        r['recorded_by'] or "—",
                        "VOID" if voided else "OK"
                    ), tags=("voided",) if voided else ())
            rent_history_result_var.set(
                L(f"{len(filtered_rows)} payment(s)", f"{len(filtered_rows)} 笔付款")
            )

        def clear_rent_history_filters():
            hist_filter_var.set("All")
            rent_hist_period_var.set("all")
            rent_hist_from_var.set(datetime.now().replace(day=1).strftime("%Y-%m-%d"))
            rent_hist_to_var.set(datetime.now().strftime("%Y-%m-%d"))
            load_rent_history()

        self._build_query_toolbar(
            rent_history_query_host, rent_history_search_var, load_rent_history,
            on_clear=clear_rent_history_filters,
            on_refresh=load_rent_history,
            on_export=lambda: self._export_treeview_csv(
                h_tree, win, "rent_payment_history.csv"
            ),
            result_var=rent_history_result_var,
            search_label=L("Tenant / stall / payment:", "租户／档口／付款：")
        )

        hist_cb.bind("<<ComboboxSelected>>", lambda e: load_rent_history())
        load_rent_history()

        # Register tab switch handler for refreshing rent data

    def show_cash_closing(self, cashier_mode=False):
        C = self.colors
        L = (lambda en, zh: zh if self.lang == "zh" else en)
        win = tk.Toplevel(self.root)
        cash_window_title = L("Cash Operations", "现金操作") if cashier_mode else L("Cash Closing", "现金日结")
        win.title(f"💰 {cash_window_title}")
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        window_w = max(760, min(1400, screen_w - 60))
        window_h = max(650, min(980, screen_h - 90))
        win.geometry(f"{window_w}x{window_h}")
        win.minsize(min(1000, window_w), min(720, window_h))
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()

        def _boost_fonts(root_widget, delta=2):
            widgets = [root_widget]
            while widgets:
                w = widgets.pop(0)
                try: widgets.extend(w.winfo_children())
                except: pass
                try:
                    fval = w.cget("font")
                    if not fval: continue
                    f = tkfont.Font(font=fval)
                    fam = f.actual("family")
                    size = int(f.actual("size") or 0)
                    if size == 0: continue
                    new_size = size + delta if size > 0 else size - delta
                    style = []
                    if str(f.actual("weight")).lower() == "bold": style.append("bold")
                    if str(f.actual("slant")).lower() == "italic": style.append("italic")
                    w.configure(font=(fam, new_size, *style))
                except: continue

        def get_today(): return today_biz_date()
        def get_biz_range():
            biz = get_today()
            s, e = business_day_range(biz)
            return biz, s, e

        def get_opening_cash():
            biz, _, _ = get_biz_range()
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT opening_cash FROM cash_opening WHERE date=? ORDER BY id DESC LIMIT 1", (biz,))
                row = c.fetchone()
                return float(row['opening_cash']) if row else 0.0

        def reload_summary():
            biz, _, _ = get_biz_range()
            with self.db.get_connection() as conn:
                pos_cash, other_income, expenses = compute_cash_summary_components(conn, biz)
            opening  = get_opening_cash()
            expected = opening + pos_cash + other_income - expenses
            lbl_pos.config(text=f"RM {pos_cash:.2f}")
            lbl_income.config(text=f"RM {other_income:.2f}")
            lbl_expense.config(text=f"RM {expenses:.2f}")
            lbl_opening.config(text=f"RM {opening:.2f}")
            lbl_expected.config(text=f"RM {expected:.2f}")
            return pos_cash, other_income, expenses, expected, opening

        hdr = tk.Frame(win, bg=C['header'], height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text=f"💰 {cash_window_title}", font=("Arial", 18, "bold"), bg=C['header'], fg='white').pack(side="left", padx=20, pady=10)
        tk.Label(hdr, text=f"Business Day: {today_biz_date()}  (06:00 ~ 05:59 next day)", font=("Arial", 11), bg=C['header'], fg='#94A3B8').pack(side="right", padx=20)

        summary_bar = tk.Frame(win, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        summary_bar.pack(fill="x", padx=16, pady=(10, 4))

        def _summary_card(parent, title, value_color):
            f = tk.Frame(parent, bg=value_color, padx=12, pady=10)
            f.pack(side="left", expand=True, fill="both", padx=4, pady=4)
            tk.Label(f, text=title, font=("Arial", 10, "bold"), bg=value_color, fg="white").pack(pady=(4, 2))
            lbl = tk.Label(f, text="RM 0.00", font=("Arial", 18, "bold"), bg=value_color, fg="white")
            lbl.pack(pady=(0, 4))
            return lbl

        lbl_opening  = _summary_card(summary_bar, L("🪙 Opening Cash", "🪙 开盘现金"),   "#F59E0B")
        lbl_pos      = _summary_card(summary_bar, L("💻 POS Cash", "💻 POS 现金销售"),  C['success'])
        lbl_income   = _summary_card(summary_bar, L("💵 Cash Receipts", "💵 其他现金收入"), "#2563EB")
        lbl_expense  = _summary_card(summary_bar, L("💸 Cash Out", "💸 现金转出"),         C['danger'])
        lbl_expected = _summary_card(summary_bar, L("✅ Expected Cash", "✅ 应有现金"), "#334155")

        nb = ttk.Notebook(win, style="AdminCashTabs.TNotebook")
        nb.pack(fill="both", expand=True, padx=16, pady=8)

        # ── CONFIGS & HELPERS ──
        cash_out_category_labels = {
            "Stock Purchase": L("Stock Purchase", "采购入库"),
            "Ingredients": L("Ingredients", "食材／原料"),
            "Small Tools & Utensils": L("Small Tools & Utensils", "小工具／餐具"),
            "Equipment & Furniture": L("Equipment & Furniture", "设备／家具"),
            "Daily Supplies": L("Daily Supplies", "日常用品"),
            "Packaging": L("Packaging", "包装用品"),
            "Repair & Maintenance": L("Repair & Maintenance", "维修保养"),
            "Transport": L("Transport", "交通运输"),
            "Utilities & Licenses": L("Utilities & Licenses", "水电与执照"),
            "Staff Welfare": L("Staff Welfare", "员工福利"),
            "Other Expense": L("Other Expense", "其他营业开销"),
            "Staff Advance": L("Staff Advance", "员工预支"),
            "Cash Transfer": L("Cash Transfer", "现金转移"),
        }
        # These protected workflows have related accounting or staff side effects.
        # Ordinary expense choices are loaded from ``expense_options`` below.
        cash_out_special_types = [
            {
                "key": "staff_advance", "label": L("Staff Advance", "员工预支"),
                "purpose": "staff_advance", "category": "Staff Advance",
                "item_name": "", "detail_mode": "staff", "quantity_mode": "month",
                "help": L(
                    "Choose the employee and salary deduction month. This is kept separate from operating expenses.",
                    "选择员工与扣薪月份；员工预支会与营业开销分开记录。"
                ),
            },
            {
                "key": "bank_deposit", "label": L("Bank Deposit", "存入银行"),
                "purpose": "cash_transfer", "category": "Cash Transfer",
                "item_name": "Bank Deposit", "detail_mode": "fixed", "quantity_mode": "none",
                "help": L("Use only when cash is physically removed from this drawer for banking.", "只在现金实际从钱箱拿去存入银行时使用。"),
            },
            {
                "key": "owner_withdrawal", "label": L("Owner Withdrawal", "业主提款"),
                "purpose": "cash_transfer", "category": "Cash Transfer",
                "item_name": "Owner Withdrawal", "detail_mode": "fixed", "quantity_mode": "none",
                "help": L("Records cash removed by the owner; it is not an operating expense.", "记录业主从钱箱提款；这不是营业开销。"),
            },
            {
                "key": "petty_cash_transfer", "label": L("Petty Cash Transfer", "备用金转移"),
                "purpose": "cash_transfer", "category": "Cash Transfer",
                "item_name": "Petty Cash Transfer", "detail_mode": "fixed", "quantity_mode": "none",
                "help": L("Records cash moved from this drawer to another cash float.", "记录从本钱箱转到其他备用金的钱。"),
            },
        ]

        # ── OPENING TAB ──
        tab_opening = tk.Frame(nb, bg=C['bg'])
        nb.add(tab_opening, text="🪙 " + L("Opening", "开盘"))
        open_card = tk.Frame(tab_opening, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        open_card.pack(fill="x", padx=16, pady=20)
        tk.Label(open_card, text="🪙 " + L("Opening Cash Registration", "开盘现金登记"), font=("Arial", 14, "bold"), bg=C['panel'], fg=C['primary']).grid(row=0, column=0, columnspan=4, padx=16, pady=(15,10), sticky="w")
        tk.Label(open_card, text=L("Opening Amount RM:", "开盘金额 RM:"), font=("Arial", 12), bg=C['panel']).grid(row=1, column=0, padx=(16,6), pady=15, sticky="w")
        open_amt_var = tk.StringVar()
        open_amt_entry = tk.Entry(open_card, textvariable=open_amt_var, font=("Arial", 18, "bold"), width=12, justify="right")
        open_amt_entry.grid(row=1, column=1, padx=6, pady=15)
        tk.Label(open_card, text=L("Notes:", "备注:"), font=("Arial", 12), bg=C['panel']).grid(row=1, column=2, padx=(16,6), pady=15)
        open_notes_var = tk.StringVar()
        tk.Entry(open_card, textvariable=open_notes_var, font=("Arial", 11), width=22).grid(row=1, column=3, padx=6, pady=15)
        open_status_lbl = tk.Label(open_card, text="", font=("Arial", 11), bg=C['panel'], fg=C['success'])
        open_status_lbl.grid(row=2, column=0, columnspan=4, padx=16, pady=(0,10), sticky="w")

        def refresh_open_status():
            opening = get_opening_cash()
            if opening > 0:
                open_status_lbl.config(text=f"✅ {L('Opening cash registered: RM', '开盘现金已登记: RM')} {opening:.2f}", fg=C['success'])
            else:
                open_status_lbl.config(text="⚠️ " + L("Opening cash not yet registered today", "今日开盘现金尚未登记"), fg=C['warning'])

        def save_opening():
            try:
                amt = float(open_amt_var.get().strip()); 
                if amt < 0: raise ValueError
            except: messagebox.showerror(L("Error", "错误"), L("Please enter a valid amount", "请输入有效金额。"), parent=win); return
            def do_save(staff):
                biz = get_today(); now = datetime.now()
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    c.execute("DELETE FROM cash_opening WHERE date=?", (biz,))
                    c.execute("INSERT INTO cash_opening (date, opening_cash, opened_by, opened_at) VALUES (?,?,?,?)", (biz, amt, staff['name'], now.strftime("%Y-%m-%d %H:%M:%S")))
                    c.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)", (now.strftime("%d-%m-%Y %H:%M:%S"), 'CASH_OPEN', f'Opening Cash RM {amt:.2f}', staff['name']))
                open_amt_var.set(""); refresh_open_status(); reload_summary()
                messagebox.showinfo(L("Registered", "已登记"), L(f"Opening RM {amt:.2f} registered\nBy: {staff['name']}", f"已登记开盘现金 RM {amt:.2f}\n员工：{staff['name']}"), parent=win)
            self._staff_login_popup(callback=do_save, title=L("👤 Staff PIN — Register Opening", "👤 员工 PIN — 登记开盘"), parent=win)
        
        RoundedButton(open_card, text="💾 " + L("Register Opening", "登记开盘"), command=save_opening, bg="#F59E0B", fg="white", font=("Arial", 12, "bold"), padx=20, pady=10, radius=8).grid(row=1, column=4, padx=16, pady=15)
        refresh_open_status()

        # ── CASH OUT TAB ──
        tab_cash_out = tk.Frame(nb, bg=C['bg'])
        nb.add(tab_cash_out, text="💸 " + L("Cash Out", "现金转出"))

        cash_out_card = tk.Frame(
            tab_cash_out, bg=C['panel'],
            highlightbackground=C['border'], highlightthickness=1
        )
        cash_out_card.pack(fill="x", padx=12, pady=10)
        for _col, _weight in enumerate((2, 3, 2, 1)):
            cash_out_card.grid_columnconfigure(_col, weight=_weight)

        tk.Label(
            cash_out_card,
            text=L("Record why cash leaves the drawer", "记录现金从钱箱转出的实际用途"),
            font=("Arial", 14, "bold"), bg=C['panel'], fg=C['primary']
        ).grid(row=0, column=0, columnspan=4, padx=14, pady=(12, 2), sticky="w")
        cash_out_help_var = tk.StringVar()
        tk.Label(
            cash_out_card, textvariable=cash_out_help_var,
            font=("Arial", 9), bg=C['panel'], fg=C['text_sub'],
            anchor="w", justify="left"
        ).grid(row=1, column=0, columnspan=4, padx=14, pady=(0, 8), sticky="ew")

        cash_out_type_var = tk.StringVar()
        cash_out_staff_var = tk.StringVar()
        quantity_var = tk.StringVar(value="1")
        cash_out_amount_var = tk.StringVar()
        cash_out_notes_var = tk.StringVar()
        cash_out_type_choice_map = {}
        staff_choice_map = {}

        tk.Label(
            cash_out_card, text=L("Expense Category:", "开销类别："),
            font=("Arial", 10, "bold"), bg=C['panel']
        ).grid(row=2, column=0, columnspan=2, padx=14, pady=(2, 3), sticky="w")
        quantity_label = tk.Label(
            cash_out_card, text=L("Quantity:", "数量："),
            font=("Arial", 10, "bold"), bg=C['panel']
        )
        quantity_label.grid(row=2, column=2, padx=14, pady=(2, 3), sticky="w")
        cash_out_amount_label = tk.Label(
            cash_out_card, text=L("Cash Amount (RM):", "现金金额 (RM)："),
            font=("Arial", 10, "bold"), bg=C['panel']
        )
        cash_out_amount_label.grid(
            row=2, column=3, padx=14, pady=(2, 3), sticky="w"
        )

        cash_out_type_combo = ttk.Combobox(
            cash_out_card, textvariable=cash_out_type_var,
            state="readonly", font=("Arial", 10), width=36
        )
        cash_out_type_combo.grid(
            row=3, column=0, columnspan=2, padx=14, pady=(0, 8), sticky="ew"
        )
        quantity_entry = tk.Entry(
            cash_out_card, textvariable=quantity_var,
            font=("Arial", 11), width=12, justify="right"
        )
        quantity_entry.grid(row=3, column=2, padx=14, pady=(0, 8), sticky="ew")
        cash_out_amount_entry = tk.Entry(
            cash_out_card, textvariable=cash_out_amount_var,
            font=("Arial", 13, "bold"), width=14, justify="right"
        )
        cash_out_amount_entry.grid(row=3, column=3, padx=14, pady=(0, 8), sticky="ew")

        cash_out_staff_label = tk.Label(
            cash_out_card, text=L("Staff Member:", "员工："),
            font=("Arial", 10, "bold"), bg=C['panel']
        )
        cash_out_staff_label.grid(row=4, column=0, padx=14, pady=(2, 3), sticky="w")
        cash_out_staff_combo = ttk.Combobox(
            cash_out_card, textvariable=cash_out_staff_var,
            state="readonly", font=("Arial", 10), width=24
        )
        cash_out_staff_combo.grid(row=5, column=0, padx=14, pady=(0, 12), sticky="ew")

        cash_out_notes_label = tk.Label(
            cash_out_card, text=L("Notes / Item:", "备注／物品："),
            font=("Arial", 10, "bold"), bg=C['panel']
        )
        cash_out_notes_label.grid(
            row=4, column=1, columnspan=2, padx=14, pady=(2, 3), sticky="w"
        )
        cash_out_notes_entry = tk.Entry(
            cash_out_card, textvariable=cash_out_notes_var,
            font=("Arial", 10)
        )
        cash_out_notes_entry.grid(
            row=5, column=1, columnspan=2, padx=14, pady=(0, 12), sticky="ew"
        )

        def reload_cash_out_choices(preferred_key=None):
            current = cash_out_type_choice_map.get(cash_out_type_var.get(), {})
            target_key = preferred_key or current.get("key")
            cash_out_type_choice_map.clear()
            staff_choice_map.clear()

            with self.db.get_connection() as conn:
                expense_rows = conn.execute(
                    """SELECT eo.id,eo.option_key,eo.name_en,eo.name_zh,eo.category,
                              eo.quantity_mode,eo.stock_item_id,eo.help_en,eo.help_zh,
                              si.name AS stock_item_name,si.stock AS stock_current,
                              si.active AS stock_active,si.track_stock AS stock_tracked
                       FROM expense_options eo
                       LEFT JOIN items si ON si.id=eo.stock_item_id
                       WHERE eo.active=1
                       ORDER BY eo.sort_order,eo.id"""
                ).fetchall()
                staff_rows = conn.execute(
                    "SELECT staff_id,name FROM staff WHERE active=1 ORDER BY name"
                ).fetchall()

            for row in staff_rows:
                display = f"{row['name']}  ({row['staff_id']})"
                staff_choice_map[display] = (str(row['staff_id']), str(row['name']))

            def add_choice(display, config):
                unique_display = display
                if unique_display in cash_out_type_choice_map:
                    suffix = config.get("item_id") or config.get("key")
                    unique_display = f"{display}  [{suffix}]"
                cash_out_type_choice_map[unique_display] = dict(config)

            for row in expense_rows:
                name_en = str(row['name_en'] or '').strip()
                name_zh = str(row['name_zh'] or '').strip() or name_en
                stock_item_id = row['stock_item_id']
                stock_item_name = str(row['stock_item_name'] or '').strip()
                linked_stock = stock_item_id is not None
                quantity_mode = str(row['quantity_mode'] or 'none').strip().lower()
                if linked_stock:
                    quantity_mode = "required"
                    fallback_en = (
                        f"Enter whole-unit quantity and total cash paid. Saving also adds stock to "
                        f"{stock_item_name or f'item #{stock_item_id}'}."
                    )
                    fallback_zh = (
                        f"填写整数数量与实付总额；保存后会同步增加"
                        f"{stock_item_name or f'商品 #{stock_item_id}'}的库存。"
                    )
                elif quantity_mode == "required":
                    fallback_en = "Enter the quantity and total cash paid. Write the actual item in Notes."
                    fallback_zh = "填写数量与实付总额，并在备注写明实际物品。"
                elif quantity_mode == "optional":
                    fallback_en = "Quantity is optional. Write the actual item or reason in Notes."
                    fallback_zh = "数量选填；请在备注写明实际物品或原因。"
                else:
                    fallback_en = "Enter the total cash paid and write the actual item or reason in Notes."
                    fallback_zh = "填写实付总额，并在备注写明实际物品或原因。"
                add_choice(L(name_en, name_zh), {
                    "key": f"expense:{row['option_key']}",
                    "purpose": "stock_purchase" if linked_stock else "expense",
                    "category": str(row['category'] or 'Other Expense').strip(),
                    "item_name": stock_item_name if linked_stock else "",
                    "item_id": int(stock_item_id) if linked_stock else None,
                    "detail_mode": "standard",
                    "quantity_mode": quantity_mode
                    if quantity_mode in ("required", "optional", "none") else "none",
                    "help": L(
                        str(row['help_en'] or '').strip() or fallback_en,
                        str(row['help_zh'] or '').strip() or fallback_zh,
                    ),
                })

            for config in cash_out_special_types:
                add_choice(config["label"], config)

            values = list(cash_out_type_choice_map)
            cash_out_type_combo['values'] = values
            selected = next((
                display for display, config in cash_out_type_choice_map.items()
                if config.get("key") == target_key
            ), values[0] if values else "")
            cash_out_type_var.set(selected)

        def update_cash_out_fields(_event=None):
            config = cash_out_type_choice_map.get(cash_out_type_var.get())
            if not config:
                cash_out_help_var.set(L(
                    "No active expense category is available.",
                    "目前没有已启用的开销类别。"
                ))
                return

            cash_out_help_var.set(config["help"])
            detail_mode = config["detail_mode"]
            quantity_mode = config["quantity_mode"]

            if detail_mode == "staff":
                staff_values = list(staff_choice_map)
                cash_out_staff_combo.configure(values=staff_values)
                cash_out_staff_var.set(staff_values[0] if staff_values else "")
                cash_out_staff_label.grid()
                cash_out_staff_combo.grid()
                cash_out_notes_label.grid_configure(column=1, columnspan=2)
                cash_out_notes_entry.grid_configure(column=1, columnspan=2)
            else:
                cash_out_staff_var.set("")
                cash_out_staff_label.grid_remove()
                cash_out_staff_combo.grid_remove()
                cash_out_notes_label.grid_configure(column=0, columnspan=3)
                cash_out_notes_entry.grid_configure(column=0, columnspan=3)

            if quantity_mode == "month":
                quantity_label.config(text=L("Deduction Month:", "扣薪月份："))
                quantity_var.set(datetime.now().strftime("%Y-%m"))
                quantity_entry.configure(state="normal")
                quantity_label.grid()
                quantity_entry.grid()
                cash_out_amount_label.grid_configure(column=3, columnspan=1)
                cash_out_amount_entry.grid_configure(column=3, columnspan=1)
            elif quantity_mode == "none":
                quantity_var.set("")
                quantity_label.grid_remove()
                quantity_entry.grid_remove()
                cash_out_amount_label.grid_configure(column=2, columnspan=2)
                cash_out_amount_entry.grid_configure(column=2, columnspan=2)
            elif quantity_mode == "optional":
                quantity_label.config(text=L("Quantity (optional):", "数量（可选）："))
                quantity_var.set("")
                quantity_entry.configure(state="normal")
                quantity_label.grid()
                quantity_entry.grid()
                cash_out_amount_label.grid_configure(column=3, columnspan=1)
                cash_out_amount_entry.grid_configure(column=3, columnspan=1)
            else:
                quantity_label.config(text=L("Quantity:", "数量："))
                quantity_var.set("1")
                quantity_entry.configure(state="normal")
                quantity_label.grid()
                quantity_entry.grid()
                cash_out_amount_label.grid_configure(column=3, columnspan=1)
                cash_out_amount_entry.grid_configure(column=3, columnspan=1)

        cash_out_type_combo.bind("<<ComboboxSelected>>", update_cash_out_fields)
        reload_cash_out_choices()
        update_cash_out_fields()

        cash_out_search_var = tk.StringVar()
        cash_out_result_var = tk.StringVar()

        def save_cash_out():
            config = cash_out_type_choice_map.get(cash_out_type_var.get())
            if not config:
                messagebox.showerror(
                    L("Missing Category", "缺少类别"),
                    L("Choose an expense category.", "请选择开销类别。"),
                    parent=win
                )
                return

            purpose = config["purpose"]
            category = config["category"]
            item_name = str(config.get("item_name") or "").strip()
            selected_advance_staff = None
            deduction_month = ""

            if config["detail_mode"] == "staff":
                staff_display = cash_out_staff_var.get().strip()
                selected_advance_staff = staff_choice_map.get(staff_display)
                if not selected_advance_staff:
                    messagebox.showerror(
                        L("Missing Staff", "缺少员工"),
                        L("Choose the staff member receiving the advance.", "请选择领取预支工资的员工。"),
                        parent=win
                    )
                    return
                deduction_month = quantity_var.get().strip()
                try:
                    datetime.strptime(deduction_month, "%Y-%m")
                except ValueError:
                    messagebox.showerror(
                        L("Invalid Month", "月份无效"),
                        L("Use YYYY-MM, for example 2026-10.", "请使用 YYYY-MM，例如 2026-10。"),
                        parent=win
                    )
                    return

            qty = None
            quantity_mode = config["quantity_mode"]
            quantity_text = quantity_var.get().strip()
            if quantity_mode in ("required", "optional") and (quantity_text or quantity_mode == "required"):
                try:
                    qty = float(quantity_text)
                    if qty <= 0:
                        raise ValueError
                    if purpose == "stock_purchase" and not qty.is_integer():
                        raise ValueError
                    if qty.is_integer():
                        qty = int(qty)
                except ValueError:
                    quantity_message = L(
                        "Enter a positive quantity. Tracked stock must use whole units.",
                        "请输入大于零的数量；追踪库存必须使用整数。"
                    ) if purpose == "stock_purchase" else L(
                        "Enter a positive quantity.", "请输入大于零的数量。"
                    )
                    messagebox.showerror(
                        L("Invalid Quantity", "数量无效"), quantity_message, parent=win
                    )
                    return

            try:
                amount = float(cash_out_amount_var.get().strip())
                if amount <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    L("Invalid Amount", "金额无效"),
                    L("Enter a positive cash amount.", "请输入大于零的现金金额。"),
                    parent=win
                )
                return

            selected_key = config["key"]
            item_id = config.get("item_id")

            def do_save(staff):
                try:
                    staff_payload = dict(staff)
                    if selected_advance_staff:
                        staff_payload['advance_staff_id'] = selected_advance_staff[0]
                        staff_payload['advance_staff_name'] = selected_advance_staff[1]
                    self._create_cash_out(
                        purpose, category, amount, staff_payload,
                        notes=cash_out_notes_var.get().strip(),
                        quantity=qty,
                        item_id=item_id,
                        deduction_month=deduction_month,
                        item_name=item_name
                    )
                    cash_out_amount_var.set("")
                    cash_out_notes_var.set("")
                    reload_cash_out_choices(selected_key)
                    update_cash_out_fields()
                    load_cash_out()
                    reload_summary()
                    refresh_closing_tab()
                    self.load_items()
                    open_cash_drawer()
                    messagebox.showinfo(
                        L("Saved", "已保存"),
                        L(
                            f"Cash out RM {amount:.2f} recorded.",
                            f"已记录现金转出 RM {amount:.2f}。"
                        ),
                        parent=win
                    )
                except Exception as exc:
                    messagebox.showerror(L("Unable to Save", "无法保存"), str(exc), parent=win)

            self._staff_login_popup(
                callback=do_save,
                title=L("👤 Staff PIN — Record Cash Out", "👤 员工 PIN — 记录现金转出"),
                parent=win
            )

        RoundedButton(
            cash_out_card, text="💾 " + L("Record Cash Out", "记录现金转出"),
            command=save_cash_out, bg=C['danger'], fg="white",
            font=("Arial", 11, "bold"), padx=18, pady=10, radius=8
        ).grid(row=5, column=3, padx=14, pady=(0, 12), sticky="e")
        cash_out_amount_entry.bind("<Return>", lambda _event: save_cash_out())

        cash_out_query_host = tk.Frame(tab_cash_out, bg=C['bg'])
        cash_out_query_host.pack(fill="x", padx=0, pady=(0, 2))

        cash_out_table_frame = tk.Frame(tab_cash_out, bg=C['bg'])
        cash_out_table_frame.pack(fill="both", expand=True, padx=12, pady=4)
        cash_out_table_frame.grid_rowconfigure(0, weight=1)
        cash_out_table_frame.grid_columnconfigure(0, weight=1)

        cash_out_columns = (
            "ID", "Time", "Category", "Notes / Item",
            "Qty / Month", "Amount", "Staff"
        )
        cash_out_tree = ttk.Treeview(
            cash_out_table_frame, columns=cash_out_columns,
            show="headings", height=12
        )
        cash_out_widths = {
            "ID": 0, "Time": 65, "Category": 180, "Notes / Item": 390,
            "Qty / Month": 95, "Amount": 95, "Staff": 110
        }
        cash_out_headings = {
            "ID": "ID", "Time": L("Time", "时间"),
            "Category": L("Category", "类别"),
            "Notes / Item": L("Notes / Item", "备注／物品"),
            "Qty / Month": L("Qty / Month", "数量／月份"),
            "Amount": L("Amount", "金额"),
            "Staff": L("Staff", "员工")
        }
        for col in cash_out_columns:
            cash_out_tree.heading(col, text=cash_out_headings[col])
            cash_out_tree.column(
                col, width=cash_out_widths[col],
                anchor="w" if col in ("Category", "Notes / Item") else "center"
            )
        cash_out_tree.column("ID", width=0, minwidth=0, stretch=False)
        cash_out_vsb = tk.Scrollbar(
            cash_out_table_frame, orient="vertical", command=cash_out_tree.yview
        )
        cash_out_hsb = tk.Scrollbar(
            cash_out_table_frame, orient="horizontal", command=cash_out_tree.xview
        )
        cash_out_tree.configure(
            yscrollcommand=cash_out_vsb.set, xscrollcommand=cash_out_hsb.set
        )
        cash_out_tree.grid(row=0, column=0, sticky="nsew")
        cash_out_vsb.grid(row=0, column=1, sticky="ns")
        cash_out_hsb.grid(row=1, column=0, sticky="ew")

        cash_out_sum_var = tk.StringVar()
        tk.Label(
            tab_cash_out, textvariable=cash_out_sum_var,
            font=("Arial", 11, "bold"), bg=C['bg'], fg=C['danger']
        ).pack(pady=3)

        def load_cash_out():
            cash_out_tree.delete(*cash_out_tree.get_children())
            keyword = cash_out_search_var.get().strip().casefold()
            with self.db.get_connection() as conn:
                rows = conn.execute(
                    """SELECT ct.*, sa.deduction_month
                       FROM cash_transactions ct
                       LEFT JOIN staff_advances sa ON sa.cash_transaction_id=ct.id
                       WHERE ct.date=?
                         AND ct.type IN ('expense','stock_purchase','staff_advance','cash_transfer')
                       ORDER BY ct.datetime DESC, ct.id DESC""",
                    (get_today(),)
                ).fetchall()

            if keyword:
                rows = [
                    row for row in rows
                    if keyword in " ".join((
                        str(row['id'] or ''), str(row['datetime'] or ''),
                        str(row['type'] or ''), str(row['item_name'] or ''),
                        str(row['category'] or ''), str(row['quantity'] or ''),
                        str(row['deduction_month'] or ''), str(row['amount'] or ''),
                        str(row['description'] or ''), str(row['recorded_by'] or '')
                    )).casefold()
                ]

            total = 0.0
            for row in rows:
                row_type = str(row['type'] or '')
                category_key = {
                    'staff_advance': 'Staff Advance',
                    'cash_transfer': 'Cash Transfer',
                }.get(row_type, str(row['category'] or ''))
                category_text = cash_out_category_labels.get(category_key, category_key or "-")
                item_text = str(row['item_name'] or '').strip()
                notes_text = str(row['description'] or '').strip()
                detail_parts = []
                if row_type in ('staff_advance', 'cash_transfer'):
                    if item_text:
                        detail_parts.append(item_text)
                    if notes_text:
                        detail_parts.append(notes_text)
                elif row_type == 'stock_purchase':
                    detail_parts.append(notes_text or item_text)
                else:
                    if item_text and item_text.casefold() != str(row['category'] or '').strip().casefold():
                        detail_parts.append(item_text)
                    if notes_text:
                        detail_parts.append(notes_text)
                detail_text = " — ".join(part for part in detail_parts if part) or "-"

                if row_type == 'staff_advance':
                    qty_or_month = row['deduction_month'] or "-"
                elif row['quantity'] is None:
                    qty_or_month = "-"
                else:
                    qty_number = float(row['quantity'])
                    qty_or_month = str(int(qty_number)) if qty_number.is_integer() else f"{qty_number:g}"
                dt_text = str(row['datetime'] or '')
                time_text = dt_text[11:16] if len(dt_text) >= 16 else dt_text
                cash_out_tree.insert("", "end", values=(
                    row['id'], time_text, category_text, detail_text,
                    qty_or_month, f"RM {float(row['amount']):.2f}",
                    row['recorded_by'] or "-"
                ))
                total += float(row['amount'] or 0)

            cash_out_sum_var.set(L(
                f"Today's total cash out: RM {total:.2f}",
                f"今日现金转出总额：RM {total:.2f}"
            ))
            cash_out_result_var.set(
                L(f"{len(rows)} cash-out record(s)", f"{len(rows)} 笔现金转出")
            )

        def void_selected_cash_out():
            selected = cash_out_tree.selection()
            if not selected:
                messagebox.showwarning(
                    L("No Selection", "未选择"),
                    L("Select a cash-out record first.", "请先选择一笔现金转出记录。"),
                    parent=win
                )
                return
            tx_id = int(cash_out_tree.item(selected[0])['values'][0])
            if not messagebox.askyesno(
                L("Confirm Void", "确认作废"),
                L(
                    "Void this cash-out record? Stock purchases also reverse the stock-in quantity.",
                    "确定作废这笔现金转出？若为采购入库，系统也会反向扣回入库数量。"
                ),
                parent=win
            ):
                return

            def do_void(staff):
                try:
                    self._void_cash_out(tx_id, staff)
                    load_cash_out()
                    reload_summary()
                    refresh_closing_tab()
                    self.load_items()
                except Exception as exc:
                    messagebox.showerror(L("Unable to Void", "无法作废"), str(exc), parent=win)

            self._staff_login_popup(
                callback=do_void,
                title=L("🔑 Admin PIN — Void Cash Out", "🔑 管理员 PIN — 作废现金转出"),
                parent=win, admin_only=True
            )

        self._build_query_toolbar(
            cash_out_query_host, cash_out_search_var, load_cash_out,
            on_refresh=load_cash_out,
            on_export=lambda: self._export_treeview_csv(
                cash_out_tree, win, "today_cash_out.csv"
            ),
            result_var=cash_out_result_var,
            search_label=L("Category / notes:", "类别／备注：")
        )
        cash_out_actions = tk.Frame(tab_cash_out, bg=C['bg'])
        cash_out_actions.pack(fill="x", padx=12, pady=(2, 6))
        RoundedButton(
            cash_out_actions,
            text="🗑 " + L("Void Selected (Admin PIN)", "作废所选（管理员 PIN）"),
            command=void_selected_cash_out, bg=C['neutral'], fg="white",
            font=("Arial", 10, "bold"), padx=14, pady=6, radius=8
        ).pack(side="left")
        load_cash_out()

        # ── KUIH DAILY SETTLEMENT TAB ──
        tab_kuih = tk.Frame(nb, bg=C['bg'])
        nb.add(tab_kuih, text="🍰 " + L("Kuih Daily Settlement", "糕点每日结算"))

        kuih_notice = tk.Frame(
            tab_kuih, bg="#FFF7ED",
            highlightbackground="#FDBA74", highlightthickness=1
        )
        kuih_notice.pack(fill="x", padx=12, pady=(10, 6))
        tk.Label(
            kuih_notice,
            text=L(
                "Cash only · Morning: batch cost · End of day: actual cash sold · Unsold kuih = internal consumption",
                "仅现金 · 早上记录本钱 · 收档记录实际卖到的现金 · 卖不完的糕点视为店内食用"
            ),
            font=("Arial", 10, "bold"), bg="#FFF7ED", fg="#9A3412",
            anchor="w", justify="left", wraplength=max(700, window_w - 100)
        ).pack(fill="x", padx=14, pady=9)

        kuih_forms = tk.Frame(tab_kuih, bg=C['bg'])
        kuih_forms.pack(fill="x", padx=12, pady=4)
        kuih_forms.grid_columnconfigure(0, weight=1)
        kuih_forms.grid_columnconfigure(1, weight=1)

        morning_card = tk.Frame(
            kuih_forms, bg=C['panel'],
            highlightbackground=C['border'], highlightthickness=1
        )
        morning_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        settle_card = tk.Frame(
            kuih_forms, bg=C['panel'],
            highlightbackground=C['border'], highlightthickness=1
        )
        settle_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        tk.Label(
            morning_card, text=L("1 · Morning Batch", "1 · 早上拿货"),
            font=("Arial", 13, "bold"), bg=C['panel'], fg=C['primary']
        ).grid(row=0, column=0, columnspan=2, padx=14, pady=(12, 8), sticky="w")
        tk.Label(
            morning_card, text=L("Cost Paid (RM):", "实际本钱 (RM)："),
            font=("Arial", 10, "bold"), bg=C['panel']
        ).grid(row=1, column=0, padx=14, pady=4, sticky="w")
        kuih_cost_var = tk.StringVar(value="20.00")
        tk.Entry(
            morning_card, textvariable=kuih_cost_var,
            font=("Arial", 12, "bold"), justify="right", width=14
        ).grid(row=1, column=1, padx=14, pady=4, sticky="ew")
        tk.Label(
            morning_card, text=L("Full-Sale Value (RM):", "若卖完金额 (RM)："),
            font=("Arial", 10, "bold"), bg=C['panel']
        ).grid(row=2, column=0, padx=14, pady=4, sticky="w")
        kuih_target_var = tk.StringVar(value="30.00")
        tk.Entry(
            morning_card, textvariable=kuih_target_var,
            font=("Arial", 12), justify="right", width=14
        ).grid(row=2, column=1, padx=14, pady=4, sticky="ew")
        tk.Label(
            morning_card, text=L("Notes:", "备注："),
            font=("Arial", 10, "bold"), bg=C['panel']
        ).grid(row=3, column=0, padx=14, pady=4, sticky="w")
        kuih_purchase_notes_var = tk.StringVar()
        tk.Entry(
            morning_card, textvariable=kuih_purchase_notes_var,
            font=("Arial", 10)
        ).grid(row=3, column=1, padx=14, pady=4, sticky="ew")
        morning_card.grid_columnconfigure(1, weight=1)

        tk.Label(
            settle_card, text=L("2 · End-of-Day Settlement", "2 · 收档结算"),
            font=("Arial", 13, "bold"), bg=C['panel'], fg="#2563EB"
        ).grid(row=0, column=0, columnspan=2, padx=14, pady=(12, 8), sticky="w")
        tk.Label(
            settle_card, text=L("Actual Cash Sales (RM):", "实际卖到现金 (RM)："),
            font=("Arial", 10, "bold"), bg=C['panel']
        ).grid(row=1, column=0, padx=14, pady=4, sticky="w")
        kuih_sales_var = tk.StringVar()
        tk.Entry(
            settle_card, textvariable=kuih_sales_var,
            font=("Arial", 12, "bold"), justify="right", width=14
        ).grid(row=1, column=1, padx=14, pady=4, sticky="ew")
        tk.Label(
            settle_card, text=L("Notes:", "备注："),
            font=("Arial", 10, "bold"), bg=C['panel']
        ).grid(row=2, column=0, padx=14, pady=4, sticky="w")
        kuih_settlement_notes_var = tk.StringVar()
        tk.Entry(
            settle_card, textvariable=kuih_settlement_notes_var,
            font=("Arial", 10)
        ).grid(row=2, column=1, padx=14, pady=4, sticky="ew")
        tk.Label(
            settle_card,
            text=L(
                "Example: cost RM20, actual sales RM25 → actual margin RM5.",
                "例如：本钱 RM20，实际卖 RM25 → 实际毛差 RM5。"
            ),
            font=("Arial", 9), bg=C['panel'], fg=C['text_sub']
        ).grid(row=3, column=0, columnspan=2, padx=14, pady=4, sticky="w")
        settle_card.grid_columnconfigure(1, weight=1)

        kuih_search_var = tk.StringVar()
        kuih_result_var = tk.StringVar()
        kuih_today_var = tk.StringVar()

        def refresh_kuih_status():
            batch = self._get_kuih_batch(get_today())
            if not batch:
                kuih_today_var.set(L(
                    "Today: no batch recorded yet.",
                    "今天：尚未记录糕点拿货。"
                ))
                return
            result = calculate_kuih_batch_summary(
                batch['cost_amount'], batch['actual_sales'] or 0,
                batch['full_sale_value'] or 0
            )
            if str(batch['status']).upper() == 'OPEN':
                kuih_today_var.set(L(
                    f"Today OPEN · Cost RM {result['cost_amount']:.2f} · "
                    f"Full-sale RM {result['full_sale_value']:.2f} · "
                    f"Expected margin RM {result['expected_margin']:.2f} · waiting for actual sales",
                    f"今天未结算 · 本钱 RM {result['cost_amount']:.2f} · "
                    f"卖完金额 RM {result['full_sale_value']:.2f} · "
                    f"预计毛差 RM {result['expected_margin']:.2f} · 等待实际销售"
                ))
            else:
                kuih_today_var.set(L(
                    f"Today CLOSED · Cost RM {result['cost_amount']:.2f} · "
                    f"Actual sales RM {result['actual_sales']:.2f} · "
                    f"Actual margin RM {result['actual_margin']:.2f} · "
                    f"Unsold value RM {result['unsold_value']:.2f} (internal consumption)",
                    f"今天已结算 · 本钱 RM {result['cost_amount']:.2f} · "
                    f"实际销售 RM {result['actual_sales']:.2f} · "
                    f"实际毛差 RM {result['actual_margin']:.2f} · "
                    f"未售价值 RM {result['unsold_value']:.2f}（店内食用）"
                ))

        def start_kuih_batch():
            try:
                cost = float(kuih_cost_var.get().strip())
                target = float(kuih_target_var.get().strip() or 0)
                if cost <= 0 or target < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    L("Invalid Amount", "金额无效"),
                    L(
                        "Enter a positive cost and a non-negative full-sale value.",
                        "请输入大于零的本钱及不小于零的卖完金额。"
                    ), parent=win
                )
                return

            def do_start(staff):
                try:
                    self._create_kuih_batch(
                        cost, target, kuih_purchase_notes_var.get().strip(),
                        staff, get_today()
                    )
                    kuih_purchase_notes_var.set("")
                    load_kuih_batches()
                    reload_summary()
                    refresh_closing_tab()
                    open_cash_drawer()
                    messagebox.showinfo(
                        L("Batch Started", "已记录拿货"),
                        L(
                            f"Morning kuih cost RM {cost:.2f} recorded as cash out.",
                            f"已把早上糕点本钱 RM {cost:.2f} 记录为现金转出。"
                        ), parent=win
                    )
                except Exception as exc:
                    messagebox.showerror(L("Unable to Start", "无法记录"), str(exc), parent=win)

            self._staff_login_popup(
                callback=do_start,
                title=L("👤 Staff PIN — Start Kuih Batch", "👤 员工 PIN — 记录糕点拿货"),
                parent=win
            )

        def settle_kuih_batch():
            try:
                sales = float(kuih_sales_var.get().strip())
                if sales < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    L("Invalid Amount", "金额无效"),
                    L(
                        "Enter actual cash sales, including 0 if none sold.",
                        "请输入实际卖到的现金；若完全没卖可填 0。"
                    ), parent=win
                )
                return

            def do_settle(staff):
                try:
                    batch = self._settle_kuih_batch(
                        sales, kuih_settlement_notes_var.get().strip(),
                        staff, get_today()
                    )
                    result = calculate_kuih_batch_summary(
                        batch['cost_amount'], batch['actual_sales'],
                        batch['full_sale_value']
                    )
                    kuih_sales_var.set("")
                    kuih_settlement_notes_var.set("")
                    load_kuih_batches()
                    reload_summary()
                    refresh_closing_tab()
                    if sales > 0:
                        open_cash_drawer()
                    messagebox.showinfo(
                        L("Settlement Saved", "结算已保存"),
                        L(
                            f"Actual sales RM {sales:.2f}\nActual margin RM {result['actual_margin']:.2f}",
                            f"实际销售 RM {sales:.2f}\n实际毛差 RM {result['actual_margin']:.2f}"
                        ), parent=win
                    )
                except Exception as exc:
                    messagebox.showerror(L("Unable to Settle", "无法结算"), str(exc), parent=win)

            self._staff_login_popup(
                callback=do_settle,
                title=L("👤 Staff PIN — Settle Kuih", "👤 员工 PIN — 糕点结算"),
                parent=win
            )

        RoundedButton(
            morning_card, text="🌅 " + L("Start Morning Batch", "记录早上拿货"),
            command=start_kuih_batch, bg=C['danger'], fg="white",
            font=("Arial", 10, "bold"), padx=16, pady=8, radius=8
        ).grid(row=4, column=0, columnspan=2, padx=14, pady=(8, 12))
        RoundedButton(
            settle_card, text="✅ " + L("Settle Actual Sales", "结算实际销售"),
            command=settle_kuih_batch, bg="#2563EB", fg="white",
            font=("Arial", 10, "bold"), padx=16, pady=8, radius=8
        ).grid(row=4, column=0, columnspan=2, padx=14, pady=(8, 12))

        kuih_status_card = tk.Frame(
            tab_kuih, bg="#EFF6FF",
            highlightbackground="#93C5FD", highlightthickness=1
        )
        kuih_status_card.pack(fill="x", padx=12, pady=6)
        tk.Label(
            kuih_status_card, textvariable=kuih_today_var,
            font=("Arial", 10, "bold"), bg="#EFF6FF", fg="#1E3A8A",
            anchor="w", justify="left", wraplength=max(700, window_w - 100)
        ).pack(fill="x", padx=14, pady=9)

        kuih_query_host = tk.Frame(tab_kuih, bg=C['bg'])
        kuih_query_host.pack(fill="x", padx=0, pady=(0, 2))

        kuih_table_frame = tk.Frame(tab_kuih, bg=C['bg'])
        kuih_table_frame.pack(fill="both", expand=True, padx=12, pady=4)
        kuih_table_frame.grid_rowconfigure(0, weight=1)
        kuih_table_frame.grid_columnconfigure(0, weight=1)
        kuih_columns = (
            "ID", "Date", "Cost", "Full Sale", "Actual Sales",
            "Expected Margin", "Actual Margin", "Status", "Opened By", "Settled By"
        )
        kuih_tree = ttk.Treeview(
            kuih_table_frame, columns=kuih_columns,
            show="headings", height=9
        )
        kuih_widths = {
            "ID": 0, "Date": 100, "Cost": 90, "Full Sale": 90,
            "Actual Sales": 95, "Expected Margin": 105, "Actual Margin": 95,
            "Status": 90, "Opened By": 100, "Settled By": 100
        }
        kuih_headings = {
            "ID": "ID", "Date": L("Date", "日期"),
            "Cost": L("Cost", "本钱"), "Full Sale": L("Full Sale", "卖完金额"),
            "Actual Sales": L("Actual Sales", "实际销售"),
            "Expected Margin": L("Expected Margin", "预计毛差"),
            "Actual Margin": L("Actual Margin", "实际毛差"),
            "Status": L("Status", "状态"),
            "Opened By": L("Opened By", "拿货记录"),
            "Settled By": L("Settled By", "结算员工")
        }
        for col in kuih_columns:
            kuih_tree.heading(col, text=kuih_headings[col])
            kuih_tree.column(col, width=kuih_widths[col], anchor="center")
        kuih_tree.column("ID", width=0, minwidth=0, stretch=False)
        kuih_vsb = tk.Scrollbar(
            kuih_table_frame, orient="vertical", command=kuih_tree.yview
        )
        kuih_hsb = tk.Scrollbar(
            kuih_table_frame, orient="horizontal", command=kuih_tree.xview
        )
        kuih_tree.configure(
            yscrollcommand=kuih_vsb.set, xscrollcommand=kuih_hsb.set
        )
        kuih_tree.grid(row=0, column=0, sticky="nsew")
        kuih_vsb.grid(row=0, column=1, sticky="ns")
        kuih_hsb.grid(row=1, column=0, sticky="ew")

        def load_kuih_batches():
            kuih_tree.delete(*kuih_tree.get_children())
            keyword = kuih_search_var.get().strip().casefold()
            with self.db.get_connection() as conn:
                rows = conn.execute(
                    "SELECT * FROM kuih_batches ORDER BY date DESC, id DESC LIMIT 90"
                ).fetchall()
            if keyword:
                rows = [
                    row for row in rows
                    if keyword in " ".join((
                        str(row['date'] or ''), str(row['status'] or ''),
                        str(row['opened_by'] or ''), str(row['settled_by'] or ''),
                        str(row['purchase_notes'] or ''),
                        str(row['settlement_notes'] or '')
                    )).casefold()
                ]
            for row in rows:
                result = calculate_kuih_batch_summary(
                    row['cost_amount'], row['actual_sales'] or 0,
                    row['full_sale_value'] or 0
                )
                closed = str(row['status']).upper() == 'CLOSED'
                kuih_tree.insert("", "end", values=(
                    row['id'], row['date'],
                    f"RM {result['cost_amount']:.2f}",
                    f"RM {result['full_sale_value']:.2f}",
                    f"RM {result['actual_sales']:.2f}" if closed else "-",
                    f"RM {result['expected_margin']:.2f}",
                    f"RM {result['actual_margin']:.2f}" if closed else "-",
                    L("Closed", "已结算") if closed else L("Open", "未结算"),
                    row['opened_by'] or "-", row['settled_by'] or "-"
                ))
            kuih_result_var.set(L(f"{len(rows)} batch(es)", f"{len(rows)} 批记录"))
            refresh_kuih_status()

        def void_selected_kuih_batch():
            selected = kuih_tree.selection()
            if not selected:
                messagebox.showwarning(
                    L("No Selection", "未选择"),
                    L("Select a kuih batch first.", "请先选择一批糕点记录。"),
                    parent=win
                )
                return
            values = kuih_tree.item(selected[0])['values']
            batch_id = int(values[0])
            batch_date = str(values[1])
            if not messagebox.askyesno(
                L("Confirm Void", "确认作废"),
                L(
                    f"Void the kuih batch for {batch_date} and both linked cash records?",
                    f"作废 {batch_date} 的糕点批次及关联的现金记录？"
                ), parent=win
            ):
                return

            def do_void(staff):
                try:
                    self._void_kuih_batch(batch_id, staff)
                    load_kuih_batches()
                    reload_summary()
                    refresh_closing_tab()
                except Exception as exc:
                    messagebox.showerror(L("Unable to Void", "无法作废"), str(exc), parent=win)

            self._staff_login_popup(
                callback=do_void,
                title=L("🔑 Admin PIN — Void Kuih Batch", "🔑 管理员 PIN — 作废糕点批次"),
                parent=win, admin_only=True
            )

        self._build_query_toolbar(
            kuih_query_host, kuih_search_var, load_kuih_batches,
            on_refresh=load_kuih_batches,
            on_export=lambda: self._export_treeview_csv(
                kuih_tree, win, "kuih_daily_settlement.csv"
            ),
            result_var=kuih_result_var,
            search_label=L("Date / status / staff:", "日期／状态／员工：")
        )
        kuih_actions = tk.Frame(tab_kuih, bg=C['bg'])
        kuih_actions.pack(fill="x", padx=12, pady=(2, 6))
        RoundedButton(
            kuih_actions,
            text="🗑 " + L("Void Selected (Admin PIN)", "作废所选（管理员 PIN）"),
            command=void_selected_kuih_batch, bg=C['neutral'], fg="white",
            font=("Arial", 10, "bold"), padx=14, pady=6, radius=8
        ).pack(side="left")
        load_kuih_batches()
        reload_summary()
        # ── END OF DAY TAB ──
        tab_close_outer = tk.Frame(nb, bg=C['bg'])
        nb.add(tab_close_outer, text="🌙 " + L("End of Day", "结账日结"))
        
        close_canvas = tk.Canvas(tab_close_outer, bg=C['bg'], highlightthickness=0)
        close_vsb = tk.Scrollbar(tab_close_outer, orient="vertical", command=close_canvas.yview)
        tab_close = tk.Frame(close_canvas, bg=C['bg'])
        close_canvas.configure(yscrollcommand=close_vsb.set)
        close_canvas.pack(side="left", fill="both", expand=True)
        close_vsb.pack(side="right", fill="y")
        close_win = close_canvas.create_window((0,0), window=tab_close, anchor="nw")
        
        def _on_canvas_config(e):
            close_canvas.itemconfig(close_win, width=e.width)
        close_canvas.bind("<Configure>", _on_canvas_config)

        def _on_tab_config(e):
            close_canvas.configure(scrollregion=close_canvas.bbox("all"))
        tab_close.bind("<Configure>", _on_tab_config)

        def _on_mousewheel(e):
            try:
                f = win.focus_get()
                if f and "tab_close" in str(f):
                    close_canvas.yview_scroll(-1*(e.delta//120), "units")
            except: pass
        tab_close.bind_all("<MouseWheel>", _on_mousewheel)

        # Header for the tab
        tk.Label(tab_close, text="✨ " + L("Daily Cash Reconciliation", "今日现金核对"), 
                 font=("Arial", 28, "bold"), bg=C['bg'], fg=C['primary']).pack(anchor="center", pady=(40, 10))
        
        main_content = tk.Frame(tab_close, bg=C['bg'])
        main_content.pack(pady=20, expand=True)

        calc_f = tk.Frame(main_content, bg=C['bg'])
        calc_f.pack()
        calc_f.columnconfigure(0, weight=1); calc_f.columnconfigure(1, weight=1)

        # --- Section 1: Physical Cash Reconciliation ---
        reconcile_frame = tk.Frame(calc_f, bg=C['bg'])
        reconcile_frame.grid(row=0, column=0, sticky="nsew", padx=(0,15))
        
        tk.Label(reconcile_frame, text="📌 " + L("Physical Cash Box", "实物现金箱"),
                 font=("Arial", 12, "bold"), bg=C['bg'], fg=C['text_sub']).pack(anchor="w", pady=(0,8))

        rec_inner = tk.Frame(reconcile_frame, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        rec_inner.pack(fill="x")

        def _card_line(parent, label, color, is_bold=False):
            f = tk.Frame(parent, bg=C['panel'], padx=20, pady=8)
            f.pack(fill="x")
            tk.Label(f, text=label, font=("Arial", 12, "bold" if is_bold else "normal"),
                     bg=C['panel'], fg=C['text_sub']).pack(side="left")
            lbl = tk.Label(f, text="RM 0.00", font=("Arial", 13, "bold" if is_bold else "normal"),
                           bg=C['panel'], fg=color)
            lbl.pack(side="right")
            return lbl

        lbl_c_open = _card_line(rec_inner, L("🪙 Opening Cash", "🪙 开盘现金"),   "#F59E0B")
        lbl_c_cash = _card_line(rec_inner, L("💻 Cash Sales", "💻 现金销售"),    C['success'])
        lbl_c_inc  = _card_line(rec_inner, L("➕ Cash Receipts", "➕ 其他现金收入"),  "#2563EB")
        lbl_c_exp  = _card_line(rec_inner, L("💸 Cash Out", "💸 现金转出"),          C['danger'])
        
        # Big Expected Box
        exp_box_f = tk.Frame(reconcile_frame, bg="#0F172A", padx=20, pady=15, 
                             highlightbackground="#1E293B", highlightthickness=1)
        exp_box_f.pack(fill="x", pady=(10,0))
        tk.Label(exp_box_f, text=L("✅ Expected Cash in Box:", "✅ 现金箱应有现金："),
                 font=("Arial", 13, "bold"), bg="#0F172A", fg="#94A3B8").pack(side="left")
        lbl_c_expected_val = tk.Label(exp_box_f, text="RM 0.00", font=("Arial", 22, "bold"),
                                      bg="#0F172A", fg="#facc15")
        lbl_c_expected_val.pack(side="right")

        # --- Section 2: Sales Metrics (Total Revenue) ---
        metrics_frame = tk.Frame(calc_f, bg=C['bg'])
        metrics_frame.grid(row=0, column=1, sticky="nsew", padx=(15,0))
        
        tk.Label(metrics_frame, text="📊 " + L("Revenue Breakdown", "销售明细"),
                 font=("Arial", 12, "bold"), bg=C['bg'], fg=C['text_sub']).pack(anchor="w", pady=(0,8))

        met_inner = tk.Frame(metrics_frame, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
        met_inner.pack(fill="x")

        def _met_card(parent, label, color, bg_acc):
            f = tk.Frame(parent, bg=bg_acc, padx=16, pady=12)
            f.pack(fill="x", pady=2, padx=2)
            tk.Label(f, text=label, font=("Arial", 11, "bold"),
                     bg=bg_acc, fg="white").pack(anchor="w")
            lbl = tk.Label(f, text="RM 0.00", font=("Arial", 18, "bold"),
                           bg=bg_acc, fg="white")
            lbl.pack(anchor="e")
            return lbl

        lbl_c_tot = _met_card(met_inner, L("🧾 Total Net Cash", "🧾 现金总净额"),   C['primary'], "#1E3A8A")
        lbl_c_qr = _met_card(met_inner, L("📱 QR Total", "📱 QR 总额"), "#0369A1", "#0E7490")

        # --- Section 3: User Input (Actual Cash & Difference) ---
        entry_f = tk.Frame(main_content, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1, padx=40, pady=30)
        entry_f.pack(pady=30, fill="x")

        input_row = tk.Frame(entry_f, bg=C['panel'])
        input_row.pack(fill="x")

        # Actual Cash Input
        left_inp = tk.Frame(input_row, bg=C['panel'])
        left_inp.pack(side="left", expand=True)
        tk.Label(left_inp, text=L("实点现金 (Actual Cash):", "实点现金 (Actual Cash):"), 
                 font=("Arial", 14, "bold"), bg=C['panel'], fg=C['primary']).pack(anchor="w")
        
        actual_var = tk.StringVar()
        actual_entry = tk.Entry(left_inp, textvariable=actual_var, font=("Arial", 28, "bold"), 
                                width=12, justify="right", relief="flat", highlightbackground=C['border'], highlightthickness=1)
        actual_entry.pack(pady=5)
        # Difference Display
        right_diff = tk.Frame(input_row, bg=C['panel'])
        right_diff.pack(side="left", expand=True)
        tk.Label(right_diff, text=L("差额 (Difference):", "差额 (Difference):"), 
                 font=("Arial", 14, "bold"), bg=C['panel'], fg=C['text_sub']).pack(anchor="w")
        
        diff_lbl = tk.Label(right_diff, text="—", font=("Arial", 32, "bold"), bg=C['panel'], fg=C['text_muted'])
        diff_lbl.pack(pady=5)

        # Notes for closing (Removed per user request)
        notes_var2 = tk.StringVar(value="")

        def on_actual_change(*_):
            try:
                val = actual_var.get().strip()
                if not val:
                    diff_lbl.config(text="—", fg=C['text_muted']); return
                actual = float(val)
                _, _, _, expected, _ = reload_summary()
                diff = actual - expected
                color = C['success'] if abs(diff) < 0.01 else (C['warning'] if abs(diff) < 5 else C['danger'])
                sign = "+" if diff >= 0 else ""
                diff_lbl.config(text=f"{sign}RM {diff:.2f}", fg=color)
            except ValueError:
                diff_lbl.config(text="Error", fg=C['danger'])

        actual_var.trace_add("write", on_actual_change)

        def refresh_closing_tab():
            _pos_cash, other_income, expenses, expected, opening = reload_summary()
            biz = get_today()
            with self.db.get_connection() as conn:
                total_sales, cash_sales, qr_sales, _ = compute_sales_payment_breakdown(conn, biz)
                qr_rent = compute_qr_rent_total(conn, biz)
            lbl_c_open.config(text=f"RM {opening:.2f}")
            lbl_c_tot.config(text=f"RM {cash_sales + other_income - expenses:.2f}")
            lbl_c_cash.config(text=f"RM {cash_sales:.2f}")
            lbl_c_qr.config(text=f"RM {qr_sales + qr_rent:.2f}")
            lbl_c_inc.config(text=f"RM {other_income:.2f}")
            lbl_c_exp.config(text=f"RM {expenses:.2f}")
            lbl_c_expected_val.config(text=f"RM {expected:.2f}")
            on_actual_change()

        # Logic for do_close removed per user request


        # Actions at bottom removed per user request

        
        refresh_closing_tab()

        # Footer Strip for global actions
        btn_strip = tk.Frame(win, bg=C['bg'])
        btn_strip.pack(fill="x", padx=16, pady=(8, 12))
        
        RoundedButton(btn_strip, text=L("🗄️ Drawer", "🗄️ 钱箱"),
                      command=self.manual_open_cash_drawer,
                      bg=C['neutral'], fg="white", font=("Arial", 11, "bold"),
                      padx=16, pady=10, radius=8).pack(side="left", padx=(0, 10))

        RoundedButton(btn_strip, text=L("🔄 Refresh Data", "🔄 刷新数据"),
                      command=lambda: [reload_summary(), refresh_closing_tab(), load_cash_out(), load_kuih_batches()],
                      bg=C['neutral'], fg="white", font=("Arial", 11, "bold"),
                      padx=16, pady=10, radius=8).pack(side="left")

        RoundedButton(btn_strip, text=L("✖ Close Cashier", "✖ 关闭收银"), command=win.destroy,
                      bg=C['header'], fg="white", font=("Arial", 11, "bold"),
                      padx=20, pady=10, radius=8).pack(side="right")

        # ═══════════════════════════════════════════════════════════════
        reload_summary()
        load_cash_out()
        load_kuih_batches()
        refresh_closing_tab()
        if cashier_mode:
            _boost_fonts(win, delta=2)

    # ================= Export Daily Report EXCEL =================
    def _export_daily_excel(self, biz_date_str, parent=None, apply_filter=False):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        import tkinter.filedialog as fd
        if not _OPENPYXL_OK:
            messagebox.showerror("Missing Dependency", "Please install openpyxl:\npip install openpyxl", parent=parent)
            return

        start_dt, end_dt = business_day_range(biz_date_str)
        s_str = start_dt.strftime("%Y-%m-%d %H:%M:%S")
        e_str = end_dt.strftime("%Y-%m-%d %H:%M:%S")

        with self.db.get_connection() as conn:
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            excluded_cats = set()
            if apply_filter:
                c.execute("SELECT name FROM product_categories WHERE exclude_export=1")
                excluded_cats = {r['name'] for r in c.fetchall()}

            c.execute("SELECT * FROM sales WHERE to_iso(datetime) >= ? AND to_iso(datetime) < ? AND amended=0 ORDER BY to_iso(datetime)", (s_str, e_str))
            sales_raw = c.fetchall()
            c.execute("SELECT description, amount, recorded_by, datetime as recorded_at FROM cash_transactions WHERE type='expense' AND date=? ORDER BY datetime", (biz_date_str,))
            expense_rows = c.fetchall()
            c.execute("SELECT description, amount, recorded_by, datetime as recorded_at FROM cash_transactions WHERE type='income' AND date=? ORDER BY datetime", (biz_date_str,))
            income_rows = c.fetchall()
            c.execute("SELECT name, category, stock FROM items WHERE active=1 AND track_stock=1 ORDER BY category, name")
            stock_rows = c.fetchall()
            c.execute("SELECT opening_cash, pos_cash, other_income, expenses, expected_cash, actual_cash, difference, cashier, closed_at, notes FROM cash_closing WHERE closing_date=? ORDER BY rowid DESC LIMIT 1", (biz_date_str,))
            closing = c.fetchone()

        sale_rows = []
        item_tally = {}
        total_sales_calc = 0.0
        for s in sales_raw:
            items_list = json.loads(s['items'])
            receipt_exported_total = 0.0
            
            for entry in items_list:
                name = entry[0]; price = entry[1]; qty = entry[2]
                cat  = entry[3] if len(entry) > 3 else ''
                
                if apply_filter and (cat in excluded_cats):
                    continue
                    
                line_val = float(Decimal(str(price)) * int(qty))
                receipt_exported_total += line_val
                
                key = (name, cat)
                if key not in item_tally:
                    item_tally[key] = {'qty': 0, 'revenue': 0.0}
                item_tally[key]['qty'] += qty
                item_tally[key]['revenue'] += line_val
                sale_rows.append((s['id'], s['datetime'], name, cat, qty, price, s['total'], s['payment'], s['cashier'] or ''))
            
            total_sales_calc += receipt_exported_total

        if apply_filter:
            total_sales = total_sales_calc
        else:
            total_sales = sum(s['total'] for s in sales_raw) # Keep original total behavior for non-filtered

        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        HDR_FILL = PatternFill("solid", start_color="1E3A5F")
        HDR_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=11)
        thin = Side(style="thin", color="CCCCCC")
        BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
        CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
        LEFT = Alignment(horizontal="left", vertical="center")
        RIGHT = Alignment(horizontal="right", vertical="center")

        def hdr_cell(ws, row, col, val):
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = HDR_FILL; cell.font = HDR_FONT
            cell.alignment = CENTER; cell.border = BORDER
            return cell

        def data_cell(ws, row, col, val, align=None, bold=False, num_fmt=None):
            cell = ws.cell(row=row, column=col, value=val)
            cell.alignment = align or LEFT; cell.border = BORDER
            cell.font = Font(name="Arial", size=10, bold=bold)
            if num_fmt: cell.number_format = num_fmt
            return cell

        ws1 = wb.create_sheet("Summary")
        ws1.column_dimensions["A"].width = 28
        ws1.column_dimensions["B"].width = 20
        ws1.cell(row=1, column=1, value=f"Daily Report {biz_date_str}").font = Font(name="Arial", size=16, bold=True)
        ws1.merge_cells("A1:B1")
        total_exp = sum(r["amount"] for r in expense_rows)
        total_inc = sum(r["amount"] for r in income_rows)
        rows_data = [("💰 POS Sales", total_sales), ("➕ Other Income", total_inc), ("➖ Expenses", total_exp), ("📦 Net Revenue", total_sales + total_inc - total_exp)]
        if closing:
            rows_data += [("🏦 Opening Cash", closing["opening_cash"]), ("✅ Expected Cash", closing["expected_cash"]), ("🔢 Actual Cash", closing["actual_cash"]), ("⚖️ Difference", closing["difference"]), ("👤 Closing Staff", closing["cashier"]), ("📝 Closing Remark", closing["notes"] or "")]
        r_idx = 3
        hdr_cell(ws1, r_idx, 1, "Item"); hdr_cell(ws1, r_idx, 2, "Amount (RM)")
        r_idx += 1
        for label, val in rows_data:
            data_cell(ws1, r_idx, 1, label)
            if isinstance(val, float):
                data_cell(ws1, r_idx, 2, val, align=RIGHT, num_fmt='"RM"#,##0.00')
            else:
                data_cell(ws1, r_idx, 2, val, align=CENTER)
            r_idx += 1

        ws2 = wb.create_sheet("Item Sales")
        cols2 = ["Item Name", "Category", "Sold Qty", "Revenue (RM)", "Current Stock"]
        for ci, h in enumerate(cols2, 1):
            hdr_cell(ws2, 1, ci, h)
            ws2.column_dimensions[get_column_letter(ci)].width = 16
        stock_map = {row['name']: row['stock'] for row in stock_rows}
        r_idx = 2
        for (nm, cat), tally in sorted(item_tally.items(), key=lambda x: (x[0][1], x[0][0])):
            qty = tally['qty']; rev = tally['revenue']
            stk = stock_map.get(nm)
            data_cell(ws2, r_idx, 1, nm)
            data_cell(ws2, r_idx, 2, cat, align=CENTER)
            data_cell(ws2, r_idx, 3, qty, align=CENTER)
            data_cell(ws2, r_idx, 4, rev, align=RIGHT, num_fmt='"RM"#,##0.00')
            data_cell(ws2, r_idx, 5, stk if stk is not None else "", align=CENTER)
            r_idx += 1

        default_name = f"DailyReport_{biz_date_str}.xlsx"
        file_path = fd.asksaveasfilename(parent=parent, initialfile=default_name, defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Save Daily Report")
        if not file_path: return
        wb.save(file_path)
        messagebox.showinfo("✅ Export Success", f"Saved to:\n{file_path}", parent=parent)

    def show_profit_center(self, parent=None):
        C = self.colors
        L = self.l
        win = tk.Toplevel(parent or self.root)
        win.title(L("💰 POS Cashflow Summary", "💰 POS 现金流汇总"))
        win.configure(bg=C['bg'])
        win.transient(self.root)
        win.grab_set()
        win.resizable(True, True)
        self._center_window(win, 1000, 680, self.root)

        # ── Header ────────────────────────────────────────────
        hdr = tk.Frame(win, bg=C['header'], height=58)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=L("💰 POS Cashflow Summary", "💰 POS 现金流汇总"), font=("Arial", 17, "bold"),
                 bg=C['header'], fg="white").pack(side="left", padx=18, pady=12)
        tk.Label(hdr, text=L("POS-recorded movement only · Not final accounting profit", "仅限 POS 已记录资金流 · 不是最终会计利润"),
                 font=("Arial", 10), bg=C['header'], fg="#94A3B8").pack(side="left", pady=16)

        # ── Filter bar ────────────────────────────────────────
        from datetime import datetime as _dt
        default_from = _dt.now().replace(day=1).strftime("%Y-%m-%d")
        default_to   = _dt.now().strftime("%Y-%m-%d")

        fbar = tk.Frame(win, bg=C['bg'])
        fbar.pack(fill="x", padx=16, pady=(10, 0))

        date_filter_row = tk.Frame(fbar, bg=C['bg'])
        date_filter_row.pack(fill="x")
        detail_filter_row = tk.Frame(fbar, bg=C['bg'])
        detail_filter_row.pack(fill="x", pady=(6, 0))

        tk.Label(date_filter_row, text=L("From:", "开始："), font=("Arial", 10, "bold"), bg=C['bg'], fg=C['text']).pack(side="left")
        from_var = tk.StringVar(value=default_from)
        dp_from = DatePicker(
            date_filter_row, textvariable=from_var,
            language=self.lang
        )
        dp_from.pack(side="left", padx=(4, 12))

        tk.Label(date_filter_row, text=L("To:", "结束："), font=("Arial", 10, "bold"), bg=C['bg'], fg=C['text']).pack(side="left")
        to_var = tk.StringVar(value=default_to)
        dp_to = DatePicker(
            date_filter_row, textvariable=to_var,
            language=self.lang
        )
        dp_to.pack(side="left", padx=(4, 12))

        all_type_label = L("All", "全部")
        type_display_to_db = {
            all_type_label: "All", L("Cash Out", "现金转出"): "Expense",
            L("Cash In", "现金收入"): "Income",
        }
        tk.Label(detail_filter_row, text=L("Type:", "类型："), font=("Arial", 10, "bold"), bg=C['bg'], fg=C['text']).pack(side="left")
        type_filter_var = tk.StringVar(value=all_type_label)
        type_combo = ttk.Combobox(detail_filter_row, textvariable=type_filter_var, values=list(type_display_to_db),
                                  state="readonly", font=("Arial", 10), width=10)
        type_combo.pack(side="left", padx=(4, 12))

        tk.Label(detail_filter_row, text=L("Category:", "类别："), font=("Arial", 10, "bold"), bg=C['bg'], fg=C['text']).pack(side="left")
        cat_filter_var = tk.StringVar(value=all_type_label)
        cat_combo = ttk.Combobox(detail_filter_row, textvariable=cat_filter_var, font=("Arial", 10), width=15)
        cat_combo.pack(side="left", padx=(4, 12))

        tk.Label(detail_filter_row, text=L("Shop:", "店家："), font=("Arial", 10, "bold"), bg=C['bg'], fg=C['text']).pack(side="left")
        shop_filter_var = tk.StringVar(value=all_type_label)
        shop_combo = ttk.Combobox(detail_filter_row, textvariable=shop_filter_var, font=("Arial", 10), width=15)
        shop_combo.pack(side="left", padx=(4, 12))

        def update_filter_dropdowns(*args):
            # Update Category list
            t_type = type_display_to_db.get(type_filter_var.get(), "All")
            cats = [all_type_label]
            try:
                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    if t_type == "All":
                        c.execute("SELECT DISTINCT name FROM cashflow_categories ORDER BY name")
                    else:
                        c.execute("SELECT name FROM cashflow_categories WHERE type=? ORDER BY name", (t_type,))
                    cats += [r['name'] for r in c.fetchall()]
            except: pass
            cat_combo['values'] = cats
            
            # Update Shop list
            shops = [all_type_label] + self._get_all_suppliers()
            shop_combo['values'] = shops

        type_filter_var.trace_add("write", lambda *args: (update_filter_dropdowns(), load()))
        cat_filter_var.trace_add("write", lambda *args: load())
        shop_filter_var.trace_add("write", lambda *args: load())
        
        update_filter_dropdowns()

        quick_bar = tk.Frame(win, bg=C['bg'])
        quick_bar.pack(fill="x", padx=16, pady=(6, 0))
        tk.Label(
            quick_bar, text=L("Quick range:", "快捷范围："),
            font=("Arial", 9, "bold"), bg=C['bg'], fg=C['text_muted']
        ).pack(side="left", padx=(0, 6))

        def set_profit_range(range_name):
            today = _dt.now()
            if range_name == "today":
                start_date = end_date = today
            elif range_name == "last_30":
                start_date = today - timedelta(days=29)
                end_date = today
            elif range_name == "last_month":
                end_date = today.replace(day=1) - timedelta(days=1)
                start_date = end_date.replace(day=1)
            else:
                start_date = today.replace(day=1)
                end_date = today
            from_var.set(start_date.strftime("%Y-%m-%d"))
            to_var.set(end_date.strftime("%Y-%m-%d"))
            load()

        for label, range_name in [
            (L("Today", "今天"), "today"),
            (L("This Month", "本月"), "this_month"),
            (L("Last Month", "上个月"), "last_month"),
            (L("Last 30 Days", "最近30天"), "last_30"),
        ]:
            tk.Button(
                quick_bar, text=label,
                command=lambda name=range_name: set_profit_range(name),
                font=("Arial", 9, "bold"), bg="#E2E8F0", fg="#334155",
                activebackground="#CBD5E1", activeforeground="#1E293B",
                relief="flat", bd=0, padx=11, pady=4, cursor="hand2"
            ).pack(side="left", padx=3)
        tk.Button(
            quick_bar, text=L("✓ Apply Dates", "✓ 应用日期"),
            command=lambda: load(),
            font=("Arial", 9, "bold"), bg=C['primary'], fg="white",
            activebackground="#1D4ED8", activeforeground="white",
            relief="flat", bd=0, padx=12, pady=4, cursor="hand2"
        ).pack(side="right", padx=3)

        profit_search_var = tk.StringVar()
        profit_result_var = tk.StringVar()
        query_host = tk.Frame(win, bg=C['bg'])
        query_host.pack(fill="x", padx=4, pady=(6, 0))

        # ── Stat cards ────────────────────────────────────────
        cards_frame = tk.Frame(win, bg=C['bg'])
        cards_frame.pack(fill="x", padx=16, pady=10)

        def make_stat_card(parent, title, color, is_profit=False):
            f = tk.Frame(parent, bg=color, padx=16, pady=12)
            if is_profit: f._is_profit_card = True
            f.pack(side="left", fill="x", expand=True, padx=5)
            tk.Label(f, text=title, font=("Arial", 10, "bold"),
                     bg=color, fg="#E2E8F0").pack(anchor="w")
            var = tk.StringVar(value="RM 0.00")
            tk.Label(f, textvariable=var, font=("Arial", 18, "bold"),
                     bg=color, fg="white").pack(anchor="w", pady=(4, 0))
            return var

        income_var  = make_stat_card(cards_frame, L("💵 Recorded Cash In", "💵 已记录流入"),  "#16A34A")
        expense_var = make_stat_card(cards_frame, L("🧾 Recorded Cash Out", "🧾 已记录流出"), "#DC2626")
        profit_var  = make_stat_card(cards_frame, L("📈 Recorded Balance", "📈 已记录差额"),     "#2563EB", is_profit=True)

        # ── Transactions table ────────────────────────────────
        tbl_frame = tk.Frame(win, bg=C['bg'])
        tbl_frame.pack(fill="both", expand=True, padx=16, pady=(0, 4))

        cols = ("ID", "Date", "Type", "Source / Category", "Shop / Company", "Payment", "Amount", "Notes")
        tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", height=14)
        col_w = {"ID": 45, "Date": 100, "Type": 80, "Source / Category": 150, "Shop / Company": 150,
                 "Payment": 80, "Amount": 100, "Notes": 220}
        headings = {
            "ID": "ID", "Date": L("Date", "日期"), "Type": L("Type", "类型"),
            "Source / Category": L("Source / Category", "来源／类别"),
            "Shop / Company": L("Shop / Company", "店家／公司"),
            "Payment": L("Payment", "付款"), "Amount": L("Amount", "金额"),
            "Notes": L("Notes", "备注"),
        }
        for col in cols:
            tree.heading(col, text=headings[col])
            tree.column(col, width=col_w.get(col, 100),
                        anchor="center" if col not in ("Source / Category", "Shop / Company", "Notes") else "w")
        sb = tk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # colour rows
        tree.tag_configure("income",  background="#F0FDF4", foreground="#15803D")
        tree.tag_configure("expense", background="#FEF2F2", foreground="#B91C1C")

        # ── Load data ─────────────────────────────────────────
        def load():
            try:
                fd_str = normalize_date_text(from_var.get())
                td_str = normalize_date_text(to_var.get())
                if td_str < fd_str:
                    raise ValueError
                from_var.set(fd_str)
                to_var.set(td_str)
            except ValueError:
                messagebox.showerror(
                    L("Invalid Date Range", "日期范围无效"),
                    L(
                        "Choose valid From and To dates; From cannot be after To.",
                        "请选择有效的开始及结束日期，开始日期不可晚于结束日期。"
                    ),
                    parent=win
                )
                return

            tree.delete(*tree.get_children())
            try:
                t_f = type_display_to_db.get(type_filter_var.get(), "All")
                c_f = cat_filter_var.get()
                s_f = shop_filter_var.get()
                keyword = profit_search_var.get().strip()

                with self.db.get_connection() as conn:
                    c = conn.cursor()
                    query = """SELECT id, date, type, source, supplier, payment_method, amount, notes
                               FROM cashflow
                               WHERE date >= ? AND date <= ?"""
                    params = [fd_str, td_str]

                    if t_f != "All":
                        query += " AND type = ?"
                        params.append(t_f)
                    if c_f != all_type_label:
                        query += " AND source = ?"
                        params.append(c_f)
                    if s_f != all_type_label:
                        query += " AND supplier = ?"
                        params.append(s_f)
                    if keyword:
                        like_keyword = f"%{keyword}%"
                        query += """ AND (
                            CAST(id AS TEXT) LIKE ? OR source LIKE ? OR supplier LIKE ?
                            OR payment_method LIKE ? OR notes LIKE ?
                        )"""
                        params.extend([like_keyword] * 5)

                    query += " ORDER BY date DESC, id DESC"
                    c.execute(query, params)
                    rows = c.fetchall()

                total_inc = 0.0
                total_exp = 0.0
                for r in rows:
                    amt   = float(r['amount'] or 0)
                    rtype = str(r['type'] or '')
                    tag   = "income" if rtype == "Income" else "expense"
                    tree.insert("", "end", iid=str(r['id']), tags=(tag,),
                                values=(r['id'], r['date'],
                                        L("Cash In", "现金收入") if rtype == "Income" else L("Cash Out", "现金转出"),
                                        r['source'] or '',
                                        r['supplier'] or '',
                                        r['payment_method'] or '',
                                        f"RM {amt:,.2f}",
                                        r['notes'] or ''))
                    if rtype == "Income":
                        total_inc += amt
                    else:
                        total_exp += amt

                net = total_inc - total_exp
                income_var.set(f"RM {total_inc:,.2f}")
                expense_var.set(f"RM {total_exp:,.2f}")
                profit_col = "#2563EB" if net >= 0 else "#DC2626"
                profit_var.set(f"RM {net:,.2f}")
                profit_result_var.set(
                    L(f"{len(rows)} transaction(s)", f"{len(rows)} 笔交易")
                )
                # update profit card colour dynamically
                for widget in cards_frame.winfo_children():
                    if hasattr(widget, '_is_profit_card'):
                        widget['bg'] = profit_col
                        for child in widget.winfo_children():
                            child['bg'] = profit_col
            except Exception as e:
                messagebox.showerror(L("Error", "错误"), str(e), parent=win)

        # ── Add transaction dialog ─────────────────────────────
        # ── Transaction dialog (Add/Edit) ──────────────────────
        def tx_dialog(edit_id=None):
            d = tk.Toplevel(win)
            d.title(L("✏️ Edit Transaction", "✏️ 编辑交易") if edit_id else L("➕ Add Transaction", "➕ 新增交易"))
            d.configure(bg=C['bg'])
            d.transient(win)
            d.grab_set()
            d.resizable(False, False)
            self._center_window(d, 420, 480, win)

            tk.Label(d, text=L("Edit Transaction", "编辑交易") if edit_id else L("Add Expense / Income", "新增支出／收入"),
                     font=("Arial", 14, "bold"), bg=C['bg'], fg=C['text']).pack(pady=(18, 12))

            form = tk.Frame(d, bg=C['bg'])
            form.pack(fill="x", padx=30)

            def row(label, widget_fn):
                f = tk.Frame(form, bg=C['bg'])
                f.pack(fill="x", pady=5)
                tk.Label(f, text=label, font=("Arial", 11, "bold"), bg=C['bg'],
                         fg=C['text'], width=16, anchor="w").pack(side="left")
                w = widget_fn(f)
                w.pack(side="left", fill="x", expand=True)
                return w

            # Initial values
            init_vals = {
                'type': "Expense",
                'source': "",
                'amount': "",
                'payment': "Cash",
                'date': datetime.now().strftime("%Y-%m-%d"),
                'notes': "",
                'supplier': ""
            }

            if edit_id:
                try:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("SELECT * FROM cashflow WHERE id=?", (edit_id,))
                        r = c.fetchone()
                        if r:
                            init_vals = {
                                'type': r['type'],
                                'source': r['source'],
                                'amount': str(r['amount']),
                                'payment': r['payment_method'],
                                'date': r['date'],
                                'notes': r['notes'] or "",
                                'supplier': r['supplier'] or ""
                            }
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load transaction: {e}", parent=win)
                    d.destroy()
                    return

            type_var = tk.StringVar(value=init_vals['type'])
            row(L("Type:", "类型："), lambda f: ttk.Combobox(f, textvariable=type_var,
                values=["Expense", "Income"], state="readonly", font=("Arial", 11)))

            source_var = tk.StringVar(value=init_vals['source'])
            source_combo = row(L("Source/Category:", "来源／类别："), lambda f: ttk.Combobox(f, textvariable=source_var,
                font=("Arial", 11)))

            def update_cats(*args):
                t_type = type_var.get()
                cats = []
                try:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        c.execute("SELECT name FROM cashflow_categories WHERE type=? ORDER BY name", (t_type,))
                        cats = [r['name'] for r in c.fetchall()]
                except Exception:
                    pass
                if not cats:
                    if t_type == "Expense":
                        cats = ["Ingredients", "Utilities", "Staff Cost", "Rental", "Repair", "Other"]
                    else:
                        cats = ["Stall Rental", "Kuih Sales"]
                source_combo['values'] = cats
                # If we are adding (not editing), set default
                if not edit_id and cats: source_var.set(cats[0])

            type_var.trace_add("write", update_cats)
            update_cats()
            if edit_id: source_var.set(init_vals['source'])

            amount_var = tk.StringVar(value=init_vals['amount'])
            row(L("Amount (RM):", "金额 (RM)："), lambda f: tk.Entry(f, textvariable=amount_var,
                font=("Arial", 11), bg="white", relief="solid", bd=1))

            pm_var = tk.StringVar(value=init_vals['payment'])
            row(L("Payment Method:", "付款方式："), lambda f: ttk.Combobox(f, textvariable=pm_var,
                values=["Cash", "Bank", "QR"], state="readonly", font=("Arial", 11)))

            date_var = tk.StringVar(value=init_vals['date'])
            row(
                L("Date:", "日期："),
                lambda f: DatePicker(
                    f, textvariable=date_var, font=("Arial", 11),
                    language=self.lang, bg=C['bg']
                )
            )

            notes_var = tk.StringVar(value=init_vals['notes'])
            row(L("Notes:", "备注："), lambda f: tk.Entry(f, textvariable=notes_var,
                font=("Arial", 11), bg="white", relief="solid", bd=1))

            shop_var = tk.StringVar(value=init_vals['supplier'])
            all_shops = self._get_all_suppliers()
            row(L("Shop/Company:", "店家／公司："), lambda f: ttk.Combobox(f, textvariable=shop_var,
                                                 values=all_shops, font=("Arial", 11)))

            def save():
                src   = source_var.get().strip()
                amt_s = amount_var.get().strip()
                date_s = date_var.get().strip()
                if not src:
                    messagebox.showerror(L("Missing", "缺少资料"), L("Source/Category is required.", "必须填写来源／类别。"), parent=d)
                    return
                try:
                    amt = float(amt_s)
                    if amt <= 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror(L("Invalid", "输入无效"), L("Enter a valid positive amount.", "请输入有效的正数金额。"), parent=d)
                    return
                try:
                    date_s = normalize_date_text(date_s)
                    date_var.set(date_s)
                except ValueError:
                    messagebox.showerror(
                        L("Invalid Date", "日期无效"),
                        L("Type or choose a valid date.", "请输入或选择有效日期。"),
                        parent=d
                    )
                    return

                try:
                    with self.db.get_connection() as conn:
                        c = conn.cursor()
                        if edit_id:
                            c.execute("""UPDATE cashflow 
                                         SET date=?, type=?, source=?, amount=?, payment_method=?, notes=?, supplier=?
                                         WHERE id=?""", 
                                      (date_s, type_var.get(), src, amt, pm_var.get(), 
                                       notes_var.get().strip(), shop_var.get().strip(), edit_id))
                        else:
                            c.execute("""INSERT INTO cashflow
                                         (date, datetime, type, source, amount, payment_method, notes, supplier)
                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                                      (date_s, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                       type_var.get(), src, amt, pm_var.get(),
                                       notes_var.get().strip(), shop_var.get().strip()))
                        conn.commit()
                    
                    if not edit_id and pm_var.get() == "Cash":
                        open_cash_drawer()
                        
                    messagebox.showinfo(
                        L("Saved", "已保存"),
                        L("Transaction updated.", "交易已更新。") if edit_id else L("Transaction recorded.", "交易已记录。"),
                        parent=d
                    )
                    d.destroy()
                    update_filter_dropdowns()
                    load()
                except Exception as e:
                    messagebox.showerror("Error", str(e), parent=d)

            RoundedButton(d, text=L("💾 Update Transaction", "💾 更新交易") if edit_id else L("💾 Save Transaction", "💾 保存交易"), command=save,
                          bg=C['success'], fg="white",
                          font=("Arial", 12, "bold"), padx=20, pady=10, radius=8
                          ).pack(pady=18)


        # ── Delete selected transaction ────────────────────────
        def delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning(L("No Selection", "未选择"), L("Select a row to delete.", "请选择要删除的记录。"), parent=win)
                return
            row_id = sel[0]   # iid is the cashflow id
            vals   = tree.item(row_id)['values']
            if not messagebox.askyesno(L("Confirm Delete", "确认删除"),
                    L("Delete this transaction?", "删除这笔交易？") + "\n\n"
                    f"  {vals[2]}  |  {vals[3]}  |  {vals[5]}",
                    parent=win):
                return
            try:
                with self.db.get_connection() as conn:
                    conn.cursor().execute("DELETE FROM cashflow WHERE id=?", (int(row_id),))
                    conn.commit()
                update_filter_dropdowns()
                load()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=win)

        # ── Edit selected handler ─────────────────────────────
        def edit_selected(e=None):
            sel = tree.selection()
            if not sel:
                if e is None: # only show warning if clicked manually
                    messagebox.showwarning(L("No Selection", "未选择"), L("Select a row to edit.", "请选择要编辑的记录。"), parent=win)
                return
            tx_dialog(edit_id=int(sel[0]))

        # This summary is intentionally read-only. Structured transactions must
        # be corrected at their source so stock, staff advances and cash stay linked.

        def clear_profit_filters():
            from_var.set(_dt.now().replace(day=1).strftime("%Y-%m-%d"))
            to_var.set(_dt.now().strftime("%Y-%m-%d"))
            type_filter_var.set(all_type_label)
            cat_filter_var.set(all_type_label)
            shop_filter_var.set(all_type_label)
            load()

        self._build_query_toolbar(
            query_host, profit_search_var, load,
            on_clear=clear_profit_filters,
            on_refresh=load,
            on_export=lambda: self._export_treeview_csv(
                tree, win, "pos_cashflow_summary.csv"
            ),
            result_var=profit_result_var,
            search_label=L("Search transactions:", "搜索交易：")
        )

        # ── Bottom button bar ─────────────────────────────────
        bbar = tk.Frame(win, bg=C['bg'])
        bbar.pack(fill="x", padx=16, pady=(0, 12))

        tk.Label(
            bbar,
            text=L(
                "Read-only: make corrections from Cash Operations, Rent or Sales Records.",
                "只读：请从现金操作、租金或销售记录修正原始交易。"
            ),
            font=("Arial", 9), bg=C['bg'], fg=C['text_sub']
        ).pack(side="left")
        RoundedButton(bbar, text=L("✖ Close", "✖ 关闭"), command=win.destroy,
                      bg=C['neutral'], fg="white", font=("Arial", 11, "bold"),
                      padx=14, pady=8, radius=8).pack(side="right")

        # Automatic refresh happens now via DatePicker callbacks.
        load()

    def _get_available_months(self):
        try:
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT DISTINCT SUBSTR(date, 1, 7) as month FROM cashflow ORDER BY month DESC")
                return [r[0] for r in c.fetchall() if r[0]]
        except Exception:
            return [datetime.now().strftime("%Y-%m")]

    def _export_accountant_sales_exclusive(self, parent=None):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        import calendar
        import tkinter.filedialog as fd

        if not _OPENPYXL_OK:
            messagebox.showerror("Missing Dependency", "Please install openpyxl:\npip install openpyxl", parent=parent)
            return

        months = self._get_available_months()
        if not months:
            months = [datetime.now().strftime("%Y-%m")]

        # Custom dialog for month selection using a dropdown
        dlg = tk.Toplevel(parent or self.root)
        dlg.title("Accountant Export")
        dlg.configure(bg="#F8FAFC")
        dlg.resizable(False, False)
        self._center_window(dlg, 400, 200, parent or self.root)
        dlg.transient(parent or self.root)
        dlg.grab_set()

        def L(en, cn):
            return self.l(en, cn)
        col = self.colors

        tk.Label(dlg, text="📅 " + L("Select month to export:", "选择导出的月份："), 
                 font=("Arial", 11, "bold"), bg="#F8FAFC").pack(pady=(25, 10))
        
        month_var = tk.StringVar(value=months[0])
        combo = ttk.Combobox(dlg, textvariable=month_var, values=months, state="readonly", font=("Arial", 12))
        combo.pack(pady=10, padx=40, fill="x")
        
        result = {"month": None}
        def on_ok():
            result["month"] = month_var.get()
            dlg.destroy()
        
        btn_f = tk.Frame(dlg, bg="#F8FAFC")
        btn_f.pack(pady=20)
        RoundedButton(btn_f, text="✅ OK", command=on_ok, bg=col['success'], fg="white", font=("Arial", 10, "bold"), padx=15, pady=8, radius=8).pack(side="left", padx=10)
        RoundedButton(btn_f, text="❌ Cancel", command=dlg.destroy, bg=col['neutral'], fg="white", font=("Arial", 10, "bold"), padx=15, pady=8, radius=8).pack(side="left")
        
        dlg.wait_window()
        month_str = result["month"]
        if not month_str:
            return

        month_str = str(month_str).strip()
        try:
            month_start = datetime.strptime(month_str, "%Y-%m")
        except ValueError:
            messagebox.showerror("Invalid Month", "Please use format YYYY-MM (example: 2026-03).", parent=parent)
            return

        year = month_start.year
        month = month_start.month
        last_day = calendar.monthrange(year, month)[1]

        with self.db.get_connection() as conn:
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            
            # 1. Get excluded categories
            c.execute("SELECT name FROM product_categories WHERE exclude_export=1")
            excluded_cats = {r['name'] for r in c.fetchall()}

            # 2. Fetch all relevant sales records once to group in Python.
            # We fetch anything that looks like it belongs to the month/year.
            # Wait, sales table has 'datetime' in DD-MM-YYYY HH:MM:SS format.
            # we need to be careful with the overnight shift (6am to 6am).
            # To be safe, we fetch all sales with the correct month-year pattern,
            # and also those from the 1st of the next month (early morning).
            
            p1 = f"%-{month:02d}-{year:04d} %"
            nm = month + 1; ny = year
            if nm > 12: nm = 1; ny += 1
            p2 = f"01-{nm:02d}-{ny:04d} %"
            
            c.execute("SELECT datetime, items, payment FROM sales WHERE (datetime LIKE ? OR datetime LIKE ?) AND amended=0", (p1, p2))
            all_sales = c.fetchall()
            
            # Group by business date
            sales_by_biz = {}
            for s in all_sales:
                bz = biz_date_of_sale(s['datetime'])
                if bz not in sales_by_biz: sales_by_biz[bz] = []
                sales_by_biz[bz].append(s)

            daily_summaries = [] # List of (date, cash, qr, total)
            order_details = []   # List of (date, time, payment, total) — one row per order

            for d in range(1, last_day + 1):
                biz_date = f"{year:04d}-{month:02d}-{d:02d}"
                day_sales = sales_by_biz.get(biz_date, [])
                
                day_cash = 0.0
                day_qr = 0.0
                day_total = 0.0

                for s in day_sales:
                    items_list = json.loads(s['items'])
                    p_method = str(s['payment'] or '').upper()
                    
                    sale_filtered_total = 0.0
                    for entry in items_list:
                        price = entry[1]; qty = entry[2]
                        cat = entry[3] if len(entry) > 3 else ''
                        if cat in excluded_cats: continue
                        sale_filtered_total += float(Decimal(str(price)) * int(qty))

                    if sale_filtered_total > 0:
                        day_total += sale_filtered_total
                        if p_method == 'CASH': day_cash += sale_filtered_total
                        elif p_method == 'QR': day_qr += sale_filtered_total
                        # One row per order
                        order_details.append({
                            'date': biz_date,
                            'time': s['datetime'],
                            'payment': s['payment'] or '',
                            'total': sale_filtered_total,
                        })

                daily_summaries.append({
                    'date': biz_date,
                    'cash': day_cash,
                    'qr': day_qr,
                    'total': day_total
                })

        # 3. Create Excel
        wb = openpyxl.Workbook()
        
        # Styles
        header_fill = PatternFill("solid", start_color="1E3A5F")
        header_font = Font(color="FFFFFF", bold=True, name="Arial", size=11)
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center")
        right = Alignment(horizontal="right", vertical="center")

        def set_header(ws, row, col, value):
            cell = ws.cell(row=row, column=col, value=value)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border
            
        def set_cell(ws, row, col, value, align=left, number=False, bold=False):
            cell = ws.cell(row=row, column=col, value=value)
            cell.alignment = align
            cell.border = border
            cell.font = Font(name="Arial", size=10, bold=bold)
            if number: cell.number_format = '"RM"#,##0.00'

        # Sheet 1: Daily Sales
        ws1 = wb.active
        ws1.title = "Daily Sales"
        
        headers1 = ["Date", "Cash Sales", "QR Sales", "Total Sales"]
        for ci, h in enumerate(headers1, 1):
            set_header(ws1, 1, ci, h)
            ws1.column_dimensions[get_column_letter(ci)].width = 18

        for ri, row in enumerate(daily_summaries, 2):
            set_cell(ws1, ri, 1, row['date'], align=center)
            set_cell(ws1, ri, 2, row['cash'], align=right, number=True)
            set_cell(ws1, ri, 3, row['qr'], align=right, number=True)
            set_cell(ws1, ri, 4, row['total'], align=right, number=True, bold=True)

        # Sheet 2: Order Summary Detail
        ws2 = wb.create_sheet("Order Summary")
        headers2 = ["Date", "Time", "Payment", "Total"]
        for ci, h in enumerate(headers2, 1):
            set_header(ws2, 1, ci, h)
        ws2.column_dimensions["A"].width = 14
        ws2.column_dimensions["B"].width = 22
        ws2.column_dimensions["C"].width = 12
        ws2.column_dimensions["D"].width = 14

        for ri, row in enumerate(order_details, 2):
            set_cell(ws2, ri, 1, row['date'], align=center)
            set_cell(ws2, ri, 2, row['time'], align=center)
            set_cell(ws2, ri, 3, row['payment'], align=center)
            set_cell(ws2, ri, 4, row['total'], align=right, number=True, bold=True)

        # 4. Save
        filename = f"919_restaurant_sales_{month_str}.xlsx"
        file_path = fd.asksaveasfilename(
            parent=parent,
            initialfile=filename,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Accountant Export"
        )
        if not file_path:
            return
            
        try:
            wb.save(file_path)
            messagebox.showinfo("Export Success", f"Saved to:\n{file_path}", parent=parent)
        except Exception as e:
            messagebox.showerror("Export Failed", str(e), parent=parent)

    def _export_monthly_income_expense_excel(self, month_str=None, parent=None, apply_filter=False):
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        import calendar
        import tkinter.filedialog as fd

        if not _OPENPYXL_OK:
            messagebox.showerror("Missing Dependency", "Please install openpyxl:\npip install openpyxl", parent=parent)
            return

        if not month_str:
            default_month = datetime.now().strftime("%Y-%m")
            month_str = simpledialog.askstring(
                "Monthly Export",
                "Enter month (YYYY-MM):",
                initialvalue=default_month,
                parent=parent,
            )
            if not month_str:
                return

        month_str = str(month_str).strip()
        try:
            month_start = datetime.strptime(month_str, "%Y-%m")
        except ValueError:
            messagebox.showerror("Invalid Month", "Please use format YYYY-MM (example: 2026-03).", parent=parent)
            return

        year = month_start.year
        month = month_start.month
        last_day = calendar.monthrange(year, month)[1]

        daily_rows = []
        totals = {
            'total_sales': 0.0,
            'cash_sales': 0.0,
            'qr_sales': 0.0,
            'card_sales': 0.0,
            'other_income': 0.0,
            'total_income': 0.0,
            'expenses': 0.0,
            'net': 0.0,
        }

        with self.db.get_connection() as conn:
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            excluded_cats = set()
            if apply_filter:
                c.execute("SELECT name FROM product_categories WHERE exclude_export=1")
                excluded_cats = {r['name'] for r in c.fetchall()}

            filtered_tx_rows = []
            for d in range(1, last_day + 1):
                biz_date = f"{year:04d}-{month:02d}-{d:02d}"
                
                f_total_sales = 0.0
                f_cash_sales = 0.0
                f_qr_sales = 0.0
                f_card_sales = 0.0
                
                if apply_filter:
                    start_dt, end_dt = business_day_range(biz_date)
                    s_str = start_dt.strftime("%Y-%m-%d %H:%M:%S")
                    e_str = end_dt.strftime("%Y-%m-%d %H:%M:%S")
                    
                    c.execute("SELECT * FROM sales WHERE to_iso(datetime) >= ? AND to_iso(datetime) < ? AND amended=0", (s_str, e_str))
                    day_sales = c.fetchall()
                    
                    for s in day_sales:
                        items_list = json.loads(s['items'])
                        p_method = str(s['payment'] or '').upper()
                        receipt_f_total = 0.0
                        for entry in items_list:
                            price = entry[1]; qty = entry[2]
                            cat  = entry[3] if len(entry) > 3 else ''
                            if cat in excluded_cats: continue
                            line_val = float(Decimal(str(price)) * int(qty))
                            receipt_f_total += line_val
                        
                        if receipt_f_total > 0:
                            f_total_sales += receipt_f_total
                            if p_method == 'CASH': f_cash_sales += receipt_f_total
                            elif p_method == 'QR': f_qr_sales += receipt_f_total
                            elif p_method == 'CARD': f_card_sales += receipt_f_total
                            filtered_tx_rows.append({
                                'date': biz_date, 'type': 'Income', 'source': f"Sales (Order #{s['id']})",
                                'payment_method': s['payment'], 'amount': receipt_f_total,
                                'recorded_by': s['cashier'], 'notes': ""
                            })

                    # Filtered exports must deduct only the refunded lines that
                    # belong to included categories. Previously refunds were not
                    # deducted at all in this mode.
                    c.execute("SELECT * FROM refunds WHERE to_iso(datetime) >= ? AND to_iso(datetime) < ?",
                              (s_str, e_str))
                    for refund in c.fetchall():
                        try:
                            refund_items = json.loads(refund['items'] or '[]')
                        except Exception:
                            refund_items = []
                        filtered_refund = 0.0
                        for entry in refund_items:
                            if not entry or len(entry) < 3:
                                continue
                            cat = entry[3] if len(entry) > 3 else ''
                            if cat in excluded_cats:
                                continue
                            filtered_refund += float(Decimal(str(entry[1])) * int(entry[2]))
                        if filtered_refund <= 0:
                            continue
                        f_total_sales -= filtered_refund
                        refund_pm = str(refund['payment_method'] or 'CASH').upper()
                        if refund_pm == 'CASH':
                            f_cash_sales -= filtered_refund
                        elif refund_pm == 'QR':
                            f_qr_sales -= filtered_refund
                        elif refund_pm == 'CARD':
                            f_card_sales -= filtered_refund
                else:
                    # Use existing fast breakdown for non-filtered
                    f_total_sales, f_cash_sales, f_qr_sales, f_card_sales = compute_sales_payment_breakdown(conn, biz_date)
                    # For non-filtered, we don't build filtered_tx_rows manually; we'll fetch from cashflow later
                
                _pos_cash, other_income, expenses = compute_cash_summary_components(conn, biz_date)
                # Net sales already includes CASH + QR + CARD and deducts refunds.
                # The previous formula used cash only, understating income whenever
                # customers paid by QR/card.
                income_summary = calculate_income_summary(
                    f_total_sales, f_cash_sales, f_qr_sales, f_card_sales,
                    other_income, expenses
                )
                total_income = income_summary['total_income']
                net = income_summary['net_profit']
                
                daily_rows.append((biz_date, f_total_sales, f_cash_sales, f_qr_sales,
                                   f_card_sales, other_income, total_income, expenses, net))
                totals['total_sales'] += f_total_sales
                totals['cash_sales']  += f_cash_sales
                totals['qr_sales']    += f_qr_sales
                totals['card_sales']  += f_card_sales
                totals['other_income'] += other_income
                totals['total_income'] += total_income
                totals['expenses']     += expenses
                totals['net']          += net

            if apply_filter:
                # Add non-sales transactions from cashflow
                c.execute(
                    """SELECT date, type, source, payment_method, amount, recorded_by, notes
                       FROM cashflow
                       WHERE date LIKE ? AND source NOT LIKE 'Sales%'
                       ORDER BY date ASC, datetime ASC, id ASC""",
                    (f"{month_str}-%",)
                )
                for r in c.fetchall():
                    filtered_tx_rows.append(dict(r))
                filtered_tx_rows.sort(key=lambda x: x['date'])
                transaction_rows = filtered_tx_rows
            else:
                # Normal behavior: fetch all from cashflow
                c.execute(
                    """SELECT date, type, source, payment_method, amount, recorded_by, notes
                       FROM cashflow
                       WHERE date LIKE ?
                       ORDER BY date ASC, datetime ASC, id ASC""",
                    (f"{month_str}-%",)
                )
                transaction_rows = c.fetchall()

        wb = openpyxl.Workbook()
        ws_summary = wb.active
        ws_summary.title = "Monthly Summary"
        ws_daily = wb.create_sheet("Daily Breakdown")
        ws_tx = wb.create_sheet("Transactions")

        header_fill = PatternFill("solid", start_color="1E3A5F")
        header_font = Font(color="FFFFFF", bold=True, name="Arial", size=11)
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center")
        right = Alignment(horizontal="right", vertical="center")

        def set_header(ws, row, col, value):
            cell = ws.cell(row=row, column=col, value=value)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border

        def set_cell(ws, row, col, value, align=left, number=False, bold=False):
            cell = ws.cell(row=row, column=col, value=value)
            cell.alignment = align
            cell.border = border
            cell.font = Font(name="Arial", size=10, bold=bold)
            if number:
                cell.number_format = '"RM"#,##0.00'

        ws_summary.merge_cells("A1:B1")
        ws_summary["A1"] = f"Income & Expense Report - {month_str}"
        ws_summary["A1"].font = Font(name="Arial", size=16, bold=True)
        ws_summary.column_dimensions["A"].width = 28
        ws_summary.column_dimensions["B"].width = 22

        set_header(ws_summary, 3, 1, "Item")
        set_header(ws_summary, 3, 2, "Amount (RM)")

        summary_data = [
            ("Net Sales After Refunds (All Payments)", totals['total_sales']),
            ("Cash Sales", totals['cash_sales']),
            ("QR Sales", totals['qr_sales']),
            ("Card Sales", totals['card_sales']),
            ("Other Income", totals['other_income']),
            ("TOTAL INCOME", totals['total_income']),
            ("Operating Expenses (Refunds Excluded)", totals['expenses']),
            ("NET PROFIT (Total Income - Expenses)", totals['net']),
        ]
        srow = 4
        for label, value in summary_data:
            set_cell(ws_summary, srow, 1, label)
            set_cell(ws_summary, srow, 2, float(value), align=right, number=True, bold=True)
            srow += 1

        daily_headers = [
            "Business Date",
            "Daily Sales (RM)",
            "Cash Sales (RM)",
            "QR Sales (RM)",
            "Card Sales (RM)",
            "Other Income (RM)",
            "Total Income (RM)",
            "Expenses (RM)",
            "Net Profit (RM)",
        ]
        for i, h in enumerate(daily_headers, 1):
            set_header(ws_daily, 1, i, h)
            ws_daily.column_dimensions[get_column_letter(i)].width = 18

        drow = 2
        for biz_date, total_sales, cash_sales, qr_sales, card_sales, other_income, total_income, expenses, net in daily_rows:
            set_cell(ws_daily, drow, 1, biz_date, align=center)
            set_cell(ws_daily, drow, 2, float(total_sales), align=right, number=True)
            set_cell(ws_daily, drow, 3, float(cash_sales), align=right, number=True)
            set_cell(ws_daily, drow, 4, float(qr_sales), align=right, number=True)
            set_cell(ws_daily, drow, 5, float(card_sales), align=right, number=True)
            set_cell(ws_daily, drow, 6, float(other_income), align=right, number=True)
            set_cell(ws_daily, drow, 7, float(total_income), align=right, number=True)
            set_cell(ws_daily, drow, 8, float(expenses), align=right, number=True)
            set_cell(ws_daily, drow, 9, float(net), align=right, number=True)
            drow += 1

        set_cell(ws_daily, drow, 1, "TOTAL", align=center, bold=True)
        set_cell(ws_daily, drow, 2, float(totals['total_sales']), align=right, number=True, bold=True)
        set_cell(ws_daily, drow, 3, float(totals['cash_sales']), align=right, number=True, bold=True)
        set_cell(ws_daily, drow, 4, float(totals['qr_sales']), align=right, number=True, bold=True)
        set_cell(ws_daily, drow, 5, float(totals['card_sales']), align=right, number=True, bold=True)
        set_cell(ws_daily, drow, 6, float(totals['other_income']), align=right, number=True, bold=True)
        set_cell(ws_daily, drow, 7, float(totals['total_income']), align=right, number=True, bold=True)
        set_cell(ws_daily, drow, 8, float(totals['expenses']), align=right, number=True, bold=True)
        set_cell(ws_daily, drow, 9, float(totals['net']), align=right, number=True, bold=True)

        # Right-side monthly quick summary for accountant view.
        right_col_title = 10  # Column J
        right_col_value = 11  # Column K
        ws_daily.column_dimensions[get_column_letter(right_col_title)].width = 26
        ws_daily.column_dimensions[get_column_letter(right_col_value)].width = 18
        set_header(ws_daily, 1, right_col_title, "Monthly Summary")
        set_header(ws_daily, 1, right_col_value, "Amount (RM)")

        quick_summary = [
            ("Total Sales", totals['total_sales']),
            ("Cash Sales", totals['cash_sales']),
            ("QR Sales", totals['qr_sales']),
            ("Card Sales", totals['card_sales']),
            ("Other Income", totals['other_income']),
            ("TOTAL INCOME", totals['total_income']),
            ("Operating Expenses", totals['expenses']),
            ("NET PROFIT", totals['net']),
        ]
        qrow = 2
        for label, value in quick_summary:
            set_cell(ws_daily, qrow, right_col_title, label, bold=True)
            set_cell(ws_daily, qrow, right_col_value, float(value), align=right, number=True, bold=True)
            qrow += 1

        tx_headers = ["Date", "Type", "Source", "Payment", "Amount (RM)", "Recorded By", "Notes"]
        for i, h in enumerate(tx_headers, 1):
            set_header(ws_tx, 1, i, h)
        tx_widths = {1: 14, 2: 12, 3: 24, 4: 12, 5: 14, 6: 16, 7: 36}
        for col_idx, width in tx_widths.items():
            ws_tx.column_dimensions[get_column_letter(col_idx)].width = width

        trow = 2
        for r in transaction_rows:
            set_cell(ws_tx, trow, 1, r['date'], align=center)
            set_cell(ws_tx, trow, 2, r['type'], align=center)
            set_cell(ws_tx, trow, 3, r['source'] or "")
            set_cell(ws_tx, trow, 4, r['payment_method'] or "", align=center)
            set_cell(ws_tx, trow, 5, float(r['amount'] or 0), align=right, number=True)
            set_cell(ws_tx, trow, 6, r['recorded_by'] or "")
            set_cell(ws_tx, trow, 7, r['notes'] or "")
            trow += 1

        default_name = f"Monthly_Income_Expense_{month_str}.xlsx"
        file_path = fd.asksaveasfilename(
            parent=parent,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Monthly Income/Expense Report",
        )
        if not file_path:
            return

        wb.save(file_path)
        messagebox.showinfo("Export Success", f"Saved to:\n{file_path}", parent=parent)

    # ================= REFUNDS =================
    def _quick_full_refund(self, sale, parent_win, refresh_callback):
        L = self.l
        sale_id = int(sale['id'])
        sale_total = float(sale['total'] or 0)
        
        if not messagebox.askyesno(L("Confirm Refund", "确认退款"), 
                                   L(f"Full refund for Order #{sale_id}?\nAmount: RM {sale_total:.2f}",
                                     f"确认对订单 #{sale_id} 进行全额退款？\n金额：RM {sale_total:.2f}"),
                                   parent=parent_win):
            return
        return_to_stock = messagebox.askyesnocancel(
            L("Return to Inventory?", "退回库存？"),
            L("Were the refunded items physically returned and sellable?\n\n"
              "Yes = add them back to stock\nNo = refund payment only",
              "退款商品是否已实际退回并可再次销售？\n\n是 = 加回库存\n否 = 只退款，不改库存"),
            parent=parent_win
        )
        if return_to_stock is None:
            return

        def _do_refund(staff):
            now = datetime.now()
            now_str = now.strftime("%d-%m-%Y %H:%M:%S")
            biz_date = biz_date_of_sale(sale['datetime'])
            payment_method = str(sale['payment'] or 'CASH').strip().upper()
            
            with self.db.get_connection() as conn:
                c = conn.cursor()
                
                # 1. Record in refunds table
                c.execute("INSERT INTO refunds (sale_id, datetime, items, refund_amount, reason, processed_by, payment_method) VALUES (?,?,?,?,?,?,?)",
                          (sale_id, now_str, sale['items'], sale_total, "Full Refund (Quick)", staff['name'], payment_method))
                refund_id = c.lastrowid
                if return_to_stock:
                    refund_items_list = json.loads(sale['items'] or '[]')
                    self._adjust_tracked_stock_for_items(
                        conn, refund_items_list, +1, source_type="REFUND",
                        reference_id=refund_id, recorded_by=staff['name'],
                        notes=f"Full refund of Sale #{sale_id}"
                    )
                
                # 2. Update daily_summary
                pm_col = self._payment_summary_column(payment_method)
                c.execute("INSERT OR IGNORE INTO daily_summary (date, total_sales, transaction_count, cash_total, qr_total, card_total, total_discount) VALUES (?,0,0,0,0,0,0)", (biz_date,))
                c.execute("UPDATE daily_summary SET total_sales = total_sales - ?, transaction_count = CASE WHEN transaction_count > 0 THEN transaction_count - 1 ELSE 0 END WHERE date=?", (sale_total, biz_date))
                if pm_col:
                    c.execute(f"UPDATE daily_summary SET {pm_col} = {pm_col} - ? WHERE date=?", (sale_total, biz_date))
                
                # 3. Record in cashflow
                self._record_cashflow(conn, "Expense", "Refund", refund_id, sale_total, payment_method, staff['name'], f"Full Refund: Sale #{sale_id}", dt=now)
                
                # 4. Audit Log
                c.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)",
                          (now_str, "REFUND", f"Full refund for Sale #{sale_id} RM{sale_total:.2f} | Restocked: {'YES' if return_to_stock else 'NO'}", staff['name']))
            
            messagebox.showinfo(L("Refunded", "已退款"), L(f"Order #{sale_id} has been fully refunded.", f"订单 #{sale_id} 已全额退款。"), parent=parent_win)
            refresh_callback()

        self._staff_login_popup(callback=_do_refund, parent=parent_win)

    def show_refunds(self):
        L = self.l
        win = tk.Toplevel(self.root)
        win.title(L("↩️ Refund Management", "↩️ 退款管理"))
        win.geometry("1300x900")
        win.configure(bg=self.colors['bg'])
        win.transient(self.root)

        C = self.colors
        hdr = tk.Frame(win, bg=C['danger'], height=52)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text=L("↩️ Refund Management", "↩️ 退款管理"), font=("Arial", 18, "bold"),
                 bg=C['danger'], fg='white').pack(side="left", padx=20, pady=10)

        period_f = tk.Frame(win, bg=C['bg'])
        period_f.pack(fill="x", padx=16, pady=(8, 2))
        tk.Label(
            period_f, text=L("Date range:", "日期范围："),
            font=("Arial", 10, "bold"), bg=C['bg'], fg=C['text']
        ).pack(side="left", padx=(0, 5))
        refund_period_var = tk.StringVar(value="today")
        for label, value in [
            (L("Today", "今天"), "today"),
            (L("Previous Day", "前一天"), "yesterday"),
            (L("Last 7 Days", "最近7天"), "last7"),
            (L("All", "全部"), "all"),
            (L("Custom", "自定义"), "custom"),
        ]:
            tk.Radiobutton(
                period_f, text=label, variable=refund_period_var, value=value,
                command=lambda: refresh_refund_lists(),
                font=("Arial", 9), bg=C['bg'], fg=C['text']
            ).pack(side="left", padx=3)

        refund_from_var = tk.StringVar(value=today_biz_date())
        refund_to_var = tk.StringVar(value=today_biz_date())
        tk.Label(period_f, text=L("From:", "开始："), font=("Arial", 9, "bold"), bg=C['bg']).pack(side="left", padx=(12, 3))
        DatePicker(
            period_f, textvariable=refund_from_var, language=self.lang,
            callback=lambda: (refund_period_var.set("custom"), refresh_refund_lists())
        ).pack(side="left")
        tk.Label(period_f, text=L("To:", "结束："), font=("Arial", 9, "bold"), bg=C['bg']).pack(side="left", padx=(8, 3))
        DatePicker(
            period_f, textvariable=refund_to_var, language=self.lang,
            callback=lambda: (refund_period_var.set("custom"), refresh_refund_lists())
        ).pack(side="left")

        search_var = tk.StringVar()
        refund_result_var = tk.StringVar()
        result_counts = {"sales": 0, "refunds": 0}
        query_host = tk.Frame(win, bg=C['bg'])
        query_host.pack(fill="x", padx=4, pady=(2, 0))

        tk.Label(win, text=L("Select Order to Refund:", "选择要退款的订单："), font=("Arial", 11, "bold"), bg=C['bg']).pack(anchor="w", padx=16, pady=(6,0))
        sales_frame = tk.Frame(win, bg=C['bg'])
        sales_frame.pack(fill="both", expand=True, padx=16, pady=4)
        scols = ("ID", "DateTime", "Total", "Payment", "Cashier", "Items")
        stree = ttk.Treeview(sales_frame, columns=scols, show="headings", height=8)
        sw = {"ID":60,"DateTime":160,"Total":100,"Payment":90,"Cashier":120,"Items":80}
        sheadings = {
            "ID": "ID", "DateTime": L("DateTime", "日期时间"),
            "Total": L("Total", "总额"), "Payment": L("Payment", "付款"),
            "Cashier": L("Cashier", "收银员"), "Items": L("Items", "商品数"),
        }
        for col in scols:
            stree.heading(col, text=sheadings[col]); stree.column(col, width=sw.get(col,100), anchor="center")
        ssb = tk.Scrollbar(sales_frame, orient="vertical", command=stree.yview)
        stree.configure(yscrollcommand=ssb.set)
        stree.pack(side="left", fill="both", expand=True)
        ssb.pack(side="right", fill="y")

        def current_refund_range():
            period = refund_period_var.get()
            today = datetime.strptime(today_biz_date(), "%Y-%m-%d")
            if period == "all":
                return None, None
            if period == "today":
                start = end = today
            elif period == "yesterday":
                start = end = today - timedelta(days=1)
            elif period == "last7":
                start, end = today - timedelta(days=6), today
            else:
                start = datetime.strptime(normalize_date_text(refund_from_var.get()), "%Y-%m-%d")
                end = datetime.strptime(normalize_date_text(refund_to_var.get()), "%Y-%m-%d")
                if end < start:
                    raise ValueError
                refund_from_var.set(start.strftime("%Y-%m-%d"))
                refund_to_var.set(end.strftime("%Y-%m-%d"))
            return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")

        def update_refund_result_count():
            refund_result_var.set(L(
                f"{result_counts['sales']} order(s) · {result_counts['refunds']} refund(s)",
                f"{result_counts['sales']} 笔订单 · {result_counts['refunds']} 笔退款"
            ))

        def load_sales():
            stree.delete(*stree.get_children())
            q = search_var.get().strip().casefold()
            try:
                start_biz, end_biz = current_refund_range()
            except ValueError:
                messagebox.showerror(
                    L("Invalid Date Range", "日期范围无效"),
                    L("Choose valid dates; From cannot be after To.", "请选择有效日期，开始日期不可晚于结束日期。"),
                    parent=win
                )
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM sales WHERE amended=0 ORDER BY id DESC LIMIT 2000")
                rows = []
                for row in c.fetchall():
                    row_biz = biz_date_of_sale(str(row['datetime'] or ''))
                    if start_biz and not (start_biz <= row_biz <= end_biz):
                        continue
                    if q:
                        searchable = " ".join((
                            str(row['id'] or ''), str(row['datetime'] or ''),
                            str(row['payment'] or ''), str(row['cashier'] or ''),
                            str(row['table_code'] or '') if 'table_code' in row.keys() else '',
                            str(row['items'] or ''),
                        )).casefold()
                        if q not in searchable:
                            continue
                    rows.append(row)
            for r in rows:
                items_list = json.loads(r['items'])
                item_count = sum(i[2] for i in items_list)
                stree.insert("", "end", values=(r['id'], r['datetime'], f"RM {r['total']:.2f}", r['payment'], r['cashier'] or "-", L(f"{item_count} items", f"{item_count} 件")))
            result_counts["sales"] = len(rows)
            update_refund_result_count()

        refund_heading = tk.Frame(win, bg=C['bg'])
        refund_heading.pack(fill="x", padx=16, pady=(8, 0))
        tk.Label(refund_heading, text=L("Refund Records:", "退款记录："), font=("Arial", 11, "bold"), bg=C['bg']).pack(side="left")
        RoundedButton(
            refund_heading, text=L("⬇ Export Refunds", "⬇ 导出退款"),
            command=lambda: self._export_treeview_csv(rtree, win, "refund_records.csv"),
            bg=C['success'], fg="white", font=("Arial", 9, "bold"),
            padx=9, pady=4, radius=7
        ).pack(side="right")
        ref_frame = tk.Frame(win, bg=C['bg'])
        ref_frame.pack(fill="both", expand=True, padx=16, pady=4)
        rcols = ("Refund ID", "Order ID", "Time", "Amount", "Reason", "Staff", "Payment")
        rtree = ttk.Treeview(ref_frame, columns=rcols, show="headings", height=5)
        rw = {"Refund ID":60,"Order ID":70,"Time":150,"Amount":90,"Reason":200,"Staff":100,"Payment":90}
        rheadings = {
            "Refund ID": L("Refund ID", "退款编号"), "Order ID": L("Order ID", "订单号"),
            "Time": L("Time", "时间"), "Amount": L("Amount", "金额"),
            "Reason": L("Reason", "原因"), "Staff": L("Staff", "员工"),
            "Payment": L("Payment", "付款"),
        }
        for col in rcols:
            rtree.heading(col, text=rheadings[col]); rtree.column(col, width=rw.get(col,100), anchor="center")
        rsb = tk.Scrollbar(ref_frame, orient="vertical", command=rtree.yview)
        rtree.configure(yscrollcommand=rsb.set)
        rtree.pack(side="left", fill="both", expand=True)
        rsb.pack(side="right", fill="y")

        def load_refunds_list():
            rtree.delete(*rtree.get_children())
            q = search_var.get().strip().casefold()
            try:
                start_biz, end_biz = current_refund_range()
            except ValueError:
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM refunds ORDER BY id DESC LIMIT 2000")
                rows = []
                for row in c.fetchall():
                    row_biz = biz_date_of_sale(str(row['datetime'] or ''))
                    if start_biz and not (start_biz <= row_biz <= end_biz):
                        continue
                    if q:
                        searchable = " ".join((
                            str(row['id'] or ''), str(row['sale_id'] or ''),
                            str(row['datetime'] or ''), str(row['reason'] or ''),
                            str(row['processed_by'] or ''), str(row['payment_method'] or ''),
                            str(row['items'] or ''),
                        )).casefold()
                        if q not in searchable:
                            continue
                    rows.append(row)
                for r in rows:
                    rtree.insert("", "end", values=(r['id'], r['sale_id'], r['datetime'], f"RM {r['refund_amount']:.2f}", r['reason'] or "-", r['processed_by'] or "-", r['payment_method'] or "-"))
            result_counts["refunds"] = len(rows)
            update_refund_result_count()

        def refresh_refund_lists():
            load_sales()
            load_refunds_list()

        def do_refund():
            sel = stree.selection()
            if not sel:
                messagebox.showwarning(L("No Selection", "未选择"), L("Please select an order to refund", "请选择要退款的订单。"), parent=win); return
            sale_id = stree.item(sel[0])['values'][0]
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT * FROM sales WHERE id=?", (sale_id,))
                sale = c.fetchone()
            items_list = json.loads(sale['items'])
            rwin = tk.Toplevel(win)
            rwin.title(L(f"Refund — Order #{sale_id}", f"退款 — 订单 #{sale_id}"))
            rwin.geometry("720x620")
            rwin.configure(bg=C['bg'])
            rwin.transient(win); rwin.grab_set()
            tk.Label(rwin, text=L(f"↩️ Refund — Order #{sale_id}", f"↩️ 退款 — 订单 #{sale_id}"), font=("Arial", 15, "bold"), bg=C['bg']).pack(pady=(16,4))
            tk.Label(rwin, text=L(f"Original Total: RM {sale['total']:.2f}  |  Payment: {sale['payment']}", f"原订单总额：RM {sale['total']:.2f}  |  付款：{sale['payment']}"), font=("Arial", 11), bg=C['bg'], fg=C['text_sub']).pack()
            tk.Label(rwin, text=L("Select items to refund:", "选择要退款的商品："), font=("Arial", 11, "bold"), bg=C['bg']).pack(anchor="w", padx=20, pady=(12,2))
            item_vars = []
            item_frame = tk.Frame(rwin, bg=C['panel'], highlightbackground=C['border'], highlightthickness=1)
            item_frame.pack(fill="x", padx=20, pady=4)
            for name, price, qty, cat in items_list:
                var = tk.IntVar(value=1)
                qty_var = tk.StringVar(value=str(qty))
                row = tk.Frame(item_frame, bg=C['panel'])
                row.pack(fill="x", padx=10, pady=3)
                tk.Checkbutton(row, variable=var, bg=C['panel']).pack(side="left")
                tk.Label(row, text=f"{name}", font=("Arial", 11), bg=C['panel'], width=22, anchor="w").pack(side="left")
                tk.Label(row, text="x", font=("Arial", 11), bg=C['panel']).pack(side="left")
                tk.Spinbox(row, from_=1, to=qty, textvariable=qty_var, font=("Arial", 11), width=4, justify="center").pack(side="left", padx=4)
                tk.Label(row, text=L(f"  RM {price:.2f} each", f"  每件 RM {price:.2f}"), font=("Arial", 11), bg=C['panel'], fg=C['text_sub']).pack(side="left")
                item_vars.append((var, qty_var, name, price, qty))
            reason_f = tk.Frame(rwin, bg=C['bg'])
            reason_f.pack(fill="x", padx=20, pady=8)
            tk.Label(reason_f, text=L("Reason:", "原因："), font=("Arial", 11, "bold"), bg=C['bg']).pack(side="left")
            reason_var = tk.StringVar()
            reason_cb = ttk.Combobox(reason_f, textvariable=reason_var, values=["Customer Unsatisfied", "Wrong Order", "Food Quality Issue", "Duplicate Charge", "Customer Cancelled", "Other"], font=("Arial", 11), width=20)
            reason_cb.set("Customer Unsatisfied")
            reason_cb.pack(side="left", padx=8)
            return_to_stock_var = tk.BooleanVar(value=True)
            tk.Checkbutton(reason_f, text=L("Return sellable items to stock", "可再销售的商品退回库存"),
                           variable=return_to_stock_var, font=("Arial", 10, "bold"),
                           bg=C['bg'], fg=C['text']).pack(side="left", padx=14)
            total_lbl = tk.Label(rwin, text=L("Refund Amount: RM 0.00", "退款金额：RM 0.00"), font=("Arial", 14, "bold"), bg=C['bg'], fg=C['danger'])
            total_lbl.pack(pady=6)
            def update_total(*args):
                total = sum(price * min(int(qty_var.get()) if qty_var.get().isdigit() else 0, max_qty) for var, qty_var, name, price, max_qty in item_vars if var.get())
                total_lbl.config(text=L(f"Refund Amount: RM {total:.2f}", f"退款金额：RM {total:.2f}"))
            for var, qty_var, name, price, max_qty in item_vars:
                var.trace('w', update_total); qty_var.trace('w', update_total)
            update_total()
            def confirm_refund():
                refund_items = []
                refund_total = Decimal('0.00')
                for var, qty_var, name, price, max_qty in item_vars:
                    if var.get():
                        q = min(int(qty_var.get()) if qty_var.get().isdigit() else 0, max_qty)
                        if q > 0:
                            refund_items.append((name, price, q))
                            refund_total += Decimal(str(price)) * q
                if not refund_items:
                    messagebox.showerror(L("Error", "错误"), L("Please select at least one item", "请至少选择一件商品。"), parent=rwin); return
                if not messagebox.askyesno(L("Confirm Refund", "确认退款"), L(f"Refund: RM {float(refund_total):.2f}\nProceed?", f"退款：RM {float(refund_total):.2f}\n是否继续？"), parent=rwin): return
                def process_refund(staff):
                    now_str = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    items_json = json.dumps(refund_items)
                    with self.db.get_connection() as conn:
                        c = conn.cursor()

                        sold_qty_by_name = {}
                        for s_name, _s_price, s_qty, _s_cat in items_list:
                            sold_qty_by_name[s_name] = sold_qty_by_name.get(s_name, 0) + int(s_qty or 0)

                        refunded_qty_by_name = {}
                        c.execute("SELECT items FROM refunds WHERE sale_id=?", (sale_id,))
                        for rr in c.fetchall():
                            try:
                                refunded_items = json.loads(rr['items'] or '[]')
                            except Exception:
                                refunded_items = []
                            for r_item in refunded_items:
                                if not r_item or len(r_item) < 3:
                                    continue
                                r_name = str(r_item[0]).strip()
                                try:
                                    r_qty = int(r_item[2] or 0)
                                except Exception:
                                    r_qty = 0
                                if r_name and r_qty > 0:
                                    refunded_qty_by_name[r_name] = refunded_qty_by_name.get(r_name, 0) + r_qty

                        requested_qty_by_name = {}
                        for r_name, _r_price, r_qty in refund_items:
                            requested_qty_by_name[r_name] = requested_qty_by_name.get(r_name, 0) + int(r_qty or 0)

                        over_refund_lines = []
                        for r_name, req_qty in requested_qty_by_name.items():
                            sold_qty = int(sold_qty_by_name.get(r_name, 0) or 0)
                            already_refunded_qty = int(refunded_qty_by_name.get(r_name, 0) or 0)
                            available_qty = max(0, sold_qty - already_refunded_qty)
                            if req_qty > available_qty:
                                over_refund_lines.append(
                                    f"{r_name}: request {req_qty}, available {available_qty} (sold {sold_qty}, refunded {already_refunded_qty})"
                                )

                        if over_refund_lines:
                            msg = L("Cannot refund more than sold quantity.", "退款数量不能超过已售数量。") + "\n\n" + "\n".join(over_refund_lines)
                            messagebox.showerror(L("Invalid Refund Quantity", "退款数量无效"), msg, parent=rwin)
                            return

                        c.execute("INSERT INTO refunds (sale_id, datetime, items, refund_amount, reason, processed_by, payment_method) VALUES (?,?,?,?,?,?,?)", (sale_id, now_str, items_json, float(refund_total), reason_var.get(), staff['name'], sale['payment']))
                        refund_id = c.lastrowid
                        if return_to_stock_var.get():
                            self._adjust_tracked_stock_for_items(
                                conn, refund_items, +1, source_type="REFUND",
                                reference_id=refund_id, recorded_by=staff['name'],
                                notes=f"Partial refund of Sale #{sale_id}: {reason_var.get()}"
                            )
                        biz_date = biz_date_of_sale(sale['datetime'])
                        sale_total = float(sale['total'] or 0)
                        payment_method = str(sale['payment'] or 'CASH').strip().upper()
                        pm_col = self._payment_summary_column(payment_method)

                        # Keep daily summary in sync for legacy/admin reports.
                        c.execute(
                            """INSERT OR IGNORE INTO daily_summary
                               (date, total_sales, transaction_count, cash_total, qr_total, card_total, total_discount)
                               VALUES (?,0,0,0,0,0,0)""",
                            (biz_date,)
                        )

                        c.execute("SELECT COALESCE(SUM(refund_amount),0) FROM refunds WHERE sale_id=? AND id<>?", (sale_id, refund_id))
                        prev_refunded = float((c.fetchone() or [0])[0] or 0)
                        new_refunded = prev_refunded + float(refund_total)
                        crossed_full_refund = (prev_refunded + 0.005 < sale_total) and (new_refunded + 0.005 >= sale_total)

                        if crossed_full_refund:
                            c.execute(
                                """UPDATE daily_summary
                                   SET total_sales = total_sales - ?,
                                       transaction_count = CASE WHEN transaction_count > 0 THEN transaction_count - 1 ELSE 0 END
                                   WHERE date=?""",
                                (float(refund_total), biz_date)
                            )
                        else:
                            c.execute(
                                "UPDATE daily_summary SET total_sales = total_sales - ? WHERE date=?",
                                (float(refund_total), biz_date)
                            )

                        if pm_col:
                            c.execute(
                                f"UPDATE daily_summary SET {pm_col} = {pm_col} - ? WHERE date=?",
                                (float(refund_total), biz_date)
                            )

                        # Cashflow is source-of-truth for new cash/sales dashboards.
                        self._record_cashflow(
                            conn,
                            "Expense",
                            "Refund",
                            refund_id,
                            float(refund_total),
                            payment_method,
                            staff['name'],
                            notes=f"Refund Order#{sale_id}: {reason_var.get()}"
                        )

                        c.execute("INSERT INTO audit_log (timestamp, action, details, user) VALUES (?,?,?,?)", (now_str, 'REFUND', f"Refund Order#{sale_id} RM{float(refund_total):.2f} | Restocked: {'YES' if return_to_stock_var.get() else 'NO'}", staff['name']))
                        conn.commit()
                    messagebox.showinfo(L("Refund Success", "退款成功"), L(f"✅ Refund RM {float(refund_total):.2f} done\nBy: {staff['name']}", f"✅ 已退款 RM {float(refund_total):.2f}\n处理人：{staff['name']}"), parent=rwin)
                    rwin.destroy(); load_sales(); load_refunds_list(); self.update_today_stats(); self.load_items()
                self._staff_login_popup(callback=process_refund)
            RoundedButton(rwin, text=L("✅ Confirm Refund", "✅ 确认退款"), command=confirm_refund, bg=C['danger'], fg="white", font=("Arial", 13, "bold"), padx=30, pady=12, radius=8).pack(pady=10)
            RoundedButton(rwin, text=L("✖ Cancel", "✖ 取消"), command=rwin.destroy, bg=C['neutral'], fg="white", font=("Arial", 11), padx=20, pady=8, radius=8).pack()

        def clear_refund_filters():
            today_text = today_biz_date()
            refund_period_var.set("today")
            refund_from_var.set(today_text)
            refund_to_var.set(today_text)
            refresh_refund_lists()

        _query_bar, search_entry = self._build_query_toolbar(
            query_host, search_var, refresh_refund_lists,
            on_clear=clear_refund_filters,
            on_refresh=refresh_refund_lists,
            on_export=lambda: self._export_treeview_csv(
                stree, win, "refundable_orders.csv"
            ),
            result_var=refund_result_var,
            search_label=L("Order / item / staff:", "订单／商品／员工：")
        )
        search_entry.focus_set()

        btn_f = tk.Frame(win, bg=C['bg'])
        btn_f.pack(fill="x", side="bottom", padx=16, pady=8)
        RoundedButton(btn_f, text=L("↩️ Process Refund", "↩️ 处理退款"), command=do_refund, bg=C['danger'], fg="white", font=("Arial", 12, "bold"), padx=20, pady=10, radius=8).pack(side="left", padx=4)
        refresh_refund_lists()


    # ================= HELPERS =================
    def cjk_truncate(self, text, max_width):
        result = ""
        width  = 0
        for ch in text:
            w = wcswidth(ch)
            if w < 0: w = 1
            if width + w > max_width: break
            result += ch
            width  += w
        return result

    def cjk_pad(self, text, max_width, align="left"):
        current_width = wcswidth(text)
        if current_width < 0: current_width = len(text)
        padding = max_width - current_width
        if padding <= 0: return text
        if align == "right": return " " * padding + text
        return text + " " * padding


    def _setup_combobox_filter(self, combo, full_list):
        """Enable search-as-you-type filtering for a ttk.Combobox."""
        def on_key(event):
            # Let the key event process first
            val = combo.get()
            if val == "":
                combo['values'] = full_list
            else:
                filtered = [x for x in full_list if val.lower() in x.lower()]
                combo['values'] = filtered
            # Open dropdown
            try:
                combo.event_generate('<Down>')
            except: pass

        combo.bind('<KeyRelease>', on_key)

# ================= MAIN =================
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    startup_screen = StartupLoadingScreen(root)
    try:
        app = POS(root, startup_progress=startup_screen.update_progress)
    except Exception as exc:
        startup_screen.close()
        root.deiconify()
        messagebox.showerror(
            "Startup Error / 启动错误",
            f"Unable to start 919 POS:\n\n{exc}",
            parent=root
        )
        root.destroy()
        raise
    startup_screen.close()
    root.deiconify()
    try:
        root.state('zoomed')
    except tk.TclError:
        pass
    root.mainloop()
