import sys
import time
import re
import traceback
import threading
import queue
import contextlib
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
from html.parser import HTMLParser
from html import unescape
from urllib import request as urlrequest, error as urlerror, parse as urlparse

try:
    import keyboard
except Exception:
    keyboard = None

try:
    import pythoncom
except Exception:
    pythoncom = None

try:
    import win32com.client
    import win32clipboard
    import win32con
    import win32gui
    import win32api
    import win32process
except Exception as exc:
    win32com = None
    win32clipboard = None
    win32con = None
    win32gui = None
    win32api = None
    win32process = None
    WIN32_IMPORT_ERROR = exc
else:
    WIN32_IMPORT_ERROR = None

_thread_state = threading.local()


# ============================================================
# COM / Windows helpers
# ============================================================

def ensure_com_initialized():
    if pythoncom is None:
        return
    if getattr(_thread_state, "com_initialized", False):
        return
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    _thread_state.com_initialized = True


def safe_call(default=None):
    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception:
                return default
        return wrapper
    return decorator


@safe_call(None)
def get_foreground_hwnd():
    if win32gui is None:
        return None
    return win32gui.GetForegroundWindow()




class ReadableHTMLTextExtractor(HTMLParser):
  

    SKIP_TAGS = {"script", "style", "noscript", "svg", "canvas", "iframe", "form", "input", "button", "select", "option"}
    NON_CONTENT_TAGS = {"nav", "header", "footer", "aside"}
    BLOCK_TAGS = {
        "p", "div", "section", "article", "main", "br", "li", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6",
        "blockquote", "tr", "td", "th", "pre"
    }

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.skip_depth = 0
        self.body_seen = False
        self.in_body = False

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "body":
            self.body_seen = True
            self.in_body = True
        if tag in self.SKIP_TAGS or tag in self.NON_CONTENT_TAGS:
            self.skip_depth += 1
            return
        if self.skip_depth == 0 and tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "body":
            self.in_body = False
        if tag in self.SKIP_TAGS or tag in self.NON_CONTENT_TAGS:
            if self.skip_depth > 0:
                self.skip_depth -= 1
            return
        if self.skip_depth == 0 and tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data):
        if self.skip_depth > 0:
            return
        # Если body найден, берем только body; если body не найден — берем весь HTML.
        if self.body_seen and not self.in_body:
            return
        data = unescape(data or "").strip()
        if data:
            self.parts.append(data)
            self.parts.append(" ")

    def get_text(self):
        return "".join(self.parts)




class ScreenReaderApp:


    MAX_READ_CHARS = 12000

    BROWSER_CLASSES = {
        "Chrome_WidgetWin_1",        # Chrome, Edge, Яндекс, Brave, Opera и др.
        "MozillaWindowClass",        # Firefox
        "ApplicationFrameWindow",    # часть системных окон Edge / WebView
        "YandexBrowserWidgetWin",
    }

    BROWSER_TITLE_WORDS = (
        "chrome",
        "google chrome",
        "edge",
        "microsoft edge",
        "firefox",
        "mozilla firefox",
        "яндекс",
        "yandex",
        "brave",
        "opera",
        "browser",
        "браузер",
    )

    BAD_COPIED_TEXT_MARKERS = (
        "не был произведен вызов coinitialize",
        "uiautomationcore.dll",
        "uiautomationinitializerinthread",
        "can not load uiautomationcore",
        "диктор веб-страниц",
        "горячие клавиши активны",
        "текст страницы получен через выделение",
        "журнал работы",
        "управление чтением",
        "настройки речи",
        "__screen_reader_clipboard_marker_",
        "javascript:(()=>",
    )

    def __init__(self, root):
        self.root = root
        self.root.title("Диктор веб-страниц")
        self.root.geometry("1160x700")
        self.root.minsize(980, 600)

        self.colors = {
            "bg": "#0f172a",
            "panel": "#111827",
            "panel2": "#1f2937",
            "text": "#f8fafc",
            "muted": "#cbd5e1",
            "accent": "#38bdf8",
            "accent_dark": "#0284c7",
            "danger": "#dc2626",
            "danger_dark": "#991b1b",
            "entry": "#020617",
            "border": "#334155",
            "ok": "#22c55e",
            "warn": "#f59e0b",
        }
        self.root.configure(bg=self.colors["bg"])

        self.action_queue = queue.Queue()
        self.last_browser_hwnd = None
        self.hotkey_error_shown = False
        self.voice = None
        self.voices = []
        self.rate = 0
        self.volume = 100

        self.setup_style()
        self.create_widgets()
        self.init_sapi()
        self.update_voices_list()
        self.setup_hotkeys()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(50, self.process_action_queue)

        if WIN32_IMPORT_ERROR:
            self.log(f"Ошибка загрузки pywin32: {WIN32_IMPORT_ERROR}")
        if keyboard is None:
            self.log("Библиотека keyboard не загружена. Горячие клавиши недоступны.")

        self.log(
            "Диктор запущен.\n"
            "Для чтения всей страницы откройте вкладку браузера и нажмите Alt+Q.\n"
            "Основной способ: копирование текста прямо из активной вкладки браузера. URL используется только как резерв."
        )

    # ============================================================
    # UI
    # ============================================================

    def setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "TCombobox",
            fieldbackground=self.colors["entry"],
            background=self.colors["panel2"],
            foreground=self.colors["text"],
            arrowcolor=self.colors["text"],
            bordercolor=self.colors["border"],
            padding=5,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", self.colors["entry"])],
            foreground=[("readonly", self.colors["text"])],
        )

    def label(self, parent, text, size=10, bold=False, color=None, anchor="w"):
        return tk.Label(
            parent,
            text=text,
            bg=parent.cget("bg"),
            fg=color or self.colors["text"],
            font=("Segoe UI", size, "bold" if bold else "normal"),
            anchor=anchor,
            justify=tk.LEFT,
        )

    def create_card(self, parent, title):
        outer = tk.Frame(
            parent,
            bg=self.colors["panel"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
        )
        outer.pack(fill=tk.X, pady=3)
        tk.Label(
            outer,
            text=title,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).pack(fill=tk.X, padx=9, pady=(5, 2))
        body = tk.Frame(outer, bg=self.colors["panel"])
        body.pack(fill=tk.X, padx=9, pady=(0, 6))
        return body

    def make_button(self, parent, text, command, variant="primary"):
        if variant == "danger":
            bg = self.colors["danger"]
            active = self.colors["danger_dark"]
        elif variant == "secondary":
            bg = self.colors["panel2"]
            active = self.colors["border"]
        else:
            bg = self.colors["accent_dark"]
            active = self.colors["accent"]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg="white",
            activebackground=active,
            activeforeground="white",
            relief=tk.FLAT,
            bd=0,
            padx=7,
            pady=4,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            takefocus=0,
        )

    def create_widgets(self):
        root_frame = tk.Frame(self.root, bg=self.colors["bg"])
        root_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=14)
        root_frame.grid_columnconfigure(0, weight=1, minsize=680)
        root_frame.grid_columnconfigure(1, weight=0, minsize=330)
        root_frame.grid_rowconfigure(1, weight=1)

        title = tk.Label(
            root_frame,
            text="Экранный диктор веб-страниц",
            bg=self.colors["bg"],
            fg=self.colors["text"],
            font=("Segoe UI", 20, "bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        left = tk.Frame(root_frame, bg=self.colors["bg"])
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 14))

        right = tk.Frame(
            root_frame,
            bg=self.colors["panel"],
            width=330,
            highlightbackground=self.colors["border"],
            highlightthickness=1,
        )
        right.grid(row=1, column=1, sticky="ns")
        right.grid_propagate(False)
        self.create_hotkeys_panel(right)

        self.create_browser_card(left)
        self.create_control_card(left)
        self.create_voice_card(left)
        self.create_log_card(left)

    def create_browser_card(self, parent):
        body = self.create_card(parent, "Браузер")

        self.browser_status = tk.StringVar(value="Браузер не выбран")
        tk.Label(
            body,
            textvariable=self.browser_status,
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            font=("Segoe UI", 10),
            anchor="w",
            wraplength=650,
            justify=tk.LEFT,
        ).pack(fill=tk.X, pady=(0, 8))

        row = tk.Frame(body, bg=self.colors["panel"])
        row.pack(fill=tk.X)

        self.url_entry = tk.Entry(
            row,
            bg=self.colors["entry"],
            fg=self.colors["text"],
            insertbackground="white",
            relief=tk.FLAT,
            font=("Segoe UI", 11),
        )
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=9, padx=(0, 8))
        self.url_entry.insert(0, "https://example.com")

        self.make_button(row, "Открыть", self.open_url, "secondary").pack(side=tk.LEFT, padx=(0, 6))
        self.make_button(row, "Найти браузер", lambda: self.enqueue_action("bind_browser"), "secondary").pack(side=tk.LEFT)

    def create_control_card(self, parent):
        body = self.create_card(parent, "Управление чтением")

        grid = tk.Frame(body, bg=self.colors["panel"])
        grid.pack(fill=tk.X)
        for i in range(3):
            grid.grid_columnconfigure(i, weight=1)

        buttons = [
            ("📖 Читать страницу", "read_page", "primary"),
            ("🖱 Под курсором", "read_cursor", "primary"),
            ("🎯 Фокус", "read_focus", "secondary"),
            ("🔗 Адрес", "read_address", "secondary"),
            ("⏹ Стоп", "stop", "danger"),
            ("❔ Подсказка", "help", "secondary"),
        ]
        for idx, (text, action, variant) in enumerate(buttons):
            btn = self.make_button(grid, text, lambda a=action: self.enqueue_action(a), variant)
            btn.grid(row=idx // 3, column=idx % 3, sticky="ew", padx=4, pady=4)

    def create_voice_card(self, parent):
        body = self.create_card(parent, "Настройки речи")

        top = tk.Frame(body, bg=self.colors["panel"])
        top.pack(fill=tk.X, pady=2)

        speed_box = tk.Frame(top, bg=self.colors["panel"])
        speed_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.label(speed_box, "Скорость", 9).pack(side=tk.LEFT, padx=(0, 6))
        self.speed_scale = tk.Scale(
            speed_box,
            from_=-10,
            to=10,
            orient=tk.HORIZONTAL,
            command=self.change_speed,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            troughcolor=self.colors["panel2"],
            highlightthickness=0,
            length=170,
            takefocus=0,
            showvalue=0,
        )
        self.speed_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.speed_label = self.label(speed_box, "0", 9, color=self.colors["muted"])
        self.speed_label.pack(side=tk.LEFT, padx=6)

        volume_box = tk.Frame(top, bg=self.colors["panel"])
        volume_box.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.label(volume_box, "Громкость", 9).pack(side=tk.LEFT, padx=(0, 6))
        self.volume_scale = tk.Scale(
            volume_box,
            from_=0,
            to=100,
            orient=tk.HORIZONTAL,
            command=self.change_volume,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            troughcolor=self.colors["panel2"],
            highlightthickness=0,
            length=170,
            takefocus=0,
            showvalue=0,
        )
        self.volume_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.volume_label = self.label(volume_box, "100%", 9, color=self.colors["muted"])
        self.volume_label.pack(side=tk.LEFT, padx=6)

        row3 = tk.Frame(body, bg=self.colors["panel"])
        row3.pack(fill=tk.X, pady=(4, 0))
        self.label(row3, "Голос", 9).pack(side=tk.LEFT, padx=(0, 8))
        self.voice_combo = ttk.Combobox(row3, state="readonly", width=45)
        self.voice_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.voice_combo.bind("<<ComboboxSelected>>", self.change_voice)

    def create_log_card(self, parent):
        outer = tk.Frame(
            parent,
            bg=self.colors["panel"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
        )
        outer.pack(fill=tk.BOTH, expand=True, pady=7)
        tk.Label(
            outer,
            text="Журнал работы",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        ).pack(fill=tk.X, padx=9, pady=(5, 2))

        body = tk.Frame(outer, bg=self.colors["panel"])
        body.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.log_text = tk.Text(
            body,
            bg=self.colors["entry"],
            fg=self.colors["text"],
            insertbackground="white",
            relief=tk.FLAT,
            font=("Consolas", 10),
            wrap=tk.WORD,
            height=18,
            takefocus=0,
            state=tk.DISABLED,
        )
        scrollbar = tk.Scrollbar(body, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def create_hotkeys_panel(self, parent):
        tk.Label(
            parent,
            text="Горячие клавиши",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        ).pack(fill=tk.X, padx=16, pady=(16, 10))

        hotkeys = [
            ("Читать страницу", "Alt + Q", "Ctrl + Alt + 1"),
            ("Под курсором", "Alt + W", "Ctrl + Alt + 2"),
            ("Элемент в фокусе", "Alt + E", "Ctrl + Alt + 3"),
            ("Остановить", "Alt + S", "Ctrl + Alt + 0"),
            ("Адрес страницы", "Alt + A", "Ctrl + Alt + 4"),
            ("Подсказка", "Alt + H", "Ctrl + Alt + 5"),
        ]

        for action, main_key, backup_key in hotkeys:
            item = tk.Frame(parent, bg=self.colors["panel2"], highlightthickness=1, highlightbackground=self.colors["border"])
            item.pack(fill=tk.X, padx=16, pady=5)
            tk.Label(
                item,
                text=action,
                bg=self.colors["panel2"],
                fg=self.colors["text"],
                font=("Segoe UI", 10, "bold"),
                anchor="w",
            ).pack(fill=tk.X, padx=10, pady=(8, 1))
            tk.Label(
                item,
                text=f"{main_key}     или     {backup_key}",
                bg=self.colors["panel2"],
                fg=self.colors["accent"],
                font=("Consolas", 10, "bold"),
                anchor="w",
            ).pack(fill=tk.X, padx=10, pady=(0, 8))

        note = (
            "Alt+Q работает корректнее, если нажимать его, когда активна вкладка браузера. "
            "Ctrl+Shift+R не используется, потому что это перезагрузка страницы."
        )
        tk.Label(
            parent,
            text=note,
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            wraplength=295,
            justify=tk.LEFT,
            anchor="w",
        ).pack(fill=tk.X, padx=16, pady=(14, 0))

    # ============================================================
    # Queue / hotkeys
    # ============================================================

    def enqueue_action(self, action, source_hwnd=None):
        if source_hwnd is None:
            source_hwnd = get_foreground_hwnd()
        self.action_queue.put((action, source_hwnd))

    def process_action_queue(self):
        try:
            while True:
                action, source_hwnd = self.action_queue.get_nowait()
                self.run_action(action, source_hwnd)
        except queue.Empty:
            pass
        self.root.after(50, self.process_action_queue)

    def setup_hotkeys(self):
        if keyboard is None:
            return

        hotkeys = [
            ("alt+q", "read_page"),
            ("ctrl+alt+1", "read_page"),
            ("alt+w", "read_cursor"),
            ("ctrl+alt+2", "read_cursor"),
            ("alt+e", "read_focus"),
            ("ctrl+alt+3", "read_focus"),
            ("alt+s", "stop"),
            ("ctrl+alt+0", "stop"),
            ("alt+a", "read_address"),
            ("ctrl+alt+4", "read_address"),
            ("alt+h", "help"),
            ("ctrl+alt+5", "help"),
        ]

        failed = []
        for combo, action in hotkeys:
            try:
                keyboard.add_hotkey(
                    combo,
                    lambda a=action: self.enqueue_action(a, get_foreground_hwnd()),
                    suppress=True,
                    trigger_on_release=True,
                )
            except Exception as exc:
                try:
                    keyboard.add_hotkey(
                        combo,
                        lambda a=action: self.enqueue_action(a, get_foreground_hwnd()),
                        suppress=False,
                        trigger_on_release=True,
                    )
                except Exception as exc2:
                    failed.append(f"{combo}: {exc2}")

        if failed:
            self.log("Не все горячие клавиши удалось установить. Возможно, нужен запуск от имени администратора.")
            for item in failed[:6]:
                self.log(item)

    def run_action(self, action, source_hwnd=None):
        ensure_com_initialized()
        try:
            if action == "read_page":
                self.read_page(source_hwnd)
            elif action == "read_cursor":
                self.read_element_under_cursor()
            elif action == "read_focus":
                self.read_focused_element()
            elif action == "read_address":
                self.read_current_address(source_hwnd)
            elif action == "stop":
                self.stop_reading()
            elif action == "help":
                self.read_hotkeys_help()
            elif action == "bind_browser":
                self.bind_browser(source_hwnd)
        except Exception as exc:
            self.log(f"Ошибка действия {action}: {exc}")
            self.log(traceback.format_exc())
            self.speak("Произошла ошибка. Подробности записаны в журнал.", log_it=False)

    # ============================================================
    # Speech / log
    # ============================================================

    def init_sapi(self):
        if win32com is None:
            self.log("SAPI недоступен: pywin32 не загружен.")
            return
        ensure_com_initialized()
        try:
            self.voice = win32com.client.Dispatch("SAPI.SpVoice")
            self.voices = list(self.voice.GetVoices())
        except Exception as exc:
            self.log(f"Не удалось запустить SAPI.SpVoice: {exc}")
            self.voice = None
            self.voices = []

    def speak(self, text, log_it=True):
        text = self.clean_text(text)
        if not text:
            return
        if log_it:
            preview = text if len(text) <= 700 else text[:700] + "..."
            self.log(f"Диктор: {preview}")
        if self.voice is None:
            return
        ensure_com_initialized()
        try:
            self.voice.Speak("", 3)
            self.voice.Speak(text, 1)
        except Exception as exc:
            self.log(f"Ошибка SAPI: {exc}")

    def stop_reading(self):
        if self.voice is not None:
            ensure_com_initialized()
            try:
                self.voice.Speak("", 3)
            except Exception:
                pass
        self.log("Чтение остановлено")

    def log(self, message):
        if not hasattr(self, "log_text"):
            return
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    # ============================================================
    # Text cleanup / validation
    # ============================================================

    def clean_text(self, text):
        if text is None:
            return ""
        text = str(text)
        text = text.replace("\r", "\n")
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        text = text.strip()
        return text

    def looks_like_url(self, text):
        text = self.clean_text(text)
        if not text:
            return False
        if re.match(r"^https?://\S+$", text, flags=re.IGNORECASE):
            return True
        if re.match(r"^(www\.)?[a-zA-Z0-9-]+\.[a-zA-Z]{2,}(/\S*)?$", text):
            return True
        return False

    def is_bad_copied_text(self, text):
        text = self.clean_text(text)
        if not text:
            return True
        low = text.lower()
        if any(marker in low for marker in self.BAD_COPIED_TEXT_MARKERS):
            return True
        if self.looks_like_url(text) and len(text) < 300:
            return True
        # Если скопировалось 1–2 коротких слова, это почти точно не содержимое страницы.
        if len(text) < 25 and len(text.split()) < 5:
            return True
        return False

    # ============================================================
    # Clipboard
    # ============================================================

    def get_clipboard_text(self):
        if win32clipboard is None:
            try:
                return self.root.clipboard_get()
            except Exception:
                return ""
        text = ""
        try:
            win32clipboard.OpenClipboard()
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        except Exception:
            text = ""
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
        return text or ""

    def set_clipboard_text(self, text):
        text = "" if text is None else str(text)
        if win32clipboard is None:
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                self.root.update()
                return True
            except Exception:
                return False
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
            return True
        except Exception:
            return False
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass

    def wait_clipboard_change(self, old_marker, timeout=1.2):
        deadline = time.time() + timeout
        last = ""
        while time.time() < deadline:
            last = self.get_clipboard_text()
            if last and last != old_marker:
                return last
            time.sleep(0.05)
        return last

    # ============================================================
    # Browser window detection
    # ============================================================

    def get_window_title(self, hwnd):
        if not hwnd or win32gui is None:
            return ""
        try:
            return win32gui.GetWindowText(hwnd) or ""
        except Exception:
            return ""

    def get_window_class(self, hwnd):
        if not hwnd or win32gui is None:
            return ""
        try:
            return win32gui.GetClassName(hwnd) or ""
        except Exception:
            return ""

    def get_root_toplevel_hwnd(self, hwnd):
        if not hwnd or win32gui is None:
            return hwnd
        try:
            root = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
            return root or hwnd
        except Exception:
            return hwnd

    def is_own_window(self, hwnd):
        if not hwnd:
            return False
        try:
            return int(hwnd) == int(self.root.winfo_id())
        except Exception:
            return False

    def is_browser_window(self, hwnd):
        if not hwnd or win32gui is None:
            return False
        hwnd = self.get_root_toplevel_hwnd(hwnd)
        if self.is_own_window(hwnd):
            return False
        try:
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return False
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if right - left < 300 or bottom - top < 250:
                return False
        except Exception:
            return False

        class_name = self.get_window_class(hwnd)
        title = self.get_window_title(hwnd).lower()
        if class_name in self.BROWSER_CLASSES:
            return True
        if any(word in title for word in self.BROWSER_TITLE_WORDS):
            return True
        return False

    def enum_browser_windows(self):
        result = []
        if win32gui is None:
            return result

        def callback(hwnd, _):
            try:
                if self.is_browser_window(hwnd):
                    result.append(hwnd)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(callback, None)
        except Exception:
            pass

        # Убираем дубли, сохраняя порядок.
        unique = []
        seen = set()
        for hwnd in result:
            root = self.get_root_toplevel_hwnd(hwnd)
            if root not in seen:
                seen.add(root)
                unique.append(root)
        return unique

    def resolve_browser_hwnd(self, source_hwnd=None):
        
        if source_hwnd:
            hwnd = self.get_root_toplevel_hwnd(source_hwnd)
            if self.is_browser_window(hwnd):
                self.last_browser_hwnd = hwnd
                return hwnd


        if self.last_browser_hwnd and self.is_browser_window(self.last_browser_hwnd):
            return self.last_browser_hwnd

       
        browsers = self.enum_browser_windows()
        if browsers:
            self.last_browser_hwnd = browsers[0]
            return browsers[0]

        return None

    def bind_browser(self, source_hwnd=None):
        hwnd = self.resolve_browser_hwnd(source_hwnd)
        if not hwnd:
            self.browser_status.set("Браузер не найден. Откройте Chrome, Edge, Firefox или Яндекс Браузер.")
            self.speak("Браузер не найден.")
            return None
        title = self.get_window_title(hwnd) or "без названия"
        cls = self.get_window_class(hwnd)
        self.last_browser_hwnd = hwnd
        self.browser_status.set(f"Выбран браузер: {title} [{cls}]")
        self.log(f"Выбран браузер: {title} [{cls}]")
        return hwnd

    def activate_window(self, hwnd):
        if not hwnd or win32gui is None:
            return False
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        except Exception:
            pass

        attached = []
        try:

            self.release_modifier_keys()

            if win32process is not None and win32api is not None:
                try:
                    current_tid = win32api.GetCurrentThreadId()
                    target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
                    fg = win32gui.GetForegroundWindow()
                    fg_tid, _ = win32process.GetWindowThreadProcessId(fg) if fg else (None, None)
                    for tid in {target_tid, fg_tid}:
                        if tid and tid != current_tid:
                            win32process.AttachThreadInput(current_tid, tid, True)
                            attached.append((current_tid, tid))
                except Exception:
                    pass

            # Alt-trick помогает Windows разрешить перевод фокуса в другое окно.
            if win32api is not None and win32con is not None:
                win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
                time.sleep(0.02)
                win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)

            win32gui.SetForegroundWindow(hwnd)
            win32gui.BringWindowToTop(hwnd)
            time.sleep(0.25)
            return int(win32gui.GetForegroundWindow()) == int(hwnd)
        except Exception:
            try:
                win32gui.BringWindowToTop(hwnd)
                time.sleep(0.25)
                return True
            except Exception:
                return False
        finally:
            if win32process is not None:
                for current_tid, tid in attached:
                    try:
                        win32process.AttachThreadInput(current_tid, tid, False)
                    except Exception:
                        pass

    def click_inside_page(self, hwnd):
        """Резервный способ перевести фокус с адресной строки на страницу."""
        if not hwnd or win32gui is None or win32api is None:
            return False
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            if width < 300 or height < 250:
                return False

            # Ниже панели вкладок и адресной строки, ближе к центру страницы.
            x = left + width // 2
            y = top + min(max(220, height // 3), height - 90)
            old_pos = win32gui.GetCursorPos()
            win32api.SetCursorPos((x, y))
            time.sleep(0.04)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.04)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.10)
            try:
                win32api.SetCursorPos(old_pos)
            except Exception:
                pass
            return True
        except Exception:
            return False

    # ============================================================
    # HTML page loading / extraction
    # ============================================================

    def normalize_url(self, url):
        url = self.clean_text(url)
        if not url:
            return ""
        low = url.lower()
        # Важно: маркеры буфера обмена нельзя превращать в https://...
        if "screen_reader_clipboard_marker" in low or low.startswith("javascript:"):
            return ""
        # Адресная строка должна быть одной строкой без пробелов.
        if " " in url or "\n" in url or "\r" in url:
            return ""
        if url.startswith("www."):
            url = "https://" + url
        parsed = urlparse.urlparse(url)
        if not parsed.scheme and "." in url.split()[0] and not url.startswith("__"):
            url = "https://" + url
            parsed = urlparse.urlparse(url)
        if parsed.scheme not in {"http", "https"}:
            return ""
        if not parsed.netloc:
            return ""
        return url

    def fetch_html_text_from_url(self, url):
        url = self.normalize_url(url)
        if not url:
            return ""

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,text/plain;q=0.8,*/*;q=0.5",
        }

        try:
            req = urlrequest.Request(url, headers=headers)
            with urlrequest.urlopen(req, timeout=8) as response:
                content_type = response.headers.get("Content-Type", "").lower()
                raw = response.read(2_500_000)
                charset = response.headers.get_content_charset() or "utf-8"
        except Exception as exc:
            self.log(f"Не удалось загрузить страницу по адресу: {exc}")
            return ""

        if not raw:
            return ""

        try:
            html = raw.decode(charset, errors="replace")
        except Exception:
            html = raw.decode("utf-8", errors="replace")

        if "html" not in content_type and "text/plain" not in content_type and "xml" not in content_type:
            self.log(f"Получен не HTML-документ: {content_type or 'тип не указан'}")
            return ""

        if "text/plain" in content_type:
            text = html
        else:
            parser = ReadableHTMLTextExtractor()
            try:
                parser.feed(html)
                text = parser.get_text()
            except Exception as exc:
                self.log(f"Ошибка разбора HTML: {exc}")
                text = re.sub(r"<[^>]+>", " ", html)

        return self.prepare_page_text_for_speech(text)

    # ============================================================
    # Browser copy strategies: page / address
    # ============================================================

    def release_modifier_keys(self):
        if win32api is None or win32con is None:
            return
        for vk in (win32con.VK_MENU, win32con.VK_CONTROL, win32con.VK_SHIFT, win32con.VK_LWIN, win32con.VK_RWIN):
            try:
                win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
            except Exception:
                pass
        time.sleep(0.03)

    def virtual_key(self, key):
        key = key.strip().lower()
        if win32con is None:
            return None
        mapping = {
            "ctrl": win32con.VK_CONTROL,
            "control": win32con.VK_CONTROL,
            "alt": win32con.VK_MENU,
            "shift": win32con.VK_SHIFT,
            "esc": win32con.VK_ESCAPE,
            "escape": win32con.VK_ESCAPE,
            "enter": win32con.VK_RETURN,
            "tab": win32con.VK_TAB,
            "space": win32con.VK_SPACE,
            "f6": win32con.VK_F6,
        }
        if key in mapping:
            return mapping[key]
        if len(key) == 1 and "a" <= key <= "z":
            return ord(key.upper())
        if len(key) == 1 and "0" <= key <= "9":
            return ord(key)
        return None

    def keyboard_press(self, combo):
        
        if win32api is None or win32con is None:
            if keyboard is None:
                return False
            try:
                keyboard.press_and_release(combo)
                time.sleep(0.08)
                return True
            except Exception as exc:
                self.log(f"Не удалось нажать {combo}: {exc}")
                return False

        try:
            self.release_modifier_keys()
            keys = [self.virtual_key(part) for part in combo.split("+")]
            keys = [vk for vk in keys if vk is not None]
            if not keys:
                return False
            for vk in keys:
                win32api.keybd_event(vk, 0, 0, 0)
                time.sleep(0.025)
            time.sleep(0.06)
            for vk in reversed(keys):
                win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.025)
            time.sleep(0.12)
            return True
        except Exception as exc:
            self.log(f"Не удалось нажать {combo}: {exc}")
            return False

    def copy_from_browser(self, hwnd, mode="page"):

        if not hwnd:
            return ""

        old_clipboard = self.get_clipboard_text()
        marker = f"__SCREEN_READER_CLIPBOARD_MARKER_{time.time()}__"
        copied = ""

        try:
            self.set_clipboard_text(marker)
            if not self.activate_window(hwnd):
                self.log("Не удалось гарантированно активировать окно браузера. Всё равно пробую отправить клавиши.")
            self.release_modifier_keys()
            self.keyboard_press("esc")

            if mode == "address":
                self.keyboard_press("ctrl+l")
                self.keyboard_press("ctrl+c")
                copied = self.clean_text(self.wait_clipboard_change(marker, timeout=1.8))
                self.keyboard_press("esc")
                if not copied or copied == marker or "SCREEN_READER_CLIPBOARD_MARKER" in copied:
                    return ""
                return copied

            # Основной способ: фокусируем область сайта мышью и копируем выделение страницы.
            attempts = [
                "клик в область страницы",
                "Esc + повторный клик в область страницы",
                "F6 + клик в область страницы",
            ]
            for idx, label in enumerate(attempts, start=1):
                self.set_clipboard_text(marker)
                self.activate_window(hwnd)
                self.release_modifier_keys()
                if idx == 2:
                    self.keyboard_press("esc")
                if idx == 3:
                    self.keyboard_press("f6")
                    self.keyboard_press("esc")
                self.click_inside_page(hwnd)
                self.keyboard_press("ctrl+a")
                self.keyboard_press("ctrl+c")
                copied = self.clean_text(self.wait_clipboard_change(marker, timeout=1.8))

                if copied and copied != marker and not self.is_bad_copied_text(copied):
                    self.log(f"Текст страницы получен через выделение вкладки браузера: попытка {idx}, {label}.")
                    return copied

                if copied and copied != marker:
                    short = copied[:180].replace("\n", " ")
                    self.log(f"Попытка {idx} отклонена: скопирован не текст страницы. Фрагмент: {short}")
                else:
                    self.log(f"Попытка {idx} не изменила буфер обмена.")

            return ""
        finally:
            try:
                self.set_clipboard_text(old_clipboard)
            except Exception:
                pass

    def copy_page_text_with_js_bookmarklet(self, hwnd):
        """Резерв: выполняет bookmarklet в адресной строке и копирует document.body.innerText."""
        if not hwnd:
            return ""
        old_clipboard = self.get_clipboard_text()
        script = (
            "javascript:(()=>{try{let t=document.body?document.body.innerText:'';"
            "let a=document.createElement('textarea');a.value=t;"
            "document.body.appendChild(a);a.select();document.execCommand('copy');a.remove();}catch(e){}})()"
        )
        try:
            if not self.activate_window(hwnd):
                self.log("JS-резерв: не удалось гарантированно активировать браузер.")
            self.release_modifier_keys()
            self.set_clipboard_text(script)
            self.keyboard_press("ctrl+l")
            self.keyboard_press("ctrl+v")
            self.keyboard_press("enter")
            deadline = time.time() + 2.5
            copied = ""
            while time.time() < deadline:
                copied = self.clean_text(self.get_clipboard_text())
                if copied and copied != script and not copied.lower().startswith("javascript:"):
                    break
                time.sleep(0.08)
            if copied and copied != script and not self.is_bad_copied_text(copied):
                self.log("Текст страницы получен через резервный JS-способ: document.body.innerText.")
                return copied
            return ""
        finally:
            try:
                self.set_clipboard_text(old_clipboard)
            except Exception:
                pass

    def copy_address_with_js_bookmarklet(self, hwnd):
        """Резерв: копирует location.href из текущей вкладки."""
        if not hwnd:
            return ""
        old_clipboard = self.get_clipboard_text()
        script = (
            "javascript:(()=>{try{let a=document.createElement('textarea');"
            "a.value=location.href;document.body.appendChild(a);a.select();"
            "document.execCommand('copy');a.remove();}catch(e){}})()"
        )
        try:
            self.activate_window(hwnd)
            self.release_modifier_keys()
            self.set_clipboard_text(script)
            self.keyboard_press("ctrl+l")
            self.keyboard_press("ctrl+v")
            self.keyboard_press("enter")
            deadline = time.time() + 2.0
            copied = ""
            while time.time() < deadline:
                copied = self.clean_text(self.get_clipboard_text())
                if copied and copied != script and not copied.lower().startswith("javascript:"):
                    break
                time.sleep(0.08)
            copied = self.normalize_url(copied)
            if copied:
                self.log("Адрес страницы получен через резервный JS-способ: location.href.")
                return copied
            return ""
        finally:
            try:
                self.set_clipboard_text(old_clipboard)
            except Exception:
                pass

    def read_page(self, source_hwnd=None):
        hwnd = self.resolve_browser_hwnd(source_hwnd)
        if not hwnd:
            self.speak("Браузер не найден. Откройте нужную страницу в Chrome, Edge, Firefox или Яндекс Браузере.")
            return

        self.bind_browser(hwnd)
        self.log("Начинаю чтение страницы. Основной способ: скопировать текст прямо из активной вкладки браузера.")

        # 1. Самый надежный для динамических сайтов способ: Ctrl+A/Ctrl+C внутри страницы.
        text = self.copy_from_browser(hwnd, mode="page")
        text = self.prepare_page_text_for_speech(text)

        # 2. Резерв: bookmarklet берет document.body.innerText без UIAutomation.
        if not text or self.is_bad_copied_text(text):
            self.log("Обычное копирование не дало нормальный текст. Пробую резерв через document.body.innerText.")
            text = self.copy_page_text_with_js_bookmarklet(hwnd)
            text = self.prepare_page_text_for_speech(text)

        # 3. Последний резерв: получить URL и скачать HTML напрямую.
        if not text or self.is_bad_copied_text(text):
            self.log("JS-резерв не дал нормальный текст. Пробую получить адрес и загрузить HTML по URL.")
            address = self.copy_from_browser(hwnd, mode="address")
            address = self.normalize_url(address) or self.copy_address_with_js_bookmarklet(hwnd)
            if address:
                self.log(f"Адрес страницы получен: {address}")
                text = self.fetch_html_text_from_url(address)
                if text and not self.is_bad_copied_text(text):
                    self.log("Текст страницы получен через HTML-загрузку по адресу.")
            else:
                self.log("Адрес страницы получить не удалось. Маркер буфера обмена не принят как URL.")

        if not text or self.is_bad_copied_text(text):
            self.speak(
                "Текст страницы не удалось получить. Откройте обычную страницу с текстом, кликните по содержимому сайта и нажмите Альт Кью. "
                "Если сайт является мессенджером, PDF, стартовой страницей или защищенным сервисом, полное чтение может не сработать."
            )
            return

        if len(text) > self.MAX_READ_CHARS:
            text = text[: self.MAX_READ_CHARS] + ". Страница длинная, поэтому прочитана первая часть."
        self.speak(text)

    def prepare_page_text_for_speech(self, text):
        text = self.clean_text(text)
        text = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", text)
        text = re.sub(r"\s+", " ", text)
        # Возвращаем переносы перед фильтрацией крупных блоков.
        text = re.sub(r"([.!?])\s+", r"\1\n", text)

        raw_lines = [line.strip(" \t\n\r|•·-–—") for line in text.splitlines()]
        lines = []
        for line in raw_lines:
            line = re.sub(r"\s+", " ", line).strip()
            if not line:
                continue
            low = line.lower()
            if len(line) <= 1:
                continue
            if low in {"reload", "back", "forward", "new tab", "downloads", "extensions", "menu", "search"}:
                continue
            if any(token in low for token in (
                "cookie", "cookies", "privacy policy", "terms of use", "sign in", "log in", "subscribe",
                "accept all", "manage settings", "enable javascript", "горячие клавиши", "журнал работы"
            )) and len(line) < 140:
                continue
            if re.match(r"^https?://\S+$", line):
                continue
            lines.append(line)

        filtered = []
        seen = set()
        for line in lines:
            key = line.lower()
            if key in seen:
                continue
            seen.add(key)
            filtered.append(line)

        # Если строк мало, не превращаем текст в бессмысленную кашу — оставляем только полезные блоки.
        return ". ".join(filtered).strip()

    def read_current_address(self, source_hwnd=None):
        hwnd = self.resolve_browser_hwnd(source_hwnd)
        if not hwnd:
            self.speak("Браузер не найден.")
            return
        self.bind_browser(hwnd)
        address = self.normalize_url(self.copy_from_browser(hwnd, mode="address"))
        if not address:
            address = self.copy_address_with_js_bookmarklet(hwnd)
        if address:
            self.speak(f"Адрес страницы: {address}")
        else:
            self.speak("Не удалось получить адрес страницы.")

    # ============================================================
    # UIAutomation only for cursor/focus reading
    # ============================================================

    def load_uiautomation(self):
        ensure_com_initialized()
        try:
            import uiautomation as auto
            return auto
        except Exception as exc:
            self.log(f"UIAutomation не загрузился: {exc}")
            return None

    def safe_get_name(self, elem):
        try:
            return elem.Name or ""
        except Exception:
            return ""

    def safe_get_type(self, elem):
        try:
            return elem.ControlTypeName or ""
        except Exception:
            return ""

    def safe_get_value(self, elem, auto):
        try:
            pattern = elem.GetPattern(auto.ValuePattern)
            if pattern:
                return pattern.CurrentValue or ""
        except Exception:
            pass
        return ""

    def extract_text_from_element(self, elem, auto, max_depth=2, depth=0):
        if elem is None or depth > max_depth:
            return ""
        parts = []
        name = self.safe_get_name(elem)
        value = self.safe_get_value(elem, auto)
        if name:
            parts.append(name)
        if value and value != name:
            parts.append(value)
        if depth < max_depth:
            try:
                children = elem.GetChildren()
            except Exception:
                children = []
            for child in children[:30]:
                child_text = self.extract_text_from_element(child, auto, max_depth, depth + 1)
                if child_text:
                    parts.append(child_text)
        unique = []
        seen = set()
        for item in parts:
            item = self.clean_text(item)
            if item and item not in seen:
                seen.add(item)
                unique.append(item)
        return " ".join(unique)

    def read_element_under_cursor(self):
        auto = self.load_uiautomation()
        if auto is None:
            self.speak("Не удалось загрузить UI Automation. Чтение элемента под курсором недоступно.")
            return
        try:
            if win32gui is not None:
                x, y = win32gui.GetCursorPos()
            else:
                x, y = auto.GetCursorPos()
            elem = auto.ControlFromPoint(x, y)
            text = self.extract_text_from_element(elem, auto, max_depth=2)
            if text:
                self.speak(text[:900])
            else:
                control_type = self.safe_get_type(elem) or "элемент"
                self.speak(f"{control_type} под курсором, текста нет.")
        except Exception as exc:
            self.log(f"Ошибка чтения под курсором: {exc}")
            self.speak("Не удалось прочитать элемент под курсором.")

    def read_focused_element(self):
        auto = self.load_uiautomation()
        if auto is None:
            self.speak("Не удалось загрузить UI Automation. Чтение фокуса недоступно.")
            return
        try:
            elem = auto.GetFocusedControl()
            text = self.extract_text_from_element(elem, auto, max_depth=2)
            if text:
                self.speak(text[:900])
            else:
                control_type = self.safe_get_type(elem) or "элемент"
                self.speak(f"{control_type} в фокусе, текста нет.")
        except Exception as exc:
            self.log(f"Ошибка чтения фокуса: {exc}")
            self.speak("Не удалось прочитать элемент в фокусе.")

    # ============================================================
    # Voice settings
    # ============================================================

    def update_voices_list(self):
        if not hasattr(self, "voice_combo"):
            return
        names = []
        for voice in self.voices:
            try:
                names.append(voice.GetDescription())
            except Exception:
                pass
        self.voice_combo["values"] = names
        if names:
            self.voice_combo.current(0)
            try:
                self.voice.Voice = self.voices[0]
            except Exception:
                pass
        self.speed_scale.set(0)
        self.volume_scale.set(100)

    def change_voice(self, event=None):
        if self.voice is None:
            return
        selected = self.voice_combo.get()
        for voice in self.voices:
            try:
                if voice.GetDescription() == selected:
                    self.voice.Voice = voice
                    self.log(f"Голос изменён: {selected}")
                    return
            except Exception:
                pass

    def change_speed(self, value):
        self.rate = int(float(value))
        self.speed_label.configure(text=str(self.rate))
        if self.voice is not None:
            try:
                self.voice.Rate = self.rate
            except Exception:
                pass

    def change_volume(self, value):
        self.volume = int(float(value))
        self.volume_label.configure(text=f"{self.volume}%")
        if self.voice is not None:
            try:
                self.voice.Volume = self.volume
            except Exception:
                pass

    # ============================================================
    # Other actions
    # ============================================================

    def open_url(self):
        url = self.url_entry.get().strip()
        if not url:
            self.log("Введите адрес сайта.")
            return
        if not url.startswith(("http://", "https://")):
            url = "https://" + url
        try:
            webbrowser.open(url)
            self.log(f"Открыт адрес: {url}")
            self.root.after(1200, lambda: self.enqueue_action("bind_browser"))
        except Exception as exc:
            self.log(f"Ошибка открытия адреса: {exc}")

    def read_hotkeys_help(self):
        self.speak(
            "Горячие клавиши. Alt Q — читать страницу. Alt W — читать под курсором. "
            "Alt E — читать элемент в фокусе. Alt A — прочитать адрес. Alt S — остановить чтение. "
            "Запасные сочетания: Control Alt 1, 2, 3, 4 и 0."
        )

    def on_close(self):
        try:
            if keyboard is not None:
                keyboard.unhook_all_hotkeys()
        except Exception:
            pass
        try:
            self.stop_reading()
        except Exception:
            pass
        self.root.destroy()


if __name__ == "__main__":
    ensure_com_initialized()
    root = tk.Tk()
    app = ScreenReaderApp(root)
    root.mainloop()
