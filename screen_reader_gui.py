import tkinter as tk
from tkinter import ttk, messagebox
import time
import webbrowser
import uiautomation as auto
import keyboard
import win32com.client
import re
import traceback
from collections import deque


class ScreenReaderApp:
    """
    Экранный диктор для чтения веб-страниц.

    Главное отличие этой версии:
    кнопка «Читать содержимое страницы» сначала ищет DocumentControl / web-view
    внутри окна браузера и читает только эту область, а не всё окно браузера
    вместе с вкладками, адресной строкой и панелью кнопок.
    """

    MAX_READ_CHARS = 8000
    MAX_TREE_NODES = 2500

    CONTENT_CONTROL_TYPES = {
        "TextControl",
        "HyperlinkControl",
        "ButtonControl",
        "ListItemControl",
        "DataItemControl",
        "EditControl",
        "HeaderControl",
        "DocumentControl",
        "ImageControl",
    }

    PAGE_ROOT_CONTROL_TYPES = {
        "DocumentControl",
        "PaneControl",
        "CustomControl",
        "GroupControl",
    }

    BROWSER_UI_WORDS = (
        "address and search bar",
        "address bar",
        "omnibox",
        "search or enter web address",
        "tab search",
        "bookmarks bar",
        "toolbar",
        "browser toolbar",
        "navigation",
        "downloads",
        "extensions",
        "profile",
        "new tab",
        "reload",
        "refresh",
        "back",
        "forward",
        "адресная строка",
        "строка адреса",
        "панель инструментов",
        "закладки",
        "новая вкладка",
        "обновить",
        "назад",
        "вперёд",
        "вперед",
    )

    BROWSER_CLASSES = [
        "Chrome_WidgetWin_1",        # Chrome, Edge, Яндекс
        "MozillaWindowClass",        # Firefox
        "ApplicationFrameWindow",    # Edge / UWP-окна
        "YandexBrowserWidgetWin",
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("Диктор веб-страниц")
        self.root.geometry("1080x780")
        self.root.minsize(920, 650)

        self.colors = {
            "bg": "#0f172a",
            "panel": "#111827",
            "panel_2": "#1f2937",
            "text": "#f8fafc",
            "muted": "#cbd5e1",
            "accent": "#38bdf8",
            "accent_dark": "#0284c7",
            "danger": "#dc2626",
            "danger_dark": "#991b1b",
            "entry": "#020617",
            "border": "#334155",
        }

        self.root.configure(bg=self.colors["bg"])

        self.voice = win32com.client.Dispatch("SAPI.SpVoice")
        self.voices = self.voice.GetVoices()
        self.current_voice_index = 0
        self.rate = 0
        self.volume = 100

        self.setup_style()
        self.create_widgets()
        self.update_voices_list()
        self.speed_scale.set(0)
        self.volume_scale.set(100)
        self.setup_hotkeys()
        self.root.bind("<Escape>", lambda event: self.stop_reading())

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.log(
            "Диктор запущен. Горячие клавиши активны.\n"
            "x — читать содержимое страницы\n"
            "c — читать элемент в фокусе\n"
            "v — читать элемент под курсором мыши\n"
            "q — остановить чтение\n"
            "w — прочитать адрес / заголовок страницы\n"
            "e — озвучить список горячих клавиш"
        )

    # ---------------------- Оформление ----------------------

    def setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "TCombobox",
            fieldbackground=self.colors["entry"],
            background=self.colors["panel_2"],
            foreground=self.colors["text"],
            arrowcolor=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            padding=6,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", self.colors["entry"])],
            foreground=[("readonly", self.colors["text"])],
        )

    def create_card(self, parent, title):
        outer = tk.Frame(
            parent,
            bg=self.colors["panel"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
            bd=0,
        )
        outer.pack(fill=tk.X, pady=8)

        header = tk.Label(
            outer,
            text=title,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        )
        header.pack(fill=tk.X, padx=16, pady=(14, 6))

        body = tk.Frame(outer, bg=self.colors["panel"])
        body.pack(fill=tk.X, padx=16, pady=(0, 16))
        return body

    def make_button(self, parent, text, command, variant="primary"):
        if variant == "danger":
            bg = self.colors["danger"]
            active = self.colors["danger_dark"]
            fg = "white"
        elif variant == "secondary":
            bg = self.colors["panel_2"]
            active = self.colors["border"]
            fg = self.colors["text"]
        else:
            bg = self.colors["accent_dark"]
            active = self.colors["accent"]
            fg = "white"

        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active,
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            relief=tk.FLAT,
            bd=0,
            padx=16,
            pady=10,
            cursor="hand2",
            takefocus=True,
            highlightthickness=2,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
        )

    def create_hotkeys_column(self, parent):
        hotkeys = [
            ("Читать страницу", "Ctrl + Alt + 1"),
            ("Читать фокус", "Ctrl + Alt + 2"),
            ("Под курсором", "Ctrl + Alt + 3"),
            ("Остановить", "Ctrl + Alt + 0"),
            ("Адрес страницы", "Ctrl + Alt + 4"),
            ("Подсказка", "Ctrl + Alt + 5"),
        ]

        card = tk.Frame(
            parent,
            bg=self.colors["panel"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill=tk.BOTH, expand=True)

        title = tk.Label(
            card,
            text="Горячие клавиши",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 14, "bold"),
            anchor="w",
        )
        title.pack(fill=tk.X, padx=14, pady=(14, 4))

        subtitle = tk.Label(
            card,
            text="Цифровые сочетания не зависят от русской/английской раскладки и не совпадают с обновлением страницы.",
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            justify=tk.LEFT,
            wraplength=230,
            anchor="w",
        )
        subtitle.pack(fill=tk.X, padx=14, pady=(0, 12))

        header = tk.Frame(card, bg=self.colors["panel_2"])
        header.pack(fill=tk.X, padx=14, pady=(0, 4))

        tk.Label(
            header,
            text="Действие",
            bg=self.colors["panel_2"],
            fg=self.colors["text"],
            font=("Segoe UI", 9, "bold"),
            anchor="w",
            width=15,
            padx=8,
            pady=7,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            header,
            text="Клавиши",
            bg=self.colors["panel_2"],
            fg=self.colors["text"],
            font=("Segoe UI", 9, "bold"),
            anchor="w",
            width=17,
            padx=8,
            pady=7,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        for action, keys in hotkeys:
            row = tk.Frame(card, bg=self.colors["panel"])
            row.pack(fill=tk.X, padx=14, pady=3)

            tk.Label(
                row,
                text=action,
                bg=self.colors["panel"],
                fg=self.colors["muted"],
                font=("Segoe UI", 9),
                anchor="w",
                width=15,
                padx=8,
                pady=6,
                wraplength=95,
                justify=tk.LEFT,
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

            tk.Label(
                row,
                text=keys,
                bg=self.colors["entry"],
                fg=self.colors["text"],
                font=("Consolas", 9, "bold"),
                anchor="w",
                width=17,
                padx=8,
                pady=6,
            ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        warning = tk.Label(
            card,
            text="Одиночные клавиши убраны. Esc останавливает чтение только когда открыто окно приложения.",
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            justify=tk.LEFT,
            wraplength=230,
            anchor="w",
        )
        warning.pack(fill=tk.X, padx=14, pady=(14, 0))

    # ---------------------- Горячие клавиши ----------------------

    def setup_hotkeys(self):
        def safe_handler(handler):
            def wrapper():
                try:
                    handler()
                except Exception as e:
                    self.log(f"Ошибка: {str(e)}")
                    print(traceback.format_exc())
                    self.speak("Произошла ошибка. Подробности записаны в лог.", log_it=False)
            return wrapper

        try:
            keyboard.add_hotkey("x", safe_handler(self.read_browser_content))
            keyboard.add_hotkey("c", safe_handler(self.read_focused_element))
            keyboard.add_hotkey("v", safe_handler(self.read_element_under_cursor))
            keyboard.add_hotkey("q", safe_handler(self.stop_reading))
            keyboard.add_hotkey("w", safe_handler(self.read_current_url))
            keyboard.add_hotkey("e", safe_handler(self.read_hotkeys_help))
        except Exception as e:
            messagebox.showwarning(
                "Предупреждение",
                "Не удалось установить горячие клавиши.\n"
                "Попробуйте запустить программу от имени администратора.\n\n"
                f"Ошибка: {e}",
            )

    # ---------------------- Речь ----------------------

    def speak(self, text, log_it=True):
        text = self.clean_text(text)
        if not text:
            return

        if log_it:
            preview = text if len(text) <= 700 else text[:700] + "..."
            self.log(f"Диктор: {preview}")

        try:
            self.voice.Speak("", 3)   # очистить очередь
            self.voice.Speak(text, 1) # асинхронно произнести
        except Exception:
            pass

    def stop_reading(self):
        try:
            self.voice.Speak("", 3)
            self.log("Чтение остановлено")
        except Exception:
            pass

    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    # ---------------------- Безопасное извлечение свойств ----------------------

    def safe_get_name(self, elem):
        try:
            return elem.Name or ""
        except Exception:
            return ""

    def safe_get_control_type(self, elem):
        try:
            return elem.ControlTypeName or ""
        except Exception:
            return ""

    def safe_get_value(self, elem):
        try:
            val_pattern = elem.GetPattern(auto.ValuePattern)
            if val_pattern:
                return val_pattern.CurrentValue or ""
        except Exception:
            pass
        return ""

    def safe_get_rect(self, elem):
        try:
            rect = elem.BoundingRectangle
            left = int(getattr(rect, "left", getattr(rect, "Left", 0)))
            top = int(getattr(rect, "top", getattr(rect, "Top", 0)))
            right = int(getattr(rect, "right", getattr(rect, "Right", 0)))
            bottom = int(getattr(rect, "bottom", getattr(rect, "Bottom", 0)))
            return left, top, right, bottom
        except Exception:
            return None

    def rect_size(self, rect):
        if not rect:
            return 0, 0
        left, top, right, bottom = rect
        return max(0, right - left), max(0, bottom - top)

    def clean_text(self, text):
        if not text:
            return ""
        text = str(text)
        text = re.sub(r"\s+", " ", text)
        text = text.replace("\u200b", "")
        return text.strip()

    def is_probably_browser_ui_text(self, text):
        text = self.clean_text(text)
        if not text:
            return True

        lower = text.lower()

        # Слишком короткие технические фрагменты часто являются кнопками панели браузера.
        if lower in {"x", "×", "+", "-", "—", "…", "⋮"}:
            return True

        return any(word in lower for word in self.BROWSER_UI_WORDS)

    # ---------------------- Поиск браузера и страницы ----------------------

    def get_browser_window(self):
        for class_name in self.BROWSER_CLASSES:
            try:
                win = auto.Control(ClassName=class_name, Name=".*", depth=1)
                if win.Exists(1, 0.3):
                    return win
            except Exception:
                continue

        try:
            for win in auto.GetRootControl().GetChildren():
                try:
                    title = self.safe_get_name(win).lower()
                    class_name = getattr(win, "ClassName", "") or ""
                    if any(b in title for b in ["chrome", "firefox", "edge", "mozilla", "yandex", "браузер"]):
                        return win
                    if class_name in self.BROWSER_CLASSES:
                        return win
                except Exception:
                    continue
        except Exception:
            pass

        return None

    def iter_controls(self, root_control, max_depth=10):
        queue = deque([(root_control, 0)])
        visited = 0

        while queue and visited < self.MAX_TREE_NODES:
            ctrl, depth = queue.popleft()
            visited += 1
            yield ctrl, depth

            if depth >= max_depth:
                continue

            try:
                children = ctrl.GetChildren()
            except Exception:
                children = []

            for child in children:
                queue.append((child, depth + 1))

    def get_page_content_root(self, browser):
        """
        Возвращает корневой элемент области веб-страницы.

        Логика:
        1. Сначала ищем DocumentControl — это обычно DOM/страница.
        2. Если DocumentControl не найден, ищем крупную Pane/Custom/Group-область
           ниже панели браузера.
        3. Окно браузера целиком не читаем, чтобы не захватывать вкладки,
           адресную строку и кнопки браузера.
        """
        browser_rect = self.safe_get_rect(browser)
        browser_width, browser_height = self.rect_size(browser_rect)

        document_candidates = []
        fallback_candidates = []

        for ctrl, depth in self.iter_controls(browser, max_depth=9):
            control_type = self.safe_get_control_type(ctrl)
            if control_type not in self.PAGE_ROOT_CONTROL_TYPES:
                continue

            name = self.safe_get_name(ctrl)
            if name and self.is_probably_browser_ui_text(name):
                continue

            rect = self.safe_get_rect(ctrl)
            width, height = self.rect_size(rect)

            if width < 300 or height < 220:
                continue

            # Если есть координаты браузера, стараемся не брать верхнюю панель.
            if browser_rect and rect:
                _, browser_top, _, _ = browser_rect
                _, top, _, _ = rect
                if top < browser_top + 45 and control_type != "DocumentControl":
                    continue

            area = width * height
            score = area - depth * 25000

            if control_type == "DocumentControl":
                score += 1_000_000
                document_candidates.append((score, ctrl, control_type, name, rect))
            else:
                # Fallback: берем только достаточно крупную область, похожую на web-view.
                if browser_width and browser_height:
                    if width < browser_width * 0.45 or height < browser_height * 0.35:
                        continue
                fallback_candidates.append((score, ctrl, control_type, name, rect))

        candidates = document_candidates or fallback_candidates
        if not candidates:
            return None

        candidates.sort(key=lambda item: item[0], reverse=True)
        best = candidates[0]
        _, ctrl, control_type, name, rect = best
        self.log(f"Найдена область страницы: {control_type}; {name or 'без названия'}; rect={rect}")
        return ctrl

    # ---------------------- АДРЕС / ЗАГОЛОВОК ----------------------

    def get_browser_address(self):
        """
        В этой версии адрес берётся из заголовка окна, если браузер отдаёт его через UIA.
        Для чтения самой страницы эта функция не используется.
        """
        browser = self.get_browser_window()
        if not browser:
            return None

        try:
            title = self.safe_get_name(browser)
            if not title:
                return None

            url_match = re.search(r"https?://[^\s]+", title)
            if url_match:
                return url_match.group(0)

            www_match = re.search(r"www\.[^\s]+", title)
            if www_match:
                return f"https://{www_match.group(0)}"

            domain_match = re.search(r"([a-zA-Z0-9-]+\.[a-zA-Z]{2,})(?:/|$|\s)", title)
            if domain_match:
                return f"https://{domain_match.group(1)}"

            return title.strip()
        except Exception as e:
            self.log(f"Ошибка чтения адреса / заголовка: {e}")
            return None

    # ---------------------- Сбор текста ----------------------

    def extract_text_from_element(self, elem, max_depth=3, current_depth=0):
        if not elem or current_depth > max_depth:
            return None

        texts = []
        control_type = self.safe_get_control_type(elem)

        name = self.clean_text(self.safe_get_name(elem))
        if name and not self.is_probably_browser_ui_text(name):
            texts.append(name)

        value = self.clean_text(self.safe_get_value(elem))
        if value and value != name and not self.is_probably_browser_ui_text(value):
            texts.append(value)

        try:
            children = elem.GetChildren()
        except Exception:
            children = []

        for child in children:
            child_text = self.extract_text_from_element(child, max_depth, current_depth + 1)
            if child_text:
                texts.append(child_text)

        return self.merge_text_fragments(texts) if texts else None

    def merge_text_fragments(self, texts):
        result = []
        seen = set()

        for text in texts:
            text = self.clean_text(text)
            if not text:
                continue

            normalized = text.lower()
            if normalized in seen:
                continue

            # Убираем повтор, когда новый фрагмент почти полностью совпадает с предыдущим.
            if result:
                prev = result[-1]
                if text == prev or text in prev:
                    continue

            seen.add(normalized)
            result.append(text)

        return " ".join(result)

    def get_text_from_page_root(self, page_root):
        texts = []

        for ctrl, depth in self.iter_controls(page_root, max_depth=14):
            control_type = self.safe_get_control_type(ctrl)

            if control_type not in self.CONTENT_CONTROL_TYPES:
                continue

            # У контейнеров имя часто совпадает с заголовком страницы, а не с текстом.
            # Поэтому для DocumentControl берём имя только на первом уровне.
            if control_type == "DocumentControl" and depth > 0:
                continue

            name = self.clean_text(self.safe_get_name(ctrl))
            value = self.clean_text(self.safe_get_value(ctrl))

            for fragment in (name, value):
                if not fragment:
                    continue
                if self.is_probably_browser_ui_text(fragment):
                    continue

                # Слишком длинные технические склейки UIA часто являются контейнерами.
                # Реальный абзац оставляем, но гигантские агрегаты пропускаем.
                if len(fragment) > 1200 and control_type not in {"TextControl", "EditControl"}:
                    continue

                texts.append(fragment)

        return self.merge_text_fragments(texts)

    # ---------------------- Функции чтения ----------------------

    def read_browser_content(self):
        browser = self.get_browser_window()
        if not browser:
            self.speak("Браузер не найден. Откройте Chrome, Firefox, Edge или Яндекс Браузер.")
            return

        time.sleep(0.2)

        page_root = self.get_page_content_root(browser)
        if not page_root:
            self.speak(
                "Не удалось отделить содержимое страницы от интерфейса браузера. "
                "Откройте страницу, дождитесь загрузки и нажмите на область сайта."
            )
            return

        full_text = self.get_text_from_page_root(page_root)

        if not full_text:
            self.speak(
                "Текст на странице не найден. Возможно, сайт скрывает текст от системной доступности "
                "или страница ещё не загрузилась."
            )
            return

        if len(full_text) > self.MAX_READ_CHARS:
            full_text = (
                full_text[: self.MAX_READ_CHARS]
                + ". Текст страницы длинный, поэтому прочитана первая часть."
            )

        self.speak(full_text)

    def read_focused_element(self):
        try:
            focused = auto.GetFocusedControl()
        except Exception:
            self.speak("Ошибка получения фокуса.")
            return

        if not focused:
            self.speak("Нет элемента в фокусе.")
            return

        text = self.extract_text_from_element(focused, max_depth=2)
        if text:
            if len(text) > 700:
                text = text[:700] + ". Текст элемента длинный, прочитана первая часть."
            self.speak(text)
        else:
            type_name = self.safe_get_control_type(focused) or "элемент"
            self.speak(f"{type_name} в фокусе. Текста нет.")

    def read_element_under_cursor(self):
        try:
            x, y = auto.GetCursorPos()
            elem = auto.ControlFromPoint(x, y)
        except Exception:
            self.speak("Ошибка получения элемента под курсором.")
            return

        if not elem:
            self.speak("Ничего не найдено под курсором.")
            return

        text = self.extract_text_from_element(elem, max_depth=2)
        if text:
            if len(text) > 700:
                text = text[:700] + ". Текст элемента длинный, прочитана первая часть."
            self.speak(text)
        else:
            type_name = self.safe_get_control_type(elem) or "элемент"
            self.speak(f"{type_name}. Текста нет.")

    def read_hotkeys_help(self):
        help_text = (
            "Горячие клавиши. "
            "Контрол Альт один — читать содержимое страницы. "
            "Контрол Альт два — читать элемент в фокусе. "
            "Контрол Альт три — читать элемент под курсором мыши. "
            "Контрол Альт ноль — остановить чтение. "
            "Контрол Альт четыре — прочитать адрес или заголовок страницы. "
            "Контрол Альт пять — озвучить эту подсказку."
        )
        self.speak(help_text)

    def read_current_url(self):
        url = self.get_browser_address()
        if url:
            self.speak(f"Текущий адрес или заголовок страницы: {url}")
        else:
            self.speak("Не удалось определить адрес или заголовок страницы. Убедитесь, что открыт браузер.")

    # ---------------------- Настройки речи ----------------------

    def update_voices_list(self):
        voice_names = [v.GetDescription() for v in self.voices]
        self.voice_combo["values"] = voice_names
        if voice_names:
            self.voice_combo.current(0)
            self.current_voice_index = 0
            self.voice.Voice = self.voices[0]

    def change_voice(self, event=None):
        selected = self.voice_combo.get()
        for i, v in enumerate(self.voices):
            if v.GetDescription() == selected:
                self.voice.Voice = v
                self.current_voice_index = i
                self.log(f"Голос изменён на: {selected}")
                break

    def change_speed(self, val):
        self.rate = int(float(val))
        self.voice.Rate = self.rate
        self.speed_value_label.config(text=str(self.rate))

    def change_volume(self, val):
        self.volume = int(float(val))
        self.voice.Volume = self.volume
        self.volume_value_label.config(text=f"{self.volume}%")

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
        except Exception as e:
            self.log(f"Ошибка открытия адреса: {e}")

    # ---------------------- GUI ----------------------

    def create_widgets(self):
        main_frame = tk.Frame(self.root, bg=self.colors["bg"])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=22, pady=18)

        title = tk.Label(
            main_frame,
            text="Диктор веб-страниц",
            font=("Segoe UI", 24, "bold"),
            bg=self.colors["bg"],
            fg=self.colors["text"],
            anchor="w",
        )
        title.pack(fill=tk.X)

        subtitle = tk.Label(
            main_frame,
            text="Высококонтрастный интерфейс. Чтение страницы без адресной строки и панели браузера.",
            font=("Segoe UI", 11),
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            anchor="w",
        )
        subtitle.pack(fill=tk.X, pady=(2, 10))

        content_area = tk.Frame(main_frame, bg=self.colors["bg"])
        content_area.pack(fill=tk.BOTH, expand=True)

        left_column = tk.Frame(content_area, bg=self.colors["bg"])
        left_column.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        hotkeys_column = tk.Frame(content_area, bg=self.colors["bg"], width=280)
        hotkeys_column.pack(side=tk.RIGHT, fill=tk.Y, padx=(14, 0))
        hotkeys_column.pack_propagate(False)
        self.create_hotkeys_column(hotkeys_column)

        # Адрес
        url_card = self.create_card(left_column, "Открыть сайт")
        url_row = tk.Frame(url_card, bg=self.colors["panel"])
        url_row.pack(fill=tk.X)

        self.url_entry = tk.Entry(
            url_row,
            font=("Segoe UI", 13),
            bg=self.colors["entry"],
            fg=self.colors["text"],
            insertbackground=self.colors["text"],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
        )
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=10, padx=(0, 10))

        open_btn = self.make_button(url_row, "Открыть", self.open_url, "primary")
        open_btn.pack(side=tk.LEFT, padx=(0, 8))

        read_url_btn = self.make_button(url_row, "Прочитать адрес", self.read_current_url, "secondary")
        read_url_btn.pack(side=tk.LEFT)

        # Управление чтением
        control_card = self.create_card(left_column, "Чтение")
        top_buttons = tk.Frame(control_card, bg=self.colors["panel"])
        top_buttons.pack(fill=tk.X)

        read_page_btn = self.make_button(
            top_buttons,
            "Читать содержимое страницы",
            self.read_browser_content,
            "primary",
        )
        read_page_btn.pack(side=tk.LEFT, padx=(0, 8), pady=(0, 8))

        read_focus_btn = self.make_button(
            top_buttons,
            "Читать фокус",
            self.read_focused_element,
            "secondary",
        )
        read_focus_btn.pack(side=tk.LEFT, padx=(0, 8), pady=(0, 8))

        read_cursor_btn = self.make_button(
            top_buttons,
            "Читать под курсором",
            self.read_element_under_cursor,
            "secondary",
        )
        read_cursor_btn.pack(side=tk.LEFT, padx=(0, 8), pady=(0, 8))

        bottom_buttons = tk.Frame(control_card, bg=self.colors["panel"])
        bottom_buttons.pack(fill=tk.X, pady=(4, 0))

        stop_btn = self.make_button(bottom_buttons, "Остановить чтение", self.stop_reading, "danger")
        stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        hotkeys_help_btn = self.make_button(
            bottom_buttons,
            "Озвучить горячие клавиши",
            self.read_hotkeys_help,
            "secondary",
        )
        hotkeys_help_btn.pack(side=tk.LEFT)

        # Настройки речи
        settings_card = self.create_card(left_column, "Настройки голоса")

        speed_row = tk.Frame(settings_card, bg=self.colors["panel"])
        speed_row.pack(fill=tk.X, pady=(0, 10))
        tk.Label(
            speed_row,
            text="Скорость",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 11, "bold"),
            width=12,
            anchor="w",
        ).pack(side=tk.LEFT)

        self.speed_scale = tk.Scale(
            speed_row,
            from_=-10,
            to=10,
            orient=tk.HORIZONTAL,
            command=self.change_speed,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            troughcolor=self.colors["panel_2"],
            activebackground=self.colors["accent"],
            highlightthickness=0,
            length=330,
            showvalue=False,
        )
        self.speed_scale.pack(side=tk.LEFT, padx=(0, 12), fill=tk.X, expand=True)

        self.speed_value_label = tk.Label(
            speed_row,
            text="0",
            bg=self.colors["panel_2"],
            fg=self.colors["text"],
            font=("Segoe UI", 11, "bold"),
            width=5,
            pady=6,
        )
        self.speed_value_label.pack(side=tk.LEFT)

        volume_row = tk.Frame(settings_card, bg=self.colors["panel"])
        volume_row.pack(fill=tk.X, pady=(0, 10))
        tk.Label(
            volume_row,
            text="Громкость",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 11, "bold"),
            width=12,
            anchor="w",
        ).pack(side=tk.LEFT)

        self.volume_scale = tk.Scale(
            volume_row,
            from_=0,
            to=100,
            orient=tk.HORIZONTAL,
            command=self.change_volume,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            troughcolor=self.colors["panel_2"],
            activebackground=self.colors["accent"],
            highlightthickness=0,
            length=330,
            showvalue=False,
        )
        self.volume_scale.pack(side=tk.LEFT, padx=(0, 12), fill=tk.X, expand=True)

        self.volume_value_label = tk.Label(
            volume_row,
            text="100%",
            bg=self.colors["panel_2"],
            fg=self.colors["text"],
            font=("Segoe UI", 11, "bold"),
            width=5,
            pady=6,
        )
        self.volume_value_label.pack(side=tk.LEFT)

        voice_row = tk.Frame(settings_card, bg=self.colors["panel"])
        voice_row.pack(fill=tk.X)
        tk.Label(
            voice_row,
            text="Голос",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 11, "bold"),
            width=12,
            anchor="w",
        ).pack(side=tk.LEFT)

        self.voice_combo = ttk.Combobox(voice_row, state="readonly", width=58, font=("Segoe UI", 11))
        self.voice_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.voice_combo.bind("<<ComboboxSelected>>", self.change_voice)

        # Лог
        log_card_outer = tk.Frame(
            left_column,
            bg=self.colors["panel"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
            bd=0,
        )
        log_card_outer.pack(fill=tk.BOTH, expand=True, pady=8)

        log_header = tk.Label(
            log_card_outer,
            text="Журнал",
            bg=self.colors["panel"],
            fg=self.colors["text"],
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        )
        log_header.pack(fill=tk.X, padx=16, pady=(14, 6))

        log_body = tk.Frame(log_card_outer, bg=self.colors["panel"])
        log_body.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))

        self.log_text = tk.Text(
            log_body,
            bg=self.colors["entry"],
            fg=self.colors["muted"],
            insertbackground=self.colors["text"],
            font=("Consolas", 10),
            wrap=tk.WORD,
            relief=tk.FLAT,
            bd=0,
            height=9,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
        )

        scrollbar = tk.Scrollbar(log_body, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        hint = tk.Label(
            left_column,
            text="Подсказка: Ctrl+Alt+1 читает страницу, Ctrl+Alt+0 останавливает, Ctrl+Alt+5 озвучивает все горячие клавиши.",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=("Segoe UI", 10),
            anchor="w",
        )
        hint.pack(fill=tk.X, pady=(4, 0))

    def on_close(self):
        try:
            self.stop_reading()
        except Exception:
            pass

        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass

        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenReaderApp(root)
    root.mainloop()
