import tkinter as tk
from tkinter import ttk, messagebox
import time
import webbrowser
import uiautomation as auto
import keyboard
import win32com.client
import re
import traceback

class ScreenReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Экранный диктор для слабовидящих")
        self.root.geometry("700x680")
        self.root.minsize(600, 500)

        self.bg_color = "#1e1e1e"
        self.fg_color = "#f0f0f0"
        self.button_bg = "#3c3f41"
        self.button_active_bg = "#5a5e62"
        self.entry_bg = "#2d2d2d"
        self.root.configure(bg=self.bg_color)

        # Инициализация SAPI5
        self.voice = win32com.client.Dispatch("SAPI.SpVoice")
        self.voices = self.voice.GetVoices()
        self.current_voice_index = 0
        self.rate = 0
        self.volume = 100

        self.setup_hotkeys()
        self.create_widgets()

        self.update_voices_list()
        self.speed_scale.set(0)
        self.volume_scale.set(100)

        self.log("Диктор запущен. Горячие клавиши активны.\n"
                 "Ctrl+Shift+R – читать страницу\n"
                 "Ctrl+Shift+F – элемент в фокусе\n"
                 "X или Ctrl+Shift+W – читать под курсором\n"
                 "Ctrl+Shift+S – остановить чтение\n"
                 "Ctrl+Shift+A – прочитать адресную строку")

    # ---------------------- Горячие клавиши ----------------------
    def setup_hotkeys(self):
        def safe_handler(handler):
            def wrapper():
                try:
                    handler()
                except Exception as e:
                    self.log(f"Ошибка: {str(e)}")
                    print(traceback.format_exc())
                    self.speak("Произошла ошибка, смотрите лог", log_it=False)
            return wrapper

        try:
            keyboard.add_hotkey('ctrl+shift+r', safe_handler(self.read_browser_content))
            keyboard.add_hotkey('ctrl+shift+f', safe_handler(self.read_focused_element))
            keyboard.add_hotkey('ctrl+shift+w', safe_handler(self.read_element_under_cursor))
            keyboard.add_hotkey('ctrl+shift+x', safe_handler(self.read_element_under_cursor))
            keyboard.add_hotkey('x', safe_handler(self.read_element_under_cursor))
            keyboard.add_hotkey('ctrl+shift+s', safe_handler(self.stop_reading))
            keyboard.add_hotkey('ctrl+shift+a', safe_handler(self.read_current_url))
        except Exception as e:
            messagebox.showwarning("Предупреждение", f"Не удалось установить горячие клавиши.\nЗапустите от имени администратора.\nОшибка: {e}")

    # ---------------------- Речь ----------------------
    def speak(self, text, log_it=True):
        if not text or not text.strip():
            return
        if log_it:
            self.log(f"Диктор: {text}")
        try:
            self.voice.Speak("", 3)
            self.voice.Speak(text, 1)
        except:
            pass

    def stop_reading(self):
        try:
            self.voice.Speak("", 3)
            self.log("Чтение остановлено")
        except:
            pass

    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    # ---------------------- Безопасное извлечение свойств ----------------------
    def safe_get_name(self, elem):
        try:
            return elem.Name or ""
        except:
            return ""

    def safe_get_control_type(self, elem):
        try:
            return elem.ControlTypeName or ""
        except:
            return ""

    def safe_get_value(self, elem):
        try:
            val_pattern = elem.GetPattern(auto.ValuePattern)
            if val_pattern:
                return val_pattern.CurrentValue or ""
        except:
            pass
        return ""

    def extract_text_from_element(self, elem, max_depth=3, current_depth=0):
        if not elem or current_depth > max_depth:
            return None
        texts = []
        name = self.safe_get_name(elem)
        if name:
            texts.append(name)
        value = self.safe_get_value(elem)
        if value:
            texts.append(value)
        if self.safe_get_control_type(elem) == 'TextControl' and name:
            texts.append(name)
        try:
            children = elem.GetChildren()
        except:
            children = []
        for child in children:
            child_text = self.extract_text_from_element(child, max_depth, current_depth+1)
            if child_text:
                texts.append(child_text)
        unique = []
        for t in texts:
            if t and t not in unique:
                unique.append(t)
        return " ".join(unique) if unique else None

    # ---------------------- Поиск браузера ----------------------
    def get_browser_window(self):
        browser_classes = [
            'Chrome_WidgetWin_1',
            'MozillaWindowClass',
            'ApplicationFrameWindow',
            'YandexBrowserWidgetWin'
        ]
        for class_name in browser_classes:
            try:
                win = auto.Control(ClassName=class_name, Name='.*', depth=1)
                if win.Exists(1, 0.3):
                    return win
            except:
                continue
        try:
            for win in auto.GetRootControl().GetChildren():
                try:
                    title = win.Name.lower() if win.Name else ""
                    if any(b in title for b in ['chrome', 'firefox', 'edge', 'mozilla', 'yandex', 'браузер']):
                        return win
                except:
                    continue
        except:
            pass
        return None

    # ---------------------- АДРЕСНАЯ СТРОКА (работает через заголовок окна) ----------------------
    def get_browser_address(self):
        """
        Читает адрес из заголовка окна браузера.
        Это работает ВСЕГДА во всех браузерах!
        """
        browser = self.get_browser_window()
        if not browser:
            return None
        
        try:
            title = self.safe_get_name(browser)
            self.log(f"Заголовок окна: {title}")  # для отладки
            
            if not title:
                return None
            
            # Способ 1: ищем прямой URL в заголовке
            url_match = re.search(r'https?://[^\s]+', title)
            if url_match:
                return url_match.group(1)
            
            # Способ 2: ищем www.домен
            www_match = re.search(r'www\.[^\s]+', title)
            if www_match:
                return f"https://{www_match.group(1)}"
            
            # Способ 3: ищем любой домен (site.com, site.ru и т.д.)
            domain_match = re.search(r'([a-zA-Z0-9-]+\.[a-zA-Z]{2,})(?:/|$|\s)', title)
            if domain_match:
                return f"https://{domain_match.group(1)}"
            
            # Способ 4: берём всё до разделителя (часто там название сайта)
            for separator in [' - ', ' – ', ' | ', ' — ']:
                if separator in title:
                    return title.split(separator)[0].strip()
            
            # Если ничего не нашли, возвращаем заголовок целиком
            return title
            
        except Exception as e:
            self.log(f"Ошибка чтения адреса: {e}")
            return None

    # ---------------------- Сбор текста страницы ----------------------
    def get_all_text_from_window(self, control):
        texts = []
        if not control:
            return texts
        def walk(ctrl):
            try:
                type_name = self.safe_get_control_type(ctrl)
                if type_name in ['TextControl', 'EditControl', 'DocumentControl',
                                 'ButtonControl', 'HyperlinkControl', 'ListItemControl']:
                    text = self.safe_get_name(ctrl)
                    if text and text.strip():
                        texts.append(text.strip())
                for child in ctrl.GetChildren():
                    walk(child)
            except:
                pass
        walk(control)
        unique = []
        prev = ""
        for t in texts:
            if t != prev:
                unique.append(t)
            prev = t
        return unique

    # ---------------------- Функции чтения ----------------------
    def read_browser_content(self):
        browser = self.get_browser_window()
        if not browser:
            self.speak("Браузер не найден. Откройте Chrome, Firefox, Edge или Яндекс Браузер.")
            return
        time.sleep(0.2)
        texts = self.get_all_text_from_window(browser)
        if not texts:
            self.speak("Не удалось найти текст на странице.")
            return
        full_text = ". ".join(texts)
        if len(full_text) > 3000:
            full_text = full_text[:3000] + "..."
        self.speak(full_text)

    def read_focused_element(self):
        try:
            focused = auto.GetFocusedControl()
        except:
            self.speak("Ошибка получения фокуса")
            return
        if not focused:
            self.speak("Нет элемента в фокусе")
            return
        text = self.extract_text_from_element(focused, max_depth=2)
        if text:
            if len(text) > 500:
                text = text[:500] + "..."
            self.speak(text)
        else:
            type_name = self.safe_get_control_type(focused) or "элемент"
            self.speak(f"{type_name} в фокусе (нет текста)")

    def read_element_under_cursor(self):
        try:
            x, y = auto.GetCursorPos()
            elem = auto.ControlFromPoint(x, y)
        except:
            self.speak("Ошибка получения элемента под курсором")
            return
        if not elem:
            self.speak("Ничего не найдено под курсором")
            return
        text = self.extract_text_from_element(elem, max_depth=2)
        if text:
            if len(text) > 500:
                text = text[:500] + "..."
            self.speak(text)
        else:
            type_name = self.safe_get_control_type(elem) or "элемент"
            self.speak(f"{type_name} (не содержит текста)")

    def read_current_url(self):
        """Читает адрес текущей страницы"""
        url = self.get_browser_address()
        if url:
            self.speak(f"Адрес текущей страницы: {url}")
        else:
            self.speak("Не удалось определить адрес страницы. Убедитесь, что открыт браузер.")

    # ---------------------- Интерфейс и настройки ----------------------
    def update_voices_list(self):
        voice_names = [v.GetDescription() for v in self.voices]
        self.voice_combo['values'] = voice_names
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
        self.speed_label.config(text=f"Скорость: {self.rate}")

    def change_volume(self, val):
        self.volume = int(float(val))
        self.voice.Volume = self.volume
        self.volume_label.config(text=f"Громкость: {self.volume}%")

    def open_url(self):
        url = self.url_entry.get().strip()
        if not url:
            self.log("Введите адрес")
            return
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
        try:
            webbrowser.open(url)
            self.log(f"Открыт адрес: {url}")
        except Exception as e:
            self.log(f"Ошибка: {e}")

    # ---------------------- GUI ----------------------
    def create_widgets(self):
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        title_label = tk.Label(main_frame, text="Экранный диктор для слабовидящих",
                               font=("Arial", 18, "bold"), bg=self.bg_color, fg=self.fg_color)
        title_label.pack(pady=(0,10))

        # Адресная строка
        url_frame = tk.LabelFrame(main_frame, text="Адресная строка", font=("Arial", 12),
                                  bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        url_frame.pack(fill=tk.X, pady=5)
        self.url_entry = tk.Entry(url_frame, font=("Arial", 12), bg=self.entry_bg, fg=self.fg_color,
                                  insertbackground='white')
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10,5), pady=10)
        open_btn = tk.Button(url_frame, text="Открыть в браузере", command=self.open_url,
                             bg=self.button_bg, fg=self.fg_color, font=("Arial", 10),
                             activebackground=self.button_active_bg)
        open_btn.pack(side=tk.RIGHT, padx=5, pady=10)
        read_url_btn = tk.Button(url_frame, text="🔗 Прочитать адрес", command=self.read_current_url,
                                 bg=self.button_bg, fg=self.fg_color, font=("Arial", 10),
                                 activebackground=self.button_active_bg)
        read_url_btn.pack(side=tk.RIGHT, padx=5, pady=10)

        # Управление чтением
        control_frame = tk.LabelFrame(main_frame, text="Управление чтением", font=("Arial", 12),
                                      bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        control_frame.pack(fill=tk.X, pady=5)

        btn_style = {"bg": self.button_bg, "fg": self.fg_color, "font": ("Arial", 11),
                     "activebackground": self.button_active_bg, "padx": 10, "pady": 5}
        btn_row1 = tk.Frame(control_frame, bg=self.bg_color)
        btn_row1.pack(pady=5)
        tk.Button(btn_row1, text="📖 Читать всю страницу", command=self.read_browser_content, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_row1, text="🎯 Читать фокус", command=self.read_focused_element, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_row1, text="🖱️ Читать под курсором (X)", command=self.read_element_under_cursor, **btn_style).pack(side=tk.LEFT, padx=5)

        btn_row2 = tk.Frame(control_frame, bg=self.bg_color)
        btn_row2.pack(pady=5)
        tk.Button(btn_row2, text="⏹️ Остановить чтение", command=self.stop_reading,
                  bg="#a13e3e", fg="white", font=("Arial", 11, "bold"),
                  activebackground="#8b2c2c", padx=10, pady=5).pack(side=tk.LEFT, padx=5)

        # Настройки речи
        settings_frame = tk.LabelFrame(main_frame, text="Настройки речи", font=("Arial", 12),
                                       bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        settings_frame.pack(fill=tk.X, pady=5)

        speed_frame = tk.Frame(settings_frame, bg=self.bg_color)
        speed_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(speed_frame, text="Скорость:", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10)).pack(side=tk.LEFT)
        self.speed_scale = tk.Scale(speed_frame, from_=-10, to=10, orient=tk.HORIZONTAL,
                                    command=self.change_speed, bg=self.bg_color, fg=self.fg_color,
                                    highlightthickness=0, length=250)
        self.speed_scale.pack(side=tk.LEFT, padx=10)
        self.speed_label = tk.Label(speed_frame, text="Скорость: 0", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10))
        self.speed_label.pack(side=tk.LEFT)

        vol_frame = tk.Frame(settings_frame, bg=self.bg_color)
        vol_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(vol_frame, text="Громкость:", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10)).pack(side=tk.LEFT)
        self.volume_scale = tk.Scale(vol_frame, from_=0, to=100, orient=tk.HORIZONTAL,
                                     command=self.change_volume, bg=self.bg_color, fg=self.fg_color,
                                     highlightthickness=0, length=250)
        self.volume_scale.pack(side=tk.LEFT, padx=10)
        self.volume_label = tk.Label(vol_frame, text="Громкость: 100%", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10))
        self.volume_label.pack(side=tk.LEFT)

        voice_frame = tk.Frame(settings_frame, bg=self.bg_color)
        voice_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(voice_frame, text="Голос:", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10)).pack(side=tk.LEFT)
        self.voice_combo = ttk.Combobox(voice_frame, state="readonly", width=40)
        self.voice_combo.pack(side=tk.LEFT, padx=10)
        self.voice_combo.bind("<<ComboboxSelected>>", self.change_voice)

        # Лог
        log_frame = tk.LabelFrame(main_frame, text="Лог произнесённого", font=("Arial", 12),
                                  bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = tk.Text(log_frame, bg=self.entry_bg, fg=self.fg_color, font=("Consolas", 10),
                                wrap=tk.WORD, height=8)
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Подсказка
        hint_label = tk.Label(main_frame, text="Горячие клавиши: Ctrl+Shift+R (страница) | F (фокус) | X (под курсором) | A (адрес) | S (стоп)",
                              bg=self.bg_color, fg="#aaaaaa", font=("Arial", 9))
        hint_label.pack(pady=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenReaderApp(root)
    root.mainloop()