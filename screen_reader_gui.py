import tkinter as tk
from tkinter import ttk, messagebox
import time
import webbrowser
import uiautomation as auto
import keyboard
import win32com.client

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

        self.log("Диктор запущен (SAPI5). Горячие клавиши активны.\n"
                 "Ctrl+Shift+R – читать страницу\n"
                 "Ctrl+Shift+F – элемент в фокусе\n"
                 "Ctrl+Shift+W / X – читать под курсором\n"
                 "Ctrl+Shift+S – остановить чтение")

    # ---------------------- Горячие клавиши ----------------------
    def setup_hotkeys(self):
        try:
            keyboard.add_hotkey('ctrl+shift+r', self.read_browser_content)
            keyboard.add_hotkey('ctrl+shift+f', self.read_focused_element)
            keyboard.add_hotkey('ctrl+shift+w', self.read_element_under_cursor)
            keyboard.add_hotkey('ctrl+shift+x', self.read_element_under_cursor)  # удобно
            keyboard.add_hotkey('x', self.read_element_under_cursor)             # одиночная X
            keyboard.add_hotkey('ctrl+shift+s', self.stop_reading)
        except Exception as e:
            messagebox.showwarning("Предупреждение", f"Не удалось установить глобальные горячие клавиши.\nВозможно, нужны права администратора.\nОшибка: {e}")

    # ---------------------- Речь ----------------------
    def speak(self, text, log_it=True):
        if not text or not text.strip():
            return
        if log_it:
            self.log(f"Диктор: {text}")
        self.voice.Speak("", 3)   # остановить предыдущую речь
        self.voice.Speak(text, 1) # асинхронно

    def stop_reading(self):
        self.voice.Speak("", 3)
        self.log("Чтение остановлено пользователем.")

    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    # ---------------------- Вспомогательная: рекурсивное извлечение текста ----------------------
    def extract_text_from_element(self, elem, max_depth=3, current_depth=0):
        """Рекурсивно собирает текст из элемента и всех его потомков."""
        if not elem or current_depth > max_depth:
            return None
        texts = []
        # Пытаемся взять Name
        if elem.Name and elem.Name.strip():
            texts.append(elem.Name.strip())
        # Пытаемся взять Value через паттерн
        try:
            val_pattern = elem.GetPattern(auto.ValuePattern)
            if val_pattern and val_pattern.CurrentValue:
                texts.append(val_pattern.CurrentValue.strip())
        except:
            pass
        # Для текстовых элементов (например, параграфы в браузере)
        if elem.ControlTypeName == 'TextControl' and elem.Name:
            texts.append(elem.Name.strip())
        # Обходим детей
        for child in elem.GetChildren():
            child_text = self.extract_text_from_element(child, max_depth, current_depth+1)
            if child_text:
                texts.append(child_text)
        # Убираем дубликаты и пустые строки
        unique = []
        for t in texts:
            if t and t not in unique:
                unique.append(t)
        return " ".join(unique) if unique else None

    # ---------------------- Поиск браузера (с поддержкой Яндекс) ----------------------
    def get_browser_window(self):
        browser_classes = [
            'Chrome_WidgetWin_1',      # Chrome, Edge, Яндекс (обычно)
            'MozillaWindowClass',      # Firefox
            'ApplicationFrameWindow',  # Edge (старый)
            'YandexBrowserWidgetWin'   # Яндекс (на случай отличия)
        ]
        for class_name in browser_classes:
            win = auto.Control(ClassName=class_name, Name='.*', depth=1)
            if win.Exists(3, 0.5):
                return win
        # Если не нашли по классу, ищем по заголовку окна
        for win in auto.GetRootControl().GetChildren():
            title = win.Name.lower() if win.Name else ""
            if any(b in title for b in ['chrome', 'firefox', 'edge', 'mozilla', 'yandex', 'браузер']):
                return win
        return None

    # ---------------------- Адресная строка ----------------------
    def get_browser_address(self):
        browser = self.get_browser_window()
        if not browser:
            return None
        address_edit = None
        def find_address(ctrl):
            nonlocal address_edit
            if ctrl.ControlTypeName == 'EditControl':
                name = ctrl.Name.lower()
                if 'адрес' in name or 'address' in name or 'search' in name or 'url' in name:
                    address_edit = ctrl
                    return True
            for child in ctrl.GetChildren():
                if find_address(child):
                    return True
            return False
        find_address(browser)
        if address_edit:
            try:
                value_pattern = address_edit.GetPattern(auto.ValuePattern)
                if value_pattern:
                    return value_pattern.CurrentValue
            except:
                pass
            if address_edit.Name and ('http' in address_edit.Name or 'www' in address_edit.Name):
                return address_edit.Name
        return None

    # ---------------------- Сбор всего текста страницы ----------------------
    def get_all_text_from_window(self, control):
        texts = []
        if not control:
            return texts
        def walk(ctrl):
            if ctrl.ControlTypeName in ['TextControl', 'EditControl', 'DocumentControl',
                                        'ButtonControl', 'HyperlinkControl', 'ListItemControl']:
                text = ctrl.Name
                if text and text.strip():
                    texts.append(text.strip())
            for child in ctrl.GetChildren():
                walk(child)
        walk(control)
        # Убираем последовательные дубликаты
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
        focused = auto.GetFocusedControl()
        if not focused:
            self.speak("Нет элемента в фокусе")
            return
        text = self.extract_text_from_element(focused, max_depth=2)
        if text:
            if len(text) > 500:
                text = text[:500] + "..."
            self.speak(text)
        else:
            type_name = focused.ControlTypeName or "элемент"
            self.speak(f"{type_name} в фокусе (нет текста)")

    def read_element_under_cursor(self):
        x, y = auto.GetCursorPos()
        elem = auto.ControlFromPoint(x, y)
        if not elem:
            self.speak("Ничего не найдено под курсором")
            return
        text = self.extract_text_from_element(elem, max_depth=2)
        if text:
            if len(text) > 500:
                text = text[:500] + "..."
            self.speak(text)
        else:
            type_name = elem.ControlTypeName or "элемент"
            self.speak(f"{type_name} (не содержит текста)")

    def read_current_url(self):
        url = self.get_browser_address()
        if url:
            self.speak(f"Адрес текущей страницы: {url}")
        else:
            self.speak("Не удалось получить адресную строку. Убедитесь, что открыт браузер.")

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
            self.log("Введите адрес (например, https://example.com)")
            return
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
        try:
            webbrowser.open(url)
            self.log(f"Открыт адрес: {url}")
        except Exception as e:
            self.log(f"Ошибка открытия браузера: {e}")

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
        read_url_btn = tk.Button(url_frame, text="Прочитать адресную строку", command=self.read_current_url,
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
        tk.Button(btn_row1, text="🎯 Читать элемент в фокусе", command=self.read_focused_element, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_row1, text="🖱️ Читать под курсором", command=self.read_element_under_cursor, **btn_style).pack(side=tk.LEFT, padx=5)

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
        hint_label = tk.Label(main_frame, text="Горячие клавиши: Ctrl+Shift+R (страница) | F (фокус) | W/X (под курсором) | S (стоп)",
                              bg=self.bg_color, fg="#aaaaaa", font=("Arial", 9))
        hint_label.pack(pady=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenReaderApp(root)
    root.mainloop()