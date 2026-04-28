import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import re
import webbrowser
import uiautomation as auto
import pyttsx3
import keyboard

# ---------------------- Класс экранного диктора с GUI ----------------------
class ScreenReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Экранный диктор для слабовидящих")
        self.root.geometry("700x600")
        self.root.minsize(600, 500)

        # Цветовая схема (тёмная тема, высокий контраст)
        self.bg_color = "#1e1e1e"
        self.fg_color = "#f0f0f0"
        self.button_bg = "#3c3f41"
        self.button_active_bg = "#5a5e62"
        self.entry_bg = "#2d2d2d"
        self.root.configure(bg=self.bg_color)

        # Инициализация синтезатора речи
        self.engine = pyttsx3.init()
        self.stop_flag = False

        # Регистрация глобальных горячих клавиш
        self.setup_hotkeys()

        # Построение интерфейса
        self.create_widgets()

        # Стартовый вывод в лог
        self.log("Диктор запущен. Горячие клавиши активны.\nCtrl+Shift+R – читать страницу\nCtrl+Shift+F – элемент в фокусе\nCtrl+Shift+W – элемент под мышью\nCtrl+Shift+S – остановить чтение")

        # Обновление списка голосов
        self.update_voices_list()
        # Загрузка начальных настроек скорости и громкости
        self.speed_scale.set(self.engine.getProperty('rate'))
        self.volume_scale.set(self.engine.getProperty('volume'))

    # ---------------------- Настройка горячих клавиш ----------------------
    def setup_hotkeys(self):
        try:
            keyboard.add_hotkey('ctrl+shift+r', self.read_browser_content)
            keyboard.add_hotkey('ctrl+shift+f', self.read_focused_element)
            keyboard.add_hotkey('ctrl+shift+w', self.read_element_under_cursor)
            keyboard.add_hotkey('ctrl+shift+s', self.stop_reading)
        except Exception as e:
            messagebox.showwarning("Предупреждение", f"Не удалось установить глобальные горячие клавиши.\nВозможно, нужны права администратора.\nОшибка: {e}")

    # ---------------------- Функции речи ----------------------
    def speak(self, text, log_it=True):
        """Озвучивание текста с возможностью логирования."""
        if not text or not text.strip():
            return
        if log_it:
            self.log(f"Диктор: {text}")
        self.stop_flag = False
        def _speak():
            self.engine.say(text)
            while self.engine.isSpeaking() and not self.stop_flag:
                time.sleep(0.05)
            if self.stop_flag:
                self.engine.stop()
        thread = threading.Thread(target=_speak, daemon=True)
        thread.start()

    def stop_reading(self):
        """Остановить текущее чтение."""
        self.stop_flag = True
        self.engine.stop()
        self.log("Чтение остановлено пользователем.")

    def log(self, message):
        """Вывод сообщения в текстовое поле лога."""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    # ---------------------- Работа с браузером (UIAutomation) ----------------------
    def get_browser_window(self):
        """Находит окно браузера (Chrome, Firefox, Edge)."""
        browser_classes = ['Chrome_WidgetWin_1', 'MozillaWindowClass', 'ApplicationFrameWindow']
        for class_name in browser_classes:
            win = auto.Control(ClassName=class_name, Name='.*', depth=1)
            if win.Exists(3, 0.5):
                return win
        # Если по классу не нашли, ищем по заголовку
        for win in auto.GetRootControl().GetChildren():
            title = win.Name
            if title and any(b in title.lower() for b in ['chrome', 'firefox', 'edge', 'mozilla', 'браузер']):
                return win
        return None

    def get_browser_address(self):
        """Извлекает текущий URL из адресной строки браузера."""
        browser = self.get_browser_window()
        if not browser:
            return None
        # Поиск адресной строки: обычно это EditControl с определёнными именами
        address_edit = None
        # Рекурсивный поиск
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
            # Пытаемся получить значение через ValuePattern
            value_pattern = address_edit.GetValuePattern()
            if value_pattern:
                return value_pattern.Value
            # Если не получилось, берём имя (иногда там сам URL)
            if address_edit.Name and ('http' in address_edit.Name or 'www' in address_edit.Name):
                return address_edit.Name
        return None

    def get_all_text_from_window(self, control):
        """Рекурсивно собирает текст из всех текстовых элементов окна."""
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

    def read_browser_content(self):
        """Прочитать всё содержимое активного окна браузера."""
        browser = self.get_browser_window()
        if not browser:
            self.speak("Браузер не найден. Откройте Chrome, Firefox или Edge.")
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
        """Озвучить элемент, находящийся в фокусе (после Tab или клика)."""
        focused = auto.GetFocusedControl()
        if focused:
            text = focused.Name
            if not text:
                val = focused.GetValuePattern()
                if val:
                    text = val.Value
            if text and text.strip():
                self.speak(text)
            else:
                self.speak("Элемент не содержит текста")
        else:
            self.speak("Нет элемента в фокусе")

    def read_element_under_cursor(self):
        """Озвучить элемент под курсором мыши."""
        cursor_pos = auto.GetCursorPos()
        elem = auto.ControlFromPoint(cursor_pos)
        if elem:
            text = elem.Name
            if not text:
                val = elem.GetValuePattern()
                if val:
                    text = val.Value
            if text and text.strip():
                self.speak(text)
            else:
                self.speak("Под курсором нет текста")
        else:
            self.speak("Не удалось определить элемент под курсором")

    def read_current_url(self):
        """Прочитать адресную строку активного браузера."""
        url = self.get_browser_address()
        if url:
            self.speak(f"Адрес текущей страницы: {url}")
        else:
            self.speak("Не удалось получить адресную строку. Убедитесь, что открыт Chrome или Edge.")

    # ---------------------- Функции управления интерфейсом ----------------------
    def update_voices_list(self):
        """Обновить список доступных голосов в комбобоксе."""
        voices = self.engine.getProperty('voices')
        self.voice_combo['values'] = [v.name for v in voices]
        if voices:
            current = self.engine.getProperty('voice')
            for i, v in enumerate(voices):
                if v.id == current:
                    self.voice_combo.current(i)
                    break
            else:
                self.voice_combo.current(0)

    def change_voice(self, event=None):
        """Сменить голос."""
        selected = self.voice_combo.get()
        voices = self.engine.getProperty('voices')
        for v in voices:
            if v.name == selected:
                self.engine.setProperty('voice', v.id)
                self.log(f"Голос изменён на: {selected}")
                break

    def change_speed(self, val):
        """Изменить скорость речи."""
        speed = int(float(val))
        self.engine.setProperty('rate', speed)
        self.speed_label.config(text=f"Скорость: {speed}")

    def change_volume(self, val):
        """Изменить громкость."""
        vol = float(val)
        self.engine.setProperty('volume', vol)
        self.volume_label.config(text=f"Громкость: {int(vol*100)}%")

    def open_url(self):
        """Открыть введённый URL в браузере по умолчанию."""
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

    # ---------------------- Построение интерфейса ----------------------
    def create_widgets(self):
        # Основной контейнер с отступами
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # --- Заголовок ---
        title_label = tk.Label(main_frame, text="Экранный диктор для слабовидящих", 
                               font=("Arial", 18, "bold"), bg=self.bg_color, fg=self.fg_color)
        title_label.pack(pady=(0,15))

        # --- Секция адресной строки ---
        url_frame = tk.LabelFrame(main_frame, text="Адресная строка", font=("Arial", 12),
                                  bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        url_frame.pack(fill=tk.X, pady=10)
        # Поле ввода URL
        self.url_entry = tk.Entry(url_frame, font=("Arial", 12), bg=self.entry_bg, fg=self.fg_color,
                                  insertbackground='white')
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10,5), pady=10)
        # Кнопка открыть
        open_btn = tk.Button(url_frame, text="Открыть в браузере", command=self.open_url,
                             bg=self.button_bg, fg=self.fg_color, font=("Arial", 10),
                             activebackground=self.button_active_bg)
        open_btn.pack(side=tk.RIGHT, padx=5, pady=10)
        # Кнопка прочитать текущий URL
        read_url_btn = tk.Button(url_frame, text="Прочитать адресную строку", command=self.read_current_url,
                                 bg=self.button_bg, fg=self.fg_color, font=("Arial", 10),
                                 activebackground=self.button_active_bg)
        read_url_btn.pack(side=tk.RIGHT, padx=5, pady=10)

        # --- Секция управления чтением (кнопки) ---
        control_frame = tk.LabelFrame(main_frame, text="Управление чтением", font=("Arial", 12),
                                      bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        control_frame.pack(fill=tk.X, pady=10)

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

        # --- Настройки речи ---
        settings_frame = tk.LabelFrame(main_frame, text="Настройки речи", font=("Arial", 12),
                                       bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        settings_frame.pack(fill=tk.X, pady=10)

        # Скорость речи
        speed_frame = tk.Frame(settings_frame, bg=self.bg_color)
        speed_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(speed_frame, text="Скорость:", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10)).pack(side=tk.LEFT)
        self.speed_scale = tk.Scale(speed_frame, from_=50, to=400, orient=tk.HORIZONTAL,
                                    command=self.change_speed, bg=self.bg_color, fg=self.fg_color,
                                    highlightthickness=0, length=250)
        self.speed_scale.pack(side=tk.LEFT, padx=10)
        self.speed_label = tk.Label(speed_frame, text="Скорость: 150", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10))
        self.speed_label.pack(side=tk.LEFT)

        # Громкость
        vol_frame = tk.Frame(settings_frame, bg=self.bg_color)
        vol_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(vol_frame, text="Громкость:", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10)).pack(side=tk.LEFT)
        self.volume_scale = tk.Scale(vol_frame, from_=0.0, to=1.0, resolution=0.05, orient=tk.HORIZONTAL,
                                     command=self.change_volume, bg=self.bg_color, fg=self.fg_color,
                                     highlightthickness=0, length=250)
        self.volume_scale.pack(side=tk.LEFT, padx=10)
        self.volume_label = tk.Label(vol_frame, text="Громкость: 100%", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10))
        self.volume_label.pack(side=tk.LEFT)

        # Голос
        voice_frame = tk.Frame(settings_frame, bg=self.bg_color)
        voice_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(voice_frame, text="Голос:", bg=self.bg_color, fg=self.fg_color, font=("Arial", 10)).pack(side=tk.LEFT)
        self.voice_combo = ttk.Combobox(voice_frame, state="readonly", width=40)
        self.voice_combo.pack(side=tk.LEFT, padx=10)
        self.voice_combo.bind("<<ComboboxSelected>>", self.change_voice)

        # --- Лог событий ---
        log_frame = tk.LabelFrame(main_frame, text="Лог произнесённого", font=("Arial", 12),
                                  bg=self.bg_color, fg=self.fg_color, bd=2, relief=tk.GROOVE)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.log_text = tk.Text(log_frame, bg=self.entry_bg, fg=self.fg_color, font=("Consolas", 10),
                                wrap=tk.WORD, height=10)
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Подсказка по горячим клавишам ---
        hint_label = tk.Label(main_frame, text="Горячие клавиши: Ctrl+Shift+R (страница) | F (фокус) | W (мышь) | S (стоп)",
                              bg=self.bg_color, fg="#aaaaaa", font=("Arial", 9))
        hint_label.pack(pady=5)

# ---------------------- Запуск приложения ----------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenReaderApp(root)
    root.mainloop()