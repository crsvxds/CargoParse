import os
import random
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import subprocess

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

from core import process_customs_data, process_cleaner_mode
from config import Config


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        
        if event:
            x = event.x_root + 15
            y = event.y_root + 15
        else:
            x = self.widget.winfo_rootx() + 25
            y = self.widget.winfo_rooty() + 20
            
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw, text=self.text, justify=tk.LEFT, background="#22252a",
            foreground="#ffffff", relief=tk.SOLID, borderwidth=1, font=("Segoe UI", 9)
        ).pack(ipadx=6, ipady=3)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()


class CustomsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CargoParse V2 — Professional Suite")

        # Загружаем конфигурацию
        self.config = Config()

        # Оптимизированные габариты для экранов ноутбуков
        self.root.geometry("860x700")
        self.root.minsize(750, 580)

        # Флаг для отмены операции
        self.is_cancelled = False
        self.is_running = False

        self.setup_styles()

        self.funny_statuses = [
            "Разгружаем контейнер из Гуанчжоу...", "Проверяем таможенную декларацию...",
            "Подкупаем инспектора шоколадкой...", "Завариваем крепкий кофе...",
            "Пересчитываем коробки вручную...", "Ищем потерявшийся артикул под столом...",
            "Ждем, пока китайская сторона подпишет доки...", "Сортируем маркировку левой пяткой..."
        ]

        self.app_mode = tk.StringVar(value="verify")
        self.articles_path = tk.StringVar(value=self.config.get("last_articles_path", ""))
        self.codes_dir = tk.StringVar(value=self.config.get("last_codes_dir", ""))
        self.output_dir = tk.StringVar(value=self.config.get("last_output_dir", ""))
        self.invoice_prefix = tk.StringVar(value=self.config.get("last_invoice_prefix", ""))

        self.auto_open = tk.BooleanVar(value=self.config.get("auto_open_results", True))
        self.check_boxes_var = tk.BooleanVar(value=self.config.get("check_boxes", True))
        self.remove_brackets_var = tk.BooleanVar(value=self.config.get("remove_brackets", True))
        self.dry_run_var = tk.BooleanVar(value=False)  # Новая опция "Только проверка"
        self.dark_theme_var = tk.BooleanVar(value=self.config.get("dark_theme", False))

        self.stat_total = tk.StringVar(value="0")
        self.stat_ok = tk.StringVar(value="0")
        self.stat_bad = tk.StringVar(value="0")
        self.stat_dups = tk.StringVar(value="0")

        # Для анимации счетчиков
        self.stat_total_target = 0
        self.stat_ok_target = 0
        self.stat_bad_target = 0
        self.stat_dups_target = 0

        # Для отслеживания текущего файла
        self.current_file_label = None
        self.eta_label = None
        self.start_time = None

        self.create_widgets()
        self.setup_hotkeys()
        self.setup_drag_drop()

        # Применяем тему при старте
        if self.dark_theme_var.get():
            self.apply_dark_theme()

        # Создаем результирующую папку если не существует
        if self.output_dir.get() and not os.path.exists(self.output_dir.get()):
            try:
                os.makedirs(self.output_dir.get(), exist_ok=True)
            except:
                pass

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Современные цвета Material Design
        self.bg_color = "#f5f7fa"
        self.card_bg = "#ffffff"
        self.primary_color = "#5468ff"
        self.primary_hover = "#3d56f0"
        self.success_color = "#00c9a7"
        self.danger_color = "#ff5c7c"
        self.warning_color = "#ffbe0b"
        self.text_main = "#1e293b"
        self.text_secondary = "#64748b"
        self.border_color = "#e2e8f0"

        # Темная тема (будет применена позже если включена)
        self.dark_bg_color = "#0f172a"
        self.dark_card_bg = "#1e293b"
        self.dark_text_main = "#e2e8f0"
        self.dark_text_secondary = "#94a3b8"
        self.dark_border_color = "#334155"

        self.root.configure(bg=self.bg_color)

        self.style.configure(".", background=self.bg_color, foreground=self.text_main, font=("Segoe UI", 9))
        self.style.configure("TLabelframe", background=self.bg_color, bordercolor=self.border_color, borderwidth=1, relief="flat")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=self.primary_color, background=self.bg_color)

        # Кнопки с градиентом (эмуляция через цвета)
        self.style.configure("TButton", font=("Segoe UI", 9, "bold"), background=self.primary_color, foreground="#ffffff", borderwidth=0, focuscolor=self.primary_color, relief="flat")
        self.style.map("TButton", background=[('active', self.primary_hover), ('disabled', '#cbd5e1')], foreground=[('disabled', '#94a3b8')])

        self.style.configure("Preset.TButton", background="#475569", foreground="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat")
        self.style.map("Preset.TButton", background=[('active', '#334155')])

        self.style.configure("Cancel.TButton", background=self.danger_color, foreground="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat")
        self.style.map("Cancel.TButton", background=[('active', '#e63946')])

        self.style.configure("TEntry", fieldbackground="#ffffff", bordercolor=self.border_color, lightcolor=self.border_color, darkcolor=self.border_color, padding=5, relief="flat")
        self.style.map("TEntry", fieldbackground=[('disabled', '#f1f5f9')], foreground=[('disabled', '#94a3b8')])

        self.style.configure("TRadiobutton", font=("Segoe UI", 9), background=self.bg_color, focuscolor=self.bg_color)
        self.style.configure("TCheckbutton", font=("Segoe UI", 9), background=self.bg_color, focuscolor=self.bg_color)
        self.style.configure("Horizontal.TProgressbar", troughcolor="#e2e8f0", bordercolor="#e2e8f0", background=self.primary_color, lightcolor=self.primary_color, darkcolor=self.primary_color, thickness=8)

    def apply_dark_theme(self):
        """Применяет темную тему"""
        self.root.configure(bg=self.dark_bg_color)
        self.style.configure(".", background=self.dark_bg_color, foreground=self.dark_text_main)
        self.style.configure("TLabelframe", background=self.dark_bg_color, bordercolor=self.dark_border_color)
        self.style.configure("TLabelframe.Label", foreground=self.primary_color, background=self.dark_bg_color)
        self.style.configure("TRadiobutton", background=self.dark_bg_color)
        self.style.configure("TCheckbutton", background=self.dark_bg_color)
        self.log_text.configure(bg="#1a1d24", fg="#f8f8f2")

    def apply_light_theme(self):
        """Применяет светлую тему"""
        self.root.configure(bg=self.bg_color)
        self.style.configure(".", background=self.bg_color, foreground=self.text_main)
        self.style.configure("TLabelframe", background=self.bg_color, bordercolor=self.border_color)
        self.style.configure("TLabelframe.Label", foreground=self.primary_color, background=self.bg_color)
        self.style.configure("TRadiobutton", background=self.bg_color)
        self.style.configure("TCheckbutton", background=self.bg_color)
        self.log_text.configure(bg="#1e222b", fg="#f8f8f2")

    def toggle_theme(self):
        """Переключает тему"""
        is_dark = self.dark_theme_var.get()
        self.config.set("dark_theme", is_dark)
        if is_dark:
            self.apply_dark_theme()
        else:
            self.apply_light_theme()

    def setup_hotkeys(self):
        """Настройка горячих клавиш"""
        # F5 - Запустить обработку
        self.root.bind('<F5>', lambda e: self.start_processing_thread())

        # Ctrl+S - Сохранить лог
        self.root.bind('<Control-s>', lambda e: self.save_log())
        self.root.bind('<Control-S>', lambda e: self.save_log())

        # Ctrl+F - Поиск по логу
        self.root.bind('<Control-f>', lambda e: self.open_log_search())
        self.root.bind('<Control-F>', lambda e: self.open_log_search())

        # Ctrl+O - Открыть папку результатов
        self.root.bind('<Control-o>', lambda e: self.open_output_folder())
        self.root.bind('<Control-O>', lambda e: self.open_output_folder())

        # Ctrl+E - Редактор articles.txt
        self.root.bind('<Control-e>', lambda e: self.open_articles_editor())
        self.root.bind('<Control-E>', lambda e: self.open_articles_editor())

        # Escape - Отмена операции
        self.root.bind('<Escape>', lambda e: self.cancel_operation())

        self.log("[ℹ] Горячие клавиши: F5-Запуск | Ctrl+S-Сохранить | Ctrl+O-Открыть результаты | Ctrl+E-Редактор | Ctrl+F-Поиск | Esc-Отмена")

    def setup_drag_drop(self):
        """Настройка drag-and-drop для полей ввода"""
        if not HAS_DND:
            return

        try:
            # Для файла артикулов
            self.ent_art.drop_target_register(DND_FILES)
            self.ent_art.dnd_bind('<<Drop>>', lambda e: self.on_drop_file(e, self.articles_path))

            # Для папки с кодами
            self.ent_codes.drop_target_register(DND_FILES)
            self.ent_codes.dnd_bind('<<Drop>>', lambda e: self.on_drop_folder(e, self.codes_dir))

            # Для папки вывода
            self.ent_out.drop_target_register(DND_FILES)
            self.ent_out.dnd_bind('<<Drop>>', lambda e: self.on_drop_folder(e, self.output_dir))
        except:
            # Ошибка при настройке DnD
            pass

    def on_drop_file(self, event, target_var):
        """Обработчик drop для файлов"""
        file_path = event.data.strip('{}')
        if os.path.isfile(file_path):
            target_var.set(file_path)
            self.log(f"[ℹ] Файл добавлен: {os.path.basename(file_path)}")

    def on_drop_folder(self, event, target_var):
        """Обработчик drop для папок"""
        folder_path = event.data.strip('{}')
        if os.path.isdir(folder_path):
            target_var.set(folder_path)
            self.log(f"[ℹ] Папка добавлена: {os.path.basename(folder_path)}")
        elif os.path.isfile(folder_path):
            folder_path = os.path.dirname(folder_path)
            target_var.set(folder_path)
            self.log(f"[ℹ] Папка добавлена: {os.path.basename(folder_path)}")

    def open_articles_editor(self):
        """Открывает встроенный редактор для articles.txt"""
        articles_file = self.articles_path.get().strip()

        # Если файл не указан, предлагаем создать новый
        if not articles_file:
            result = messagebox.askyesno(
                "Файл не указан",
                "Файл articles.txt не указан.\nСоздать новый файл?"
            )
            if result:
                articles_file = filedialog.asksaveasfilename(
                    title="Создать articles.txt",
                    defaultextension=".txt",
                    filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
                    initialfile="articles.txt"
                )
                if articles_file:
                    self.articles_path.set(articles_file)
                    # Создаем пустой файл
                    with open(articles_file, "w", encoding="utf-8") as f:
                        f.write("")
                else:
                    return
            else:
                return

        # Создаем окно редактора
        editor_window = tk.Toplevel(self.root)
        editor_window.title(f"Редактор - {os.path.basename(articles_file)}")
        editor_window.geometry("900x600")

        # Создаем фрейм для инструкций
        info_frame = ttk.Frame(editor_window, padding="10")
        info_frame.pack(fill=tk.X)

        info_text = "Формат: [Артикул] [Наименование...] [Описание] [Грузовые места] [Количество] [Ед. изм]"
        info_label = ttk.Label(info_frame, text=info_text, font=("Segoe UI", 9), foreground="#718096")
        info_label.pack(side=tk.LEFT)

        # Счетчик строк
        self.line_count_var = tk.StringVar(value="Строк: 0")
        line_count_label = ttk.Label(info_frame, textvariable=self.line_count_var, font=("Segoe UI", 9, "bold"), foreground=self.primary_color)
        line_count_label.pack(side=tk.RIGHT)

        # Создаем текстовое поле с прокруткой
        text_frame = ttk.Frame(editor_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Нумерация строк
        line_numbers = tk.Text(text_frame, width=4, padx=5, takefocus=0, border=0, background="#f0f0f0", state='disabled', wrap='none', font=("Consolas", 10))
        line_numbers.pack(side=tk.LEFT, fill=tk.Y)

        text_editor = tk.Text(text_frame, wrap=tk.NONE, font=("Consolas", 10), undo=True, maxundo=-1)
        text_editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Вертикальная прокрутка
        v_scrollbar = ttk.Scrollbar(text_frame, command=text_editor.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_editor.configure(yscrollcommand=v_scrollbar.set)

        # Горизонтальная прокрутка
        h_scrollbar = ttk.Scrollbar(editor_window, orient=tk.HORIZONTAL, command=text_editor.xview)
        h_scrollbar.pack(fill=tk.X, padx=10, pady=(0, 10))
        text_editor.configure(xscrollcommand=h_scrollbar.set)

        # Функция обновления нумерации строк
        def update_line_numbers(event=None):
            line_numbers.configure(state='normal')
            line_numbers.delete('1.0', tk.END)

            # Получаем количество строк
            line_count = int(text_editor.index('end-1c').split('.')[0])
            line_numbers_string = "\n".join(str(i) for i in range(1, line_count))
            line_numbers.insert('1.0', line_numbers_string)
            line_numbers.configure(state='disabled')

            # Обновляем счетчик
            self.line_count_var.set(f"Строк: {line_count - 1}")

        # Синхронизация прокрутки нумерации со скроллом текста
        def on_scroll(*args):
            line_numbers.yview_moveto(args[0])
            v_scrollbar.set(*args)

        text_editor.configure(yscrollcommand=on_scroll)

        # Загружаем содержимое файла
        try:
            if os.path.exists(articles_file):
                with open(articles_file, "r", encoding="utf-8") as f:
                    content = f.read()
                    text_editor.insert("1.0", content)
                    update_line_numbers()
            else:
                messagebox.showwarning("Файл не найден", f"Файл {articles_file} не существует.\nБудет создан новый файл при сохранении.")
        except Exception as e:
            messagebox.showerror("Ошибка чтения", f"Не удалось прочитать файл:\n{e}")
            editor_window.destroy()
            return

        # Обновляем нумерацию при изменении текста
        text_editor.bind('<KeyRelease>', update_line_numbers)

        # Функция сохранения
        def save_file():
            try:
                content = text_editor.get("1.0", tk.END).rstrip('\n')
                with open(articles_file, "w", encoding="utf-8") as f:
                    f.write(content)
                messagebox.showinfo("Успешно", f"Файл сохранен:\n{articles_file}")
                self.log(f"[✓] Файл articles.txt сохранен: {line_count - 1} строк")
            except Exception as e:
                messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл:\n{e}")

        # Функция сохранения и закрытия
        def save_and_close():
            save_file()
            editor_window.destroy()

        # Кнопки управления
        btn_frame = ttk.Frame(editor_window, padding="10")
        btn_frame.pack(fill=tk.X)

        save_close_btn = ttk.Button(btn_frame, text="💾 Сохранить и закрыть", command=save_and_close, style="TButton")
        save_close_btn.pack(side=tk.LEFT, padx=(0, 10), ipadx=10, ipady=5)

        save_btn = ttk.Button(btn_frame, text="💾 Сохранить", command=save_file)
        save_btn.pack(side=tk.LEFT, padx=(0, 10), ipadx=10, ipady=5)

        cancel_btn = ttk.Button(btn_frame, text="❌ Закрыть без сохранения", command=editor_window.destroy)
        cancel_btn.pack(side=tk.LEFT, ipadx=10, ipady=5)

        # Горячие клавиши
        editor_window.bind('<Control-s>', lambda e: save_file())
        editor_window.bind('<Control-S>', lambda e: save_file())
        editor_window.bind('<Control-w>', lambda e: editor_window.destroy())
        editor_window.bind('<Control-W>', lambda e: editor_window.destroy())

        # Делаем окно модальным
        editor_window.transient(self.root)
        editor_window.grab_set()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. ПЕРЕКЛЮЧАТЕЛЬ РЕЖИМОВ
        mode_frame = ttk.LabelFrame(main_frame, text=" Режим работы приложения ", padding="6")
        mode_frame.pack(fill=tk.X, pady=(0, 6))
        
        rb1 = ttk.Radiobutton(mode_frame, text="📦  Верификация по артикулам (Сверка с articles.txt)", variable=self.app_mode, value="verify", command=self.toggle_mode_ui)
        rb1.pack(side=tk.LEFT, padx=(10, 20))
        
        rb2 = ttk.Radiobutton(mode_frame, text="⚡  Быстрая очистка файлов (Форматтер столбцов)", variable=self.app_mode, value="clean", command=self.toggle_mode_ui)
        rb2.pack(side=tk.LEFT, padx=10)

        # 2. ПАНЕЛЬ МОНИТОРИНГА (Уменьшенная высота карточек)
        self.dash_frame = ttk.LabelFrame(main_frame, text=" Панель мониторинга ", padding="6")
        self.dash_frame.pack(fill=tk.X, pady=(0, 6))
        self.dash_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="equal")

        self.dash_blocks = []
        stats_config = [
            ("Всего артикулов", self.stat_total, self.primary_color, "Обнаружено строк/файлов для обработки"),
            ("Успешно (OK)", self.stat_ok, self.success_color, "Данные сошлись и успешно обработаны"),
            ("Расхождения", self.stat_bad, self.danger_color, "Артикулы с ошибками или нехваткой кодов"),
            ("Исключено дублей", self.stat_dups, self.warning_color, "Повторяющиеся коды, отсеянные за сессию")
        ]

        for idx, (title, var, color, tip) in enumerate(stats_config):
            card = tk.Canvas(self.dash_frame, bg=self.card_bg, highlightthickness=1, highlightbackground=self.border_color, height=52)
            card.grid(row=0, column=idx, padx=4, sticky="nsew")
            
            lbl_t = tk.Label(card, text=title, font=("Segoe UI", 8, "bold"), fg="#718096", bg=self.card_bg)
            lbl_t.pack(pady=(4, 0))
            
            lbl_v = tk.Label(card, textvariable=var, font=("Segoe UI", 14, "bold"), fg=color, bg=self.card_bg)
            lbl_v.pack(pady=(0, 4))
            
            ToolTip(card, tip)
            self.dash_blocks.append(lbl_t)

        # 3. НАСТРОЙКИ ПУТЕЙ
        self.path_frame = ttk.LabelFrame(main_frame, text=" Настройки путей к данным ", padding="8")
        self.path_frame.pack(fill=tk.X, pady=(0, 6))
        self.path_frame.columnconfigure(1, weight=1)

        self.btn_preset = ttk.Button(self.path_frame, text="⚡  ЗАПОЛНИТЬ РАБОЧУЮ ПАПКУ (БЫСТРЫЙ НАБОР V2)", style="Preset.TButton", command=self.apply_quick_preset)
        self.btn_preset.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=(0, 6))

        # Строка 1: Файл артикулов
        self.lbl_art = ttk.Label(self.path_frame, text="Файл артикулов:")
        self.lbl_art.grid(row=1, column=0, sticky=tk.W, pady=3, padx=(0, 8))
        self.ent_art = ttk.Entry(self.path_frame, textvariable=self.articles_path)
        self.ent_art.grid(row=1, column=1, sticky=tk.EW, pady=3, padx=4)

        art_btn_frame = ttk.Frame(self.path_frame)
        art_btn_frame.grid(row=1, column=2, pady=3, padx=(4, 0))

        self.btn_art = ttk.Button(art_btn_frame, text="Обзор...", width=10, command=lambda: self.articles_path.set(filedialog.askopenfilename(initialdir=os.path.dirname(self.articles_path.get()) if self.articles_path.get() else os.path.expanduser("~")) or self.articles_path.get()))
        self.btn_art.pack(side=tk.LEFT, padx=(0, 2))

        self.btn_edit_art = ttk.Button(art_btn_frame, text="📝", width=3, command=self.open_articles_editor)
        self.btn_edit_art.pack(side=tk.LEFT)
        ToolTip(self.btn_edit_art, "Редактировать articles.txt")

        # Строка 2: Папка с кодами
        self.lbl_source = ttk.Label(self.path_frame, text="Папка с кодами:")
        self.lbl_source.grid(row=2, column=0, sticky=tk.W, pady=3, padx=(0, 8))
        self.ent_codes = ttk.Entry(self.path_frame, textvariable=self.codes_dir)
        self.ent_codes.grid(row=2, column=1, sticky=tk.EW, pady=3, padx=4)
        self.btn_codes = ttk.Button(self.path_frame, text="Обзор...", width=10, command=lambda: self.codes_dir.set(filedialog.askdirectory(initialdir=BASE_DIR) or self.codes_dir.get()))
        self.btn_codes.grid(row=2, column=2, pady=3, padx=(4, 0))

        # Строка 3: Папка вывода
        self.lbl_out = ttk.Label(self.path_frame, text="Папка назначения:")
        self.lbl_out.grid(row=3, column=0, sticky=tk.W, pady=3, padx=(0, 8))
        self.ent_out = ttk.Entry(self.path_frame, textvariable=self.output_dir)
        self.ent_out.grid(row=3, column=1, sticky=tk.EW, pady=3, padx=4)
        self.btn_out = ttk.Button(self.path_frame, text="Обзор...", width=10, command=lambda: self.output_dir.set(filedialog.askdirectory(initialdir=BASE_DIR) or self.output_dir.get()))
        self.btn_out.grid(row=3, column=2, pady=3, padx=(4, 0))

        # Строка 4: Инвойс
        self.lbl_inv = ttk.Label(self.path_frame, text="Номер инвойса (префикс):")
        self.lbl_inv.grid(row=4, column=0, sticky=tk.W, pady=3, padx=(0, 8))
        self.ent_inv = ttk.Entry(self.path_frame, textvariable=self.invoice_prefix)
        self.ent_inv.grid(row=4, column=1, sticky=tk.EW, pady=3, padx=4)
        ToolTip(self.ent_inv, "Например: UI776. Будет добавлен в начало имён файлов (UI776 УралТрейд.xlsx)")

        # Строка 5: Чекбоксы в одну строчку для экономии высоты
        cb_line = ttk.Frame(self.path_frame)
        cb_line.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(4, 0))

        self.box_chk = ttk.Checkbutton(cb_line, text="Учитывать коробки (УралТрейд)", variable=self.check_boxes_var)
        self.box_chk.pack(side=tk.LEFT, padx=(0, 15))
        ToolTip(self.box_chk, "Сверять штуки и количество коробок для Трейда.")

        self.bracket_chk = ttk.Checkbutton(cb_line, text="Удалять скобки (01), (00), (21)", variable=self.remove_brackets_var)
        self.bracket_chk.pack(side=tk.LEFT, padx=(0, 15))
        ToolTip(self.bracket_chk, "Опускает скобки вокруг (01), (00), (21) преобразовывая их в голые цифры.")

        self.dry_run_chk = ttk.Checkbutton(cb_line, text="🔍 Только проверка (без сохранения)", variable=self.dry_run_var)
        self.dry_run_chk.pack(side=tk.LEFT)
        ToolTip(self.dry_run_chk, "Выполнить проверку без создания файлов Excel (предпросмотр результатов)")

        # 4. ЖУРНАЛ РАБОТЫ (Автоматически сжимается и тянется)
        log_frame = ttk.LabelFrame(main_frame, text=" Подробный журнал работы ", padding="4")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        
        self.log_text = tk.Text(log_frame, height=5, width=80, state=tk.NORMAL, font=("Consolas", 9), bg="#1e222b", fg="#f8f8f2", insertbackground="#ffffff", relief=tk.FLAT, padx=6, pady=6)
        sb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Делаем лог только для чтения, но с возможностью выделения и копирования
        # Блокируем только ввод текста, но разрешаем Ctrl+C
        def block_input(event):
            # Разрешаем только Ctrl+C, Ctrl+A и выделение мышкой
            if event.keysym in ('c', 'C', 'a', 'A') and (event.state & 0x4):  # Ctrl нажат
                return None
            return "break"

        self.log_text.bind("<Key>", block_input)

        # 5. ИНДИКАТОРЫ ПРОГРЕССА
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 6))

        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode='determinate', style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=(0, 3))

        sub_frame = ttk.Frame(progress_frame)
        sub_frame.pack(fill=tk.X)
        self.status_label = ttk.Label(sub_frame, text="Система готова к запуску", font=("Segoe UI", 8, "italic"), foreground=self.text_secondary)
        self.status_label.pack(side=tk.LEFT)
        self.progress_label = ttk.Label(sub_frame, text="0%", font=("Segoe UI", 8, "bold"), foreground=self.primary_color, width=5, anchor=tk.E)
        self.progress_label.pack(side=tk.RIGHT)

        # Детальная информация о прогрессе
        detail_frame = ttk.Frame(progress_frame)
        detail_frame.pack(fill=tk.X, pady=(3, 0))

        self.current_file_label = ttk.Label(detail_frame, text="", font=("Segoe UI", 8), foreground=self.text_secondary)
        self.current_file_label.pack(side=tk.LEFT)

        self.eta_label = ttk.Label(detail_frame, text="", font=("Segoe UI", 8), foreground=self.text_secondary)
        self.eta_label.pack(side=tk.RIGHT)

        # 6. КНОПКИ (Гарантированно видны внизу окна)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)

        self.run_btn = ttk.Button(btn_frame, text="⚡ СТАРТ ОБРАБОТКИ", style="TButton", command=self.start_processing_thread)
        self.run_btn.pack(side=tk.LEFT, ipadx=12, ipady=4)

        self.cancel_btn = ttk.Button(btn_frame, text="✖ ОТМЕНА", style="Cancel.TButton", command=self.cancel_operation, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=10, ipadx=12, ipady=4)

        self.save_btn = tk.Button(btn_frame, text="💾 Сохранить лог", font=("Segoe UI", 9), bg="#e2e8f0", fg=self.text_main, borderwidth=0, activebackground="#cbd5e1", command=self.save_log)
        self.save_btn.pack(side=tk.LEFT, padx=10, ipadx=8, ipady=4)

        ttk.Checkbutton(btn_frame, text="Открыть папку результатов", variable=self.auto_open).pack(side=tk.LEFT, padx=10)

        ttk.Checkbutton(btn_frame, text="🌙 Темная тема", variable=self.dark_theme_var, command=self.toggle_theme).pack(side=tk.RIGHT, padx=10)

        # 7. СТАТУС-БАР внизу окна
        self.statusbar_frame = tk.Frame(self.root, bg="#334155", height=24)
        self.statusbar_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # Левая часть: основной статус
        self.statusbar_left = tk.Label(
            self.statusbar_frame,
            text="● Готов к работе",
            font=("Segoe UI", 8),
            bg="#334155",
            fg="#e2e8f0",
            anchor=tk.W,
            padx=10
        )
        self.statusbar_left.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Правая часть: дополнительная информация
        self.statusbar_right = tk.Label(
            self.statusbar_frame,
            text="",
            font=("Segoe UI", 8),
            bg="#334155",
            fg="#94a3b8",
            anchor=tk.E,
            padx=10
        )
        self.statusbar_right.pack(side=tk.RIGHT)

    def update_statusbar(self, status_text, info_text=""):
        """Обновляет статус-бар"""
        self.statusbar_left.configure(text=status_text)
        self.statusbar_right.configure(text=info_text)

    def apply_quick_preset(self):
        # Предлагаем пользователю выбрать рабочую папку
        initial_dir = self.output_dir.get() if self.output_dir.get() else os.path.expanduser("~")
        workzone_dir = filedialog.askdirectory(title="Выберите рабочую папку (workzone)", initialdir=initial_dir)

        if not workzone_dir:
            return

        # Устанавливаем пути относительно выбранной папки
        self.articles_path.set(os.path.join(workzone_dir, "articles.txt"))
        self.output_dir.set(os.path.join(workzone_dir, "result"))

        if self.app_mode.get() == "clean":
            self.codes_dir.set(os.path.join(workzone_dir, "result"))
        else:
            self.codes_dir.set(os.path.join(workzone_dir, "codes"))

        # Создаем папки если не существуют
        for path in [self.codes_dir.get(), self.output_dir.get()]:
            if not os.path.exists(path):
                try:
                    os.makedirs(path, exist_ok=True)
                except:
                    pass

        self.log(f"[ℹ] Рабочая папка установлена: {workzone_dir}")
        self.save_current_settings()

    def save_current_settings(self):
        """Сохраняет текущие настройки в конфиг"""
        self.config.set("last_articles_path", self.articles_path.get())
        self.config.set("last_codes_dir", self.codes_dir.get())
        self.config.set("last_output_dir", self.output_dir.get())
        self.config.set("last_invoice_prefix", self.invoice_prefix.get())
        self.config.set("auto_open_results", self.auto_open.get())
        self.config.set("check_boxes", self.check_boxes_var.get())
        self.config.set("remove_brackets", self.remove_brackets_var.get())

    def cancel_operation(self):
        """Отменяет текущую операцию"""
        if self.is_running:
            self.is_cancelled = True
            self.log("[⚠] Запрошена отмена операции... Ожидание завершения текущего файла.")
            self.status_label.configure(text="Отмена операции...")
            self.cancel_btn.configure(state=tk.DISABLED)

    def toggle_mode_ui(self):
        if self.app_mode.get() == "clean":
            self.ent_art.configure(state=tk.DISABLED)
            self.btn_art.configure(state=tk.DISABLED)
            self.box_chk.configure(state=tk.DISABLED)

            # В режиме очистки codes_dir = result
            if self.output_dir.get():
                self.codes_dir.set(self.output_dir.get())

            self.lbl_source.configure(text="Папка-источник (result):")
            self.dash_blocks[0].configure(text="Всего файлов")
            self.dash_blocks[1].configure(text="Очищено файлов")
            self.dash_frame.configure(text=" Панель мониторинга (Режим Клинера) ")
        else:
            self.ent_art.configure(state=tk.NORMAL)
            self.btn_art.configure(state=tk.NORMAL)
            self.box_chk.configure(state=tk.NORMAL)

            # Восстанавливаем codes из конфига или предлагаем выбрать
            saved_codes = self.config.get("last_codes_dir", "")
            if saved_codes and os.path.exists(saved_codes):
                self.codes_dir.set(saved_codes)

            self.lbl_source.configure(text="Папка с кодами:")
            self.dash_blocks[0].configure(text="Всего артикулов")
            self.dash_blocks[1].configure(text="Успешно (OK)")
            self.dash_frame.configure(text=" Панель мониторинга ")

    def save_log(self):
        content = self.log_text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Лог пуст", "Нет данных для сохранения.")
            return
        file_path = filedialog.asksaveasfilename(
            title="Сохранить лог", initialdir=self.output_dir.get(), defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")], initialfile="cargoparse_log.txt"
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Готово", f"Лог успешно сохранён:\n{file_path}")

    def open_output_folder(self):
        """Открывает папку с результатами в проводнике"""
        output_path = self.output_dir.get().strip()
        if not output_path:
            messagebox.showwarning("Папка не указана", "Укажите папку для результатов")
            return

        if not os.path.exists(output_path):
            messagebox.showwarning("Папка не найдена", f"Папка не существует:\n{output_path}")
            return

        try:
            os.startfile(output_path)
            self.log(f"[ℹ] Открыта папка: {output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть папку:\n{e}")

    def open_log_search(self):
        """Открывает окно поиска по логу"""
        if not hasattr(self, 'search_window') or not self.search_window.winfo_exists():
            self.search_window = tk.Toplevel(self.root)
            self.search_window.title("Поиск в логе")
            self.search_window.geometry("450x150")
            self.search_window.resizable(False, False)

            # Поле поиска
            search_frame = ttk.Frame(self.search_window, padding="10")
            search_frame.pack(fill=tk.X)

            ttk.Label(search_frame, text="Найти:").pack(side=tk.LEFT, padx=(0, 5))
            self.search_entry = ttk.Entry(search_frame, width=35)
            self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.search_entry.focus()

            # Переменные для отслеживания
            self.all_matches = []  # Все найденные позиции
            self.current_match_index = -1  # Текущее совпадение

            # Автопоиск при вводе
            self.search_entry.bind('<KeyRelease>', lambda e: self.perform_search())

            # Кнопки
            btn_frame = ttk.Frame(self.search_window, padding="10")
            btn_frame.pack(fill=tk.X)

            ttk.Button(btn_frame, text="⬇️ Следующее", command=self.goto_next_match).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="⬆️ Предыдущее", command=self.goto_previous_match).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="❌ Закрыть", command=self.close_search).pack(side=tk.RIGHT, padx=5)

            # Счетчик совпадений
            self.match_info_label = ttk.Label(self.search_window, text="", font=("Segoe UI", 9), foreground=self.text_secondary)
            self.match_info_label.pack(pady=5)

            # Горячие клавиши
            self.search_entry.bind('<Return>', lambda e: self.goto_next_match())
            self.search_entry.bind('<Shift-Return>', lambda e: self.goto_previous_match())
            self.search_window.bind('<Escape>', lambda e: self.close_search())
            self.search_window.protocol("WM_DELETE_WINDOW", self.close_search)

        else:
            self.search_window.lift()
            self.search_entry.focus()

    def perform_search(self):
        """Выполняет поиск и находит все совпадения"""
        search_text = self.search_entry.get()

        # Очищаем предыдущие подсветки
        self.log_text.tag_remove("search", "1.0", tk.END)
        self.log_text.tag_remove("search_current", "1.0", tk.END)

        if not search_text:
            self.all_matches = []
            self.current_match_index = -1
            self.match_info_label.configure(text="")
            return

        # Находим все совпадения
        self.all_matches = []
        start_pos = "1.0"

        while True:
            pos = self.log_text.search(search_text, start_pos, tk.END, nocase=True)
            if not pos:
                break

            end_pos = f"{pos}+{len(search_text)}c"
            self.all_matches.append((pos, end_pos))

            # Подсвечиваем все найденные
            self.log_text.tag_add("search", pos, end_pos)

            start_pos = end_pos

        # Настройка подсветки
        self.log_text.tag_config("search", background="#4a5568", foreground="#ffffff")
        self.log_text.tag_config("search_current", background="#ffeb3b", foreground="#000000")

        if self.all_matches:
            self.current_match_index = 0
            self.highlight_current_match()
            self.match_info_label.configure(text=f"Найдено: {len(self.all_matches)} совпадений")
        else:
            self.current_match_index = -1
            self.match_info_label.configure(text="Совпадений не найдено")

    def highlight_current_match(self):
        """Подсвечивает текущее совпадение"""
        if not self.all_matches or self.current_match_index < 0:
            return

        # Убираем подсветку с предыдущего
        self.log_text.tag_remove("search_current", "1.0", tk.END)

        # Подсвечиваем текущее
        pos, end_pos = self.all_matches[self.current_match_index]
        self.log_text.tag_add("search_current", pos, end_pos)
        self.log_text.see(pos)

        # Обновляем счетчик
        self.match_info_label.configure(
            text=f"Совпадение {self.current_match_index + 1} из {len(self.all_matches)}"
        )

    def goto_next_match(self):
        """Переход к следующему совпадению"""
        if not self.all_matches:
            return

        self.current_match_index = (self.current_match_index + 1) % len(self.all_matches)
        self.highlight_current_match()

    def goto_previous_match(self):
        """Переход к предыдущему совпадению"""
        if not self.all_matches:
            return

        self.current_match_index = (self.current_match_index - 1) % len(self.all_matches)
        self.highlight_current_match()

    def close_search(self):
        """Закрывает окно поиска и очищает подсветку"""
        self.log_text.tag_remove("search", "1.0", tk.END)
        self.log_text.tag_remove("search_current", "1.0", tk.END)
        if hasattr(self, 'search_window') and self.search_window.winfo_exists():
            self.search_window.destroy()

    def update_status_phrases_loop(self):
        if getattr(self, 'is_running', False):
            self.status_label.configure(text=random.choice(self.funny_statuses))
            self.root.after(2000, self.update_status_phrases_loop)

    def set_stat(self, stat_type, value):
        """Устанавливает статистику"""
        if stat_type == "total":
            self.stat_total.set(str(value))
        elif stat_type == "ok":
            self.stat_ok.set(str(value))
        elif stat_type == "bad":
            self.stat_bad.set(str(value))
        elif stat_type == "dups":
            self.stat_dups.set(str(value))

    def set_current_file(self, filename):
        """Обновляет информацию о текущем обрабатываемом файле"""
        if self.current_file_label:
            self.current_file_label.configure(text=f"📄 Обрабатывается: {filename}")

    def update_eta(self, current, total, elapsed_time):
        """Обновляет ETA (оставшееся время)"""
        if self.eta_label and current > 0 and total > 0:
            avg_time = elapsed_time / current
            remaining = (total - current) * avg_time

            if remaining < 60:
                eta_text = f"⏱️ Осталось: ~{int(remaining)} сек"
            else:
                eta_text = f"⏱️ Осталось: ~{int(remaining / 60)} мин"

            self.eta_label.configure(text=eta_text)

    def update_progress(self, current, total):
        percent = int((current / total) * 100) if total > 0 else 0
        self.progress_bar['value'] = percent
        self.progress_label.configure(text=f"{percent}%")
        self.root.update_idletasks()

    def log(self, message):
        self.root.after(0, self._unsafe_log, message)

    def _unsafe_log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def start_processing_thread(self):
        if self.is_running:
            messagebox.showwarning("Внимание", "Обработка уже выполняется!")
            return

        # ===== ВАЛИДАЦИЯ ПЕРЕД ЗАПУСКОМ =====
        validation_errors = []
        validation_warnings = []

        # Проверка папки с кодами
        code_dir = self.codes_dir.get().strip()
        if not code_dir:
            validation_errors.append("❌ Папка с кодами не указана")
        elif not os.path.exists(code_dir):
            validation_errors.append(f"❌ Папка с кодами не найдена:\n   {code_dir}")
        elif not os.path.isdir(code_dir):
            validation_errors.append(f"❌ Указанный путь не является папкой:\n   {code_dir}")
        else:
            # Проверяем, есть ли в папке xlsx файлы
            xlsx_files = [f for f in os.listdir(code_dir) if f.endswith('.xlsx')]
            if not xlsx_files:
                validation_warnings.append(f"⚠️  В папке с кодами нет .xlsx файлов:\n   {code_dir}")

        # Проверка папки результатов
        out_dir = self.output_dir.get().strip()
        if not out_dir:
            validation_errors.append("❌ Папка для результатов не указана")
        elif not os.path.exists(out_dir):
            # Спрашиваем создать ли папку
            if messagebox.askyesno("Создать папку?", f"Папка не существует:\n{out_dir}\n\nСоздать её?"):
                try:
                    os.makedirs(out_dir, exist_ok=True)
                    self.log(f"[✓] Создана папка: {out_dir}")
                except Exception as e:
                    validation_errors.append(f"❌ Не удалось создать папку результатов:\n   {e}")
            else:
                validation_errors.append("❌ Папка для результатов не существует")

        # Проверка файла артикулов (только для режима verify)
        if self.app_mode.get() == "verify":
            art = self.articles_path.get().strip()
            if not art:
                validation_errors.append("❌ Файл артикулов не указан")
            elif not os.path.exists(art):
                validation_errors.append(f"❌ Файл артикулов не найден:\n   {art}")
            elif not os.path.isfile(art):
                validation_errors.append(f"❌ Указанный путь не является файлом:\n   {art}")

        # Показываем ошибки
        if validation_errors:
            error_msg = "Обнаружены ошибки:\n\n" + "\n\n".join(validation_errors)
            messagebox.showerror("Ошибка валидации", error_msg)
            self.log("[✗] Валидация не пройдена. Исправьте ошибки.")
            return

        # Показываем предупреждения (но не блокируем запуск)
        if validation_warnings:
            warning_msg = "Обнаружены предупреждения:\n\n" + "\n\n".join(validation_warnings) + "\n\nПродолжить?"
            if not messagebox.askyesno("Предупреждение", warning_msg):
                self.log("[ℹ] Запуск отменён пользователем")
                return

        self.log("[✓] Валидация пройдена успешно")
        # ===== КОНЕЦ ВАЛИДАЦИИ =====

        self.is_running = True
        self.is_cancelled = False
        self.run_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)

        for var in [self.stat_total, self.stat_ok, self.stat_bad, self.stat_dups]:
            var.set("0")
        self.log_text.delete("1.0", tk.END)

        # Сохраняем время старта для ETA
        import time
        self.start_time = time.time()

        # Обновляем статус-бар
        self.update_statusbar("● Обработка запущена...", "")

        # Сохраняем настройки перед запуском
        self.save_current_settings()

        self.update_status_phrases_loop()
        threading.Thread(target=self.run_core, daemon=True).start()

    def run_core(self):
        art = self.articles_path.get().strip()
        code_dir = self.codes_dir.get().strip()
        out_dir = self.output_dir.get().strip()
        prefix = self.invoice_prefix.get().strip()
        dry_run = self.dry_run_var.get()

        if not os.path.exists(code_dir):
            messagebox.showerror("Ошибка", f"Папка {code_dir} не найдена на диске!")
            self.is_running = False
            self.run_btn.configure(state=tk.NORMAL)
            self.cancel_btn.configure(state=tk.DISABLED)
            return

        try:
            is_remove_brackets_enabled = self.remove_brackets_var.get()

            if self.app_mode.get() == "verify":
                if not os.path.exists(art):
                    messagebox.showerror("Ошибка", f"Файл {art} не найден!")
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL)
                    self.cancel_btn.configure(state=tk.DISABLED)
                    return

                if dry_run:
                    self.log("=== РЕЖИМ ПРОВЕРКИ (БЕЗ СОХРАНЕНИЯ ФАЙЛОВ) ===")
                else:
                    self.log("=== ЗАПУСК ПОЛНОЙ ВЕРИФИКАЦИИ ПО АРТИКУЛАМ ===")

                is_check_boxes_enabled = self.check_boxes_var.get()

                # Передаем callback для проверки отмены и dry_run режим
                process_customs_data(
                    art, code_dir, out_dir,
                    self.log, self.update_progress, self.set_stat,
                    check_boxes=is_check_boxes_enabled,
                    remove_brackets=is_remove_brackets_enabled,
                    invoice_prefix=prefix,
                    cancel_callback=lambda: self.is_cancelled,
                    dry_run=dry_run,
                    file_callback=self.set_current_file
                )
            else:
                if dry_run:
                    self.log("=== РЕЖИМ ПРОВЕРКИ ОЧИСТКИ (БЕЗ СОХРАНЕНИЯ ФАЙЛОВ) ===")
                else:
                    self.log("=== ЗАПУСК РЕЖИМА БЫСТРОЙ ОЧИСТКИ (ФОРМАТТЕР СТОЛБЦОВ) ===")

                process_cleaner_mode(
                    code_dir, out_dir, self.log, self.update_progress, self.set_stat,
                    remove_brackets=is_remove_brackets_enabled,
                    invoice_prefix=prefix,
                    cancel_callback=lambda: self.is_cancelled,
                    dry_run=dry_run,
                    file_callback=self.set_current_file
                )

            if self.is_cancelled:
                self.status_label.configure(text="Операция отменена пользователем")
                self.update_statusbar("● Операция отменена", "")
                messagebox.showwarning("Отменено", "Операция была отменена пользователем.")
            else:
                self.status_label.configure(text="Обработка завершена успешно!")

                # Вычисляем время обработки
                import time
                elapsed = time.time() - self.start_time if self.start_time else 0
                elapsed_str = f"{elapsed:.1f} сек" if elapsed < 60 else f"{elapsed/60:.1f} мин"

                # Формируем статистику для статус-бара
                total = int(self.stat_total.get())
                ok = int(self.stat_ok.get())
                bad = int(self.stat_bad.get())
                info_str = f"Обработано: {total} | OK: {ok} | BAD: {bad} | Время: {elapsed_str}"
                self.update_statusbar("● Завершено успешно", info_str)

                if dry_run:
                    messagebox.showinfo("Проверка завершена", f"Проверка выполнена!\nФайлы не были созданы (режим проверки).")
                else:
                    messagebox.showinfo("Успех", f"Операция выполнена!\nВсе результаты сохранены в:\n{out_dir}")

                    if self.auto_open.get() and not dry_run:
                        if sys.platform == 'win32':
                            os.startfile(out_dir)
                        elif sys.platform == 'darwin':
                            subprocess.call(['open', out_dir])
                        else:
                            subprocess.call(['xdg-open', out_dir])

        except Exception as ex:
            messagebox.showerror("Критический сбой", f"Ошибка: {ex}")
            self.update_statusbar("● Ошибка выполнения", str(ex)[:50])
        finally:
            self.is_running = False
            self.is_cancelled = False
            self.run_btn.configure(state=tk.NORMAL)
            self.cancel_btn.configure(state=tk.DISABLED)
            self.current_file_label.configure(text="")
            self.eta_label.configure(text="")

    def show_smart_tips(self):
        """Показывает умные подсказки на основе результатов"""
        try:
            total = int(self.stat_total.get())
            ok = int(self.stat_ok.get())
            bad = int(self.stat_bad.get())
            dups = int(self.stat_dups.get())

            tips = []

            # Если много расхождений
            if total > 0 and bad > total * 0.3:
                tips.append("⚠️ Более 30% артикулов с расхождениями. Проверьте файлы кодов.")

            # Если много дубликатов
            if dups > total * 0.1:
                tips.append("🔄 Обнаружено много дубликатов. Возможно, файлы обрабатывались ранее.")

            # Если всё отлично
            if bad == 0 and total > 0:
                tips.append("✨ Отлично! Все артикулы обработаны без расхождений!")

            # Если мало успешных
            if total > 0 and ok < total * 0.5:
                tips.append("💡 Менее 50% успешных. Возможно, не хватает файлов деклараций.")

            # Показываем подсказки в логе
            if tips:
                self.log("\n" + "="*60)
                self.log("🤖 УМНЫЕ ПОДСКАЗКИ:")
                self.log("="*60)
                for tip in tips:
                    self.log(tip)
                self.log("="*60 + "\n")

        except:
            pass