import os
import random
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import subprocess

# Импортируем оба режима из нашего движка core.py
from core import process_customs_data, process_cleaner_mode

# Базовые константы путей для быстрого набора
BASE_DIR = r"W:\relises\CargoParse V2\CargoParse V2"
DEFAULT_ART = os.path.join(BASE_DIR, "articles.txt")
DEFAULT_CODES = os.path.join(BASE_DIR, "codes")
DEFAULT_RESULT = os.path.join(BASE_DIR, "result")

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text: return
        
        if event:
            x = event.x_root + 15
            y = event.y_root + 15
        else:
            x = self.widget.winfo_rootx() + 25
            y = self.widget.winfo_rooty() + 20
            
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify=tk.LEFT, background="#22252a", foreground="#ffffff", relief=tk.SOLID, borderwidth=1, font=("Segoe UI", "9")).pack(ipadx=6, ipady=3)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw: tw.destroy()

class CustomsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CargoParse V2 — Professional Suite")
        self.root.geometry("900x900")
        self.root.minsize(850, 750)
        
        # Создаем папку результатов, если её ещё нет
        if not os.path.exists(DEFAULT_RESULT):
            os.makedirs(DEFAULT_RESULT, exist_ok=True)
        
        self.setup_styles()

        self.funny_statuses = [
            "Разгружаем контейнер из Гуанчжоу...", "Проверяем таможенную декларацию...",
            "Подкупаем инспектора шоколадкой...", "Завариваем крепкий кофе...",
            "Пересчитываем коробки вручную...", "Ищем потерявшийся артикул под столом...",
            "Ждем, пока китайская сторона подпишет доки...", "Сортируем маркировку левой пяткой..."
        ]

        self.app_mode = tk.StringVar(value="verify") 

        # Изначально поля можно оставить пустыми или дефолтными
        self.articles_path = tk.StringVar(value=DEFAULT_ART)
        self.codes_dir = tk.StringVar(value=DEFAULT_CODES)
        self.output_dir = tk.StringVar(value=DEFAULT_RESULT)
        
        self.auto_open = tk.BooleanVar(value=True)

        self.stat_total = tk.StringVar(value="0")
        self.stat_ok = tk.StringVar(value="0")
        self.stat_bad = tk.StringVar(value="0")
        self.stat_dups = tk.StringVar(value="0")

        self.create_widgets()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.bg_color = "#f4f6f9"
        self.card_bg = "#ffffff"
        self.primary_color = "#4a6cf7"
        self.success_color = "#2ecc71"
        self.danger_color = "#e74c3c"
        self.warning_color = "#f39c12"
        self.text_main = "#2d3748"
        self.border_color = "#e2e8f0"

        self.root.configure(bg=self.bg_color)

        self.style.configure(".", background=self.bg_color, foreground=self.text_main, font=("Segoe UI", 10))
        self.style.configure("TLabelframe", background=self.bg_color, bordercolor=self.border_color, borderwidth=1, relief="solid")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=self.primary_color, background=self.bg_color)
        
        self.style.configure("TButton", font=("Segoe UI", 10, "bold"), background=self.primary_color, foreground="#ffffff", borderwidth=0, focuscolor=self.primary_color)
        self.style.map("TButton", background=[('active', '#3b5bdb'), ('disabled', '#cbd5e1')], foreground=[('disabled', '#94a3b8')])

        # Стиль для кнопки быстрого набора
        self.style.configure("Preset.TButton", background="#2d3748", foreground="#ffffff")
        self.style.map("Preset.TButton", background=[('active', '#1a202c')])

        self.style.configure("TEntry", fieldbackground="#ffffff", bordercolor=self.border_color, lightcolor=self.border_color, darkcolor=self.border_color, padding=5)
        self.style.map("TEntry", fieldbackground=[('disabled', '#e2e8f0')], foreground=[('disabled', '#718096')])
        
        self.style.configure("TRadiobutton", font=("Segoe UI", 10), background=self.bg_color, focuscolor=self.bg_color)
        self.style.configure("TCheckbutton", font=("Segoe UI", 10), background=self.bg_color, focuscolor=self.bg_color)
        self.style.configure("Horizontal.TProgressbar", troughcolor="#e2e8f0", bordercolor="#e2e8f0", background=self.primary_color, lightcolor=self.primary_color, darkcolor=self.primary_color, thickness=8)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. ПЕРЕКЛЮЧАТЕЛЬ РЕЖИМОВ
        mode_frame = ttk.LabelFrame(main_frame, text=" Режим работы приложения ", padding="15")
        mode_frame.pack(fill=tk.X, pady=(0, 15))
        
        rb1 = ttk.Radiobutton(mode_frame, text="📦  Верификация по артикулам (Сверка с articles.txt)", variable=self.app_mode, value="verify", command=self.toggle_mode_ui)
        rb1.pack(side=tk.LEFT, padx=(10, 30))
        
        rb2 = ttk.Radiobutton(mode_frame, text="⚡  Быстрая очистка файлов (Форматтер столбцов)", variable=self.app_mode, value="clean", command=self.toggle_mode_ui)
        rb2.pack(side=tk.LEFT, padx=10)

        # 2. ПАНЕЛЬ МОНИТОРИНГА
        self.dash_frame = ttk.LabelFrame(main_frame, text=" Панель мониторинга ", padding="15")
        self.dash_frame.pack(fill=tk.X, pady=(0, 15))
        self.dash_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="equal")

        self.dash_blocks = []
        stats_config = [
            ("Всего артикулов", self.stat_total, self.primary_color, "Обнаружено строк/файлов для обработки"),
            ("Успешно (OK)", self.stat_ok, self.success_color, "Данные сошлись и успешно обработаны"),
            ("Расхождения", self.stat_bad, self.danger_color, "Артикулы с ошибками или нехваткой кодов"),
            ("Исключено дублей", self.stat_dups, self.warning_color, "Повторяющиеся коды, отсеянные за сессию")
        ]

        for idx, (title, var, color, tip) in enumerate(stats_config):
            card = tk.Canvas(self.dash_frame, bg=self.card_bg, highlightthickness=1, highlightbackground=self.border_color, height=85)
            card.grid(row=0, column=idx, padx=6, sticky="nsew")
            
            lbl_t = tk.Label(card, text=title, font=("Segoe UI", 9, "bold"), fg="#718096", bg=self.card_bg)
            lbl_t.pack(pady=(12, 2))
            
            lbl_v = tk.Label(card, textvariable=var, font=("Segoe UI", 18, "bold"), fg=color, bg=self.card_bg)
            lbl_v.pack(pady=(0, 10))
            
            ToolTip(card, tip)
            self.dash_blocks.append(lbl_t)

        # 3. НАСТРОЙКИ ПУТЕЙ (С КНОПКОЙ БЫСТРОГО НАБОРА)
        self.path_frame = ttk.LabelFrame(main_frame, text=" Настройки путей к данным ", padding="15")
        self.path_frame.pack(fill=tk.X, pady=(0, 15))
        self.path_frame.columnconfigure(1, weight=1)

        # КНОПКА БЫСТРОГО НАБОРА (ПРЕСЕТ) ВСТАВЛЕНА НАВЕРХ ПАНЕЛИ ПУТЕЙ
        self.btn_preset = ttk.Button(self.path_frame, text="⚡  ЗАПОЛНИТЬ РАБОЧУЮ ПАПКУ (БЫСТРЫЙ НАБОР V2)", style="Preset.TButton", command=self.apply_quick_preset)
        self.btn_preset.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=(0, 10))
        ToolTip(self.btn_preset, "Один клик — и программа мгновенно пропишет пути к статьям, кодам и папке result из ТЗ.")

        # Строка 1: Файл артикулов
        self.lbl_art = ttk.Label(self.path_frame, text="Файл артикулов:")
        self.lbl_art.grid(row=1, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        self.ent_art = ttk.Entry(self.path_frame, textvariable=self.articles_path)
        self.ent_art.grid(row=1, column=1, sticky=tk.EW, pady=8, padx=5)
        self.btn_art = ttk.Button(self.path_frame, text="Обзор...", command=lambda: self.articles_path.set(filedialog.askopenfilename(initialdir=BASE_DIR) or self.articles_path.get()))
        self.btn_art.grid(row=1, column=2, pady=8, padx=(5, 0))

        # Строка 2: Папка с кодами / Источник
        self.lbl_source = ttk.Label(self.path_frame, text="Папка с кодами:")
        self.lbl_source.grid(row=2, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        self.ent_codes = ttk.Entry(self.path_frame, textvariable=self.codes_dir)
        self.ent_codes.grid(row=2, column=1, sticky=tk.EW, pady=8, padx=5)
        self.btn_codes = ttk.Button(self.path_frame, text="Обзор...", command=lambda: self.codes_dir.set(filedialog.askdirectory(initialdir=BASE_DIR) or self.codes_dir.get()))
        self.btn_codes.grid(row=2, column=2, pady=8, padx=(5, 0))

        # Строка 3: Папка вывода результатов
        self.lbl_out = ttk.Label(self.path_frame, text="Папка назначения:")
        self.lbl_out.grid(row=3, column=0, sticky=tk.W, pady=8, padx=(0, 10))
        self.ent_out = ttk.Entry(self.path_frame, textvariable=self.output_dir)
        self.ent_out.grid(row=3, column=1, sticky=tk.EW, pady=8, padx=5)
        self.btn_out = ttk.Button(self.path_frame, text="Обзор...", command=lambda: self.output_dir.set(filedialog.askdirectory(initialdir=BASE_DIR) or self.output_dir.get()))
        self.btn_out.grid(row=3, column=2, pady=8, padx=(5, 0))

        # 4. ЖУРНАЛ РАБОТЫ (ТЕМНЫЙ ТЕРМИНАЛ)
        log_frame = ttk.LabelFrame(main_frame, text=" Подробный журнал работы ", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        self.log_text = tk.Text(log_frame, height=12, width=85, state=tk.DISABLED, font=("Consolas", 10), bg="#1e222b", fg="#f8f8f2", insertbackground="#ffffff", relief=tk.FLAT, padx=10, pady=10)
        sb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # 5. ИНДИКАТОРЫ ПРОГРЕССА
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode='determinate', style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        sub_frame = ttk.Frame(progress_frame)
        sub_frame.pack(fill=tk.X)
        self.status_label = ttk.Label(sub_frame, text="Система готова к запуску", font=("Segoe UI", 9, "italic"), foreground="#718096")
        self.status_label.pack(side=tk.LEFT)
        self.progress_label = ttk.Label(sub_frame, text="0%", font=("Segoe UI", 9, "bold"), foreground=self.primary_color, width=6, anchor=tk.E)
        self.progress_label.pack(side=tk.RIGHT)

        # 6. НИЖНЯЯ ПАНЕЛЬ С КНОПКАМИ
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        
        self.run_btn = ttk.Button(btn_frame, text="⚡ СТАРТ ОБРАБОТКИ", style="TButton", command=self.start_processing_thread)
        self.run_btn.pack(side=tk.LEFT, ipadx=15, ipady=8)
        
        self.save_btn = tk.Button(btn_frame, text="Сохранить лог", font=("Segoe UI", 10), bg="#e2e8f0", fg=self.text_main, borderwidth=0, activebackground="#cbd5e1", command=self.save_log)
        self.save_btn.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=6)
        
        ttk.Checkbutton(btn_frame, text="Открыть папку результатов", variable=self.auto_open).pack(side=tk.LEFT, padx=15)

    # ЛОГИКА КНОПКИ БЫСТРОГО НАБОРА
    def apply_quick_preset(self):
        self.articles_path.set(DEFAULT_ART)
        self.output_dir.set(DEFAULT_RESULT)
        
        # Папку-источник ставим в зависимости от активного режима
        if self.app_mode.get() == "clean":
            self.codes_dir.set(DEFAULT_RESULT)
        else:
            self.codes_dir.set(DEFAULT_CODES)
            
        self.log("[ℹ] Сработал Быстрый набор: пути сброшены на стандартную конфигурацию V2.")

    def toggle_mode_ui(self):
        if self.app_mode.get() == "clean":
            self.ent_art.configure(state=tk.DISABLED)
            self.btn_art.configure(state=tk.DISABLED)
            # Принудительно меняем источник на result для удобства очистки
            self.codes_dir.set(DEFAULT_RESULT)
            
            self.lbl_source.configure(text="Папка-источник (result):")
            self.dash_blocks[0].configure(text="Всего файлов")
            self.dash_blocks[1].configure(text="Очищено файлов")
            self.dash_frame.configure(text=" Панель мониторинга (Режим Клинера) ")
        else:
            self.ent_art.configure(state=tk.NORMAL)
            self.btn_art.configure(state=tk.NORMAL)
            # Возвращаем источник на коды
            self.codes_dir.set(DEFAULT_CODES)
            
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
            title="Сохранить лог",
            initialdir=self.output_dir.get(),
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
            initialfile="cargoparse_log.txt"
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Готово", f"Лог успешно сохранён:\n{file_path}")

    def update_status_phrases_loop(self):
        if getattr(self, 'is_running', False):
            self.status_label.configure(text=random.choice(self.funny_statuses))
            self.root.after(2000, self.update_status_phrases_loop)

    def set_stat(self, stat_type, value):
        if stat_type == "total": self.stat_total.set(str(value))
        elif stat_type == "ok": self.stat_ok.set(str(value))
        elif stat_type == "bad": self.stat_bad.set(str(value))
        elif stat_type == "dups": self.stat_dups.set(str(value))

    def update_progress(self, current, total):
        percent = int((current / total) * 100) if total > 0 else 0
        self.progress_bar['value'] = percent
        self.progress_label.configure(text=f"{percent}%")
        self.root.update_idletasks()

    def log(self, message):
        self.root.after(0, self._unsafe_log, message)

    def _unsafe_log(self, message):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def start_processing_thread(self):
        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED)
        for var in [self.stat_total, self.stat_ok, self.stat_bad, self.stat_dups]: var.set("0")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

        self.update_status_phrases_loop()
        threading.Thread(target=self.run_core, daemon=True).start()

    def run_core(self):
        art = self.articles_path.get().strip()
        code_dir = self.codes_dir.get().strip()
        out_dir = self.output_dir.get().strip()

        if not os.path.exists(code_dir):
            messagebox.showerror("Ошибка", f"Папка {code_dir} не найдена на диске!")
            self.is_running = False
            self.run_btn.configure(state=tk.NORMAL)
            return

        try:
            if self.app_mode.get() == "verify":
                if not os.path.exists(art):
                    messagebox.showerror("Ошибка", f"Файл {art} не найден!")
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL)
                    return
                self.log("=== ЗАПУСК ПОЛНОЙ ВЕРИФИКАЦИИ ПО АРТИКУЛАМ ===")
                process_customs_data(art, code_dir, out_dir, self.log, self.update_progress, self.set_stat)
            else:
                self.log("=== ЗАПУСК РЕЖИМА БЫСТРОЙ ОЧИСТКИ (ФОРМАТТЕР СТОЛБЦОВ) ===")
                process_cleaner_mode(code_dir, out_dir, self.log, self.update_progress, self.set_stat)

            self.status_label.configure(text="Обработка завершена успешно!")
            messagebox.showinfo("Успех", f"Операция выполнена!\nВсе результаты сохранены в:\n{out_dir}")
            
            if self.auto_open.get():
                if sys.platform == 'win32':
                    os.startfile(out_dir)
                elif sys.platform == 'darwin':
                    subprocess.call(['open', out_dir])
                else:
                    subprocess.call(['xdg-open', out_dir])
                    
        except Exception as ex:
            messagebox.showerror("Критический сбой", f"Ошибка: {ex}")
        finally:
            self.is_running = False
            self.run_btn.configure(state=tk.NORMAL)