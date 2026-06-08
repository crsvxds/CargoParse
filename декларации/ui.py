import os
import random
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import subprocess

# Импортируем оба режима из нашего движка
from core import process_customs_data, process_cleaner_mode

WORK_DIR = r"W:\relises\CargoParse V2\CargoParse V2"

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text: return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 20
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify=tk.LEFT, background="#ffffe0", relief=tk.SOLID, borderwidth=1, font=("Segoe UI", "9")).pack(ipadx=5, ipady=2)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw: tw.destroy()

class CustomsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CargoParse V2 - Профессиональный Комплекс")
        self.root.geometry("850x780")
        self.root.minsize(800, 700)

        self.funny_statuses = [
            "Разгружаем контейнер из Гуанчжоу...", "Проверяем таможенную декларацию...",
            "Подкупаем инспектора шоколадкой...", "Завариваем крепкий кофе...",
            "Пересчитываем коробки вручную...", "Ищем потерявшийся артикул под столом...",
            "Ждем, пока китайская сторона подпишет доки...", "Сортируем маркировку левой пяткой..."
        ]

        # ПЕРЕМЕННАЯ РЕЖИМА
        self.app_mode = tk.StringVar(value="verify") 

        self.articles_path = tk.StringVar(value=os.path.join(WORK_DIR, "articles.txt"))
        self.codes_dir = tk.StringVar(value=os.path.join(WORK_DIR, "codes"))
        self.output_dir = tk.StringVar(value=WORK_DIR)
        
        self.auto_open = tk.BooleanVar(value=True)

        self.stat_total = tk.StringVar(value="0")
        self.stat_ok = tk.StringVar(value="0")
        self.stat_bad = tk.StringVar(value="0")
        self.stat_dups = tk.StringVar(value="0")

        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.configure("StatTitle.TLabel", font=("Segoe UI", 9), foreground="#555555")
        style.configure("StatVal.TLabel", font=("Segoe UI", 16, "bold"), foreground="#0056b3")
        style.configure("StatValBad.TLabel", font=("Segoe UI", 16, "bold"), foreground="#d9534f")
        style.configure("StatValDups.TLabel", font=("Segoe UI", 16, "bold"), foreground="#f0ad4e")

        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ПАНЕЛЬ ВЫБОРА РЕЖИМА
        mode_frame = ttk.LabelFrame(main_frame, text=" Режим работы приложения ", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        rb1 = ttk.Radiobutton(mode_frame, text="📦 Полная верификация по артикулам (Сверка articles.txt)", variable=self.app_mode, value="verify", command=self.toggle_mode_ui)
        rb1.pack(side=tk.LEFT, padx=20)
        
        rb2 = ttk.Radiobutton(mode_frame, text="⚡ Быстрая очистка файлов (Форматтер столбцов)", variable=self.app_mode, value="clean", command=self.toggle_mode_ui)
        rb2.pack(side=tk.LEFT, padx=20)

        # ПАНЕЛЬ МОНИТОРИНГА
        self.dash_frame = ttk.LabelFrame(main_frame, text=" Панель мониторинга ", padding="10")
        self.dash_frame.pack(fill=tk.X, pady=(0, 10))
        self.dash_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="equal")

        self.dash_blocks = []
        for idx, (title, var, style_name, tip) in enumerate([
            ("Всего артикулов", self.stat_total, "StatVal.TLabel", "Обнаружено объектов обработки"),
            ("Успешно (OK)", self.stat_ok, "StatVal.TLabel", "Успешно сведенные данные"),
            ("Расхождения", self.stat_bad, "StatValBad.TLabel", "Обнаруженные ошибки"),
            ("Исключено дублей", self.stat_dups, "StatValDups.TLabel", "Повторы кодов на сессии")
        ]):
            c = ttk.Frame(self.dash_frame, relief="groove", padding="5")
            c.grid(row=0, column=idx, padx=5, sticky="nsew")
            lbl_t = ttk.Label(c, text=title, style="StatTitle.TLabel")
            lbl_t.pack()
            ttk.Label(c, textvariable=var, style=style_name).pack()
            ToolTip(c, tip)
            self.dash_blocks.append(lbl_t)

        # НАСТРОЙКИ ПУТЕЙ
        self.path_frame = ttk.LabelFrame(main_frame, text=" Настройки путей к данным ", padding="10")
        self.path_frame.pack(fill=tk.X, pady=5)

        self.lbl_art = ttk.Label(self.path_frame, text="Файл артикулов:")
        self.lbl_art.grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ent_art = ttk.Entry(self.path_frame, textvariable=self.articles_path, width=65)
        self.ent_art.grid(row=0, column=1, padx=5, pady=5)
        self.btn_art = ttk.Button(self.path_frame, text="Обзор...", command=lambda: self.articles_path.set(filedialog.askopenfilename(initialdir=WORK_DIR) or self.articles_path.get()))
        self.btn_art.grid(row=0, column=2, pady=5)

        ttk.Label(self.path_frame, text="Папка с кодами:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.path_frame, textvariable=self.codes_dir, width=65).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.path_frame, text="Обзор...", command=lambda: self.codes_dir.set(filedialog.askdirectory(initialdir=WORK_DIR) or self.codes_dir.get())).grid(row=1, column=2, pady=5)

        ttk.Label(self.path_frame, text="Сохранить результат в:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.path_frame, textvariable=self.output_dir, width=65).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.path_frame, text="Обзор...", command=lambda: self.output_dir.set(filedialog.askdirectory(initialdir=WORK_DIR) or self.output_dir.get())).grid(row=2, column=2, pady=5)

        # ЛОГ
        log_frame = ttk.LabelFrame(main_frame, text=" Подробный журнал работы ", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = tk.Text(log_frame, height=12, width=85, state=tk.DISABLED, font=("Consolas", 10))
        sb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # ПРОГРЕСС
        progress_frame = ttk.Frame(main_frame, padding="5")
        progress_frame.pack(fill=tk.X, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 2))

        sub_frame = ttk.Frame(progress_frame)
        sub_frame.pack(fill=tk.X)
        self.status_label = ttk.Label(sub_frame, text="Система готова к запуску", font=("Segoe UI", 9, "italic"))
        self.status_label.pack(side=tk.LEFT)
        self.progress_label = ttk.Label(sub_frame, text="0%", width=6, anchor=tk.E)
        self.progress_label.pack(side=tk.RIGHT)

        # КНОПКИ
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        self.run_btn = ttk.Button(btn_frame, text="СТАРТ ОБРАБОТКИ", command=self.start_processing_thread)
        self.run_btn.pack(side=tk.LEFT, padx=5, ipadx=12, ipady=6)
        ttk.Button(btn_frame, text="Сохранить лог", command=self.save_log).pack(side=tk.LEFT, padx=5, ipadx=8, ipady=6)
        ttk.Checkbutton(btn_frame, text="Открыть папку после завершения", variable=self.auto_open).pack(side=tk.LEFT, padx=15)

    # Динамическая смена интерфейса
    def toggle_mode_ui(self):
        if self.app_mode.get() == "clean":
            self.ent_art.configure(state=tk.DISABLED)
            self.btn_art.configure(state=tk.DISABLED)
            self.dash_blocks[0].configure(text="Всего файлов")
            self.dash_blocks[1].configure(text="Очищено файлов")
            self.dash_frame.configure(text=" Панель мониторинга (Режим Клинера) ")
        else:
            self.ent_art.configure(state=tk.NORMAL)
            self.btn_art.configure(state=tk.NORMAL)
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
            initialdir=WORK_DIR,
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
            initialfile="cargoparse_log.txt"
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Готово", f"Лог сохранён:\n{file_path}")

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
        # Безопасный вызов из фонового потока
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
            messagebox.showerror("Ошибка", "Путь к папке с кодами не найден!")
            self.is_running = False
            self.run_btn.configure(state=tk.NORMAL)
            return

        try:
            # Маршрутизация в зависимости от выбранного режима
            if self.app_mode.get() == "verify":
                if not os.path.exists(art):
                    messagebox.showerror("Ошибка", "Файл articles.txt не найден!")
                    self.is_running = False
                    self.run_btn.configure(state=tk.NORMAL)
                    return
                self.log("=== ЗАПУСК ПОЛНОЙ ВЕРИФИКАЦИИ ПО АРТИКУЛАМ ===")
                process_customs_data(art, code_dir, out_dir, self.log, self.update_progress, self.set_stat)
            else:
                self.log("=== ЗАПУСК РЕЖИМА БЫСТРОЙ ОЧИСТКИ (ФОРМАТТЕР) ===")
                # Передаем out_dir как параметр, чтобы чистые файлы сохранялись куда указал юзер
                process_cleaner_mode(code_dir, out_dir, self.log, self.update_progress, self.set_stat)

            self.status_label.configure(text="Обработка завершена успешно!")
            messagebox.showinfo("Успех", f"Операция выполнена!\nСохранено в:\n{out_dir}")
            
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