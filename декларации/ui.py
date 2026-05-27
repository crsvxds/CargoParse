import sys
import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from core import process_customs_data  # Импортируем нашу логику

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
        label = tk.Label(tw, text=self.text, justify=tk.LEFT, background="#ffffe0", relief=tk.SOLID, borderwidth=1, font=("Segoe UI", "9", "normal"))
        label.pack(ipadx=5, ipady=2)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw: tw.destroy()

class CustomsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Customs Data Consolidator (ВЭД Китай)")
        self.root.geometry("800x720")
        self.root.minsize(750, 650)

        self.funny_statuses = [
            "Разгружаем контейнер из Гуанчжоу...", "Проверяем таможенную декларацию...",
            "Подкупаем инспектора шоколадкой...", "Завариваем крепкий кофе...",
            "Пересчитываем коробки вручную...", "Ищем потерявшийся артикул под столом...",
            "Ждем, пока китайская сторона подпишет доки...", "Сортируем маркировку левой пяткой..."
        ]

        self.articles_path = tk.StringVar(value="articles.txt")
        self.codes_dir = tk.StringVar(value="codes")
        self.output_path = tk.StringVar(value="result.xlsx")
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

        # Dashboard
        dash_frame = ttk.LabelFrame(main_frame, text=" Панель мониторинга (Текущая сессия) ", padding="10")
        dash_frame.pack(fill=tk.X, pady=(0, 10))
        dash_frame.columnconfigure((0, 1, 2, 3), weight=1, uniform="equal")

        for idx, (title, var, style_name, tip) in enumerate([
            ("Всего артикулов", self.stat_total, "StatVal.TLabel", "Строк обнаружено в файле"),
            ("Успешно (OK)", self.stat_ok, "StatVal.TLabel", "Сумма кодов совпала с заявленной"),
            ("Расхождения", self.stat_bad, "StatValBad.TLabel", "Артикулы с ошибками/нехваткой"),
            ("Исключено дублей", self.stat_dups, "StatValDups.TLabel", "Повторы кодов в эксельках")
        ]):
            c = ttk.Frame(dash_frame, relief="groove", padding="5")
            c.grid(row=0, column=idx, padx=5, sticky="nsew")
            ttk.Label(c, text=title, style="StatTitle.TLabel").pack()
            ttk.Label(c, textvariable=var, style=style_name).pack()
            ToolTip(c, tip)

        # Настройки путей
        file_frame = ttk.LabelFrame(main_frame, text=" Настройки путей к данным ", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Label(file_frame, text="Файл артикулов:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.articles_path, width=55).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Обзор...", command=lambda: self.articles_path.set(filedialog.askopenfilename() or self.articles_path.get())).grid(row=0, column=2, pady=5)

        ttk.Label(file_frame, text="Папка с кодами:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.codes_dir, width=55).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Обзор...", command=lambda: self.codes_dir.set(filedialog.askdirectory() or self.codes_dir.get())).grid(row=1, column=2, pady=5)

        ttk.Label(file_frame, text="Сохранить результат:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_path, width=55).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Обзор...", command=lambda: self.output_path.set(filedialog.asksaveasfilename(defaultextension=".xlsx") or self.output_path.get())).grid(row=2, column=2, pady=5)

        # Лог
        log_frame = ttk.LabelFrame(main_frame, text=" Подробный журнал верификации ", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = tk.Text(log_frame, height=10, width=85, state=tk.DISABLED, font=("Consolas", 9))
        sb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Прогресс
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

        # Кнопки управления
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        self.run_btn = ttk.Button(btn_frame, text="СТАРТ ОБРАБОТКИ", command=self.start_processing_thread)
        self.run_btn.pack(side=tk.LEFT, padx=5, ipadx=12, ipady=6)
        ttk.Checkbutton(btn_frame, text="Открыть Excel после завершения", variable=self.auto_open).pack(side=tk.LEFT, padx=15)

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
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def start_processing_thread(self):
        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED)
        for var in [self.stat_total, self.stat_ok, self.stat_bad, self.stat_dups]: var.set("0")
        self.log_text.configure(state=tk.NORMAL); self.log_text.delete("1.0", tk.END); self.log_text.configure(state=tk.DISABLED)
        
        self.update_status_phrases_loop()
        threading.Thread(target=self.run_core, daemon=True).start()

    def run_core(self):
        art = self.articles_path.get().strip()
        code_dir = self.codes_dir.get().strip()
        out = self.output_path.get().strip()

        if not os.path.exists(art) or not os.path.exists(code_dir) or not out:
            messagebox.showerror("Ошибка", "Проверьте корректность заполнения всех путей!")
            self.is_running = False
            self.run_btn.configure(state=tk.NORMAL)
            return

        try:
            self.log("=== ЗАПУСК ВЕРИФИКАЦИИ ДЕКЛАРАЦИЙ ===")
            # Передаем ссылки на функции UI внутрь логики ядра
            process_customs_data(art, code_dir, out, self.log, self.update_progress, self.set_stat)
            
            self.status_label.configure(text="Обработка завершена успешно!")
            messagebox.showinfo("Успех", f"Обработка завершена!\nФайл сохранен: {out}")
            if self.auto_open.get():
                os.startfile(out) if sys.platform == 'win32' else subprocess.call(['open', out])
        except Exception as ex:
            messagebox.showerror("Критический сбой", f"Ошибка: {ex}")
        finally:
            self.is_running = False
            self.run_btn.configure(state=tk.NORMAL)