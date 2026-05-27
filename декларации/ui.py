import sys
import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from core import process_customs_data 

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
        self.result_wb = None # Хранилище результата

        self.funny_statuses = ["Разгружаем контейнер...", "Подкупаем инспектора...", "Ищем артикул под столом..."]
        self.articles_path = tk.StringVar(value="articles.txt")
        self.codes_dir = tk.StringVar(value="codes")
        self.auto_open = tk.BooleanVar(value=True)

        self.stat_total = tk.StringVar(value="0")
        self.stat_ok = tk.StringVar(value="0")
        self.stat_bad = tk.StringVar(value="0")
        self.stat_dups = tk.StringVar(value="0")

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ... (Код создания панелей мониторинга оставляем как был) ...
        dash_frame = ttk.LabelFrame(main_frame, text=" Панель мониторинга ", padding="10")
        dash_frame.pack(fill=tk.X, pady=(0, 10))
        for idx, (title, var, style_name, tip) in enumerate([
            ("Всего", self.stat_total, "StatVal.TLabel", ""), ("OK", self.stat_ok, "StatVal.TLabel", ""),
            ("Расхождения", self.stat_bad, "StatValBad.TLabel", ""), ("Дубли", self.stat_dups, "StatValDups.TLabel", "")]):
            c = ttk.Frame(dash_frame, relief="groove", padding="5")
            c.grid(row=0, column=idx, padx=5, sticky="nsew")
            ttk.Label(c, text=title).pack()
            ttk.Label(c, textvariable=var).pack()

        # Настройки путей (без поля сохранения, так как теперь оно динамическое)
        file_frame = ttk.LabelFrame(main_frame, text=" Пути ", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Label(file_frame, text="Файл артикулов:").grid(row=0, column=0)
        ttk.Entry(file_frame, textvariable=self.articles_path, width=50).grid(row=0, column=1)
        
        ttk.Label(file_frame, text="Папка с кодами:").grid(row=1, column=0)
        ttk.Entry(file_frame, textvariable=self.codes_dir, width=50).grid(row=1, column=1)

        # Лог и прогресс...
        self.log_text = tk.Text(main_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=5)

        # Кнопки
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        self.run_btn = ttk.Button(btn_frame, text="СТАРТ", command=self.start_processing_thread)
        self.run_btn.pack(side=tk.LEFT)
        
        # НОВАЯ КНОПКА
        self.save_btn = ttk.Button(btn_frame, text="СОХРАНИТЬ ОТЧЕТ", command=self.save_report, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=10)

    def save_report(self):
        if not self.result_wb: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path:
            self.result_wb.save(path)
            messagebox.showinfo("Успех", "Отчет сохранен!")
            if self.auto_open.get(): os.startfile(path)

    def start_processing_thread(self):
        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED)
        self.save_btn.configure(state=tk.DISABLED)
        threading.Thread(target=self.run_core, daemon=True).start()

    def run_core(self):
        art = self.articles_path.get()
        code_dir = self.codes_dir.get()
        # Вызываем логику и получаем результат в переменную
        self.result_wb = process_customs_data(art, code_dir, self.log, self.update_progress, self.set_stat)
        
        self.run_btn.configure(state=tk.NORMAL)
        self.save_btn.configure(state=tk.NORMAL)
        messagebox.showinfo("Готово", "Обработка завершена! Нажмите кнопку сохранения.")

    def log(self, msg):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.configure(state=tk.DISABLED)
        
    def update_progress(self, curr, total):
        self.progress_bar['value'] = (curr / total) * 100
        self.root.update_idletasks()
        
    def set_stat(self, type, val):
        if type == "total": self.stat_total.set(val)
        # (Остальные статы аналогично)