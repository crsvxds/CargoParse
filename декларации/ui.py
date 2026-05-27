import sys
import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from core import process_customs_data

class CustomsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Customs Data Consolidator (ВЭД Китай)")
        self.root.geometry("800x720")
        
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
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Dashboard
        dash_frame = ttk.LabelFrame(main_frame, text=" Панель мониторинга ", padding="10")
        dash_frame.pack(fill=tk.X, pady=(0, 10))
        
        for idx, (title, var) in enumerate([("Всего", self.stat_total), ("OK", self.stat_ok), 
                                          ("Расхождения", self.stat_bad), ("Дубли", self.stat_dups)]):
            c = ttk.Frame(dash_frame, relief="groove", padding="5")
            c.grid(row=0, column=idx, padx=5, sticky="nsew")
            ttk.Label(c, text=title).pack()
            ttk.Label(c, textvariable=var, font=("Segoe UI", 14, "bold")).pack()

        # Настройки путей
        file_frame = ttk.LabelFrame(main_frame, text=" Настройки путей ", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Label(file_frame, text="Файл артикулов:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.articles_path, width=50).grid(row=0, column=1)
        ttk.Button(file_frame, text="...", command=lambda: self.articles_path.set(filedialog.askopenfilename() or self.articles_path.get())).grid(row=0, column=2)

        ttk.Label(file_frame, text="Папка с кодами:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.codes_dir, width=50).grid(row=1, column=1)
        ttk.Button(file_frame, text="...", command=lambda: self.codes_dir.set(filedialog.askdirectory() or self.codes_dir.get())).grid(row=1, column=2)

        ttk.Label(file_frame, text="Сохранить как:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.output_path, width=50).grid(row=2, column=1)
        ttk.Button(file_frame, text="...", command=lambda: self.output_path.set(filedialog.asksaveasfilename(defaultextension=".xlsx") or self.output_path.get())).grid(row=2, column=2)

        # Лог и прогресс
        self.log_text = tk.Text(main_frame, height=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X)

        # Кнопки
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.run_btn = ttk.Button(btn_frame, text="СТАРТ ОБРАБОТКИ", command=self.start_processing_thread)
        self.run_btn.pack(side=tk.LEFT, padx=5, ipadx=10)
        
        ttk.Checkbutton(btn_frame, text="Открыть Excel", variable=self.auto_open).pack(side=tk.LEFT, padx=10)

    def log(self, message):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def update_progress(self, current, total):
        self.progress_bar['value'] = (current / total) * 100
        self.root.update_idletasks()

    def set_stat(self, stat_type, value):
        if stat_type == "total": self.stat_total.set(value)
        elif stat_type == "ok": self.stat_ok.set(value)
        elif stat_type == "bad": self.stat_bad.set(value)
        elif stat_type == "dups": self.stat_dups.set(value)

    def start_processing_thread(self):
        threading.Thread(target=self.run_core, daemon=True).start()

    def run_core(self):
        self.run_btn.configure(state=tk.DISABLED)
        try:
            process_customs_data(self.articles_path.get(), self.codes_dir.get(), 
                                 self.output_path.get(), self.log, self.update_progress, self.set_stat)
            if self.auto_open.get(): os.startfile(self.output_path.get())
        finally:
            self.run_btn.configure(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = CustomsApp(root)
    root.mainloop()