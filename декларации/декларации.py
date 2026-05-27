import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment


class CustomsApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Customs Data Consolidator (ВЭД Китай)")
        self.root.geometry("1000x480")
        self.root.minsize(600, 400)

        # Переменные путей (по умолчанию ищут в текущей папке)
        self.articles_path = tk.StringVar(value="articles.txt")
        self.codes_dir = tk.StringVar(value="codes")
        self.output_path = tk.StringVar(value="result.xlsx")

        self.create_widgets()

    def create_widgets(self):
        # Главный контейнер с отступами
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- БЛОК ВЫБОРА ФАЙЛОВ ---
        file_frame = ttk.LabelFrame(main_frame, text=" Настройки путей ", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        # Файл артикулов
        ttk.Label(file_frame, text="Файл артикулов:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(file_frame, textvariable=self.articles_path, width=55).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(
            file_frame, text="Обзор...", command=self.browse_articles
        ).grid(row=0, column=2, pady=5)

        # Папка с кодами
        ttk.Label(file_frame, text="Папка с кодами:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(file_frame, textvariable=self.codes_dir, width=55).grid(
            row=1, column=1, padx=5, pady=5
        )
        ttk.Button(file_frame, text="Обзор...", command=self.browse_codes).grid(
            row=1, column=2, pady=5
        )

        # Выходной файл
        ttk.Label(file_frame, text="Сохранить результат как:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(file_frame, textvariable=self.output_path, width=55).grid(
            row=2, column=1, padx=5, pady=5
        )
        ttk.Button(file_frame, text="Обзор...", command=self.browse_output).grid(
            row=2, column=2, pady=5
        )

        # --- БЛОК ЛОГОВ (КОНСОЛЬ В ИНТЕРФЕЙСЕ) ---
        log_frame = ttk.LabelFrame(main_frame, text=" Лог выполнения ", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.log_text = tk.Text(log_frame, height=12, width=80, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(
            log_frame, orient=tk.VERTICAL, command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # --- КНОПКА ЗАПУСКА ---
        self.run_btn = ttk.Button(
            main_frame, text="СТАРТ ОБРАБОТКИ", command=self.start_processing
        )
        self.run_btn.pack(pady=5, ipadx=10, ipady=5)

    # --- МЕТОДЫ ОБЗОРЩИКА ОТКРЫТИЯ ФАЙЛОВ ---
    def browse_articles(self):
        file = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if file:
            self.articles_path.set(file)

    def browse_codes(self):
        directory = filedialog.askdirectory()
        if directory:
            self.codes_dir.set(directory)

    def browse_output(self):
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
        )
        if file:
            self.output_path.set(file)

    def log(self, message):
        """Вывод сообщений в текстовое поле GUI"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    # --- ГЛАВНАЯ ЛОГИКА ОБРАБОТКИ ДЕКЛАРАЦИЙ ---
    def start_processing(self):
        self.clear_log()

        art_file = self.articles_path.get()
        c_folder = self.codes_dir.get()
        out_file = self.output_path.get()

        # ЖЕСТКАЯ ПРОВЕРКА НА ЖИВУЧЕСТЬ ПРОГРАММЫ
        errors = []
        if not os.path.exists(art_file):
            errors.append(f"Не найден файл артикулов: '{art_file}'")
        if not os.path.exists(c_folder):
            errors.append(f"Не найдена папка с кодами: '{c_folder}'")

        if errors:
            error_msg = "\n".join(errors)
            messagebox.showerror("Критическая ошибка", error_msg)
            self.log("[ОШИБКА ЗАПУСКА] Проверьте пути к файлам и папкам.")
            return

        self.run_btn.configure(state=tk.DISABLED)
        self.log("=== ЗАПУСК ВЕРИФИКАЦИИ ДЕКЛАРАЦИЙ ===")

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Результат"
            current_row = 1

            not_found_articles = []
            mismatched_articles = []

            with open(art_file, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if line.strip()]

            for line in lines:
                parts = line.split()
                if len(parts) < 6:
                    self.log(f"[ПРОПУСК] Неверный формат строки: {line[:20]}...")
                    continue

                article = parts[0]
                unit = parts[-1]
                quantity = parts[-2]
                cargo_places = parts[-3]
                description = parts[-4]
                name = " ".join(parts[1:-4])

                try:
                    quantity_int = int(quantity)
                except ValueError:
                    self.log(f"[ОШИБКА] Неверное число количества у {article}")
                    continue

                # Поиск файлов-суффиксов (-1...-10)
                possible_names = [article] + [f"{article}-{i}" for i in range(1, 11)]
                matched_files = []
                all_files = os.listdir(c_folder)

                for target in possible_names:
                    for file_name in all_files:
                        name_without_ext = os.path.splitext(file_name)[0]
                        if name_without_ext == target:
                            matched_files.append(os.path.join(c_folder, file_name))
                            break

                header_row = current_row
                ws[f"C{header_row}"] = article
                ws[f"D{header_row}"] = name
                ws[f"E{header_row}"] = description
                ws[f"F{header_row}"] = cargo_places
                ws[f"G{header_row}"] = quantity
                ws[f"H{header_row}"] = unit
                current_row += 1

                if not matched_files:
                    ws[f"A{current_row}"] = "НЕТУ"
                    not_found_articles.append(article)
                    self.log(f"Артикул {article}: Файлы декларации НЕ найдены")
                    current_row += 2
                    continue

                # Сбор кодов из внутренностей файлов
                all_product_codes = []
                files_loaded_names = []

                for file_path in matched_files:
                    try:
                        codes_df = pd.read_excel(file_path, header=None)
                        files_loaded_names.append(os.path.basename(file_path))
                        for value in codes_df[0].tolist():
                            if pd.isna(value):
                                continue
                            cleaned_code = str(value).strip()
                            if cleaned_code:
                                all_product_codes.append(cleaned_code)
                    except Exception as e:
                        self.log(f"[ОШИБКА ЧТЕНИЯ] файла {os.path.basename(file_path)}")

                total_codes_count = len(all_product_codes)
                files_list_str = ", ".join(files_loaded_names)

                if total_codes_count != quantity_int:
                    diff = abs(quantity_int - total_codes_count)
                    status_text = "меньше" if total_codes_count < quantity_int else "больше"
                    ws[f"F{header_row + 1}"] = f"В файлах суммарно: {total_codes_count}, {status_text.capitalize()} на {diff}"
                    
                    # Собираем данные в точном соответствии с твоим шаблоном вывода
                    mismatched_articles.append({
                        "article": article,
                        "expected": quantity_int,
                        "actual": total_codes_count,
                        "diff": diff,
                        "status": status_text,
                        "files": files_list_str
                    })
                    self.log(f"Артикул {article}: РАСХОЖДЕНИЕ (Ожидалось {quantity_int}, собрано {total_codes_count})")
                else:
                    self.log(f"Артикул {article}: OK (Совпало {total_codes_count} шт.)")

                if total_codes_count == 0:
                    ws[f"A{current_row}"] = "НЕТУ"
                    current_row += 2
                    continue

                for code in all_product_codes:
                    ws[f"A{current_row}"] = code
                    current_row += 1
                current_row += 1

            # Выравнивание колонок в Excel
            for col in ws.columns:
                max_length = max(len(str(cell.value or '')) for cell in col)
                column = col[0].column_letter
                for cell in col:
                    cell.alignment = Alignment(vertical="top")
                ws.column_dimensions[column].width = min(max_length + 5, 60)

            wb.save(out_file)

            # ============================================================
            # КРАСИВЫЙ ФИНАЛЬНЫЙ ОТЧЕТ В ОКНО GUI ПО ТВОЕМУ ПРИМЕРУ
            # ============================================================
            self.log("\n============================================================")
            self.log("ФИНАЛЬНЫЙ ОТЧЕТ ПО ОШИБКАМ И РАСХОЖДЕНИЯМ:")
            self.log("============================================================\n")

            # 1. Вывод по ненайденным файлам
            if not_found_articles:
                self.log(f"[!] НЕ НАЙДЕНЫ ФАЙЛЫ ДЛЯ АРТИКУЛОВ ({len(not_found_articles)} шт.):")
                for art in not_found_articles:
                    self.log(f"    - Артикул {art}: файлы отсутствуют в папке.")
                self.log("")
            else:
                self.log("[✓] Для каждого артикула из списка найден хотя бы один файл.\n")

            # 2. Вывод расхождений по количеству
            if mismatched_articles:
                self.log(f"[!] РАСХОЖДЕНИЕ ОБЩЕЙ СУММЫ КОДОВ ВНУТРИ ФАЙЛОВ ({len(mismatched_articles)} шт.):")
                for item in mismatched_articles:
                    self.log(
                        f"    - Артикул {item['article']}: "
                        f"В текстовике указано {item['expected']}, "
                        f"суммарно in файлах [{item['files']}] найдено {item['actual']}. "
                        f"Итог: {item['status']} на {item['diff']} шт."
                    )
            else:
                self.log("[✓] Расхождений по количеству кодов во внутрянке не обнаружено.")
            
            self.log("\n============================================================")
            
            messagebox.showinfo("Успех", f"Обработка завершена!\nФайл сохранен: {out_file}")

        except Exception as ex:
            messagebox.showerror("Критический сбой", f"Произошла ошибка во время сборки:\n{ex}")
            self.log(f"[КРАШ СИСТЕМЫ] {ex}")
        finally:
            self.run_btn.configure(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = CustomsApp(root)
    root.mainloop()