import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment

def process_customs_data(art_file, c_folder, log_callback, progress_callback, stats_callback):
    wb = Workbook()
    ws = wb.active
    ws.title = "Результат"
    current_row = 1

    not_found_articles = []
    mismatched_articles = []
    duplicate_summary = {} 

    with open(art_file, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    total_lines = len(lines)
    stats_callback("total", total_lines)
    
    ok_counter = 0
    bad_counter = 0
    total_dups_counter = 0
    
    for index, line in enumerate(lines):
        parts = line.split()
        if len(parts) < 6:
            log_callback(f"[ПРОПУСК] Неверный формат строки: {line[:20]}...")
            bad_counter += 1
            stats_callback("bad", bad_counter)
            progress_callback(index + 1, total_lines)
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
            log_callback(f"[ОШИБКА] Неверное число количества у {article}")
            bad_counter += 1
            stats_callback("bad", bad_counter)
            progress_callback(index + 1, total_lines)
            continue

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
            log_callback(f"Артикул {article}: Файлы декларации НЕ найдены")
            bad_counter += 1
            stats_callback("bad", bad_counter)
            current_row += 2
            progress_callback(index + 1, total_lines)
            continue

        all_product_codes = []
        files_loaded_names = []
        dup_count_for_article = 0

        for file_path in matched_files:
            try:
                codes_df = pd.read_excel(file_path, header=None)
                file_name_short = os.path.basename(file_path)
                files_loaded_names.append(file_name_short)
                
                code_source_map = {}  
                
                for value in codes_df[0].tolist():
                    if pd.isna(value): continue
                    cleaned_code = str(value).strip()
                    if cleaned_code:
                        if cleaned_code in code_source_map:
                            dup_count_for_article += 1
                            total_dups_counter += 1
                            stats_callback("dups", total_dups_counter)
                        else:
                            code_source_map[cleaned_code] = file_name_short
                            all_product_codes.append(cleaned_code)
            except Exception as e:
                log_callback(f"[ОШИБКА ЧТЕНИЯ] файла {os.path.basename(file_path)}")

        if dup_count_for_article > 0:
            duplicate_summary[article] = dup_count_for_article

        total_codes_count = len(all_product_codes)
        files_list_str = ", ".join(files_loaded_names)

        if total_codes_count != quantity_int:
            diff = abs(quantity_int - total_codes_count)
            status_text = "меньше" if total_codes_count < quantity_int else "больше"
            ws[f"F{header_row + 1}"] = f"В файлах суммарно (без дублей): {total_codes_count}, {status_text.capitalize()} на {diff}"
            
            mismatched_articles.append({
                "article": article, "expected": quantity_int, "actual": total_codes_count,
                "diff": diff, "status": status_text, "files": files_list_str
            })
            log_callback(f"Артикул {article}: РАСХОЖДЕНИЕ")
            bad_counter += 1
            stats_callback("bad", bad_counter)
        else:
            log_callback(f"Артикул {article}: OK")
            ok_counter += 1
            stats_callback("ok", ok_counter)

        if total_codes_count > 0:
            for code in all_product_codes:
                ws[f"A{current_row}"] = code
                current_row += 1
        current_row += 1
        progress_callback(index + 1, total_lines)

    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        column = col[0].column_letter
        for cell in col:
            cell.alignment = Alignment(vertical="top")
        ws.column_dimensions[column].width = min(max_length + 5, 60)

    # Вместо сохранения возвращаем объект книги
    return wb