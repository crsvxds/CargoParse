import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment

# --- ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ ОЧИСТКИ МУСОРА ---
def clean_raw_value(val):
    v = str(val).strip()
    if not v or v.lower() == 'nan' or v == 'НЕТУ': 
        return ""
    
    # Отрезаем текстовые приписки
    if 'код упаковки:' in v.lower():
        parts = v.split(':', 1)
        if len(parts) > 1:
            v = parts[1].strip()
            
    # Игнорируем любые текстовые шапки
    lower_v = v.lower()
    if 'код маркировки' in lower_v or 'код упаковки' in lower_v or 'артикул' in lower_v:
        return ""
        
    return v

# =================================================================
# РЕЖИМ 1: ОСНОВНАЯ ВЕРИФИКАЦИЯ (Сверка по articles.txt)
# =================================================================
def process_customs_data(art_file, c_folder, out_dir, log_callback, progress_callback, stats_callback):
    wb_import = Workbook()
    ws_import = wb_import.active
    ws_import.title = "УралИмпорт"
    curr_import = 1

    wb_trade = Workbook()
    ws_trade = wb_trade.active
    ws_trade.title = "УралТрейд"
    curr_trade = 1

    not_found_articles = []
    mismatched_articles = []
    duplicate_summary = {} 
    article_routing_summary = []

    # Глобальная база для сквозной проверки уникальности
    global_seen_codes = {}

    with open(art_file, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    total_lines = len(lines)
    stats_callback("total", total_lines)
    
    ok_counter = 0
    bad_counter = 0
    total_dups_counter = 0
    all_files = os.listdir(c_folder)

    for index, line in enumerate(lines):
        parts = line.split()
        if len(parts) < 6:
            art_name = parts[0] if parts else "НЕИЗВЕСТНО"
            log_callback(f"[ПРОПУСК] Неверный формат строки: {line[:20]}...")
            article_routing_summary.append((art_name, "❌ Ошибка формата строки"))
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
            article_routing_summary.append((article, "❌ Ошибка количества"))
            bad_counter += 1
            stats_callback("bad", bad_counter)
            progress_callback(index + 1, total_lines)
            continue

        # Умный поиск файлов по префиксу
        matched_files = []
        for file_name in all_files:
            name_without_ext = os.path.splitext(file_name)[0]
            if name_without_ext == article or \
               name_without_ext.startswith(f"{article}-") or \
               name_without_ext.startswith(f"{article}_") or \
               name_without_ext.startswith(f"{article} "):
                matched_files.append(os.path.join(c_folder, file_name))

        if not matched_files:
            not_found_articles.append(article)
            article_routing_summary.append((article, "❌ Файлы не найдены"))
            log_callback(f"[ПРОПУСК] ... Артикул {article}: Файлы декларации НЕ найдены")
            bad_counter += 1
            stats_callback("bad", bad_counter)
            progress_callback(index + 1, total_lines)
            continue

        # ШАГ 1: Считываем сырые данные
        file_data_list = []
        for file_path in matched_files:
            try:
                df = pd.read_excel(file_path, header=None)
                file_name_short = os.path.basename(file_path)
                num_cols = df.shape[1]
                
                f_type = "import"
                f_codes_raw = []
                
                if num_cols >= 4:
                    f_type = "trade"
                    last_pack_code = "БЕЗ_КОДА_УПАКОВКИ"
                    for idx_row in range(len(df)):
                        mark_val = str(df.iloc[idx_row, 1]).strip() if pd.notna(df.iloc[idx_row, 1]) else ""
                        pack_val = str(df.iloc[idx_row, 3]).strip() if pd.notna(df.iloc[idx_row, 3]) else ""
                        
                        if pack_val and pack_val.lower() != 'nan' and not pack_val.lower().startswith('код'):
                            last_pack_code = pack_val
                            
                        if not mark_val or mark_val.lower() == 'nan' or mark_val.lower().startswith('код'):
                            continue
                        if mark_val.startswith("(00)"):
                            continue
                            
                        f_codes_raw.append((mark_val, last_pack_code))
                else:
                    if len(df) > 1:
                        for val in df.iloc[1:, 0].dropna():
                            v = str(val).strip()
                            if v and v.lower() != 'nan':
                                f_codes_raw.append(v)
                                
                file_data_list.append({
                    'name': file_name_short,
                    'type': f_type,
                    'raw_codes': f_codes_raw
                })
            except Exception as e:
                log_callback(f"[ОШИБКА ЧТЕНИЯ] файла {os.path.basename(file_path)}. Детали: {e}")

        def get_unique_count(files_subset):
            seen = set()
            for f in files_subset:
                if f['type'] == 'trade':
                    for mark, pkg in f['raw_codes']: seen.add(mark)
                else:
                    for mark in f['raw_codes']: seen.add(mark)
            return len(seen)

        # ШАГ 2: Автокоррекция конфликтов версий файлов
        combined_count = get_unique_count(file_data_list)
        selected_files = file_data_list
        
        if combined_count != quantity_int:
            perfect_matches = []
            for f in file_data_list:
                if get_unique_count([f]) == quantity_int:
                    perfect_matches.append(f)
            
            if len(perfect_matches) == 1:
                selected_files = perfect_matches
                log_callback(f"   [АВТОКОРРЕКЦИЯ] Артикул {article}: Общая сумма ({combined_count}) не совпала.")
                log_callback(f"   [АВТОКОРРЕКЦИЯ] Взят ТОЛЬКО файл '{selected_files[0]['name']}' (ровно {quantity_int} шт.).")

        # ШАГ 3: Финальная обработка и сквозной контроль дублей
        all_product_codes = []
        files_loaded_names = []
        dup_count_for_article = 0
        trade_groups = {}
        article_type = "import"
        local_art_seen = set()

        for fdata in selected_files:
            current_file_name = fdata['name']
            files_loaded_names.append(current_file_name)
            
            if fdata['type'] == 'trade':
                article_type = 'trade'
                for mark_val, pkg_code in fdata['raw_codes']:
                    if mark_val in local_art_seen:
                        dup_count_for_article += 1
                        total_dups_counter += 1
                        stats_callback("dups", total_dups_counter)
                    elif mark_val in global_seen_codes:
                        dup_count_for_article += 1
                        total_dups_counter += 1
                        stats_callback("dups", total_dups_counter)
                        log_callback(f"   [МЕЖФАЙЛОВЫЙ ДУБЛЬ] Артикул {article}: Код {mark_val} уже был в {global_seen_codes[mark_val]}")
                    else:
                        local_art_seen.add(mark_val)
                        global_seen_codes[mark_val] = current_file_name
                        all_product_codes.append(mark_val)
                        
                        if pkg_code not in trade_groups:
                            trade_groups[pkg_code] = []
                        trade_groups[pkg_code].append(mark_val)
            else:
                for mark_val in fdata['raw_codes']:
                    if mark_val in local_art_seen:
                        dup_count_for_article += 1
                        total_dups_counter += 1
                        stats_callback("dups", total_dups_counter)
                    elif mark_val in global_seen_codes:
                        dup_count_for_article += 1
                        total_dups_counter += 1
                        stats_callback("dups", total_dups_counter)
                        log_callback(f"   [МЕЖФАЙЛОВЫЙ ДУБЛЬ] Артикул {article}: Код {mark_val} уже был в {global_seen_codes[mark_val]}")
                    else:
                        local_art_seen.add(mark_val)
                        global_seen_codes[mark_val] = current_file_name
                        all_product_codes.append(mark_val)

        if dup_count_for_article > 0:
            duplicate_summary[article] = dup_count_for_article

        total_codes_count = len(all_product_codes)
        type_label = "УралТрейд" if article_type == "trade" else "УралИмпорт"
        article_routing_summary.append((article, f"✅ {type_label}"))

        if total_codes_count != quantity_int:
            diff = abs(quantity_int - total_codes_count)
            status_text = "меньше" if total_codes_count < quantity_int else "больше"
            diff_str = f"В файлах: {total_codes_count}, {status_text.capitalize()} на {diff}"
            
            files_list_str = ", ".join(files_loaded_names)
            mismatched_articles.append({
                "article": article, "expected": quantity_int, "actual": total_codes_count,
                "diff": diff, "status": status_text, "files": files_list_str
            })
            log_callback(f"[{type_label}] Артикул {article}: РАСХОЖДЕНИЕ (Ожидалось {quantity_int}, собрано {total_codes_count})")
            bad_counter += 1
            stats_callback("bad", bad_counter)
        else:
            diff_str = f"В файлах: {total_codes_count}, Совпало"
            log_callback(f"[{type_label}] ... Артикул {article}: OK (Совпало {total_codes_count} шт.)")
            ok_counter += 1
            stats_callback("ok", ok_counter)

        # Вывод в Excel (Сводная логика)
        if article_type == "trade":
            ws = ws_trade
            if total_codes_count == 0:
                header_row = curr_trade
                ws[f"A{header_row}"] = "Код маркировки"; ws[f"B{header_row}"] = article; ws[f"C{header_row}"] = name; ws[f"D{header_row}"] = description; ws[f"E{header_row}"] = cargo_places; ws[f"F{header_row}"] = diff_str; ws[f"G{header_row}"] = quantity; ws[f"H{header_row}"] = unit; ws[f"I{header_row}"] = "Код упаковки: НЕТУ"
                curr_trade += 1
                ws[f"A{curr_trade}"] = "НЕТУ"
                curr_trade += 2
            else:
                for pkg_code, codes_list in trade_groups.items():
                    if not codes_list: continue
                    header_row = curr_trade
                    ws[f"A{header_row}"] = "Код маркировки"; ws[f"B{header_row}"] = article; ws[f"C{header_row}"] = name; ws[f"D{header_row}"] = description; ws[f"E{header_row}"] = cargo_places; ws[f"F{header_row}"] = diff_str; ws[f"G{header_row}"] = quantity; ws[f"H{header_row}"] = unit; ws[f"I{header_row}"] = f"Код упаковки: {pkg_code}"
                    curr_trade += 1
                    for code in codes_list:
                        ws[f"A{curr_trade}"] = code
                        curr_trade += 1
                    curr_trade += 1 
        else:
            ws = ws_import
            header_row = curr_import
            ws[f"A{header_row}"] = "Код маркировки"; ws[f"B{header_row}"] = article; ws[f"C{header_row}"] = name; ws[f"D{header_row}"] = description; ws[f"E{header_row}"] = cargo_places; ws[f"F{header_row}"] = diff_str; ws[f"G{header_row}"] = quantity; ws[f"H{header_row}"] = unit
            curr_import += 1
            if total_codes_count == 0:
                ws[f"A{curr_import}"] = "НЕТУ"
                curr_import += 2
            else:
                for code in all_product_codes:
                    ws[f"A{curr_import}"] = code
                    curr_import += 1
                curr_import += 1 

        progress_callback(index + 1, total_lines)

    for work_sheet in [ws_import, ws_trade]:
        for col in work_sheet.columns:
            max_length = max(len(str(cell.value or '')) for cell in col)
            column = col[0].column_letter
            for cell in col: cell.alignment = Alignment(vertical="top")
            work_sheet.column_dimensions[column].width = min(max_length + 5, 60)

    out_import = os.path.join(out_dir, "УралИмпорт.xlsx")
    out_trade = os.path.join(out_dir, "УралТрейд.xlsx")
    files_created = []
    
    if curr_import > 1:
        wb_import.save(out_import)
        files_created.append(out_import)
    if curr_trade > 1:
        wb_trade.save(out_trade)
        files_created.append(out_trade)

    log_callback("\n============================================================")
    log_callback("ФИНАЛЬНЫЙ ОТЧЕТ ПО ОШИБКАМ И РАСХОЖДЕНИЯМ:")
    log_callback("============================================================\n")

    if not files_created:
        log_callback("[!] ВНИМАНИЕ: Не создано ни одного файла.")
    else:
        for f_path in files_created: log_callback(f"[✓] Успешно сохранен файл: {os.path.basename(f_path)}")

    if duplicate_summary:
        log_callback(f"\n[!] ОБНАРУЖЕНЫ СОВПАДЕНИЯ (ДУБЛИКАТЫ) КОДОВ МАРКИРОВКИ:")
        log_callback("    *Все типы дубликатов были полностью исключены из итоговых Excel-файлов.*")
    
    log_callback("\n============================================================")
    log_callback("СВОДНАЯ ТАБЛИЦА РАСПРЕДЕЛЕНИЯ АРТИКУЛОВ:")
    log_callback("============================================================")
    log_callback(f"{'АРТИКУЛ'.ljust(30)} | {'НАПРАВЛЕНИЕ / СТАТУС'}")
    log_callback("-" * 60)
    for art, route in article_routing_summary:
        log_callback(f"{art.ljust(30)} | {route}")
    log_callback("============================================================\n")


# =================================================================
# РЕЖИМ 2: БЫСТРАЯ ОЧИСТКА (Форматтер / Клинер)
# =================================================================
def process_cleaner_mode(c_folder, out_dir, log_callback, progress_callback, stats_callback):
    all_files = [f for f in os.listdir(c_folder) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]
    trade_files = [f for f in all_files if 'трейд' in f.lower() and 'clean' not in f.lower()]
    import_files = [f for f in all_files if 'импорт' in f.lower() and 'clean' not in f.lower()]

    total_files = len(trade_files) + len(import_files)
    stats_callback("total", total_files)
    
    if total_files == 0:
        log_callback("⚠️ В папке не найдены файлы, содержащие слова 'трейд' или 'импорт'.")
        return

    # Глобальные базы для сквозной проверки в режиме очистки
    global_seen_marks = set()
    global_seen_packs = set()
    total_duplicates_removed = 0
    processed_count = 0

    # Обработка Трейд
    for f_name in trade_files:
        log_callback(f"⏳ Обработка ТРЕЙД: {f_name}")
        file_path = os.path.join(c_folder, f_name)
        df = pd.read_excel(file_path, header=None)

        mark_codes = []
        pack_codes = []
        local_dups = 0

        for i in range(len(df)):
            if df.shape[1] > 0:
                val_a = clean_raw_value(df.iloc[i, 0])
                if val_a:
                    if val_a not in global_seen_marks:
                        global_seen_marks.add(val_a)
                        mark_codes.append(val_a)
                    else:
                        local_dups += 1
                        total_duplicates_removed += 1
                        
            if df.shape[1] > 8:
                val_i = clean_raw_value(df.iloc[i, 8])
                if val_i:
                    if val_i not in global_seen_packs:
                        global_seen_packs.add(val_i)
                        pack_codes.append(val_i)
                    else:
                        local_dups += 1
                        total_duplicates_removed += 1

        wb = Workbook()
        ws_mark = wb.active
        ws_mark.title = "Коды маркировки"
        for idx, code in enumerate(mark_codes, start=1):
            ws_mark[f"A{idx}"] = code
        ws_mark.column_dimensions['A'].width = 45

        if pack_codes:
            ws_pack = wb.create_sheet(title="Коды упаковки")
            for idx, code in enumerate(pack_codes, start=1):
                ws_pack[f"A{idx}"] = code
            ws_pack.column_dimensions['A'].width = 30

        out_name = f"CLEAN_{f_name}"
        wb.save(os.path.join(out_dir, out_name))
        log_callback(f"  ✅ Создан: {out_name} (Маркировок: {len(mark_codes)}, Упаковок: {len(pack_codes)} | Удалено дублей: {local_dups})\n")
        
        processed_count += 1
        progress_callback(processed_count, total_files)
        stats_callback("ok", processed_count)
        stats_callback("dups", total_duplicates_removed)

    # Обработка Импорт
    for f_name in import_files:
        log_callback(f"⏳ Обработка ИМПОРТ: {f_name}")
        file_path = os.path.join(c_folder, f_name)
        df = pd.read_excel(file_path, header=None)

        mark_codes = []
        local_dups = 0
        
        for i in range(len(df)):
            if df.shape[1] > 0:
                val_a = clean_raw_value(df.iloc[i, 0])
                if val_a:
                    if val_a not in global_seen_marks:
                        global_seen_marks.add(val_a)
                        mark_codes.append(val_a)
                    else:
                        local_dups += 1
                        total_duplicates_removed += 1

        wb = Workbook()
        ws_mark = wb.active
        ws_mark.title = "Коды маркировки"
        for idx, code in enumerate(mark_codes, start=1):
            ws_mark[f"A{idx}"] = code
        ws_mark.column_dimensions['A'].width = 45

        out_name = f"CLEAN_{f_name}"
        wb.save(os.path.join(out_dir, out_name))
        log_callback(f"  ✅ Создан: {out_name} (Маркировок: {len(mark_codes)} | Удалено дублей: {local_dups})\n")

        processed_count += 1
        progress_callback(processed_count, total_files)
        stats_callback("ok", processed_count)
        stats_callback("dups", total_duplicates_removed)

    log_callback(f"🎉 ВСЕ ФАЙЛЫ ОЧИЩЕНЫ! (Всего отсеяно дубликатов за сессию: {total_duplicates_removed})")