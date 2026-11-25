import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
import hashlib
import json
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ УМНОГО ПРЕОБРАЗОВАНИЯ В СТРОКУ (УБИРАЕТ .0) ---
def safe_str(val):
    if pd.isna(val):
        return ""
    # Если Pandas прочитал целое число как float (1200.0), делаем его целым (1200)
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


# --- ФУНКЦИИ КЭШИРОВАНИЯ ---
def get_file_hash(file_path):
    """Вычисляет хэш файла для определения изменений"""
    try:
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            buf = f.read(65536)  # Читаем блоками
            while len(buf) > 0:
                hasher.update(buf)
                buf = f.read(65536)
        return hasher.hexdigest()
    except:
        return None

def load_cache(cache_file):
    """Загружает кэш из файла"""
    if os.path.exists(cache_file):
        try:
            with open(cache_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_cache(cache_file, cache_data):
    """Сохраняет кэш в файл"""
    try:
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, indent=2)
    except:
        pass

# --- ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ ОЧИСТКИ МУСОРА ---
def clean_raw_value(val):
    v = safe_str(val)
    if not v or v.lower() == 'nan' or v == 'НЕТУ': 
        return ""
    
    if 'код упаковки:' in v.lower():
        parts = v.split(':', 1)
        if len(parts) > 1:
            v = parts[1].strip()
            
    lower_v = v.lower()
    if 'код маркировки' in lower_v or 'код упаковки' in lower_v or 'артикул' in lower_v:
        return ""
        
    if not any(char.isdigit() for char in v):
        return ""
        
    return v

# --- ТОЧЕЧНОЕ УДАЛЕНИЕ СКОБОК ТОЛЬКО ДЛЯ 00, 01, 21 ---
def remove_specific_brackets(val):
    if pd.isna(val):
        return ""
    v_str = safe_str(val)
    if v_str.lower() == "nan" or v_str == "":
        return ""
    v_str = re.sub(r'\((00|01|21)\)', r'\1', v_str)
    return v_str

# =================================================================
# РЕЖИМ 1: ОСНОВНАЯ ВЕРИФИКАЦИЯ (Сверка по articles.txt)
# =================================================================
def process_customs_data(art_file, c_folder, out_dir, log_callback, progress_callback, stats_callback, check_boxes=True, remove_brackets=False, invoice_prefix="", cancel_callback=None, dry_run=False, file_callback=None):
    wb_import = Workbook()
    ws_import = wb_import.active
    ws_import.title = "УралИмпорт"
    curr_import = 1

    wb_trade = Workbook()
    ws_trade = wb_trade.active
    ws_trade.title = "УралТрейд"
    curr_trade = 1

    # Инициализация файла-рассписки (отключено)
    # wb_receipt = Workbook()
    # ws_receipt = wb_receipt.active
    # ws_receipt.title = "Рассписка"
    # ws_receipt.append(["Артикул", "План (шт)", "Факт (шт)", "Принадлежность"])

    not_found_articles = []
    mismatched_articles = []
    duplicate_summary = {}
    article_routing_summary = []
    global_seen_codes = {}

    # Кандидаты для КЛИН-версии (коды сходятся, коробки не сходятся)
    klin_candidates = []

    with open(art_file, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    # 1. АГРЕГАЦИЯ ТЕКСТОВОГО ФАЙЛА (Суммируем одинаковые артикулы)
    articles_dict = {}
    ok_counter = 0
    bad_counter = 0
    
    for line in lines:
        parts = line.split()
        if len(parts) < 6:
            art_name = parts[0] if parts else "НЕИЗВЕСТНО"
            log_callback(f"[ПРОПУСК] Неверный формат строки: {line[:20]}...")
            bad_counter += 1
            continue

        article = parts[0]
        unit = parts[-1]
        try:
            quantity_int = int(parts[-2])
        except ValueError:
            log_callback(f"[ОШИБКА] Неверное число количества у {article}")
            bad_counter += 1
            continue

        try:
            cargo_places_int = int(parts[-3])
        except ValueError:
            cargo_places_int = 0

        description = parts[-4]
        name = " ".join(parts[1:-4])

        if article not in articles_dict:
            articles_dict[article] = {
                'name': name,
                'desc': description,
                'qty': 0,
                'boxes': 0,
                'unit': unit
            }
        
        # Суммируем количество и коробки для дублирующихся строк
        articles_dict[article]['qty'] += quantity_int
        articles_dict[article]['boxes'] += cargo_places_int

    total_articles = len(articles_dict)
    stats_callback("total", total_articles)
    total_dups_counter = 0
    
    all_files = os.listdir(c_folder)
    unused_files = {f for f in all_files if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')}

    # 2. УМНЫЙ ПОИСК ФАЙЛОВ ДЛЯ КАЖДОГО АРТИКУЛА
    file_to_article = {}
    for file_name in unused_files.copy():
        name_without_ext = os.path.splitext(file_name)[0]
        
        best_match = None
        for art in articles_dict.keys():
            if name_without_ext.startswith(art):
                if best_match is None or len(art) > len(best_match):
                    best_match = art
                    
        if best_match:
            if best_match not in file_to_article:
                file_to_article[best_match] = []
            file_to_article[best_match].append(file_name)
            unused_files.remove(file_name)

    # 3. ОСНОВНОЙ ЦИКЛ ОБРАБОТКИ
    for index, (article, adata) in enumerate(articles_dict.items()):
        # Проверка на отмену операции
        if cancel_callback and cancel_callback():
            log_callback("[⚠] Операция отменена пользователем.")
            return

        # Уведомляем о текущем артикуле
        if file_callback:
            file_callback(f"Артикул {article}")

        name = adata['name']
        description = adata['desc']
        quantity_int = adata['qty']
        cargo_places_int = adata['boxes']
        unit = adata['unit']

        matched_files_names = file_to_article.get(article, [])
        matched_files_names.sort(key=lambda x: (os.path.splitext(x)[0] != article, x))
        matched_files = [os.path.join(c_folder, fn) for fn in matched_files_names]

        if not matched_files:
            not_found_articles.append(article)
            article_routing_summary.append((article, "❌ Файлы не найдены"))
            log_callback(f"[ПРОПУСК] ... Артикул {article}: Файлы декларации НЕ найдены (План: {quantity_int} шт)")

            # Запись в рассписку (отключено)
            # ws_receipt.append([article, quantity_int, 0, "НЕТ ФАЙЛОВ"])

            bad_counter += 1
            stats_callback("bad", bad_counter)
            progress_callback(index + 1, total_articles)
            continue

        file_data_list = []
        for file_path in matched_files:
            try:
                df = pd.read_excel(file_path, header=None)
                # Извлекаем чистое имя без расширения .xlsx
                file_name_short = os.path.splitext(os.path.basename(file_path))[0]
                num_cols = df.shape[1]

                f_type = "import"
                f_codes_raw = []

                if num_cols >= 4:
                    f_type = "trade"
                    last_pack_code = "БЕЗ_КОДА_УПАКОВКИ"
                    for idx_row in range(len(df)):
                        mark_val = clean_raw_value(df.iloc[idx_row, 1])
                        raw_pack = df.iloc[idx_row, 3]
                        pack_val = safe_str(raw_pack)

                        if pack_val and pack_val.lower() != 'nan' and not pack_val.lower().startswith('код'):
                            last_pack_code = pack_val

                        if not mark_val or len(mark_val) < 10: continue
                        if remove_brackets: mark_val = remove_specific_brackets(mark_val)
                        if mark_val.startswith("(00)") or mark_val.startswith("00"): continue

                        f_codes_raw.append((mark_val, last_pack_code))
                else:
                    if len(df) > 0:
                        for col_idx in range(df.shape[1]):
                            col_has_codes = False
                            for val in df.iloc[:, col_idx].dropna():
                                v = clean_raw_value(val)
                                if v and len(v) >= 10:
                                    col_has_codes = True
                                    if remove_brackets: v = remove_specific_brackets(v)
                                    f_codes_raw.append(v)
                            if col_has_codes: break

                file_data_list.append({'name': file_name_short, 'type': f_type, 'raw_codes': f_codes_raw})
            except Exception as e:
                log_callback(f"[ОШИБКА ЧТЕНИЯ] файла {os.path.basename(file_path)}. Детали: {e}")

        trade_groups = {}  
        import_groups = {} 
        files_loaded_names = []
        dup_count_for_article = 0
        local_art_seen = set()
        
        trade_codes_count = 0
        import_codes_count = 0
        total_codes_count = 0
        actual_packages_set = set()

        for fdata in file_data_list:
            current_file_name = fdata['name']
            files_loaded_names.append(current_file_name)

            if fdata['type'] == 'trade':
                if current_file_name not in trade_groups: trade_groups[current_file_name] = {}
                for mark_val, pkg_code in fdata['raw_codes']:
                    if mark_val in local_art_seen or mark_val in global_seen_codes:
                        dup_count_for_article += 1
                        total_dups_counter += 1
                        stats_callback("dups", total_dups_counter)
                    else:
                        local_art_seen.add(mark_val)
                        global_seen_codes[mark_val] = current_file_name
                        trade_codes_count += 1
                        total_codes_count += 1

                        if not pkg_code.startswith("БЕЗ_КОДА"): actual_packages_set.add(pkg_code)
                        if pkg_code not in trade_groups[current_file_name]: trade_groups[current_file_name][pkg_code] = []
                        trade_groups[current_file_name][pkg_code].append(mark_val)
            else:
                if current_file_name not in import_groups: import_groups[current_file_name] = []
                for mark_val in fdata['raw_codes']:
                    if mark_val in local_art_seen or mark_val in global_seen_codes:
                        dup_count_for_article += 1
                        total_dups_counter += 1
                        stats_callback("dups", total_dups_counter)
                    else:
                        local_art_seen.add(mark_val)
                        global_seen_codes[mark_val] = current_file_name
                        import_codes_count += 1
                        total_codes_count += 1
                        import_groups[current_file_name].append(mark_val)

        if dup_count_for_article > 0: duplicate_summary[article] = dup_count_for_article

        actual_packages_count = len(actual_packages_set)

        status_parts = []
        diff_parts = []  # Для вывода расхождений

        if trade_groups and import_groups:
            type_label = "Смешанный (Трейд+Импорт)"
            status_parts.append(f"Штук: факт {total_codes_count}, план {quantity_int} (Трейд: {trade_codes_count}, Импорт: {import_codes_count})")
            receipt_firm = f"СМЕШАННЫЙ (Трейд: {trade_codes_count}, Импорт: {import_codes_count})"

            # Расхождение по штукам
            pieces_diff = total_codes_count - quantity_int
            if pieces_diff != 0:
                sign = "+" if pieces_diff > 0 else ""
                diff_parts.append(f"Штук: {sign}{pieces_diff}")

            if check_boxes and cargo_places_int > 0:
                status_parts.append(f"Коробок(Трейд): факт {actual_packages_count}, план {cargo_places_int}")
                boxes_diff = actual_packages_count - cargo_places_int
                if boxes_diff != 0:
                    sign = "+" if boxes_diff > 0 else ""
                    diff_parts.append(f"Коробок: {sign}{boxes_diff}")
        elif trade_groups:
            type_label = "УралТрейд"
            status_parts.append(f"Штук: факт {trade_codes_count}, план {quantity_int}")
            receipt_firm = "УРАЛТРЕЙД"

            # Расхождение по штукам
            pieces_diff = trade_codes_count - quantity_int
            if pieces_diff != 0:
                sign = "+" if pieces_diff > 0 else ""
                diff_parts.append(f"Штук: {sign}{pieces_diff}")

            if check_boxes and cargo_places_int > 0:
                status_parts.append(f"Коробок: факт {actual_packages_count}, план {cargo_places_int}")
                boxes_diff = actual_packages_count - cargo_places_int
                if boxes_diff != 0:
                    sign = "+" if boxes_diff > 0 else ""
                    diff_parts.append(f"Коробок: {sign}{boxes_diff}")
        else:
            type_label = "УралИмпорт"
            status_parts.append(f"Штук: факт {import_codes_count}, план {quantity_int}")
            receipt_firm = "УРАЛИМПОРТ"

            # Расхождение по штукам
            pieces_diff = import_codes_count - quantity_int
            if pieces_diff != 0:
                sign = "+" if pieces_diff > 0 else ""
                diff_parts.append(f"Штук: {sign}{pieces_diff}")

        diff_str = " | ".join(status_parts)
        diff_summary = " | ".join(diff_parts) if diff_parts else ""

        # Запись в файл-рассписку (отключено)
        # ws_receipt.append([article, quantity_int, total_codes_count, receipt_firm])

        if check_boxes and cargo_places_int > 0 and trade_groups:
            is_mismatch = (total_codes_count != quantity_int) or (actual_packages_count != cargo_places_int)
        else:
            is_mismatch = (total_codes_count != quantity_int)

        if is_mismatch:
            # Формируем сообщение с расхождением
            mismatch_msg = f"⚠️ Расхождение ({diff_str})"
            if diff_summary:
                mismatch_msg += f" | Разница: {diff_summary}"

            log_callback(f"[{type_label}] Артикул {article}: РАСХОЖДЕНИЕ! ({diff_str}). Разница: {diff_summary}. Но данные добавлены.")
            article_routing_summary.append((article, mismatch_msg))
            mismatched_articles.append({
                "article": article, "expected": quantity_int, "actual": total_codes_count,
                "diff": abs(quantity_int - total_codes_count), "status": "Расхождение", "files": ", ".join(files_loaded_names)
            })

            # КЛИН-кандидаты: коды сходятся, коробки не сходятся, это Трейд
            if trade_groups and check_boxes and cargo_places_int > 0:
                codes_match = (total_codes_count == quantity_int)
                boxes_mismatch = (actual_packages_count != cargo_places_int)

                if codes_match and boxes_mismatch:
                    # Собираем все коды маркировки и упаковки для этого артикула
                    all_mark_codes = []
                    all_pack_codes = []

                    for file_name, p_groups in trade_groups.items():
                        for pkg_code, codes_list in p_groups.items():
                            all_mark_codes.extend(codes_list)
                            if not pkg_code.startswith("БЕЗ_КОДА"):
                                all_pack_codes.append(pkg_code)

                    klin_candidates.append({
                        "article": article,
                        "name": name,
                        "description": description,
                        "quantity_plan": quantity_int,
                        "quantity_fact": total_codes_count,
                        "boxes_plan": cargo_places_int,
                        "boxes_fact": actual_packages_count,
                        "mark_codes": all_mark_codes,
                        "pack_codes": all_pack_codes,
                        "unit": unit
                    })

            bad_counter += 1
            stats_callback("bad", bad_counter)
        else:
            article_routing_summary.append((article, f"✅ {type_label} (Совпало)"))
            log_callback(f"[{type_label}] ... Артикул {article}: OK ({diff_str})")
            ok_counter += 1
            stats_callback("ok", ok_counter)

        # Проверяем, является ли это KLIN-кандидатом (коды сходятся, коробки не сходятся)
        is_klin_candidate = any(k['article'] == article for k in klin_candidates)

        if trade_groups:
            for file_name, p_groups in trade_groups.items():
                for pkg_code, codes_list in p_groups.items():
                    if not codes_list: continue
                    ws_trade[f"A{curr_trade}"] = "Код маркировки"
                    ws_trade[f"B{curr_trade}"] = file_name # Теперь без .xlsx
                    ws_trade[f"C{curr_trade}"] = name
                    ws_trade[f"D{curr_trade}"] = description
                    ws_trade[f"E{curr_trade}"] = str(cargo_places_int)
                    ws_trade[f"F{curr_trade}"] = diff_str
                    ws_trade[f"G{curr_trade}"] = str(quantity_int)
                    ws_trade[f"H{curr_trade}"] = unit

                    # Если это KLIN-кандидат, пишем специальную надпись вместо кода упаковки
                    if is_klin_candidate:
                        ws_trade[f"I{curr_trade}"] = "ДЕКЛАРИРУЕМ ТОЛЬКО КОДЫ"
                    else:
                        display_pkg_code = remove_specific_brackets(pkg_code) if remove_brackets else pkg_code
                        ws_trade[f"I{curr_trade}"] = f"Код упаковки: {display_pkg_code}"

                    curr_trade += 1
                    for code in codes_list:
                        ws_trade[f"A{curr_trade}"] = code
                        curr_trade += 1
                    curr_trade += 1

        if import_groups:
            for file_name, codes_list in import_groups.items():
                if not codes_list: continue
                ws_import[f"A{curr_import}"] = "Код маркировки"
                ws_import[f"B{curr_import}"] = file_name # Теперь без .xlsx
                ws_import[f"C{curr_import}"] = name
                ws_import[f"D{curr_import}"] = description
                ws_import[f"E{curr_import}"] = str(cargo_places_int)
                ws_import[f"F{curr_import}"] = diff_str
                ws_import[f"G{curr_import}"] = str(quantity_int)
                ws_import[f"H{curr_import}"] = unit
                curr_import += 1
                for code in codes_list:
                    ws_import[f"A{curr_import}"] = code
                    curr_import += 1
                curr_import += 1 

        progress_callback(index + 1, total_articles)

    # Выравнивание ширины столбцов
    for work_sheet in [ws_import, ws_trade]:  # Убрали ws_receipt
        for col in work_sheet.columns:
            max_length = max(len(str(cell.value or '')) for cell in col)
            column = col[0].column_letter
            for cell in col: cell.alignment = Alignment(vertical="top")
            work_sheet.column_dimensions[column].width = min(max_length + 5, 60)

    # ДОБАВЛЕНО: Формирование префикса для имен файлов
    prefix = f"{invoice_prefix.strip()} " if invoice_prefix.strip() else ""

    out_import = os.path.join(out_dir, f"{prefix}УралИмпорт.xlsx")
    out_trade = os.path.join(out_dir, f"{prefix}УралТрейд.xlsx")
    # out_receipt = os.path.join(out_dir, f"{prefix}рассписка.xlsx")  # Отключено

    files_created = []

    # Сохраняем файлы только если НЕ dry_run режим
    if not dry_run:
        if curr_import > 1:
            wb_import.save(out_import)
            files_created.append(out_import)
        if curr_trade > 1:
            wb_trade.save(out_trade)
            files_created.append(out_trade)

        # Сохраняем рассписку всегда (ОТКЛЮЧЕНО)
        # wb_receipt.save(out_receipt)
        # files_created.append(out_receipt)
    else:
        log_callback("\n[ℹ] РЕЖИМ ПРОВЕРКИ: Файлы не были сохранены.")
        if curr_import > 1:
            files_created.append(f"(предпросмотр) {os.path.basename(out_import)}")
        if curr_trade > 1:
            files_created.append(f"(предпросмотр) {os.path.basename(out_trade)}")
        # files_created.append(f"(предпросмотр) {os.path.basename(out_receipt)}")  # Отключено


    log_callback("\n============================================================")
    log_callback("ФИНАЛЬНЫЙ ОТЧЕТ ПО РАСХОЖДЕНИЯМ:")
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

    log_callback("============================================================")
    log_callback(f"📊 ОБЩАЯ СУММА УНИКАЛЬНЫХ КОДОВ (ТРЕЙД + ИМПОРТ): {len(global_seen_codes)}")
    log_callback("============================================================\n")

    log_callback("============================================================")
    log_callback("НЕИСПОЛЬЗОВАННЫЕ ФАЙЛЫ В ПАПКЕ (ОСТАЛИСЬ НЕВЕРИФИЦИРОВАННЫМИ):")
    log_callback("============================================================")
    if unused_files:
        for f_unused in sorted(unused_files):
            log_callback(f" ⚠️ {f_unused}")
    else:
        log_callback(" 🎉 Отлично! Все Excel-файлы из папки были успешно задействованы.")
    log_callback("============================================================\n")

    # Возвращаем кандидатов для КЛИН-версии (коды сходятся, коробки не сходятся)
    return klin_candidates


# =================================================================
# ГЕНЕРАЦИЯ КЛИН-ФИЛЬА (CLEAN_версия)
# =================================================================
def create_klin_file(selected_articles, out_dir, invoice_prefix="", remove_brackets=False, log_callback=None):
    """
    Создает файл CLEAN_УралТрейд.xlsx с 3 вкладками:
    1. Коды маркировки - по артикулам
    2. Коды упаковки - по артикулам
    3. ТОЛЬКО КОДЫ - все коды маркировки без артикулов
    selected_articles: список словарей из klin_candidates
    Возвращает путь к созданному файлу или None.
    """
    if not selected_articles:
        return None

    from openpyxl.styles import Alignment

    wb = Workbook()

    # Удаляем дефолтный лист
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    ws_mark = wb.create_sheet(title="Коды маркировки", index=0)
    ws_pack = wb.create_sheet(title="Коды упаковки", index=1)
    ws_only_codes = wb.create_sheet(title="ТОЛЬКО КОДЫ", index=2)

    # Лист 1: Коды маркировки (по артикулам)
    row_mark = 1
    total_marks = 0
    for art in selected_articles:
        # Заголовок артикула
        ws_mark[f"A{row_mark}"] = f"Артикул: {art['article']} | {art['name']}"
        ws_mark[f"B{row_mark}"] = f"Коробок: план {art['boxes_plan']}, факт {art['boxes_fact']}"
        row_mark += 1

        # Коды маркировки
        for code in art["mark_codes"]:
            display = remove_specific_brackets(code) if remove_brackets else code
            ws_mark[f"A{row_mark}"] = display
            row_mark += 1

        # Пустая строка между артикулами
        row_mark += 1
        total_marks += len(art["mark_codes"])

    # Лист 2: Коды упаковки (по артикулам)
    row_pack = 1
    total_packs = 0
    for art in selected_articles:
        # Заголовок артикула
        ws_pack[f"A{row_pack}"] = f"Артикул: {art['article']} | {art['name']}"
        row_pack += 1

        # Коды упаковки
        if art["pack_codes"]:
            for code in art["pack_codes"]:
                display = remove_specific_brackets(code) if remove_brackets else code
                ws_pack[f"A{row_pack}"] = display
                row_pack += 1
        else:
            ws_pack[f"A{row_pack}"] = "БЕЗ КОДА УПАКОВКИ"
            row_pack += 1

        # Пустая строка между артикулами
        row_pack += 1
        total_packs += len(art["pack_codes"])

    # Лист 3: ТОЛЬКО КОДЫ (все коды маркировки без артикулов)
    row_only = 1
    for art in selected_articles:
        for code in art["mark_codes"]:
            display = remove_specific_brackets(code) if remove_brackets else code
            ws_only_codes[f"A{row_only}"] = display
            row_only += 1

    # Форматирование
    for work_sheet in [ws_mark, ws_pack, ws_only_codes]:
        for col in work_sheet.columns:
            max_length = max(len(str(cell.value or '')) for cell in col)
            column = col[0].column_letter
            for cell in col:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            work_sheet.column_dimensions[column].width = min(max_length + 5, 60)

    prefix = f"{invoice_prefix.strip()} " if invoice_prefix.strip() else ""
    out_path = os.path.join(out_dir, f"{prefix}CLEAN_УралТрейд.xlsx")

    wb.save(out_path)

    if log_callback:
        log_callback(f"  ✅ Создан: {os.path.basename(out_path)} (Маркировок: {total_marks}, Упаковок: {total_packs})")

    return out_path


# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ФОРМАТТЕРА ---
def extract_trade_codes_for_cleaner(df, remove_brackets=False):
    """
    Для Форматтера: извлекает коды маркировки и упаковки из Трейд файлов.
    Возвращает список пар (код_маркировки, код_упаковки).
    """
    result = []
    if len(df) == 0:
        return result

    for idx_row in range(len(df)):
        mark_val = clean_raw_value(df.iloc[idx_row, 1])
        raw_pack = df.iloc[idx_row, 3] if df.shape[1] > 3 else None
        pack_val = safe_str(raw_pack) if raw_pack is not None else ""

        if mark_val and len(mark_val) >= 10:
            if remove_brackets:
                mark_val = remove_specific_brackets(mark_val)
            if not mark_val.startswith("(00)") and not mark_val.startswith("00"):
                pack_code = pack_val if pack_val and pack_val.lower() != 'nan' and not pack_val.lower().startswith('код') else "БЕЗ_КОДА_УПАКОВКИ"
                result.append((mark_val, pack_code))

    return result


def extract_import_codes_for_cleaner(df, remove_brackets=False):
    """
    Для Форматтера: извлекает коды маркировки из Импорт файлов.
    Возвращает список кодов.
    """
    result = []
    if len(df) == 0:
        return result

    for col_idx in range(df.shape[1]):
        col_has_codes = False
        for val in df.iloc[:, col_idx].dropna():
            v = clean_raw_value(val)
            if v and len(v) >= 10:
                col_has_codes = True
                if remove_brackets:
                    v = remove_specific_brackets(v)
                if not v.startswith("(00)") and not v.startswith("00"):
                    result.append(v)
        if col_has_codes:
            break

    return result


# =================================================================
# РЕЖИМ 2: БЫСТРАЯ ОЧИСТКА (Форматтер / Клинер)
# =================================================================
def process_cleaner_mode(c_folder, out_dir, log_callback, progress_callback, stats_callback, remove_brackets=False, invoice_prefix="", cancel_callback=None, dry_run=False, file_callback=None, exclude_articles=None):
    all_files = [f for f in os.listdir(c_folder) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]

    # Исключаем файлы, содержащие названия артикулов из exclude_articles
    if exclude_articles:
        all_files = [f for f in all_files if not any(art in f for art in exclude_articles)]

    trade_files = [f for f in all_files if 'трейд' in f.lower() and 'clean' not in f.lower() and 'декларации' not in f.lower()]
    import_files = [f for f in all_files if 'импорт' in f.lower() and 'clean' not in f.lower() and 'декларации' not in f.lower()]

    total_files = len(trade_files) + len(import_files)
    stats_callback("total", total_files)
    
    if total_files == 0:
        log_callback("⚠️ В папке не найдены файлы, содержащие слова 'трейд' или 'импорт' (или они уже очищены).")
        return

    global_seen_marks = set()
    global_seen_packs = set()
    total_duplicates_removed = 0
    processed_count = 0

    # Создаем папку для Риммы
    rimma_dir = os.path.join(out_dir, "Для декларации")
    os.makedirs(rimma_dir, exist_ok=True)

    # ДОБАВЛЕНО: Префикс из GUI
    prefix = f"{invoice_prefix.strip()} " if invoice_prefix.strip() else ""

    # ------------------ ОБРАБОТКА ТРЕЙДА ------------------
    for f_name in trade_files:
        # Проверка на отмену операции
        if cancel_callback and cancel_callback():
            log_callback("[⚠] Операция отменена пользователем.")
            return

        # Уведомляем о текущем файле
        if file_callback:
            file_callback(f_name)

        log_callback(f"⏳ Обработка ТРЕЙД: {f_name}")
        file_path = os.path.join(c_folder, f_name)
        df = pd.read_excel(file_path, header=None)

        mark_codes = []
        pack_codes = []
        local_dups = 0

        codes_pairs = extract_trade_codes_for_cleaner(df, remove_brackets)

        for mark_code, pack_code in codes_pairs:
            # Проверяем дубликаты маркировки - это ГЛАВНЫЙ счетчик
            if mark_code in global_seen_marks:
                local_dups += 1
                total_duplicates_removed += 1
            else:
                global_seen_marks.add(mark_code)
                mark_codes.append(mark_code)

            # Проверяем дубликаты упаковки - отдельно
            if pack_code != "БЕЗ_КОДА_УПАКОВКИ":
                if pack_code not in global_seen_packs:
                    global_seen_packs.add(pack_code)
                    pack_codes.append(pack_code)

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

        # ДОБАВЛЕНО: убираем дублирование инвойса из имени
        base_f_name = f_name
        if invoice_prefix.strip() and base_f_name.startswith(invoice_prefix.strip()):
            base_f_name = base_f_name[len(invoice_prefix.strip()):].lstrip(" _-")

        out_name = f"{prefix}CLEAN_{base_f_name}"

        if not dry_run:
            wb.save(os.path.join(out_dir, out_name))
            log_callback(f"  ✅ Создан: {out_name} (Маркировок: {len(mark_codes)}, Упаковок: {len(pack_codes)} | Удалено дублей: {local_dups})")
        else:
            log_callback(f"  🔍 Проверка: {out_name} (Маркировок: {len(mark_codes)}, Упаковок: {len(pack_codes)} | Удалено дублей: {local_dups})")

        # --- ФАЙЛ ДЛЯ РИММЫ (УРАЛ ТРЕЙД) ---
        wb_rimma = Workbook()
        ws_rimma = wb_rimma.active
        ws_rimma.title = "Для декларации"
        
        for r_idx in range(len(df)):
            row_data = []
            for c_idx in range(df.shape[1]):
                cell_val = df.iloc[r_idx, c_idx]
                if pd.isna(cell_val):
                    row_data.append("")
                else:
                    # Применяем умную конвертацию в строку (чтобы не было 1200.0)
                    str_val = safe_str(cell_val)
                    str_val = re.sub(r'(?i)код упаковки\s*:\s*', '', str_val).strip()
                    row_data.append(str_val)
            ws_rimma.append(row_data)

        for col in ws_rimma.columns:
            max_length = max(len(str(cell.value or '')) for cell in col)
            ws_rimma.column_dimensions[col[0].column_letter].width = min(max_length + 5, 50)

        # ДОБАВЛЕНО: Префикс для файла декларации
        base_name = os.path.splitext(base_f_name)[0]
        out_rimma = f"{prefix}{base_name}_для_декларации.xlsx"

        # Сохраняем в отдельную папку!
        if not dry_run:
            wb_rimma.save(os.path.join(rimma_dir, out_rimma))
            log_callback(f"  📁 Создан: {out_rimma} (В папке 'Для декларации')\n")
        else:
            log_callback(f"  🔍 Проверка: {out_rimma} (В папке 'Для декларации')\n")
        
        processed_count += 1
        progress_callback(processed_count, total_files)
        stats_callback("ok", processed_count)
        stats_callback("dups", total_duplicates_removed)

    # ------------------ ОБРАБОТКА ИМПОРТА ------------------
    for f_name in import_files:
        # Проверка на отмену операции
        if cancel_callback and cancel_callback():
            log_callback("[⚠] Операция отменена пользователем.")
            return

        # Уведомляем о текущем файле
        if file_callback:
            file_callback(f_name)

        log_callback(f"⏳ Обработка ИМПОРТ: {f_name}")
        file_path = os.path.join(c_folder, f_name)
        df = pd.read_excel(file_path, header=None)

        mark_codes = []
        local_dups = 0

        codes = extract_import_codes_for_cleaner(df, remove_brackets)

        for code in codes:
            if code not in global_seen_marks:
                global_seen_marks.add(code)
                mark_codes.append(code)
            else:
                local_dups += 1
                total_duplicates_removed += 1

        wb = Workbook()
        ws_mark = wb.active
        ws_mark.title = "Коды маркировки"
        for idx, code in enumerate(mark_codes, start=1):
            ws_mark[f"A{idx}"] = code
        ws_mark.column_dimensions['A'].width = 45

        # ДОБАВЛЕНО: убираем дублирование инвойса из имени
        base_f_name = f_name
        if invoice_prefix.strip() and base_f_name.startswith(invoice_prefix.strip()):
            base_f_name = base_f_name[len(invoice_prefix.strip()):].lstrip(" _-")

        out_name = f"{prefix}CLEAN_{base_f_name}"

        if not dry_run:
            wb.save(os.path.join(out_dir, out_name))
            log_callback(f"  ✅ Создан: {out_name} (Маркировок: {len(mark_codes)} | Удалено дублей: {local_dups})")
        else:
            log_callback(f"  🔍 Проверка: {out_name} (Маркировок: {len(mark_codes)} | Удалено дублей: {local_dups})")

        # --- ФАЙЛ ДЛЯ РИММЫ (УРАЛ ИМПОРТ) ---
        wb_rimma = Workbook()
        ws_rimma = wb_rimma.active
        ws_rimma.title = "Для декларации"
        
        ws_rimma.append(["Код маркировки", "артикул", "наименование", "артикул 2", "места", "кол", "шт"])
        
        current_art = ""
        current_name = ""
        current_desc = ""
        current_places = ""
        current_qty = ""
        current_unit = ""
        is_first_code = False
        
        for r_idx in range(len(df)):
            # Применяем умную конвертацию в строку (отсекаем .0 у чисел)
            row_vals = [safe_str(x) for x in df.iloc[r_idx, :]]
            
            while row_vals and not row_vals[-1]:
                row_vals.pop()
                
            if not row_vals:
                continue
                
            val1 = row_vals[0].lower()
            
            if "код маркировки" in val1:
                current_art = row_vals[1] if len(row_vals) > 1 else ""
                current_name = row_vals[2] if len(row_vals) > 2 else ""
                current_desc = row_vals[3] if len(row_vals) > 3 else ""
                current_places = row_vals[4] if len(row_vals) > 4 else ""
                current_qty = row_vals[6] if len(row_vals) > 6 else ""
                current_unit = row_vals[7] if len(row_vals) > 7 else ""
                is_first_code = True
                continue
                
            code_val = row_vals[0]
            if code_val:
                if is_first_code:
                    ws_rimma.append([
                        code_val, current_art, current_name, current_desc, 
                        current_places, current_qty, current_unit
                    ])
                    is_first_code = False
                else:
                    ws_rimma.append([
                        code_val, current_art, current_name, "", "", "", ""
                    ])

        for col in ws_rimma.columns:
            max_length = max(len(str(cell.value or '')) for cell in col)
            ws_rimma.column_dimensions[col[0].column_letter].width = min(max_length + 5, 50)

        # ДОБАВЛЕНО: Префикс для файла декларации
        base_name = os.path.splitext(base_f_name)[0]
        out_rimma = f"{prefix}{base_name}_для_декларации.xlsx"

        # Сохраняем в отдельную папку!
        if not dry_run:
            wb_rimma.save(os.path.join(rimma_dir, out_rimma))
            log_callback(f"  📁 Создан: {out_rimma} (В папке 'Для декларации')\n")
        else:
            log_callback(f"  🔍 Проверка: {out_rimma} (В папке 'Для декларации')\n")

        processed_count += 1
        progress_callback(processed_count, total_files)
        stats_callback("ok", processed_count)
        stats_callback("dups", total_duplicates_removed)

    log_callback(f"\n📊 ОБЩАЯ СУММА ОЧИЩЕННЫХ УНИКАЛЬНЫХ КОДОВ ЗА СЕССИЮ: {len(global_seen_marks)}")
    log_callback(f"🎉 ВСЕ ФАЙЛЫ ОЧИЩЕНЫ! (Всего отсеяно дубликатов за сессию: {total_duplicates_removed})")