import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment

# =========================================
# НАСТРОЙКИ
# =========================================

ARTICLES_FILE = "articles.txt"
CODES_FOLDER = "codes"
OUTPUT_FILE = "result.xlsx"

# =========================================
# СОЗДАНИЕ EXCEL И СПИСКОВ ДЛЯ ОТЧЕТА
# =========================================

wb = Workbook()
ws = wb.active
ws.title = "Результат"

current_row = 1

# Списки для финального отчета в консоли
not_found_articles = []
mismatched_articles = []

# =========================================
# ЧТЕНИЕ articles.txt
# =========================================

if not os.path.exists(ARTICLES_FILE):
    print(f"[КРИТИЧЕСКАЯ ОШИБКА] Не найден файл {ARTICLES_FILE}")
    exit()

with open(ARTICLES_FILE, "r", encoding="utf-8") as f:
    lines = [line.strip() for line in f if line.strip()]

print("==================================================")
print("ОБРАБОТКА АРТИКУЛОВ:")
print("==================================================")

for line in lines:
    parts = line.split()

    if len(parts) < 6:
        print(f"[ОШИБКА СТРОКИ] Неверный формат: {line}")
        continue

    # =========================================
    # РАЗБОР СТРОКИ ТЕКСТОВИКА
    # =========================================

    article = parts[0]
    unit = parts[-1]
    quantity = parts[-2]
    cargo_places = parts[-3]
    description = parts[-4]

    # Всё между артикулом и описанием = название
    name = " ".join(parts[1:-4])

    try:
        quantity_int = int(quantity)
    except ValueError:
        print(f"[ОШИБКА КОЛИЧЕСТВА] Не число для артикула {article}")
        continue

    print(f"Обработка артикула: {article} ...", end=" ")

    # =========================================
    # ПОИСК ФАЙЛА
    # =========================================

    matched_file = None
    if os.path.exists(CODES_FOLDER):
        for file_name in os.listdir(CODES_FOLDER):
            if article in file_name:
                matched_file = os.path.join(CODES_FOLDER, file_name)
                break

    # Запоминаем строку, где пишется шапка товара
    header_row = current_row

    # Запись информации о товаре (Строка-шапка)
    ws[f"C{header_row}"] = article
    ws[f"D{header_row}"] = name
    ws[f"E{header_row}"] = description
    ws[f"F{header_row}"] = cargo_places
    ws[f"G{header_row}"] = quantity
    ws[f"H{header_row}"] = unit

    current_row += 1

    # =========================================
    # ЕСЛИ ФАЙЛА НЕТ
    # =========================================

    if matched_file is None:
        ws[f"A{current_row}"] = "НЕТУ"
        not_found_articles.append(article)  # Добавляем в список потерянных
        print("ФАЙЛ НЕ НАЙДЕН")
        current_row += 2
        continue

    # =========================================
    # ЧТЕНИЕ КОДОВ ЧЕРЕЗ PANDAS (считаем только заполненные строки)
    # =========================================

    codes = []
    try:
        codes_df = pd.read_excel(matched_file, header=None)
        
        for value in codes_df[0].tolist():
            if pd.isna(value):
                continue  # Пропускаем пустые строки (сверху, снизу, в середине)
                
            cleaned_code = str(value).strip()
            if cleaned_code:  # Проверяем, что в строке есть символы
                codes.append(cleaned_code)
                
    except Exception as e:
        ws[f"A{current_row}"] = "ОШИБКА ЧТЕНИЯ"
        print(f"ОШИБКА ЧТЕНИЯ ФАЙЛА ({e})")
        current_row += 2
        continue

    codes_count = len(codes)

    # =========================================
    # ПРОВЕРКА КОЛИЧЕСТВА И РАСЧЕТ РАЗНИЦЫ
    # =========================================

    if codes_count != quantity_int:
        diff = abs(quantity_int - codes_count)
        status_text = "Меньше" if codes_count < quantity_int else "Больше"
        
        # Запись подробной разницы в Excel (в столбец F под грузовые места)
        ws[f"F{header_row + 1}"] = f"В файле: {codes_count}, {status_text} на {diff}"
        
        # Сохраняем информацию для итогового отчета в консоли
        mismatched_articles.append({
            "article": article,
            "expected": quantity_int,
            "actual": codes_count,
            "diff": diff,
            "status": status_text
        })
        print(f"НЕСОВПАДЕНИЕ (Ожидалось {quantity_int}, по факту {codes_count})")
    else:
        print("ОК (Количество совпало)")

    # =========================================
    # ЗАПИСЬ КОДОВ В СТОЛБЕЦ А
    # =========================================

    if codes_count == 0:
        ws[f"A{current_row}"] = "НЕТУ"
        current_row += 2
        continue

    for code in codes:
        ws[f"A{current_row}"] = code
        current_row += 1

    # Пустая строка между блоками товаров
    current_row += 1

# =========================================
# ШИРИНА СТОЛБЦОВ И ВЫРАВНИВАНИЕ
# =========================================

for col in ws.columns:
    max_length = 0
    column = col[0].column_letter

    for cell in col:
        if cell.value is not None:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        cell.alignment = Alignment(vertical="top")

    adjusted_width = min(max_length + 5, 60)
    ws.column_dimensions[column].width = adjusted_width

# Сохранение итогового файла Excel
wb.save(OUTPUT_FILE)

# =========================================
# ФИНАЛЬНЫЙ ОТЧЕТ В КОНСОЛЬ
# =========================================

print("\n" + "="*50)
print("ФИНАЛЬНЫЙ ОТЧЕТ ПО ОШИБКАМ И РАСХОЖДЕНИЯМ:")
print("="*50)

# 1. Вывод не найденных файлов
if not_found_articles:
    print(f"\n[!] НЕ НАЙДЕНЫ ФАЙЛЫ ДЛЯ АРТИКУЛОВ ({len(not_found_articles)} шт.):")
    for art in not_found_articles:
        print(f"    - Артикул: {art}")
else:
    print("\n[✓] Все файлы артикулов успешно найдены в папке.")

# 2. Вывод расхождений по количеству
if mismatched_articles:
    print(f"\n[!] РАСХОЖДЕНИЕ КОЛИЧЕСТВА КОДОВ ({len(mismatched_articles)} шт.):")
    for item in mismatched_articles:
        print(
            f"    - Артикул {item['article']}: "
            f"В текстовике указано {item['expected']}, "
            f"в файле найдено {item['actual']}. "
            f"Разница: {item['status'].lower()} на {item['diff']} шт."
        )
else:
    print("[✓] Расхождений по количеству кодов не обнаружено.")

print("\n" + "="*50)
print(f"Успешно завершено! Итоговый файл: {OUTPUT_FILE}")
print("="*50)
