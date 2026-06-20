import os
import json
import glob
import re
import pandas as pd

def parse_float(val):
    """Безопасное преобразование строки в float"""
    if val is None or str(val).strip().lower() in ('null', 'none', ''):
        return None
    try:
        return float(str(val).replace(',', '.'))
    except ValueError:
        return None

def format_float(val):
    """Форматирование float в строку с запятой"""
    if val is None:
        return "----"
    if val < 0.001 or val > 1000:
        return f"{val:.2e}".replace('.', ',')
    else:
        return f"{val:g}".replace('.', ',')

def load_pdk_data(json_filepath):
    """Загрузка справочника ПДК из JSON"""
    if not os.path.exists(json_filepath):
        raise FileNotFoundError(f"Файл '{json_filepath}' не найден.")
    
    with open(json_filepath, 'r', encoding='utf-8') as f:
        substances = json.load(f)
    
    pdk_lookup = {}
    for sub in substances:
        code = str(sub.get('code', '')).strip()
        if len(code) >= 4 and code[:4].isdigit():
            code = code[:4]
            pdk_lookup[code] = {
                'pdk_ss': parse_float(sub.get('pdk_ss')),
                'pdk_sg': parse_float(sub.get('pdk_sg'))
            }
    return pdk_lookup

def recalculate_cell(cell_value, code, pdk_lookup):
    """Пересчет доли ПДК в ячейке"""
    if pd.isna(cell_value) or str(cell_value).strip() in ('', '----', '—', '-', 'null'):
        return cell_value
    
    cell_str = str(cell_value).strip()
    
    if '/' in cell_str:
        parts = cell_str.split('/')
        part1 = parts[0].strip()
        part2 = parts[1].strip()
        has_slash = True
    else:
        part1 = cell_str
        part2 = None
        has_slash = False

    pdk_data = pdk_lookup.get(code, {})
    pdk_ss = pdk_data.get('pdk_ss')
    pdk_sg = pdk_data.get('pdk_sg')

    def process_part(part):
        if part in ('----', '—', '-', ''):
            return part
        val = parse_float(part)
        if val is None:
            return part
        
        # Если ПДКсс нет — не пересчитываем (вещество будет отфильтровано)
        if pdk_ss is None:
            return part
        
        # Если ПДКсг есть, пересчитываем: old_share * (ПДКсг / ПДКсс)
        if pdk_sg is not None and pdk_sg > 0:
            new_val = val * (pdk_sg / pdk_ss)
        else:
            # Если ПДКсг нет, программа уже поделила на ПДКсс
            new_val = val
            
        return format_float(new_val)

    new_part1 = process_part(part1)
    
    if has_slash and part2 is not None:
        new_part2 = process_part(part2)
        return f"{new_part1} / {new_part2}"
    else:
        return new_part1

def extract_code(row):
    """Извлечение 4-значного кода вещества"""
    val = str(row.iloc[0])
    match = re.search(r'^(\d{4})', val.strip())
    if match:
        return match.group(1)
    return None

def has_pdk_ss(code, pdk_lookup):
    """Проверка наличия ПДКсс у вещества"""
    pdk_data = pdk_lookup.get(code, {})
    return pdk_data.get('pdk_ss') is not None

def main():
    # 1. Загрузка справочника
    json_file = 'output.json'
    print(f"Загрузка справочника из {json_file}...")
    pdk_lookup = load_pdk_data(json_file)
    print(f"Загружено нормативов для {len(pdk_lookup)} веществ.")

    # 2. Поиск входного Excel-файла
    excel_files = [f for f in glob.glob('*.xls*') if not f.startswith('recalculated_') and not f.startswith('~')]
    if not excel_files:
        raise FileNotFoundError("Не найден входной файл Excel.")
    
    input_file = excel_files[0]
    print(f"Найден входной файл: {input_file}")

    # 3. Чтение данных
    print("Чтение данных из Excel...")
    df = pd.read_excel(input_file, header=None)

    # 4. Фильтрация: оставляем только вещества с ПДКсс
    print("Фильтрация веществ (только с ПДКсс)...")
    filtered_rows = []
    for idx, row in df.iterrows():
        code = extract_code(row)
        if code is None:
            # Шапка или пустая строка — пропускаем
            continue
        if has_pdk_ss(code, pdk_lookup):
            filtered_rows.append(idx)
    
    df_filtered = df.loc[filtered_rows].reset_index(drop=True)
    print(f"Осталось {len(df_filtered)} строк с веществами, имеющими ПДКсс.")

    # 5. Пересчет столбцов (индексы 2, 3, 4, 5 = колонки 3, 4, 5, 6)
    print("Выполнение пересчета...")
    cols_to_process = [2, 3, 4, 5]
    
    for col in cols_to_process:
        df_filtered[col] = df_filtered.apply(
            lambda row: recalculate_cell(row.iloc[col], extract_code(row), pdk_lookup), 
            axis=1
        )

    # 6. Сохранение результата
    os.makedirs('output', exist_ok=True)
    output_filename = os.path.join('output', f'recalculated_{input_file}')
    
    df_filtered.to_excel(output_filename, index=False, header=False)
    print(f"✅ Готово! Файл сохранен: {output_filename}")
    print(f"📊 Исключено веществ без ПДКсс: {len(df) - len(df_filtered)}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ Ошибка: {e}")