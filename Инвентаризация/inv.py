# main.py
import os
import json
import importlib
from pathlib import Path

def find_json_files(root_dir="."):
    """Находит все .json файлы в корневой папке (не в подпапках)."""
    json_files = []
    for file in os.listdir(root_dir):
        if file.endswith(".json") and os.path.isfile(os.path.join(root_dir, file)):
            json_files.append(file)
    return json_files

def load_json(filepath):
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        return json.load(f)

def main():
    # Создаём папку для результатов, если её нет
    Path("output").mkdir(exist_ok=True)
    
    json_files = find_json_files()
    if not json_files:
        print("Нет JSON файлов в текущей папке.")
        return
    
    # Загружаем все модули из папки modules
    modules = []
    for module_file in os.listdir("modules"):
        if module_file.endswith(".py") and not module_file.startswith("__"):
            module_name = module_file[:-3]
            module = importlib.import_module(f"modules.{module_name}")
            modules.append(module)
    
    for json_file in json_files:
        base_name = json_file[:-5]  # "INV_1"
        data = load_json(json_file)
        
        # Создаём папку для этого файла
        output_dir = f"output/{base_name}"
        Path(output_dir).mkdir(exist_ok=True)
        
        # Передаём управление каждому модулю
        for mod in modules:
            if hasattr(mod, "export"):
                print(f"Обработка {json_file} модулем {mod.__name__}...")
                mod.export(data, output_dir, base_name)
            else:
                print(f"В модуле {mod.__name__} нет функции export()")

if __name__ == "__main__":
    main()
