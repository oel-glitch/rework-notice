#!/usr/bin/env python3

import os
import sys
import shutil
import subprocess
from pathlib import Path


def generate_version_file():
    """Generate version.txt file with build information."""
    try:
        from src.version_info import APP_VERSION, get_git_hash, get_build_timestamp
        
        git_hash = get_git_hash()
        build_time = get_build_timestamp()
        
        with open('version.txt', 'w', encoding='utf-8') as f:
            f.write(f"VERSION={APP_VERSION}\n")
            f.write(f"GIT_HASH={git_hash}\n")
            f.write(f"BUILD_TIME={build_time}\n")
        
        print(f"✓ Создан version.txt: v{APP_VERSION} ({git_hash})")
        return True
    except Exception as e:
        print(f"⚠ Не удалось создать version.txt: {e}")
        return False


def build_executable():
    print("=" * 60)
    print("  Сборка .exe файла для PDF Парсера")
    print("=" * 60)
    print()
    
    print("Проверка зависимостей...")
    try:
        import PyInstaller
        print("✓ PyInstaller установлен")
    except ImportError:
        print("✗ PyInstaller не установлен")
        print("Установка PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    print("\nОчистка предыдущих сборок...")
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("main.spec"):
        os.remove("main.spec")
    
    print("\nГенерация файла версии...")
    generate_version_file()
    
    print("\nСборка .exe файла...")
    
    separator = ';' if sys.platform == 'win32' else ':'
    
    cmd = [
        "pyinstaller",
        "--name=PDF_Parser_OATI",
        "--windowed",
        "--onefile",
        f"--add-data=templates{separator}templates",
        f"--add-data=config{separator}config",
        f"--add-data=src{separator}src",
        f"--add-data=version.txt{separator}.",
        f"--add-data=setup_scanner.bat{separator}.",
        f"--add-data=SCANNER_SETUP_RU.txt{separator}.",
        "--hidden-import=tkinterdnd2",
        "--hidden-import=pymorphy3",
        "--hidden-import=pymorphy3.opencorpora_dict",
        "--hidden-import=dawg2",
        "--hidden-import=PyPDF2",
        "--hidden-import=docx",
        "--hidden-import=fitz",
        "--hidden-import=PIL",
        "--hidden-import=zeep",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=requests",
        "--collect-all=pymorphy3_dicts_ru",
        "--collect-all=tkinterdnd2",
        "main.py"
    ]
    
    print(f"Платформа: {sys.platform}")
    print(f"Разделитель для --add-data: '{separator}'")
    
    result = subprocess.run(cmd)
    
    if result.returncode == 0:
        dist_file = Path("dist/PDF_Parser_OATI")
        exe_file = Path("dist/PDF_Parser_OATI.exe")
        
        if dist_file.exists() and not exe_file.exists():
            print(f"\nДобавление расширения .exe к файлу...")
            dist_file.rename(exe_file)
            print(f"✓ Файл переименован: {exe_file}")
        
        print("\nКопирование файлов установки сканера...")
        scanner_files = [
            ("setup_scanner.bat", "dist/setup_scanner.bat"),
            ("SCANNER_SETUP_RU.txt", "dist/SCANNER_SETUP_RU.txt")
        ]
        
        for src, dst in scanner_files:
            if os.path.exists(src):
                shutil.copy2(src, dst)
                print(f"  ✓ {src} → {dst}")
            else:
                print(f"  ⚠ {src} не найден, пропущен")
        
        print("\n" + "=" * 60)
        print("  ✓ Сборка завершена успешно!")
        print("=" * 60)
        print(f"\n.exe файл: dist/PDF_Parser_OATI.exe")
        print("\n✓ Всё упаковано внутри .exe файла:")
        print("  • Python интерпретатор")
        print("  • Все библиотеки (PyPDF2, python-docx, pymorphy3, и т.д.)")
        print("  • Шаблоны Word документов")
        print("  • Исходный код приложения")
        print("\n🚀 ГОТОВО К РАСПРОСТРАНЕНИЮ!")
        print("\nДля использования:")
        print("  1. Скопируйте PDF_Parser_OATI.exe на любой компьютер с Windows")
        print("  2. Запустите двойным кликом")
        print("  3. Приложение автоматически создаст нужные папки и базу данных")
        print("\nНикаких дополнительных действий не требуется!")
    else:
        print("\n✗ Ошибка при сборке")
        sys.exit(1)


if __name__ == "__main__":
    build_executable()
