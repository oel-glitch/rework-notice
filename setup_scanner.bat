@echo off
chcp 65001 >nul
color 0A
title Установка компонентов сканера Canon DR-M260

echo.
echo ═══════════════════════════════════════════════════════════════════════════
echo  УСТАНОВКА КОМПОНЕНТОВ ДЛЯ МОДУЛЯ "СКАНИРОВАНИЕ"
echo  PDF Parser ОАТИ v4.1
echo ═══════════════════════════════════════════════════════════════════════════
echo.
echo Проверяю установленные компоненты...
echo.

:: ═══════════════════════════════════════════════════════════════════════════
:: ПРОВЕРКА 1: Tesseract OCR
:: ═══════════════════════════════════════════════════════════════════════════
echo [1/3] Tesseract OCR (распознавание текста)
echo ─────────────────────────────────────────────────────────────────────────

set TESSERACT_FOUND=0

:: Проверка через команду where
where tesseract >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set TESSERACT_FOUND=1
    echo [√] УСТАНОВЛЕН - найден в PATH
    tesseract --version 2>nul | findstr /C:"tesseract"
    goto :check_twain
)

:: Проверка в стандартных путях установки
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    set TESSERACT_FOUND=1
    echo [√] УСТАНОВЛЕН - C:\Program Files\Tesseract-OCR\tesseract.exe
    "C:\Program Files\Tesseract-OCR\tesseract.exe" --version 2>nul | findstr /C:"tesseract"
    goto :check_twain
)

if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" (
    set TESSERACT_FOUND=1
    echo [√] УСТАНОВЛЕН - C:\Program Files (x86)\Tesseract-OCR\tesseract.exe
    "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" --version 2>nul | findstr /C:"tesseract"
    goto :check_twain
)

:: Tesseract не найден
echo [×] НЕ НАЙДЕН
echo.
echo     ► Tesseract OCR необходим для распознавания текста на отсканированных
echo       документах. Без него сканер будет создавать только изображения PDF.
echo.
echo     УСТАНОВКА:
echo     1. Скачайте установщик Tesseract OCR для Windows (русский язык):
echo        https://github.com/UB-Mannheim/tesseract/wiki
echo        
echo        Рекомендуемый файл: tesseract-ocr-w64-setup-v5.3.3.20231005.exe
echo.
echo     2. Запустите установщик с правами администратора
echo.
echo     3. При установке ОБЯЗАТЕЛЬНО отметьте:
echo        [√] Русский язык (Russian language data)
echo        [√] Добавить в PATH (Add to PATH)
echo.
echo     4. Перезапустите этот скрипт после установки
echo.

:check_twain
:: ═══════════════════════════════════════════════════════════════════════════
:: ПРОВЕРКА 2: TWAIN Драйвер
:: ═══════════════════════════════════════════════════════════════════════════
echo.
echo [2/3] TWAIN драйвер Canon DR-M260
echo ─────────────────────────────────────────────────────────────────────────

set TWAIN_FOUND=0

:: Проверка ключей реестра TWAIN
reg query "HKLM\SOFTWARE\TWAIN_32" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set TWAIN_FOUND=1
)

reg query "HKLM\SOFTWARE\WOW6432Node\TWAIN_32" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set TWAIN_FOUND=1
)

if %TWAIN_FOUND% EQU 1 (
    echo [√] ДРАЙВЕР TWAIN УСТАНОВЛЕН
    echo     Найдены записи в реестре Windows
    goto :check_wia
)

:: TWAIN не найден
echo [×] НЕ НАЙДЕН
echo.
echo     ► TWAIN - стандартный интерфейс для работы со сканерами в Windows.
echo       Без драйвера сканер Canon DR-M260 не будет работать.
echo.
echo     УСТАНОВКА:
echo     1. Подключите сканер Canon DR-M260 к компьютеру
echo.
echo     2. Скачайте драйверы с официального сайта Canon:
echo        https://www.canon.ru/support/business-product-support/
echo        
echo        Поиск: DR-M260
echo        Раздел: Драйверы и программное обеспечение
echo.
echo     3. Установите пакет "ISIS/TWAIN Driver" для Windows
echo        (требуются права администратора)
echo.
echo     4. Перезагрузите компьютер после установки драйвера
echo.
echo     5. Проверьте работу сканера в стандартных приложениях Windows
echo        (Сканер, Paint, и т.д.)
echo.

:check_wia
:: ═══════════════════════════════════════════════════════════════════════════
:: ПРОВЕРКА 3: WIA (Windows Image Acquisition)
:: ═══════════════════════════════════════════════════════════════════════════
echo.
echo [3/3] WIA (Windows Image Acquisition)
echo ─────────────────────────────────────────────────────────────────────────

:: WIA встроен в Windows, проверяем наличие службы
sc query stisvc >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [√] СЛУЖБА WIA ДОСТУПНА
    echo     Windows Image Acquisition работает корректно
) else (
    echo [×] СЛУЖБА WIA НЕ НАЙДЕНА
    echo.
    echo     ► WIA - встроенная служба Windows для работы со сканерами.
    echo       Обычно присутствует во всех версиях Windows.
    echo.
    echo     РЕШЕНИЕ:
    echo     1. Откройте "Службы" (services.msc)
    echo     2. Найдите "Служба загрузки изображений Windows (WIA)"
    echo     3. Установите тип запуска: "Автоматически"
    echo     4. Запустите службу
)

:: ═══════════════════════════════════════════════════════════════════════════
:: ИТОГОВАЯ СВОДКА
:: ═══════════════════════════════════════════════════════════════════════════
echo.
echo.
echo ═══════════════════════════════════════════════════════════════════════════
echo  ИТОГИ ПРОВЕРКИ
echo ═══════════════════════════════════════════════════════════════════════════

set ALL_OK=1

if %TESSERACT_FOUND% EQU 0 (
    echo [×] Tesseract OCR не установлен
    set ALL_OK=0
) else (
    echo [√] Tesseract OCR установлен
)

if %TWAIN_FOUND% EQU 0 (
    echo [×] TWAIN драйвер не найден
    set ALL_OK=0
) else (
    echo [√] TWAIN драйвер установлен
)

echo.
if %ALL_OK% EQU 1 (
    echo ┌─────────────────────────────────────────────────────────────────────┐
    echo │  ✓ ВСЕ КОМПОНЕНТЫ УСТАНОВЛЕНЫ!                                      │
    echo │                                                                     │
    echo │  Модуль "Сканирование" готов к работе.                             │
    echo │  Запустите PDF_Parser_OATI.exe и проверьте статус сканера.         │
    echo └─────────────────────────────────────────────────────────────────────┘
) else (
    echo ┌─────────────────────────────────────────────────────────────────────┐
    echo │  ! ТРЕБУЕТСЯ УСТАНОВКА КОМПОНЕНТОВ                                  │
    echo │                                                                     │
    echo │  Установите недостающие компоненты согласно инструкциям выше,       │
    echo │  затем повторно запустите этот скрипт для проверки.                 │
    echo │                                                                     │
    echo │  Подробная инструкция: SCANNER_SETUP_RU.txt                         │
    echo └─────────────────────────────────────────────────────────────────────┘
)

echo.
echo ═══════════════════════════════════════════════════════════════════════════
echo.

:: Python библиотеки встроены в .exe
echo ПРИМЕЧАНИЕ: Python библиотеки (Pillow, pytesseract, pywin32) уже встроены
echo             в PDF_Parser_OATI.exe и не требуют установки.
echo.

pause
