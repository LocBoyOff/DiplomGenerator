# Генератор дипломов

GUI-приложение для автоматической генерации PDF-дипломов из Excel-данных и шаблона PPTX.

## Установка (из кода)
pip install -r requirements.txt
python diploma_generator.py
## Использование
1. Выбери **Excel-файл** с данными (столбцы: ФИО, Дата, Часы и т.д.).
2. Выбери **шаблон PPTX** с плейсхолдерами `{ФИО}`, `{ДАТА}`.
3. Выбери **папку для PDF**.
4. **Сопоставь поля** (кнопка "Сопоставление").
5. **Запусти генерацию**!

## Зависимости
- wxPython (GUI)
- python-pptx (работа с PPTX)
- openpyxl (Excel)
- comtypes (PowerPoint COM)
- psutil (управление процессами)

## Releases
Скачай готовый `.exe` из [Releases](https://github.com/LocBoyOff/DiplomGenerator/releases).

## Автор
Восстановлено из .exe с помощью PyInstaller + декомпиляции.  
© 2025 LocBoyOff
