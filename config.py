"""
Конфигурация приложения.

Для переноса на другой ПК правьте только `settings.json` рядом с программой.
Если `settings.json` отсутствует — используются значения по умолчанию (от папки программы).
"""

from __future__ import annotations

import json
import os
from pathlib import Path


def _app_dir() -> Path:
    """
    Папка, рядом с которой лежат файлы программы.
    Работает и для обычного запуска `python app.py`, и для сборки в EXE (PyInstaller).
    """
    if getattr(__import__("sys"), "frozen", False):
        return Path(__import__("sys").executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = _app_dir()

# Значения по умолчанию (от папки программы)
_DEFAULTS = {
    "excel_file": str(BASE_DIR / "пример таблицы.xlsx"),
    "word_template": str(BASE_DIR / "шаблон.docx"),
    "acts_folder": str(BASE_DIR / "Acts"),
    "database_file": str(BASE_DIR / "raw_materials.db"),
}


def _load_settings() -> dict:
    settings_path = BASE_DIR / "settings.json"
    if not settings_path.exists():
        return {}
    try:
        with settings_path.open("r", encoding="utf-8") as f:
            data = json.load(f) or {}
        if not isinstance(data, dict):
            return {}
        return data
    except Exception:
        # Не блокируем запуск из-за битого settings.json
        return {}


def _resolve_path(value: str | None, fallback: str) -> str:
    """
    Разрешает путь: если относительный — считаем относительно BASE_DIR.
    """
    raw = (value or "").strip()
    if not raw:
        raw = fallback
    p = Path(raw)
    if not p.is_absolute():
        p = (BASE_DIR / p).resolve()
    return str(p)


_settings = _load_settings()

# Пути к файлам/папкам (переносимые)
# Excel-журнал, куда добавляется новая строка при создании акта
EXCEL_FILE = _resolve_path(_settings.get("excel_file"), _DEFAULTS["excel_file"])
WORD_TEMPLATE = _resolve_path(_settings.get("word_template"), _DEFAULTS["word_template"])
ACTS_FOLDER = _resolve_path(_settings.get("acts_folder"), _DEFAULTS["acts_folder"])
DATABASE_FILE = _resolve_path(_settings.get("database_file"), _DEFAULTS["database_file"])

# Цвета статусов для Excel (RGB)
STATUS_COLORS = {
    "РАЗРЕШЕНО": "92D050",   # Зеленый
    "КАРАНТИН": "FFFF00",    # Желтый
    "БРАК": "FF0000",        # Красный
    "КОНТРОЛЬ": "FFC000"     # Оранжевый
}

# Список полей формы (соответствует столбцам Excel)
FORM_FIELDS = [
    "Наименование",
    "Внешний вид заявлено",
    "Внешний вид факт",
    "Соответствие внешнего вида",
    "Поставщик",
    "Производитель",
    "Дата поступления",
    "Дата проверки",
    "№ партии",
    "Дата изготовления",
    "Срок годности",
    "Фактическая масса (кг)",
    "Проверяемые показатели",
    "Результат исследований",
    "Норматив по паспорту",
    "Заключение по проверяемым показателям",
    "Плотность измеренная г/см³, насыпная плотность кг/м³",
    "Плотность по паспорту, кг/м³",
    "Заключение по плотности",
    "Влажность измеренная, %",
    "Влажность по паспорту, %",
    "Заключение по влажности",
    "Метталомагнитные примеси, мг/кг",
    "Металломагнитные примеси по паспорту, мг/кг",
    "Заключение по металломагнитным примесям",
    "ФИО",
    "Коментарии"
]

# Поля для таблицы в Word-акте (плотность, влажность, примеси)
WORD_TABLE_FIELDS = [
    ("Плотность измеренная г/см³, насыпная плотность кг/м³", "Плотность по паспорту, кг/м³", "Заключение по плотности"),
    ("Влажность измеренная, %", "Влажность по паспорту, %", "Заключение по влажности"),
    ("Метталомагнитные примеси, мг/кг", "Металломагнитные примеси по паспорту, мг/кг", "Заключение по металломагнитным примесям")
]
