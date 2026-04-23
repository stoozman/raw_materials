import sqlite3
import re
from datetime import datetime, timezone, timedelta
from config import DATABASE_FILE


def init_database():
    """Инициализация базы данных и создание таблиц"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    # Таблица для записей о сырье (все поля из FORM_FIELDS)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS raw_materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            act_number TEXT NOT NULL,
            status TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            -- Поля формы (соответствуют столбцам Excel)
            name TEXT,                                    -- Наименование
            appearance_claimed TEXT,                      -- Внешний вид заявлено
            appearance_actual TEXT,                       -- Внешний вид факт
            appearance_match TEXT,                        -- Соответствие внешнего вида
            supplier TEXT,                                -- Поставщик
            manufacturer TEXT,                            -- Производитель
            arrival_date TEXT,                            -- Дата поступления
            check_date TEXT,                              -- Дата проверки
            batch_number TEXT,                            -- № партии
            manufacture_date TEXT,                        -- Дата изготовления
            expiry_date TEXT,                             -- Срок годности
            actual_mass TEXT,                             -- Фактическая масса (кг)
            test_indicators TEXT,                         -- Проверяемые показатели
            research_result TEXT,                         -- Результат исследований
            passport_norm TEXT,                           -- Норматив по паспорту
            test_conclusion TEXT,                         -- Заключение по проверяемым показателям
            density_measured TEXT,                        -- Плотность измеренная
            density_passport TEXT,                        -- Плотность по паспорту
            density_conclusion TEXT,                      -- Заключение по плотности
            humidity_measured TEXT,                       -- Влажность измеренная
            humidity_passport TEXT,                       -- Влажность по паспорту
            humidity_conclusion TEXT,                     -- Заключение по влажности
            metal_impurities_measured TEXT,               -- Металломагнитные примеси
            metal_impurities_passport TEXT,                 -- Металломагнитные примеси по паспорту
            metal_impurities_conclusion TEXT,             -- Заключение по металломагнитным примесям
            fio TEXT,                                     -- ФИО
            comments TEXT,                                -- Коментарии

            -- Дополнительные поля для акта
            check_time TEXT,                              -- Время проверки (МСК)
            excel_row INTEGER,                            -- Номер строки в Excel
            word_path TEXT                                -- Путь к созданному акту (docx)
        )
    ''')

    # Миграции (если база уже создана ранее и не содержит новых колонок)
    cursor.execute("PRAGMA table_info(raw_materials)")
    existing_cols = {row[1] for row in cursor.fetchall()}
    if "word_path" not in existing_cols:
        cursor.execute("ALTER TABLE raw_materials ADD COLUMN word_path TEXT")

    # Таблица для хранения текущего номера акта
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS act_counter (
            id INTEGER PRIMARY KEY,
            last_number INTEGER DEFAULT 0
        )
    ''')

    # Инициализация счетчика, если таблица пустая
    cursor.execute("SELECT COUNT(*) FROM act_counter")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO act_counter (id, last_number) VALUES (1, 265)")

    conn.commit()
    conn.close()


def get_next_act_number():
    """Получение следующего номера акта (формат: 266П)"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT last_number FROM act_counter WHERE id = 1")
    result = cursor.fetchone()
    last_number = result[0] if result else 265

    next_number = last_number + 1
    act_number = f"{next_number}П"

    conn.close()
    return act_number, next_number


def increment_act_number():
    """Инкремент номера акта в базе данных"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    cursor.execute("UPDATE act_counter SET last_number = last_number + 1 WHERE id = 1")

    conn.commit()
    conn.close()


def save_record(data, status, act_number=None, record_id=None):
    """
    Сохранение записи в базу данных

    Args:
        data: словарь с данными формы
        status: статус записи (РАЗРЕШЕНО, КАРАНТИН, БРАК, КОНТРОЛЬ)
        act_number: номер акта (если None - генерируется новый)
        record_id: ID записи для обновления (если None - создается новая)

    Returns:
        tuple: (record_id, act_number)
    """
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    # Генерация нового номера акта, если нужно
    if act_number is None:
        act_number, _ = get_next_act_number()
        increment_act_number()

    # Текущее время в МСК (UTC+3)
    msk_tz = timezone(timedelta(hours=3))
    check_time = datetime.now(msk_tz).strftime("%H:%M:%S")

    # Mapping формы на поля БД
    name = data.get("Наименование", "")
    appearance_claimed = data.get("Внешний вид заявлено", "")
    appearance_actual = data.get("Внешний вид факт", "")
    appearance_match = data.get("Соответствие внешнего вида", "")
    supplier = data.get("Поставщик", "")
    manufacturer = data.get("Производитель", "")
    arrival_date = data.get("Дата поступления", "")
    check_date = data.get("Дата проверки", "")
    batch_number = data.get("№ партии", "")
    manufacture_date = data.get("Дата изготовления", "")
    expiry_date = data.get("Срок годности", "")
    actual_mass = data.get("Фактическая масса (кг)", "")
    test_indicators = data.get("Проверяемые показатели", "")
    research_result = data.get("Результат исследований", "")
    passport_norm = data.get("Норматив по паспорту", "")
    test_conclusion = data.get("Заключение по проверяемым показателям", "")
    density_measured = data.get("Плотность измеренная г/см³, насыпная плотность кг/м³", "")
    density_passport = data.get("Плотность по паспорту, кг/м³", "")
    density_conclusion = data.get("Заключение по плотности", "")
    humidity_measured = data.get("Влажность измеренная, %", "")
    humidity_passport = data.get("Влажность по паспорту, %", "")
    humidity_conclusion = data.get("Заключение по влажности", "")
    metal_impurities_measured = data.get("Метталомагнитные примеси, мг/кг", "")
    metal_impurities_passport = data.get("Металломагнитные примеси по паспорту, мг/кг", "")
    metal_impurities_conclusion = data.get("Заключение по металломагнитным примесям", "")
    fio = data.get("ФИО", "")
    comments = data.get("Коментарии", "")

    if record_id:
        # Обновление существующей записи
        cursor.execute('''
            UPDATE raw_materials SET
                status = ?,
                updated_at = CURRENT_TIMESTAMP,
                name = ?,
                appearance_claimed = ?,
                appearance_actual = ?,
                appearance_match = ?,
                supplier = ?,
                manufacturer = ?,
                arrival_date = ?,
                check_date = ?,
                batch_number = ?,
                manufacture_date = ?,
                expiry_date = ?,
                actual_mass = ?,
                test_indicators = ?,
                research_result = ?,
                passport_norm = ?,
                test_conclusion = ?,
                density_measured = ?,
                density_passport = ?,
                density_conclusion = ?,
                humidity_measured = ?,
                humidity_passport = ?,
                humidity_conclusion = ?,
                metal_impurities_measured = ?,
                metal_impurities_passport = ?,
                metal_impurities_conclusion = ?,
                fio = ?,
                comments = ?,
                check_time = ?
            WHERE id = ?
        ''', (
            status, name, appearance_claimed, appearance_actual, appearance_match,
            supplier, manufacturer, arrival_date, check_date, batch_number,
            manufacture_date, expiry_date, actual_mass, test_indicators,
            research_result, passport_norm, test_conclusion, density_measured,
            density_passport, density_conclusion, humidity_measured, humidity_passport,
            humidity_conclusion, metal_impurities_measured, metal_impurities_passport,
            metal_impurities_conclusion, fio, comments, check_time, record_id
        ))
    else:
        # Создание новой записи
        cursor.execute('''
            INSERT INTO raw_materials (
                act_number, status, name, appearance_claimed, appearance_actual,
                appearance_match, supplier, manufacturer, arrival_date, check_date,
                batch_number, manufacture_date, expiry_date, actual_mass,
                test_indicators, research_result, passport_norm, test_conclusion,
                density_measured, density_passport, density_conclusion,
                humidity_measured, humidity_passport, humidity_conclusion,
                metal_impurities_measured, metal_impurities_passport,
                metal_impurities_conclusion, fio, comments, check_time
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            act_number, status, name, appearance_claimed, appearance_actual,
            appearance_match, supplier, manufacturer, arrival_date, check_date,
            batch_number, manufacture_date, expiry_date, actual_mass, test_indicators,
            research_result, passport_norm, test_conclusion, density_measured,
            density_passport, density_conclusion, humidity_measured, humidity_passport,
            humidity_conclusion, metal_impurities_measured, metal_impurities_passport,
            metal_impurities_conclusion, fio, comments, check_time
        ))
        record_id = cursor.lastrowid

    conn.commit()
    conn.close()

    return record_id, act_number


def get_record_by_id(record_id):
    """Получение записи по ID со всеми полями"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    cursor.execute('''
        SELECT * FROM raw_materials WHERE id = ?
    ''', (record_id,))

    row = cursor.fetchone()
    conn.close()

    if row:
        # row[0]=id, row[1]=act_number, row[2]=status, row[3]=created_at, row[4]=updated_at
        # row[5]=name, row[6]=appearance_claimed, row[7]=appearance_actual, row[8]=appearance_match
        # row[9]=supplier, row[10]=manufacturer, row[11]=arrival_date, row[12]=check_date
        # row[13]=batch_number, row[14]=manufacture_date, row[15]=expiry_date, row[16]=actual_mass
        # row[17]=test_indicators, row[18]=research_result, row[19]=passport_norm, row[20]=test_conclusion
        # row[21]=density_measured, row[22]=density_passport, row[23]=density_conclusion
        # row[24]=humidity_measured, row[25]=humidity_passport, row[26]=humidity_conclusion
        # row[27]=metal_impurities_measured, row[28]=metal_impurities_passport, row[29]=metal_impurities_conclusion
        # row[30]=fio, row[31]=comments, row[32]=check_time, row[33]=excel_row, row[34]=word_path
        return {
            "id": row[0],
            "act_number": row[1],
            "status": row[2],
            "Наименование": row[5] or "",
            "Внешний вид заявлено": row[6] or "",
            "Внешний вид факт": row[7] or "",
            "Соответствие внешнего вида": row[8] or "",
            "Поставщик": row[9] or "",
            "Производитель": row[10] or "",
            "Дата поступления": row[11] or "",
            "Дата проверки": row[12] or "",
            "№ партии": row[13] or "",
            "Дата изготовления": row[14] or "",
            "Срок годности": row[15] or "",
            "Фактическая масса (кг)": row[16] or "",
            "Проверяемые показатели": row[17] or "",
            "Результат исследований": row[18] or "",
            "Норматив по паспорту": row[19] or "",
            "Заключение по проверяемым показателям": row[20] or "",
            "Плотность измеренная г/см³, насыпная плотность кг/м³": row[21] or "",
            "Плотность по паспорту, кг/м³": row[22] or "",
            "Заключение по плотности": row[23] or "",
            "Влажность измеренная, %": row[24] or "",
            "Влажность по паспорту, %": row[25] or "",
            "Заключение по влажности": row[26] or "",
            "Метталомагнитные примеси, мг/кг": row[27] or "",
            "Металломагнитные примеси по паспорту, мг/кг": row[28] or "",
            "Заключение по металломагнитным примесям": row[29] or "",
            "ФИО": row[30] or "",
            "Коментарии": row[31] or "",
            "check_time": row[32] or "",
            "excel_row": row[33],
            "word_path": row[34] if len(row) > 34 else None
        }
    return None


def get_all_records():
    """Получение всех записей"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    cursor.execute('''
        SELECT id, act_number, status, name, supplier, created_at
        FROM raw_materials ORDER BY created_at DESC
    ''')

    rows = cursor.fetchall()
    conn.close()

    records = []
    for row in rows:
        records.append({
            "id": row[0],
            "act_number": row[1],
            "status": row[2],
            "name": row[3],
            "supplier": row[4],
            "created_at": row[5]
        })

    return records


def update_excel_row(record_id, excel_row):
    """Обновление номера строки в Excel"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    cursor.execute(
        "UPDATE raw_materials SET excel_row = ? WHERE id = ?",
        (excel_row, record_id)
    )

    conn.commit()
    conn.close()


def update_word_path(record_id, word_path):
    """Обновление пути к акту (docx)"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE raw_materials SET word_path = ? WHERE id = ?",
        (word_path, record_id)
    )
    conn.commit()
    conn.close()


def delete_record(record_id):
    """Удаление записи из БД"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM raw_materials WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()


def shift_excel_rows_after(deleted_row):
    """
    После удаления строки из Excel нужно сдвинуть сохранённые номера строк в БД.
    Все записи, у которых excel_row > deleted_row, уменьшаем на 1.
    """
    if not deleted_row:
        return
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE raw_materials SET excel_row = excel_row - 1 WHERE excel_row IS NOT NULL AND excel_row > ?",
        (deleted_row,)
    )
    conn.commit()
    conn.close()


def extract_number_from_act(act_number):
    """Извлечение числа из номера акта (например, '266П' -> 266)"""
    match = re.match(r'(\d+)П', act_number)
    if match:
        return int(match.group(1))
    return 0


def sync_act_number_from_records():
    """Синхронизация счетчика актов с существующими записями"""
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT act_number FROM raw_materials ORDER BY id DESC LIMIT 1")
    result = cursor.fetchone()

    if result:
        last_act_number = extract_number_from_act(result[0])
        cursor.execute(
            "UPDATE act_counter SET last_number = ? WHERE id = ?",
            (last_act_number, 1)
        )
        conn.commit()

    conn.close()
