import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import sys

# Добавляем текущую директорию в путь для импорта
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from database import (
    init_database, save_record, get_record_by_id,
    get_all_records, sync_act_number_from_records,
    delete_record, shift_excel_rows_after
)
from documents import generate_documents, delete_excel_row, delete_act_files_for_record
from config import FORM_FIELDS, STATUS_COLORS


class RecordsListWindow(tk.Toplevel):
    """Окно списка записей для выбора"""

    def __init__(self, parent, on_select_callback, on_delete_callback):
        super().__init__(parent)
        self.title("Выбор записи")
        self.geometry("800x500")
        self.on_select_callback = on_select_callback
        self.on_delete_callback = on_delete_callback

        # Заголовок
        self.label_title = tk.Label(
            self, text="Существующие записи", font=("Arial", 16, "bold")
        )
        self.label_title.pack(pady=10)

        # Список записей (Canvas с scrollbar)
        self.canvas = tk.Canvas(self, width=750, height=380)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas)

        self.scroll_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        self.scrollbar.pack(side="right", fill="y")

        # Кнопка обновления
        self.btn_refresh = tk.Button(
            self, text="Обновить список", command=self.load_records
        )
        self.btn_refresh.pack(pady=10)

        self.load_records()

    def load_records(self):
        """Загрузка и отображение записей"""
        # Очищаем текущий список
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        records = get_all_records()

        if not records:
            tk.Label(
                self.scroll_frame,
                text="Нет записей в базе данных",
                font=("Arial", 12)
            ).pack(pady=20)
            return

        # Заголовки таблицы
        headers_frame = tk.Frame(self.scroll_frame)
        headers_frame.pack(fill="x", padx=5, pady=2)

        headers = ["№ акта", "Наименование", "Поставщик", "Статус", "Действие"]
        for header in headers:
            tk.Label(
                headers_frame, text=header, font=("Arial", 11, "bold"), width=15
            ).pack(side="left", padx=2)

        # Данные записей
        for record in records:
            row_frame = tk.Frame(self.scroll_frame)
            row_frame.pack(fill="x", padx=5, pady=2)

            tk.Label(row_frame, text=record["act_number"], width=15).pack(side="left", padx=2)
            tk.Label(row_frame, text=record["name"][:20], width=15).pack(side="left", padx=2)
            tk.Label(row_frame, text=record["supplier"][:15], width=15).pack(side="left", padx=2)

            # Цветной индикатор статуса
            status_color = STATUS_COLORS.get(record["status"], "gray")
            status_label = tk.Label(
                row_frame, text=record["status"], width=15,
                bg=f"#{status_color}"
            )
            status_label.pack(side="left", padx=2)

            tk.Button(
                row_frame, text="Выбрать", width=10,
                command=lambda r=record: self.select_record(r["id"])
            ).pack(side="left", padx=2)

            tk.Button(
                row_frame, text="Удалить", width=10,
                fg="white", bg="#C00000",
                command=lambda r=record: self.delete_record_ui(r["id"])
            ).pack(side="left", padx=2)

    def select_record(self, record_id):
        """Выбор записи"""
        self.on_select_callback(record_id)
        self.destroy()

    def delete_record_ui(self, record_id):
        """Удаление записи"""
        if messagebox.askyesno(
            "Подтверждение удаления",
            "Удалить запись из базы данных, строку из Excel и файл акта?\n\nОтменить действие будет нельзя."
        ):
            self.on_delete_callback(record_id)
            self.load_records()


class RawMaterialsApp(tk.Tk):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()

        self.title("Учет лабораторного сырья")
        self.geometry("900x750")

        # Текущая запись (для режима редактирования)
        self.current_record_id = None
        self.current_act_number = None

        self.create_widgets()

    def _bind_clipboard_menu(self, widget):
        """
        Контекстное меню (ПКМ) для копирования/вставки в полях.
        """
        menu = tk.Menu(self, tearoff=0)

        def do_cut():
            try:
                widget.event_generate("<<Cut>>")
            except Exception:
                pass

        def do_copy():
            try:
                widget.event_generate("<<Copy>>")
            except Exception:
                try:
                    self.clipboard_clear()
                    self.clipboard_append(widget.get())
                except Exception:
                    pass

        def do_paste():
            try:
                widget.event_generate("<<Paste>>")
            except Exception:
                pass

        def do_select_all():
            try:
                widget.selection_range(0, "end")
                widget.icursor("end")
            except Exception:
                pass

        menu.add_command(label="Вырезать", command=do_cut)
        menu.add_command(label="Копировать", command=do_copy)
        menu.add_command(label="Вставить", command=do_paste)
        menu.add_separator()
        menu.add_command(label="Выделить всё", command=do_select_all)

        def show_menu(event):
            try:
                widget.focus_set()
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    menu.grab_release()
                except Exception:
                    pass

        widget.bind("<Button-3>", show_menu)  # Windows
        widget.bind("<Control-Button-1>", show_menu)  # на всякий случай

    def create_widgets(self):
        """Создание элементов интерфейса"""
        # Заголовок
        self.frame_header = tk.Frame(self)
        self.frame_header.pack(fill="x", padx=10, pady=5)

        self.label_title = tk.Label(
            self.frame_header,
            text="ЖУРНАЛ УЧЕТА ПОСТУПЛЕНИЯ СЫРЬЯ",
            font=("Arial", 18, "bold")
        )
        self.label_title.pack(pady=5)

        # Кнопка выбора существующей записи
        self.btn_select_record = tk.Button(
            self.frame_header,
            text="Выбрать существующую запись",
            command=self.open_records_list,
            width=30
        )
        self.btn_select_record.pack(pady=5)

        # Индикатор режима
        self.label_mode = tk.Label(
            self.frame_header,
            text="Режим: НОВАЯ ЗАПИСЬ",
            font=("Arial", 12),
            fg="green"
        )
        self.label_mode.pack(pady=2)

        # Форма ввода (Canvas со scrollbar)
        self.form_canvas = tk.Canvas(self)
        self.form_scrollbar = tk.Scrollbar(self, orient="vertical", command=self.form_canvas.yview)
        self.frame_form = tk.Frame(self.form_canvas)

        self.frame_form.bind(
            "<Configure>",
            lambda e: self.form_canvas.configure(scrollregion=self.form_canvas.bbox("all"))
        )

        self.form_canvas.create_window((0, 0), window=self.frame_form, anchor="nw")
        self.form_canvas.configure(yscrollcommand=self.form_scrollbar.set)

        self.form_canvas.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        self.form_scrollbar.pack(side="right", fill="y")

        # Словарь для хранения полей ввода
        self.entry_fields = {}

        choice_fields = {
            "Соответствие внешнего вида",
            "Заключение по проверяемым показателям",
            "Заключение по плотности",
            "Заключение по влажности",
            "Заключение по металломагнитным примесям",
        }
        choice_values = ["—", "соответствует", "не соответствует"]

        # Создаем поля ввода
        for i, field_name in enumerate(FORM_FIELDS):
            row_frame = tk.Frame(self.frame_form)
            row_frame.pack(fill="x", padx=10, pady=3)

            label = tk.Label(
                row_frame, text=f"{field_name}:",
                font=("Arial", 12), width=25, anchor="e"
            )
            label.pack(side="left", padx=5)

            if field_name in choice_fields:
                entry = ttk.Combobox(
                    row_frame,
                    values=choice_values,
                    state="readonly",
                    font=("Arial", 12),
                    width=48,
                )
                entry.set("—")
            else:
                entry = tk.Entry(row_frame, width=50, font=("Arial", 12))

            entry.pack(side="left", padx=5, fill="x", expand=True)
            self._bind_clipboard_menu(entry)

            self.entry_fields[field_name] = entry

        # Кнопки статусов (помещаем в frame_form чтобы скроллировались вместе)
        self.frame_buttons = tk.Frame(self.frame_form)
        self.frame_buttons.pack(fill="x", padx=10, pady=10)

        self.label_buttons = tk.Label(
            self.frame_buttons,
            text="Выберите статус и сохраните:",
            font=("Arial", 14, "bold")
        )
        self.label_buttons.pack(pady=5)

        self.frame_status_buttons = tk.Frame(self.frame_buttons)
        self.frame_status_buttons.pack(pady=10)

        # Кнопки с цветами статусов
        self.btn_razresheno = tk.Button(
            self.frame_status_buttons,
            text="РАЗРЕШЕНО",
            command=lambda: self.save_with_status("РАЗРЕШЕНО"),
            bg=f"#{STATUS_COLORS['РАЗРЕШЕНО']}",
            fg="black",
            font=("Arial", 12, "bold"),
            width=15,
            height=2
        )
        self.btn_razresheno.pack(side="left", padx=10)

        self.btn_karantin = tk.Button(
            self.frame_status_buttons,
            text="КАРАНТИН",
            command=lambda: self.save_with_status("КАРАНТИН"),
            bg=f"#{STATUS_COLORS['КАРАНТИН']}",
            fg="black",
            font=("Arial", 12, "bold"),
            width=15,
            height=2
        )
        self.btn_karantin.pack(side="left", padx=10)

        self.btn_brak = tk.Button(
            self.frame_status_buttons,
            text="БРАК",
            command=lambda: self.save_with_status("БРАК"),
            bg=f"#{STATUS_COLORS['БРАК']}",
            fg="white",
            font=("Arial", 12, "bold"),
            width=15,
            height=2
        )
        self.btn_brak.pack(side="left", padx=10)

        self.btn_control = tk.Button(
            self.frame_status_buttons,
            text="КОНТРОЛЬ",
            command=lambda: self.save_with_status("КОНТРОЛЬ"),
            bg=f"#{STATUS_COLORS['КОНТРОЛЬ']}",
            fg="black",
            font=("Arial", 12, "bold"),
            width=15,
            height=2
        )
        self.btn_control.pack(side="left", padx=10)

        # Кнопка очистки формы
        self.btn_clear = tk.Button(
            self.frame_form,
            text="Очистить форму (новая запись)",
            command=self.clear_form,
            width=30
        )
        self.btn_clear.pack(pady=10)

        # Статусная строка
        self.label_status = tk.Label(
            self.frame_form,
            text="Готово к работе",
            font=("Arial", 11)
        )
        self.label_status.pack(pady=5)

    def get_form_data(self):
        """Получение данных из формы"""
        return {
            field: entry.get().strip()
            for field, entry in self.entry_fields.items()
        }

    def set_form_data(self, data):
        """Заполнение формы данными"""
        for field, entry in self.entry_fields.items():
            value = (data.get(field, "") or "").strip()
            if isinstance(entry, ttk.Combobox):
                entry.set(value if value else "—")
            else:
                entry.delete(0, "end")
                entry.insert(0, value)

    def clear_form(self):
        """Очистка формы и сброс режима"""
        for entry in self.entry_fields.values():
            if isinstance(entry, ttk.Combobox):
                entry.set("—")
            else:
                entry.delete(0, "end")

        self.current_record_id = None
        self.current_act_number = None

        self.label_mode.configure(
            text="Режим: НОВАЯ ЗАПИСЬ",
            fg="green"
        )
        self.label_status.configure(text="Готово к работе")

    def open_records_list(self):
        """Открытие окна списка записей"""
        RecordsListWindow(self, self.load_record_for_edit, self.delete_record_full)

    def delete_record_full(self, record_id):
        """Полное удаление: БД + Excel + акт"""
        try:
            record = get_record_by_id(record_id)
            if not record:
                messagebox.showwarning("Внимание", "Запись не найдена.")
                return

            # Удаляем акт(ы)
            removed_acts, failed_acts = delete_act_files_for_record(record)

            # Удаляем строку в Excel и сдвигаем excel_row в БД
            excel_row = record.get("excel_row")
            if excel_row:
                delete_excel_row(excel_row)
                shift_excel_rows_after(excel_row)

            # Удаляем запись из БД
            delete_record(record_id)

            # Если сейчас редактировали именно эту запись — сбрасываем форму
            if self.current_record_id == record_id:
                self.clear_form()

            if failed_acts:
                details = "\n".join([f"- {p}\n  причина: {err}" for p, err in failed_acts[:5]])
                more = "" if len(failed_acts) <= 5 else f"\n...и ещё {len(failed_acts) - 5} шт."
                messagebox.showwarning(
                    "Удаление выполнено частично",
                    f"Запись удалена.\nУдалено актов: {removed_acts}\n"
                    f"Не удалось удалить акт(ы): {len(failed_acts)}\n\n"
                    f"Закройте Word/файл и попробуйте ещё раз.\n\n{details}{more}"
                )
            else:
                messagebox.showinfo("Готово", f"Запись удалена.\nУдалено актов: {removed_acts}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось удалить запись:\n{str(e)}")

    def load_record_for_edit(self, record_id):
        """Загрузка записи для редактирования"""
        record = get_record_by_id(record_id)
        if record:
            self.current_record_id = record["id"]
            self.current_act_number = record["act_number"]

            # Заполняем форму
            self.set_form_data(record)

            # Обновляем индикатор режима
            self.label_mode.configure(
                text=f"Режим: РЕДАКТИРОВАНИЕ (Акт: {record['act_number']})",
                fg="orange"
            )

            self.label_status.configure(
                text=f"Загружена запись ID: {record_id}. Внесите изменения и выберите статус."
            )

    def save_with_status(self, status):
        """Сохранение записи с выбранным статусом"""
        try:
            # Получаем данные формы
            form_data = self.get_form_data()

            # Проверяем обязательные поля
            if not form_data["Наименование"]:
                messagebox.showwarning(
                    "Внимание",
                    "Поле 'Наименование' обязательно для заполнения!"
                )
                return

            # Сохраняем в базу данных
            record_id, act_number = save_record(
                data=form_data,
                status=status,
                act_number=self.current_act_number,
                record_id=self.current_record_id
            )

            # Генерируем документы
            result = generate_documents(form_data, record_id, status)

            # Обновляем текущие значения (для режима редактирования)
            self.current_record_id = record_id
            self.current_act_number = act_number

            # Показываем результат
            messagebox.showinfo(
                "Успех",
                f"Запись сохранена!\n\n"
                f"№ акта: {act_number}\n"
                f"Статус: {status}\n"
                f"Строка в Excel: {result['excel_row']}\n"
                f"Акт сохранен: {result['word_path']}"
            )

            # Обновляем статусную строку
            mode_text = "РЕДАКТИРОВАНИЕ" if self.current_record_id else "НОВАЯ ЗАПИСЬ"
            self.label_mode.configure(
                text=f"Режим: {mode_text} (Акт: {act_number})",
                fg="orange" if self.current_record_id else "green"
            )

            self.label_status.configure(
                text=f"Сохранено: Акт {act_number}, Статус: {status}"
            )

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")


def main():
    """Главная функция запуска приложения"""
    # Инициализация базы данных
    init_database()

    # Синхронизация счетчика актов
    sync_act_number_from_records()

    # Запуск приложения
    app = RawMaterialsApp()
    app.mainloop()


if __name__ == "__main__":
    main()
