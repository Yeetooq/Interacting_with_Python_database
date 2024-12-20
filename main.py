import tkinter as tk
from tkinter import messagebox, ttk
import pyodbc


class DatabaseApp:
    def __init__(self, root, db_path):
        self.root = root
        self.root.title("Работа с Базой Данных Access")
        self.db_path = db_path

        # Список доступных таблиц
        self.tables = ["Студент", "Льгота", "Группа", "Факультет", "РодственникСтудента", "ВидыРодственников"]

        self.create_widgets()

    def create_widgets(self):
        # Выбор таблицы из выпадающего списка
        self.table_label = tk.Label(self.root, text="Выберите таблицу:")
        self.table_label.grid(row=0, column=0, padx=10, pady=10)

        self.table_combobox = ttk.Combobox(self.root, values=self.tables)
        self.table_combobox.grid(row=0, column=1, padx=10, pady=10)
        self.table_combobox.current(0)  # Устанавливаем таблицу по умолчанию

        # Кнопка для отображения данных
        self.display_button = tk.Button(self.root, text="Показать данные", command=self.show_data)
        self.display_button.grid(row=1, column=0, padx=10, pady=10)

        # Кнопки для других операций
        self.add_button = tk.Button(self.root, text="Добавить запись", command=self.add_data)
        self.add_button.grid(row=2, column=0, padx=10, pady=10)

        self.update_button = tk.Button(self.root, text="Изменить запись", command=self.request_row_id_for_update)
        self.update_button.grid(row=3, column=0, padx=10, pady=10)

        self.delete_button = tk.Button(self.root, text="Удалить запись", command=self.delete_data)
        self.delete_button.grid(row=4, column=0, padx=10, pady=10)

        # Текстовое поле для вывода результатов
        self.result_text = tk.Text(self.root, height=22, width=130)
        self.result_text.grid(row=0, column=2, rowspan=6, padx=10, pady=10)

        # Поля для ввода данных при добавлении/изменении
        self.input_fields_frame = tk.Frame(self.root)
        self.input_fields_frame.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

    def connect_to_db(self):
        """Подключение к базе данных MS Access."""
        try:
            conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + self.db_path)
            conn = pyodbc.connect(conn_str)
            return conn
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к базе данных: {e}")
            return None

    def show_data(self):
        """Показать данные из выбранной таблицы."""
        table_name = self.table_combobox.get()  # Получаем имя выбранной таблицы
        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                cursor.execute(f"SELECT * FROM {table_name}")
                columns = [column[0] for column in cursor.description]  # Получаем имена столбцов
                rows = cursor.fetchall()

                # Очистка текстового поля
                self.result_text.delete(1.0, tk.END)

                # Заголовки столбцов
                self.result_text.insert(tk.END, " | ".join(columns) + "\n")
                self.result_text.insert(tk.END, "-" * (len(" | ".join(columns)) + 2) + "\n")

                # Вывод строк данных
                for row in rows:
                    self.result_text.insert(tk.END, " | ".join(str(value) for value in row) + "\n")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при выполнении запроса: {e}")
            finally:
                conn.close()

    def add_data(self):
        """Отображаем поля для ввода данных для добавления записи."""
        table_name = self.table_combobox.get()

        # Очистка предыдущих полей ввода
        for widget in self.input_fields_frame.winfo_children():
            widget.destroy()

        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Получаем имена столбцов для выбранной таблицы
                cursor.execute(f"SELECT * FROM {table_name} WHERE 1=0")  # Не выбираем данные, а только структуру
                columns = [column[0] for column in cursor.description]

                # Динамически создаем поля ввода для каждого столбца
                self.inputs = {}
                for i, column in enumerate(columns):
                    label = tk.Label(self.input_fields_frame, text=column)
                    label.grid(row=i, column=0, padx=10, pady=5)

                    entry = tk.Entry(self.input_fields_frame)
                    entry.grid(row=i, column=1, padx=10, pady=5)
                    self.inputs[column] = entry

                # Кнопка для добавления данных
                self.submit_add_button = tk.Button(self.input_fields_frame, text="Добавить", command=self.submit_add)
                self.submit_add_button.grid(row=len(columns), column=0, columnspan=2, padx=10, pady=10)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при получении структуры таблицы: {e}")
            finally:
                conn.close()
    
    def submit_add(self):
        """Отправить данные на добавление в таблицу."""
        table_name = self.table_combobox.get()
        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Собираем данные из полей ввода
                columns = list(self.inputs.keys())
                values = [self.inputs[column].get() for column in columns]

                # Формируем SQL запрос для вставки данных
                placeholders = ", ".join(["?" for _ in columns])
                cursor.execute(f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({placeholders})", values)
                conn.commit()

                messagebox.showinfo("Успех", "Данные успешно добавлены!")
                self.clear_input_fields()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при добавлении данных: {e}")
            finally:
                conn.close()

    def request_row_id_for_update(self):
        """Запросить первичный ключ для изменения записи, а затем загрузить её данные для редактирования."""
        # Очистка предыдущих полей ввода
        for widget in self.input_fields_frame.winfo_children():
            widget.destroy()

        table_name = self.table_combobox.get()

        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Получаем имена столбцов для выбранной таблицы
                cursor.execute(f"SELECT * FROM {table_name} WHERE 1=0")  # Не выбираем данные, а только структуру
                columns = [column[0] for column in cursor.description]

                # Предполагаем, что первый столбец — это первичный ключ
                self.primary_key_column = columns[0]

                # Запрашиваем значение первичного ключа
                label = tk.Label(self.input_fields_frame, text=f"Введите {self.primary_key_column} для изменения:")
                label.grid(row=0, column=0, padx=10, pady=5)

                self.primary_key_entry = tk.Entry(self.input_fields_frame)
                self.primary_key_entry.grid(row=0, column=1, padx=10, pady=5)

                # Кнопка для загрузки данных по первичному ключу
                load_button = tk.Button(self.input_fields_frame, text="Загрузить данные",
                                        command=self.load_data_for_update)
                load_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при получении структуры таблицы: {e}")
            finally:
                conn.close()

    def load_data_for_update(self):
        """Загрузить данные для выбранной строки по первичному ключу и отобразить их в полях для редактирования."""
        table_name = self.table_combobox.get()
        primary_key_value = self.primary_key_entry.get()  # Получаем значение первичного ключа

        if not primary_key_value:
            messagebox.showerror("Ошибка", "Введите значение первичного ключа.")
            return

        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Формируем запрос, чтобы выбрать строку по первичному ключу
                cursor.execute(f"SELECT * FROM {table_name} WHERE {self.primary_key_column} = ?", (primary_key_value,))
                row = cursor.fetchone()

                if row:
                    # Получаем имена столбцов
                    columns = [column[0] for column in cursor.description]

                    # Очистка предыдущих полей ввода
                    for widget in self.input_fields_frame.winfo_children():
                        widget.grid_forget()

                    # Создание полей ввода для данных строки
                    self.inputs = {}
                    for i, column in enumerate(columns):
                        label = tk.Label(self.input_fields_frame, text=column)
                        label.grid(row=i + 1, column=0, padx=10, pady=5)

                        entry = tk.Entry(self.input_fields_frame)
                        entry.grid(row=i + 1, column=1, padx=10, pady=5)
                        entry.insert(tk.END, row[i])  # Вставляем текущие данные в поля ввода
                        self.inputs[column] = entry

                    # Кнопка для обновления данных
                    update_button = tk.Button(self.input_fields_frame, text="Обновить", command=self.submit_update)
                    update_button.grid(row=len(columns) + 1, column=0, columnspan=2, padx=10, pady=10)

                else:
                    messagebox.showerror("Ошибка", f"Запись с первичным ключом {primary_key_value} не найдена.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при загрузке данных: {e}")
            finally:
                conn.close()

    def submit_update(self):
        """Отправить данные на обновление записи в таблице."""
        table_name = self.table_combobox.get()
        primary_key_value = self.primary_key_entry.get()  # Получаем первичный ключ

        if not primary_key_value:
            messagebox.showerror("Ошибка", "Введите значение первичного ключа.")
            return

        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Собираем данные из полей ввода
                columns = list(self.inputs.keys())
                values = [self.inputs[column].get() for column in columns]

                # Формируем SQL запрос для обновления данных
                set_values = ", ".join([f"{columns[i]} = ?" for i in range(len(columns))])
                cursor.execute(f"UPDATE {table_name} SET {set_values} WHERE {self.primary_key_column} = ?",
                               values + [primary_key_value])
                conn.commit()

                messagebox.showinfo("Успех", "Данные успешно обновлены!")
                self.clear_input_fields()

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при обновлении данных: {e}")
            finally:
                conn.close()

    def submit_update(self):
        """Отправить данные на обновление записи в таблице."""
        table_name = self.table_combobox.get()
        primary_key_value = self.primary_key_entry.get()  # Получаем первичный ключ

        if not primary_key_value:
            messagebox.showerror("Ошибка", "Введите значение первичного ключа.")
            return

        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Собираем данные из полей ввода
                columns = list(self.inputs.keys())
                values = [self.inputs[column].get() for column in columns]

                # Формируем SQL запрос для обновления данных
                set_values = ", ".join([f"{columns[i]} = ?" for i in range(len(columns))])
                cursor.execute(f"UPDATE {table_name} SET {set_values} WHERE {self.primary_key_column} = ?",
                               values + [primary_key_value])
                conn.commit()

                messagebox.showinfo("Успех", "Данные успешно обновлены!")
                self.clear_input_fields()

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при обновлении данных: {e}")
            finally:
                conn.close()

    def delete_data(self):
        """Удалить запись из выбранной таблицы."""
        table_name = self.table_combobox.get()

        # Очистка предыдущих полей ввода
        for widget in self.input_fields_frame.winfo_children():
            widget.destroy()

        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Получаем имена столбцов для выбранной таблицы
                cursor.execute(f"SELECT * FROM {table_name} WHERE 1=0")  # Не выбираем данные, а только структуру
                columns = [column[0] for column in cursor.description]

                # Динамически создаем поля ввода для первичного ключа (для удаления)
                self.inputs = {}
                primary_key_column = columns[0]  # Предполагаем, что первый столбец - это первичный ключ

                label = tk.Label(self.input_fields_frame, text=f"Введите {primary_key_column} для удаления")
                label.grid(row=0, column=0, padx=10, pady=5)

                entry = tk.Entry(self.input_fields_frame)
                entry.grid(row=0, column=1, padx=10, pady=5)
                self.inputs[primary_key_column] = entry

                # Кнопка для удаления
                self.submit_delete_button = tk.Button(self.input_fields_frame, text="Удалить",
                                                      command=self.submit_delete)
                self.submit_delete_button.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при получении структуры таблицы: {e}")
            finally:
                conn.close()

    def submit_delete(self):
        """Удалить запись из таблицы по первичному ключу."""
        table_name = self.table_combobox.get()
        conn = self.connect_to_db()
        if conn:
            cursor = conn.cursor()
            try:
                # Получаем данные из поля ввода (предполагаем, что это первичный ключ)
                primary_key_value = self.inputs[list(self.inputs.keys())[0]].get()

                cursor.execute(f"DELETE FROM {table_name} WHERE {list(self.inputs.keys())[0]} = ?",
                               (primary_key_value,))
                conn.commit()

                messagebox.showinfo("Успех", "Запись успешно удалена!")
                self.clear_input_fields()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при удалении данных: {e}")
            finally:
                conn.close()

    def clear_input_fields(self):
        """Очистить все поля ввода."""
        for widget in self.input_fields_frame.winfo_children():
            widget.destroy()


def main():
    # Путь к базе данных MS Access
    db_path = r"C:\Users\roma2\PycharmProjects\bd_6lab\bd_3lab (1).accdb"

    # Создание окна
    root = tk.Tk()
    app = DatabaseApp(root, db_path)
    root.mainloop()


if __name__ == "__main__":
    main()
