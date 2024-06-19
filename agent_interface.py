import sys
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, Toplevel, Label, Entry, Button, scrolledtext
from tkinter import ttk
import subprocess
import psycopg2
import imaplib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
from email.parser import BytesParser
from datetime import datetime,date
import pandas as pd
import tkinter.font as tkFont
import openpyxl
# Подключение к базе данных PostgreSQL
conn = psycopg2.connect(
    dbname="diplom2",
    user="postgres",
    password="123",
    host="localhost",
    port="5432"
)
current_user_id = int(sys.argv[1])
# Конфигурация почтового сервера
IMAP_SERVER = "imap.mail.ru"
SMTP_SERVER = "smtp.mail.ru"
EMAIL = "durandil2@mail.ru"
PASSWORD = "rpE8GUAGemuLEv6CzhrY"
def show_profile():
    try:
        # Создаем курсор для выполнения SQL-запросов
        cursor = conn.cursor()

        # Выполняем SQL-запрос для получения информации о профиле агента с id=1
        cursor.execute("SELECT id, login,ФИО,email,номер_телефона,отдел FROM агенты WHERE id = %s", (current_user_id,))

        # Получаем результат запроса
        agent_profile = cursor.fetchone()

        # Закрываем курсор
        cursor.close()

        # Очищаем содержимое правой области
        clear_right_frame()

        # Создаем внутреннюю рамку для отображения содержимого
        inner_frame = tk.Frame(right_frame, bg="white")
        inner_frame.pack(expand=True, fill="both")

        # Создаем рамку для таблицы информации о профиле
        profile_frame = tk.Frame(inner_frame, bg="white", bd=2, relief="groove")
        profile_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # Создаем стиль для жирного шрифта
        bold_font = ("Helvetica", 12, )

        # Создаем таблицу с заголовками и информацией
        for i, label in enumerate(["Агент ID", "Логин", "ФИО", "Email", "Телефон", "Отдел"]):
            # Создаем Label для заголовка
            header_label = tk.Label(profile_frame, text=label, bg="white", font=bold_font, padx=20, pady=10, anchor="w")
            header_label.grid(row=i, column=0, sticky="ew")

            # Создаем Entry для информации
            profile_info = tk.Entry(profile_frame, bg="white", font=("Helvetica", 16), bd=0)
            profile_info.insert(0, agent_profile[i])
            profile_info.grid(row=i, column=1, sticky="ew")

            # Применяем зебру для каждой строки
            # if i % 2 == 0:
            #     header_label.config(bg="#f0f0f0")
            #     profile_info.config(bg="#f0f0f0")

        # Настраиваем размеры столбцов
        profile_frame.grid_columnconfigure(0, weight=1)
        profile_frame.grid_columnconfigure(1, weight=1)

    except Exception as e:
        # Если произошла ошибка, выводим сообщение об ошибке
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def clear_right_frame():
    # Очищаем содержимое правой области
    for widget in right_frame.winfo_children():
        widget.destroy()
# get_accessible_tables возвращает список таблиц, доступных для текущего пользователя
def get_accessible_tables(user_id):
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT доступ FROM агенты WHERE id = %s", (user_id,))
        result = cursor.fetchone()
        cursor.close()
        if result:
            return result[0].split(',')
        return []
    except Exception as e:
        print(f"Ошибка при получении доступных таблиц: {e}")
        return []

def show_databases():
    try:
        accessible_tables = get_accessible_tables(current_user_id)
        if not accessible_tables:
            print("У текущего пользователя нет доступа к таблицам.")
            return

        cursor = conn.cursor()
        cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'")
        all_tables = cursor.fetchall()
        cursor.close()

        clear_right_frame()

        inner_frame = tk.Frame(right_frame, bg="white")
        inner_frame.pack(expand=True, fill="both")

        canvas = tk.Canvas(inner_frame, bg="white", bd=0, highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)

        database_frame = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=database_frame, anchor="nw", width=right_frame.winfo_width())

        v_scrollbar = tk.Scrollbar(inner_frame, orient="vertical", command=canvas.yview)
        v_scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=v_scrollbar.set)

        accessible_tables_set = set(accessible_tables)
        for table in all_tables:
            table_name = table[0]
            if table_name in accessible_tables_set:
                button = tk.Button(database_frame, text=table_name, command=lambda tn=table_name: show_table_content(tn), bg="#e7e7e7", font=("Helvetica", 12))
                button.pack(fill="both", padx=10, pady=5)

        button_style = {
            "bg": "#8EC6C5",
            "fg": "black",
            "font": ("Helvetica", 12),
            "relief": "flat",
            "padx": 10,
            "pady": 5,
            "bd": 0,
            "highlightthickness": 0,
            "highlightbackground": "white"
        }

        add_button = tk.Button(database_frame, text="Создать таблицу", command=open_create_table_window, **button_style)
        add_button.pack(fill=tk.X, padx=10, pady=5)

        database_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    except Exception as e:
        print(f"Произошла ошибка: {e}")

def show_table_content(table_name):
    try:
        clear_right_frame()

        # Buttons creation
        button_frame = tk.Frame(right_frame, bg="white")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        button_style = {
            "bg": "#8EC6C5",
            "fg": "black",
            "font": ("Helvetica", 12),
            "relief": "flat",
            "padx": 10,
            "pady": 5
        }

        add_button = tk.Button(button_frame, text="Добавить запись", command=lambda: add_record(table_name, column_names), **button_style)
        add_button.pack(side=tk.LEFT, padx=10, pady=5)

        edit_button = tk.Button(button_frame, text="Редактировать запись", command=lambda: edit_record(tree, table_name), **button_style)
        edit_button.pack(side=tk.LEFT, padx=10, pady=5)

        delete_button = tk.Button(button_frame, text="Удалить запись", command=lambda: remove_record(tree, table_name), **button_style)
        delete_button.pack(side=tk.LEFT, padx=10, pady=5)

        import_button = tk.Button(button_frame, text="Импорт из Excel", command=lambda: import_from_excel(table_name, column_names), **button_style)
        import_button.pack(side=tk.LEFT, padx=10, pady=5)

        sort_button = tk.Button(button_frame, text="Сортировать данные", command=lambda: sort_records(tree, table_name, column_names), **button_style)
        sort_button.pack(side=tk.LEFT, padx=10, pady=5)

        search_button = tk.Button(button_frame, text="Поиск записей", command=lambda: search_records(tree, table_name, column_names), **button_style)
        search_button.pack(side=tk.LEFT, padx=10, pady=5)

        cursor = conn.cursor()
        cursor.execute(f"SELECT column_name FROM information_schema.columns WHERE table_name = %s ORDER BY ordinal_position", (table_name,))
        column_names = [row[0] for row in cursor.fetchall()]

        cursor.execute(f"SELECT {', '.join(column_names)} FROM {table_name}")
        records = cursor.fetchall()
        cursor.close()

        if not records:
            label_no_records = tk.Label(right_frame, text="No records found", font=("Helvetica", 14), bg="white")
            label_no_records.pack(pady=10)
            return

        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica", 12), rowheight=25)
        style.configure("Treeview.Heading", font=("Helvetica", 14, "bold"))
        style.map("Treeview", background=[("selected", "#347083")], foreground=[("selected", "white")])

        style.configure("Treeview.oddrow", background="white")
        style.configure("Treeview.evenrow", background="#f2f2f2")

        tree = ttk.Treeview(right_frame, columns=column_names, show="headings", selectmode="browse")
        for col in column_names:
            tree.heading(col, text=col)
            tree.column(col, anchor="w", width=tkFont.Font().measure(col) + 20)

        for i, record in enumerate(records):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", "end", values=record, tags=(tag,))

        tree.pack(expand=True, fill=tk.BOTH)

    except Exception as e:
        error_label = tk.Label(right_frame, text=f"Error: {e}", font=("Helvetica", 14), bg="white", fg="red")
        error_label.pack(pady=10)



    except Exception as e:
        print("An error occurred:", e)

        # Add expand/collapse functionality
        def toggle_column(column):
            current_width = tree.column(column, 'width')
            if current_width > 0:
                tree.column(column, width=0)
            else:
                tree.column(column, width=tkFont.Font().measure(column) + 20)

        for col in column_names:
            tree.heading(col, text=col, command=lambda _col=col: toggle_column(_col))

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def open_create_table_window():
    create_table_window = tk.Toplevel(root)
    create_table_window.title("Создать новую таблицу")

    tk.Label(create_table_window, text="Название таблицы:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    table_name_entry = tk.Entry(create_table_window)
    table_name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    columns = []

    def add_column():
        row = len(columns) + 2  # Учитываем уже существующие элементы (эти кнопки)
        column_name = tk.Entry(create_table_window)
        column_name.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        column_type = ttk.Combobox(create_table_window, values=["INTEGER", "TEXT", "DATE", "BOOLEAN", "FLOAT"])
        column_type.grid(row=row, column=1, padx=10, pady=5, sticky="w")
        columns.append((column_name, column_type))

        # Скрытие и перемещение кнопок
        add_column_button.grid_forget()
        create_table_button.grid_forget()
        add_column_button.grid(row=row + 1, column=0, columnspan=2, padx=10, pady=5, sticky="sw")
        create_table_button.grid(row=row + 1, column=0, columnspan=2, padx=10, pady=5, sticky="se")

        # Увеличение высоты окна при добавлении столбца
        create_table_window.geometry(f"{create_table_window.winfo_width()}x{create_table_window.winfo_height() + 40}")
    def create_table():
        table_name = table_name_entry.get()
        if not table_name:
            messagebox.showwarning("Warning", "Введите название таблицы")
            return

        column_definitions = []
        # Помещаем столбец id в начало списка
        column_definitions.append("id SERIAL PRIMARY KEY")
        for column_name, column_type in columns:
            name = column_name.get().strip()  # Удаляем лишние пробелы
            type_ = column_type.get()
            if not name or not type_:
                messagebox.showwarning("Warning", "Заполните все поля столбцов")
                return
            column_definitions.append(f"{name} {type_}")

        # Добавление необходимых столбцов в конец таблицы
        additional_columns = [
            "обработано TEXT",
            "комментарий TEXT",
            "id_агента INTEGER",
            "дата TIMESTAMP WITHOUT TIME ZONE"
        ]
        column_definitions.extend(additional_columns)

        cursor = conn.cursor()
        cursor.execute(f"CREATE TABLE {table_name} ({', '.join(column_definitions)})")
        conn.commit()

        # Добавление названия новой таблицы в таблицу агентов
        cursor.execute("SELECT доступ FROM агенты WHERE id = %s", (current_user_id,))
        access = cursor.fetchone()[0]
        if access:
            access += f",{table_name}"  # Удаляем пробел после запятой
        else:
            access = table_name
        cursor.execute("UPDATE агенты SET доступ = %s WHERE id = %s", (access, current_user_id))
        conn.commit()
        cursor.close()

        messagebox.showinfo("Info", "Таблица успешно создана")
        create_table_window.destroy()
        show_databases()
    add_column_button = tk.Button(create_table_window, text="Добавить столбец", command=add_column)
    add_column_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="sw")

    create_table_button = tk.Button(create_table_window, text="Создать таблицу", command=create_table)
    create_table_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="se")


def add_record(table_name, column_names):
    record_window = tk.Toplevel(root)
    record_window.title("Добавить запись")

    entries = {}
    cursor = conn.cursor()
    cursor.execute(f"SELECT MAX(id) FROM {table_name}")
    max_id = cursor.fetchone()[0]
    cursor.close()

    for idx, col in enumerate(column_names):
        tk.Label(record_window, text=col).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(record_window)
        entry.grid(row=idx, column=1, padx=10, pady=5, sticky="w")
        if col == 'id':
            entry.insert(0, max_id + 1 if max_id is not None else 1)  # Заполнение id автоматически
            entry.config(state='readonly')  # Делаем поле только для чтения
        elif col == 'id_агента':
            entry.insert(0, current_user_id)  # Заполнение id_агента текущим пользователем
            entry.config(state='readonly')  # Делаем поле только для чтения
        elif col == 'дата':
            # Заполнение даты текущей датой с временным промежутком
            now = datetime.now()
            entry.insert(0, now.strftime("%Y-%m-%d %H:%M:%S"))
            entry.config(state='readonly')

        entries[col] = entry

    def save_record():
        values = {col: entry.get() for col, entry in entries.items()}
        cursor = conn.cursor()
        columns_str = ', '.join([col for col, value in values.items() if value])
        placeholders_str = ', '.join(['%s'] * len(columns_str.split(', ')))
        cursor.execute(f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders_str})", [values[col] for col in columns_str.split(', ')])
        conn.commit()
        cursor.close()
        messagebox.showinfo("Info", "Запись добавлена успешно")
        record_window.destroy()
        show_table_content(table_name)

    save_button = tk.Button(record_window, text="Сохранить", command=save_record)
    save_button.grid(row=len(column_names), column=0, columnspan=2, padx=10, pady=5, sticky="ew")
def edit_record(tree, table_name):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите запись для редактирования")
        return

    item_values = tree.item(selected_item, "values")
    edit_window = tk.Toplevel(root)
    edit_window.title("Редактировать запись")

    entries = {}
    column_names = tree["columns"]
    for idx, col in enumerate(column_names):
        tk.Label(edit_window, text=col).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(edit_window)
        entry.grid(row=idx, column=1, padx=10, pady=5, sticky="w")
        entry.insert(0, item_values[idx])
        entries[col] = entry

    def update_record():
        values = {col: entry.get() for col, entry in entries.items()}
        cursor = conn.cursor()
        cursor.execute(f"UPDATE {table_name} SET {', '.join([f'{col} = %s' for col in values.keys()])} WHERE id = %s", list(values.values()) + [item_values[0]])
        conn.commit()
        cursor.close()
        messagebox.showinfo("Info", "Запись обновлена успешно")
        edit_window.destroy()
        show_table_content(table_name)

    save_button = tk.Button(edit_window, text="Сохранить", command=update_record)
    save_button.grid(row=len(column_names), column=0, columnspan=2, padx=10, pady=5, sticky="ew")

def remove_record(tree, table_name):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите запись для удаления")
        return

    item_values = tree.item(selected_item, "values")
    cursor = conn.cursor()
    cursor.execute(f"DELETE FROM {table_name} WHERE id = %s", (item_values[0],))
    conn.commit()
    cursor.close()
    messagebox.showinfo("Info", "Запись удалена успешно")
    show_table_content(table_name)


def import_from_excel(table_name, column_names):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    df = pd.read_excel(file_path, engine='openpyxl')
    columns_in_file = df.columns.tolist()
    columns_to_import = [col for col in columns_in_file if col in column_names[1:]]

    if not columns_to_import:
        messagebox.showwarning("Warning", "No matching columns found in the Excel file")
        return

    cursor = conn.cursor()
    for _, row in df.iterrows():
        values = [row[col] for col in columns_to_import]
        cursor.execute(f"INSERT INTO {table_name} ({', '.join(columns_to_import)}) VALUES ({', '.join(['%s'] * len(columns_to_import))})", values)

    conn.commit()
    cursor.close()

    messagebox.showinfo("Success", "Data imported successfully")
    show_table_content(table_name)

def sort_records(tree, table_name, column_names):
    sort_window = tk.Toplevel(root)
    sort_window.title("Сортировка данных")

    tk.Label(sort_window, text="Выберите столбец для сортировки:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    column_combo = ttk.Combobox(sort_window, values=column_names)
    column_combo.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    tk.Label(sort_window, text="Тип сортировки:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    sort_type_combo = ttk.Combobox(sort_window, values=["По возрастанию", "По убыванию"])
    sort_type_combo.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    def sort_data():
        column = column_combo.get()
        sort_type = sort_type_combo.get()
        if not column or not sort_type:
            messagebox.showwarning("Warning", "Выберите столбец и тип сортировки")
            return

        sort_order = "ASC" if sort_type == "По возрастанию" else "DESC"
        cursor = conn.cursor()
        cursor.execute(f"SELECT {', '.join(column_names)} FROM {table_name} ORDER BY {column} {sort_order}")
        sorted_records = cursor.fetchall()
        cursor.close()

        tree.delete(*tree.get_children())
        for i, record in enumerate(sorted_records):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", "end", values=record, tags=(tag,))

        sort_window.destroy()

    sort_button = tk.Button(sort_window, text="Сортировать", command=sort_data)
    sort_button.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

def search_records(tree, table_name, column_names):
    search_window = tk.Toplevel(root)
    search_window.title("Поиск записей")

    tk.Label(search_window, text="Выберите столбец для поиска:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    column_combo = ttk.Combobox(search_window, values=column_names)
    column_combo.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    tk.Label(search_window, text="Введите значение для поиска:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    search_value_entry = tk.Entry(search_window)
    search_value_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    def search_data():
        column = column_combo.get()
        search_value = search_value_entry.get()
        if not column or not search_value:
            messagebox.showwarning("Warning", "Выберите столбец и введите значение для поиска")
            return

        cursor = conn.cursor()
        cursor.execute(f"SELECT {', '.join(column_names)} FROM {table_name} WHERE {column}::text ILIKE %s", (f"%{search_value}%",))
        search_results = cursor.fetchall()
        cursor.close()

        tree.delete(*tree.get_children())
        for i, record in enumerate(search_results):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", "end", values=record, tags=(tag,))

        search_window.destroy()

    search_button = tk.Button(search_window, text="Поиск", command=search_data)
    search_button.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

def logout():
    if messagebox.askokcancel("Logout", "Are you sure you want to logout?"):
        root.destroy()
        subprocess.Popen(["python", "login.py"])

def show_manager_info():
    try:
        # Создаем курсор для выполнения SQL-запросов
        cursor = conn.cursor()


        # получаем id_руководителя из таблицы агенты
        cursor.execute("SELECT id_руководителя FROM агенты WHERE id = %s", (current_user_id,))
        id_руководителя = cursor.fetchone()[0]

        # используем id_руководителя для получения информации из таблицы руководители
        cursor.execute("SELECT id,ФИО,email,номер_телефона FROM руководители WHERE id = %s", (id_руководителя,))

        # Получаем результат запроса
        manager_profile = cursor.fetchone()

        # Закрываем курсор
        cursor.close()

        # Очищаем содержимое правой области
        clear_right_frame()

        # Создаем внутреннюю рамку для отображения содержимого
        inner_frame = tk.Frame(right_frame, bg="white")
        inner_frame.pack(expand=True, fill="both")

        # Создаем рамку для таблицы информации о профиле
        profile_frame = tk.Frame(inner_frame, bg="white", bd=2, relief="groove")
        profile_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # Создаем стиль для жирного шрифта
        bold_font = ("Helvetica", 14, "bold")

        # Создаем таблицу с заголовками и информацией
        for i, label in enumerate(["ID Руководителя", "ФИО", "Email", "Телефон"]):
            # Создаем Label для заголовка
            header_label = tk.Label(profile_frame, text=label, bg="white", font=bold_font, padx=20, pady=10, anchor="w")
            header_label.grid(row=i, column=0, sticky="ew")

            # Создаем Entry для информации
            profile_info = tk.Entry(profile_frame, bg="white", font=("Helvetica", 14), bd=0)
            profile_info.insert(0, manager_profile[i])
            profile_info.grid(row=i, column=1, sticky="ew")

        # Настраиваем размеры столбцов
        profile_frame.grid_columnconfigure(0, weight=1)
        profile_frame.grid_columnconfigure(1, weight=1)

    except Exception as e:
        # Если произошла ошибка, выводим сообщение об ошибке
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

sent_emails = []  # Список для хранения отправленных писем

def fetch_emails():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        result, data = mail.search(None, "ALL")
        email_ids = data[0].split()

        emails = []
        for email_id in email_ids[-10:]:  # Fetch last 10 emails
            result, msg_data = mail.fetch(email_id, "(RFC822)")
            msg = BytesParser().parsebytes(msg_data[0][1])
            subject, encoding = decode_header(msg["subject"])[0]
            if isinstance(subject, bytes):
                try:
                    subject = subject.decode(encoding if encoding else "utf-8", errors='replace')
                except LookupError:
                    subject = subject.decode('utf-8', errors='replace')
            from_, encoding = decode_header(msg.get("From"))[0]
            if isinstance(from_, bytes):
                try:
                    from_ = from_.decode(encoding if encoding else "utf-8", errors='replace')
                except LookupError:
                    from_ = from_.decode('utf-8', errors='replace')
            emails.append((email_id, subject, from_, msg))

        mail.logout()
        return emails
    except Exception as e:
        messagebox.showerror("Error", f"Failed to fetch emails: {str(e)}")
        return []

def show_emails():
    clear_right_frame()

    emails = fetch_emails()

    # Создаем внутреннюю рамку для отображения содержимого
    inner_frame = tk.Frame(right_frame, bg="white")
    inner_frame.pack(expand=True, fill="both")

    # Создаем Canvas для отображения содержимого
    canvas = tk.Canvas(inner_frame, bg="white", bd=0, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    # Создаем рамку для отображения писем
    email_frame = tk.Frame(canvas, bg="white")
    canvas.create_window((0, 0), window=email_frame, anchor="nw")

    # Создаем вертикальную прокрутку для Canvas
    scrollbar = tk.Scrollbar(inner_frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Привязываем область прокрутки к холсту
    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas.bind("<Configure>", on_configure)

    bold_font = ("Helvetica", 14, "bold")

    for i, (email_id, subject, from_, msg) in enumerate(emails):
        subject_label = tk.Label(email_frame, text=f"{subject} (от: {from_})", bg="white", font=bold_font, padx=20, pady=10, anchor="w")
        subject_label.grid(row=i, column=0, sticky="ew")

        if i % 2 == 0:
            subject_label.config(bg="#f0f0f0")

        subject_label.bind("<Button-1>", lambda event, email_id=email_id, msg=msg, subject=subject, from_=from_: show_email_content(email_id, msg, subject, from_))

    email_frame.grid_columnconfigure(0, weight=1)

    # Добавляем кнопку "Отправить письмо" и "Отправленные"
    send_email_button = tk.Button(right_frame, text="Отправить письмо", command=open_send_email_window, bg="lightgrey", font=("Helvetica", 14), padx=10, pady=10)
    send_email_button.pack(side="left", pady=10, padx=5)

    sent_button = tk.Button(right_frame, text="Отправленные", command=show_sent_emails, bg="lightgrey", font=("Helvetica", 14), padx=10, pady=10)
    sent_button.pack(side="left", pady=10, padx=5)

def show_email_content(email_id, msg, subject, from_):
    clear_right_frame()

    inner_frame = tk.Frame(right_frame, bg="white")
    inner_frame.pack(expand=True, fill="both")

    subject_label = tk.Label(inner_frame, text=subject, bg="white", font=("Helvetica", 14, "bold"), padx=20, pady=10, anchor="w")
    subject_label.pack()

    from_label = tk.Label(inner_frame, text=f"От: {from_}", bg="white", font=("Helvetica", 12), padx=20, pady=5, anchor="w")
    from_label.pack()

    # Создаем Canvas для отображения содержимого
    canvas = tk.Canvas(inner_frame, bg="white", bd=0, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    # Создаем рамку для отображения содержимого письма
    email_content_frame = tk.Frame(canvas, bg="white")
    canvas.create_window((0, 0), window=email_content_frame, anchor="nw")

    # Создаем вертикальную прокрутку для Canvas
    scrollbar = tk.Scrollbar(inner_frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Привязываем область прокрутки к холсту
    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas.bind("<Configure>", on_configure)

    bold_font = ("Helvetica", 14, "bold")

    for part in msg.walk():
        if part.get_content_type() == "text/plain":
            body = part.get_payload(decode=True)
            charset = part.get_content_charset()
            if charset is None:
                charset = 'utf-8'
            try:
                body = body.decode(charset, errors='replace')
            except LookupError:
                body = body.decode('utf-8', errors='replace')

            body_label = tk.Label(email_content_frame, text=body, bg="white", font=("Helvetica", 12), padx=20, pady=10, anchor="w", justify="left", wraplength=500)
            body_label.pack(expand=True, fill="both")

def open_send_email_window():
    send_email_window = tk.Toplevel(root)
    send_email_window.title("Отправить письмо")
    send_email_window.geometry("600x600")

    tk.Label(send_email_window, text="Получатель:", font=("Helvetica", 12)).pack(pady=5)
    recipient_entry = tk.Entry(send_email_window, width=50)
    recipient_entry.pack(pady=5)

    tk.Label(send_email_window, text="Тема:", font=("Helvetica", 12)).pack(pady=5)
    subject_entry = tk.Entry(send_email_window, width=50)
    subject_entry.pack(pady=5)

    tk.Label(send_email_window, text="Сообщение:", font=("Helvetica", 12)).pack(pady=5)
    body_text = scrolledtext.ScrolledText(send_email_window, wrap=tk.WORD, width=50, height=10)
    body_text.pack(pady=5)

    def send_email_action():
        recipient = recipient_entry.get()
        subject = subject_entry.get()
        body = body_text.get("1.0", tk.END)

        if not recipient or not subject or not body:
            messagebox.showerror("Error", "All fields are required!")
            return

        msg = MIMEMultipart()
        msg["From"] = EMAIL
        msg["To"] = recipient
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        try:
            server = smtplib.SMTP(SMTP_SERVER, 587)
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.sendmail(EMAIL, recipient, msg.as_string())
            server.quit()

            sent_emails.append((recipient, subject, body))

            messagebox.showinfo("Success", "Email sent successfully!")
            send_email_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")

    send_button = tk.Button(send_email_window, text="Отправить", command=send_email_action, bg="lightgrey", font=("Helvetica", 12), padx=10, pady=10)
    send_button.pack(pady=10)

def show_sent_emails():
    clear_right_frame()

    # Создаем внутреннюю рамку для отображения содержимого
    inner_frame = tk.Frame(right_frame, bg="white")
    inner_frame.pack(expand=True, fill="both")

    # Создаем Canvas для отображения содержимого
    canvas = tk.Canvas(inner_frame, bg="white", bd=0, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    # Создаем рамку для отображения отправленных писем
    sent_email_frame = tk.Frame(canvas, bg="white")
    canvas.create_window((0, 0), window=sent_email_frame, anchor="nw")

    # Создаем вертикальную прокрутку для Canvas
    scrollbar = tk.Scrollbar(inner_frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Привязываем область прокрутки к холсту
    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas.bind("<Configure>", on_configure)

    bold_font = ("Helvetica", 14, "bold")

    for i, (recipient, subject, body) in enumerate(sent_emails):
        subject_label = tk.Label(sent_email_frame, text=f"{subject} (кому: {recipient})", bg="white", font=bold_font, padx=20, pady=10, anchor="w")
        subject_label.grid(row=i, column=0, sticky="ew")

        if i % 2 == 0:
            subject_label.config(bg="#f0f0f0")

        subject_label.bind("<Button-1>", lambda event, subject=subject, recipient=recipient, body=body: show_sent_email_content(subject, recipient, body))

    sent_email_frame.grid_columnconfigure(0, weight=1)

def show_sent_email_content(subject, recipient, body):
    clear_right_frame()

    inner_frame = tk.Frame(right_frame, bg="white")
    inner_frame.pack(expand=True, fill="both")

    subject_label = tk.Label(inner_frame, text=subject, bg="white", font=("Helvetica", 14, "bold"), padx=20, pady=10, anchor="w")
    subject_label.pack()

    recipient_label = tk.Label(inner_frame, text=f"Кому: {recipient}", bg="white", font=("Helvetica", 12), padx=20, pady=5, anchor="w")
    recipient_label.pack()

    body_label = tk.Label(inner_frame, text=body, bg="white", font=("Helvetica", 12), padx=20, pady=10, anchor="w", justify="left", wraplength=500)
    body_label.pack(expand=True, fill="both")


def fetch_notifications():
    cursor = conn.cursor()
    cursor.execute("SELECT type, description, start_date FROM tasks_notifications")
    notifications = cursor.fetchall()
    cursor.close()
    return notifications


def calculate_time_difference(start_date):
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, '%Y-%m-%d %H:%M:%S')
    now = datetime.now()
    diff = now - start_date
    days = diff.days
    hours, remainder = divmod(diff.seconds, 3600)
    if days > 0:
        return f"{days} дней {hours} ч назад"
    elif hours > 0:
        return f"{hours} ч назад"
    else:
        minutes, _ = divmod(remainder, 60)
        return f"{minutes} мин назад"


def check_task_completion(current_user_id, start_date, end_date):
    cursor = conn.cursor()
    cursor.execute("SELECT доступ FROM агенты WHERE id = %s", (current_user_id,))
    access_list = cursor.fetchone()[0].split(',')

    total_completed = 0
    for table in access_list:
        # Проверка наличия столбца "обработано" и "дата" в таблице
        cursor.execute(f"""
            SELECT EXISTS (
                SELECT 1 
                FROM information_schema.columns 
                WHERE table_name = %s AND column_name = 'обработано'
            )
        """, (table,))
        has_processed_column = cursor.fetchone()[0]

        cursor.execute(f"""
            SELECT EXISTS (
                SELECT 1 
                FROM information_schema.columns 
                WHERE table_name = %s AND column_name = 'дата'
            )
        """, (table,))
        has_date_column = cursor.fetchone()[0]

        if has_processed_column and has_date_column:
            cursor.execute(f"""
                SELECT COUNT(*)
                FROM {table}
                WHERE обработано = 'Да'
                AND дата BETWEEN %s AND %s
            """, (start_date, end_date))
            total_completed += cursor.fetchone()[0]

    cursor.close()
    return total_completed

def show_agent_tasks_notifications():
    clear_right_frame()

    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tasks_notifications WHERE agent_id = %s", (current_user_id,))
    tasks_notifications = cursor.fetchall()
    cursor.close()

    for task_notification in tasks_notifications:
        task_id, task_type, task_title, task_description, task_start_date, task_end_date, is_permanent, agent_id, record_count = task_notification

        bg_color = "lightblue" if task_type == 1 else "lightgreen"
        frame = tk.Frame(right_frame, bg=bg_color, bd=2, relief="solid", padx=10, pady=10)
        frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(frame, text=f"Тип: {'Задача' if task_type == 1 else 'Оповещение'}", font=("Helvetica", 14, "bold"),
                 bg=bg_color).pack(anchor="w")
        tk.Label(frame, text=f"Название: {task_title}", font=("Helvetica", 12), bg=bg_color).pack(anchor="w")
        tk.Label(frame, text=f"Описание: {task_description}", font=("Helvetica", 12), bg=bg_color).pack(anchor="w")
        tk.Label(frame, text=f"Дата появления: {task_start_date}", font=("Helvetica", 12), bg=bg_color).pack(anchor="w")

        if task_type == 1:
            completed_count = check_task_completion(agent_id, task_start_date, task_end_date)
            progress = min(completed_count / record_count, 1.0) * 100
            progress_var = tk.DoubleVar(value=progress)
            progress_bar = ttk.Progressbar(frame, variable=progress_var, maximum=100)
            progress_bar.pack(fill=tk.X, pady=5)
            tk.Label(frame, text=f"Выполнение задачи: {completed_count}/{record_count}", font=("Helvetica", 12),
                     bg=bg_color).pack(anchor="w")


# Создаем основное окно
root = tk.Tk()
root.title("Интерфейс агента")
root.geometry("1000x600")
root.configure(bg="white")

# Создаем верхнюю панель с заголовком
top_frame = tk.Frame(root, bg="#3b5998", height=50)
top_frame.pack(side="top", fill="x")

title_label = tk.Label(top_frame, text="Оповещения", bg="#3b5998", fg="white", font=("Helvetica", 16, "bold"))
title_label.pack(pady=10)

# Создаем левое меню
left_menu = tk.Frame(root, bg="#2c2c2c", width=200)
left_menu.pack(side="left", fill="y")

# Создаем кнопки в левом меню
buttons = [
    ("Профиль", show_profile),
    ("Мои таблицы", show_databases),
    ("Руководитель", show_manager_info),
    ("Почта", show_emails),
    ("Оповещения", show_agent_tasks_notifications),
    ("Выход", logout)
]

for button_text, button_command in buttons:
    button = tk.Button(left_menu, text=button_text, bg="#2c2c2c", fg="white", font=("Helvetica", 12), bd=0, anchor="w", command=button_command)
    button.pack(fill="x", pady=2)

# Создаем область для отображения информации справа
right_frame = tk.Canvas(root, bg="white")
right_frame.pack(side="right", expand=True, fill="both")

# Создаем внутреннюю рамку для отображения содержимого
inner_frame = tk.Frame(right_frame, bg="white")
inner_frame.pack(expand=True, fill="both")

# Функция для изменения размеров Canvas при изменении размеров окна
def resize_canvas(event):
    right_frame.configure(scrollregion=right_frame.bbox("all"))

root.bind("<Configure>", resize_canvas)

root.mainloop()

