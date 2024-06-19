import hashlib
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, Toplevel, Label, Entry, Button, scrolledtext
from tkinter import ttk
import tkinter.font as tkFont
import psycopg2
import subprocess
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import datetime as dt
from tkcalendar import DateEntry
import imaplib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import decode_header
from email.parser import BytesParser

# Конфигурация почтового сервера
IMAP_SERVER = "imap.mail.ru"
SMTP_SERVER = "smtp.mail.ru"
def clear_right_frame():
    for widget in right_frame.winfo_children():
        widget.destroy()
# Функция для отображения списка агентов в правой области
def show_agents():
    clear_right_frame()

    cursor = conn.cursor()
    cursor.execute("SELECT * FROM агенты ORDER BY id")
    agents = cursor.fetchall()
    cursor.close()

    if not agents:
        label_no_agents = tk.Label(right_frame, text="No agents found", font=("Helvetica", 14), bg="white")
        label_no_agents.pack(pady=10)
        return

    style = ttk.Style()
    style.configure("Treeview", font=("Helvetica", 12), rowheight=25)
    style.configure("Treeview.Heading", font=("Helvetica", 14, "bold"))
    style.map("Treeview", background=[("selected", "#347083")], foreground=[("selected", "white")])

    tree = ttk.Treeview(right_frame, columns=("id", "ФИО", "login", "password", "номер_телефона",
                                              "email", "отдел", "id_руководителя", "доступ"), show="headings")
    style_treeview(tree)

    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, anchor="w", width=tk.font.Font().measure(col) + 20)

    for i, agent in enumerate(agents):
        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        tree.insert("", "end", values=agent, tags=(tag,))

    tree.pack(expand=True, fill=tk.BOTH)

    button_frame = tk.Frame(right_frame, bg="white")
    button_frame.pack(fill=tk.X, pady=10)

    add_button = tk.Button(button_frame, text="Добавить агента", command=add_agent, **button_style)
    add_button.pack(side=tk.LEFT, padx=10, pady=5)

    edit_button = tk.Button(button_frame, text="Редактировать агента", command=lambda: edit_agent(tree), **button_style)
    edit_button.pack(side=tk.LEFT, padx=10, pady=5)

    delete_button = tk.Button(button_frame, text="Удалить агента", command=lambda: remove_agent(tree), **button_style)
    delete_button.pack(side=tk.LEFT, padx=10, pady=5)

    access_button = tk.Button(button_frame, text="Добавить доступ", command=lambda: add_access(tree), **button_style)
    access_button.pack(side=tk.LEFT, padx=10, pady=5)
def style_treeview(tree):
    tree.tag_configure('evenrow', background='lightgrey')
    tree.tag_configure('oddrow', background='white')
    tree.tag_configure('header', font=('Helvetica', 14, 'bold'))
    tree.tag_configure('cell', font=('Helvetica', 12))

def add_access(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите агента.")
        return

    agent_id = tree.item(selected_item[0], 'values')[0]
    access_window = tk.Toplevel()
    access_window.title("Добавить доступ агенту")

    available_tables = get_available_tables()

    listbox = tk.Listbox(access_window, selectmode=tk.MULTIPLE)
    for table in available_tables:
        listbox.insert(tk.END, table)
    listbox.pack(padx=10, pady=10)

    def save_access():
        selected_tables = [listbox.get(idx) for idx in listbox.curselection()]
        if selected_tables:
            update_agent_access(agent_id, selected_tables)
            access_window.destroy()
            show_agents()
        else:
            messagebox.showwarning("Warning", "Please select at least one table.")

    save_button = tk.Button(access_window, text="Сохранить", command=save_access, bg="#4CAF50", fg="white")
    save_button.pack(pady=10)

def get_available_tables():
    cursor = conn.cursor()
    cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")
    all_tables = [row[0] for row in cursor.fetchall()]
    cursor.close()

    excluded_tables = {"агенты", "руководители", "analiz", "tasks_notifications", "ad_compaigns"}
    available_tables = [table for table in all_tables if table not in excluded_tables]
    return available_tables

def update_agent_access(agent_id, selected_tables):
    access_list = ",".join(selected_tables)
    cursor = conn.cursor()
    cursor.execute("UPDATE агенты SET доступ = %s WHERE id = %s", (access_list, agent_id))
    conn.commit()
    cursor.close()
def add_agent():
    add_window = Toplevel(root)
    add_window.title("Добавить агента")

    labels = ["ФИО", "login", "password", "номер_телефона", "email", "отдел"]
    entries = {}

    for i, label in enumerate(labels):
        tk.Label(add_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(add_window)
        entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")
        entries[label] = entry

    def save_agent():
        agent_data = {label: entry.get() for label, entry in entries.items()}
        agent_data["password"] = hashlib.sha256(agent_data["password"].encode()).hexdigest()
        agent_data["id_руководителя"] = get_current_user_manager_id()

        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO агенты (ФИО, login, password, номер_телефона, email, отдел, id_руководителя) VALUES (%s, %s, %s, %s, %s, %s, %s)",
            tuple(agent_data.values())
        )
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Агент добавлен успешно")
        add_window.destroy()
        show_agents()

    save_button = tk.Button(add_window, text="Save", command=save_agent, bg="#4CAF50", fg="white")
    save_button.grid(row=len(labels), column=0, columnspan=2, pady=10)

def edit_agent(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите агента")
        return

    agent_data = tree.item(selected_item)["values"]
    agent_id = agent_data[0]

    edit_window = Toplevel(root)
    edit_window.title("Редактировать агента")

    labels = ["ФИО", "login","password", "номер_телефона", "email", "отдел"]
    entries = {}

    for i, (label, value) in enumerate(zip(labels, agent_data[1:])):
        tk.Label(edit_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(edit_window)
        entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")
        entry.insert(0, value)
        entries[label] = entry

    def save_agent():
        updated_data = {label: entry.get() for label, entry in entries.items()}
        cursor = conn.cursor()
        update_query = f"UPDATE агенты SET {', '.join([f'{label} = %s' for label in updated_data])} WHERE id = %s"
        cursor.execute(update_query, (*updated_data.values(), agent_id))
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Агент сохранен успешно")
        edit_window.destroy()
        show_agents()

    save_button = tk.Button(edit_window, text="Save", command=save_agent, bg="#4CAF50", fg="white")
    save_button.grid(row=len(labels), column=0, columnspan=2, pady=10)

def remove_agent(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите агента")
        return

    agent_data = tree.item(selected_item)["values"]
    agent_id = agent_data[0]

    if messagebox.askokcancel("Delete Agent", f"Вы уверены что хотите удалить агента ID {agent_id}?"):
        cursor = conn.cursor()
        cursor.execute("DELETE FROM агенты WHERE id = %s", (agent_id,))
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Агент удален успешно")
        show_agents()

def show_db_tables():
    clear_right_frame()

    canvas = tk.Canvas(right_frame)
    scrollbar = tk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
    table_frame = tk.Frame(canvas)

    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    canvas.create_window((0, 0), window=table_frame, anchor="nw", width=right_frame.winfo_width())

    cursor = conn.cursor()
    exclude_tables = ('агенты', 'руководители', 'tasks_notifications', 'analiz', 'ad_campaigns')
    query = "SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' AND table_name NOT IN %s"
    cursor.execute(query, (exclude_tables,))
    tables = cursor.fetchall()
    cursor.close()

    if not tables:
        label_no_tables = tk.Label(table_frame, text="Не найдено таблиц", font=("Helvetica", 12), bg="white")
        label_no_tables.pack(pady=10)
        return

    for table in tables:
        table_name = table[0]
        button = tk.Button(table_frame, text=table_name, command=lambda tn=table_name: show_table_content(tn), bg="#e7e7e7", font=("Helvetica", 12))
        button.pack(fill="both", padx=10, pady=5)


    add_button = tk.Button(table_frame, text="Создать таблицу", command=open_create_table_window, **button_style)
    add_button.pack(fill="both", padx=10, pady=5)

    table_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

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

        messagebox.showinfo("Info", "Таблица успешно создана")
        create_table_window.destroy()
        show_db_tables()
    add_column_button = tk.Button(create_table_window, text="Добавить столбец", command=add_column)
    add_column_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="sw")

    create_table_button = tk.Button(create_table_window, text="Создать таблицу", command=create_table)
    create_table_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="se")

def show_table_content(table_name):
    try:
        clear_right_frame()

        # Buttons creation
        button_frame = tk.Frame(right_frame, bg="white")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

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

        tree = ttk.Treeview(right_frame, columns=column_names, show="headings", selectmode="browse")
        for col in column_names:
            tree.heading(col, text=col)
            tree.column(col, anchor="w", width=tkFont.Font().measure(col) + 20)

        for i, record in enumerate(records):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            tree.insert("", "end", values=record, tags=(tag,))

        style_treeview(tree)
        tree.pack(expand=True, fill=tk.BOTH)

    except Exception as e:
        print("An error occurred:", e)

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

def add_record(table_name, column_names):
    add_window = Toplevel(root)
    add_window.title("Добавить запись")

    entries = {}

    for i, label in enumerate(column_names[1:]): 
        tk.Label(add_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(add_window)
        entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")
        entries[label] = entry

    def save_record():
        record_data = {label: entry.get() for label, entry in entries.items()}

        cursor = conn.cursor()
        insert_query = f"INSERT INTO {table_name} ({', '.join(record_data.keys())}) VALUES ({', '.join(['%s'] * len(record_data))})"
        cursor.execute(insert_query, tuple(record_data.values()))
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Запись добавлена успешно")
        add_window.destroy()
        show_table_content(table_name)

    save_button = tk.Button(add_window, text="Save", command=save_record, bg="#4CAF50", fg="white")
    save_button.grid(row=len(column_names), column=0, columnspan=2, pady=10)

def edit_record(tree, table_name):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите запись для редактирования")
        return

    record_data = tree.item(selected_item)["values"]
    record_id = record_data[0]

    edit_window = Toplevel(root)
    edit_window.title("Редактировать запись")

    entries = {}

    for i, (label, value) in enumerate(zip(tree["columns"][1:], record_data[1:])): 
        tk.Label(edit_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(edit_window)
        entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")
        entry.insert(0, value)
        entries[label] = entry

    def save_record():
        updated_data = {label: entry.get() for label, entry in entries.items()}

        cursor = conn.cursor()
        update_query = f"UPDATE {table_name} SET {', '.join([f'{label} = %s' for label in updated_data])} WHERE {tree['columns'][0]} = %s"
        cursor.execute(update_query, (*updated_data.values(), record_id))
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Запись успешно сохранена")
        edit_window.destroy()
        show_table_content(table_name)

    save_button = tk.Button(edit_window, text="Save", command=save_record, bg="#4CAF50", fg="white")
    save_button.grid(row=len(tree["columns"]), column=0, columnspan=2, pady=10)

def remove_record(tree, table_name):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите запись для удаления")
        return

    record_data = tree.item(selected_item)["values"]
    record_id = record_data[0]

    if messagebox.askokcancel("Delete Record", f"Вы уверены что хотите удалить запись ID {record_id}?"):
        cursor = conn.cursor()
        cursor.execute(f"DELETE FROM {table_name} WHERE {tree['columns'][0]} = %s", (record_id,))
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Запись удалена успешно")
        show_table_content(table_name)

# Функция для импорта данных из файла Excel
def import_from_excel(table_name, column_names):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    df = pd.read_excel(file_path, engine='openpyxl')
    columns_in_file = df.columns.tolist()
    columns_to_import = [col for col in columns_in_file if col in column_names[1:]]

    if not columns_to_import:
        messagebox.showwarning("Warning", "Не найдено схожих столбцов")
        return

    cursor = conn.cursor()
    for _, row in df.iterrows():
        values = [row[col] for col in columns_to_import]
        cursor.execute(f"INSERT INTO {table_name} ({', '.join(columns_to_import)}) VALUES ({', '.join(['%s'] * len(columns_to_import))})", values)
    conn.commit()
    cursor.close()

    messagebox.showinfo("Success", "Информация импортирована успешно")
    show_table_content(table_name)

def sort_records(tree, table_name, column_names):
    sort_window = Toplevel(root)
    sort_window.title("Сортировать данные")

    tk.Label(sort_window, text="Выберите столбец:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    column_name = ttk.Combobox(sort_window, values=column_names)
    column_name.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    tk.Label(sort_window, text="Порядок:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    sort_order = ttk.Combobox(sort_window, values=["ASC", "DESC"])
    sort_order.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    def sort():
        col = column_name.get()
        order = sort_order.get()
        if not col or not order:
            messagebox.showwarning("Warning", "Выберите столбец и порядок сортировки")
            return

        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM {table_name} ORDER BY {col} {order}")
        sorted_records = cursor.fetchall()
        cursor.close()

        for item in tree.get_children():
            tree.delete(item)

        for record in sorted_records:
            tree.insert("", "end", values=record)

        sort_window.destroy()

    sort_button = tk.Button(sort_window, text="Sort", command=sort, bg="#4CAF50", fg="white")
    sort_button.grid(row=2, column=0, columnspan=2, pady=10)

def search_records(tree, table_name, column_names):
    search_window = Toplevel(root)
    search_window.title("Поиск записей")

    tk.Label(search_window, text="Выберите столбец:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    column_name = ttk.Combobox(search_window, values=column_names)
    column_name.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    tk.Label(search_window, text="Значение:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    search_value = tk.Entry(search_window)
    search_value.grid(row=1, column=1, padx=10, pady=5, sticky="w")

    def search():
        col = column_name.get()
        val = search_value.get()
        if not col or not val:
            messagebox.showwarning("Warning", "Выберите столбец и введите значение для поиска")
            return

        cursor = conn.cursor()
        search_query = f"SELECT * FROM {table_name} WHERE {col}::TEXT LIKE %s"
        cursor.execute(search_query, (f"%{val}%",))
        search_results = cursor.fetchall()
        cursor.close()

        for item in tree.get_children():
            tree.delete(item)

        for record in search_results:
            tree.insert("", "end", values=record)

        search_window.destroy()

    search_button = tk.Button(search_window, text="Search", command=search, bg="#4CAF50", fg="white")
    search_button.grid(row=2, column=0, columnspan=2, pady=10)

# Проверка выполнения задачи
def check_task_completion(agent_id, required_count):
    cursor = conn.cursor()
    cursor.execute("SELECT доступ FROM агенты WHERE id = %s", (agent_id,))
    access_list = cursor.fetchone()[0].split(',')

    total_completed = 0
    for table in access_list:
        # Проверка наличия столбца "обработано" в таблице
        cursor.execute(f"""
            SELECT EXISTS (
                SELECT 1 
                FROM information_schema.columns 
                WHERE table_name = %s AND column_name = 'обработано'
            )
        """, (table,))
        if cursor.fetchone()[0]:  # Если столбец существует
            cursor.execute(f"SELECT COUNT(*) FROM {table} WHERE обработано = 'Да'")
            total_completed += cursor.fetchone()[0]

    cursor.close()
    return total_completed >= required_count

def add_task_notification():
    def save_task_notification():
        task_type = entry_type.get()
        task_title = entry_title.get()
        task_description = entry_description.get()
        task_start_date = start_date_entry.get()
        task_end_date = end_date_entry.get()
        is_permanent = is_permanent_var.get()
        record_count = entry_record_count.get() if task_type == "1" else None

        agent_ids = [agent_listbox.get(idx).split(":")[0] for idx in agent_listbox.curselection()]

        cursor = conn.cursor()
        for agent_id in agent_ids:
            cursor.execute(
                "INSERT INTO tasks_notifications (type, title, description, start_date, end_date, is_permanent, agent_id, record_count) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                (task_type, task_title, task_description, task_start_date, task_end_date, is_permanent, agent_id, record_count)
            )
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Оповещение добавлено успешно")
        task_window.destroy()
        show_tasks_notifications()

    task_window = tk.Toplevel()
    task_window.title("Добавить задачу/уведомление")

    labels = ["Тип", "Заголовок", "Описание", "Дата начала", "Дата окончания", "Постоянное", "Количество записей"]
    entries = {}

    is_permanent_var = tk.BooleanVar()  # Инициализация переменной до создания виджетов

    for idx, label_text in enumerate(labels):
        label = tk.Label(task_window, text=label_text)
        label.grid(row=idx, column=0, padx=10, pady=5)
        if label_text in ["Дата начала", "Дата окончания"]:
            entry = tk.Entry(task_window)
        elif label_text == "Постоянное":
            entry = tk.Checkbutton(task_window, variable=is_permanent_var)
        elif label_text == "Тип":
            entry = ttk.Combobox(task_window, values=["1", "2"])
        else:
            entry = tk.Entry(task_window)
        entry.grid(row=idx, column=1, padx=10, pady=5)
        entries[label_text] = entry

    entry_type = entries["Тип"]
    entry_title = entries["Заголовок"]
    entry_description = entries["Описание"]
    start_date_entry = entries["Дата начала"]
    end_date_entry = entries["Дата окончания"]
    entry_record_count = entries["Количество записей"]

    agent_listbox = tk.Listbox(task_window, selectmode=tk.MULTIPLE)
    agent_listbox.grid(row=len(labels), column=0, columnspan=2, padx=10, pady=5)
    cursor = conn.cursor()
    cursor.execute("SELECT id, ФИО FROM агенты")
    agents = cursor.fetchall()
    cursor.close()
    for agent in agents:
        agent_listbox.insert(tk.END, f"{agent[0]}: {agent[1]}")

    save_button = tk.Button(task_window, text="Сохранить", command=save_task_notification, bg="#4CAF50", fg="white")
    save_button.grid(row=len(labels) + 1, column=0, columnspan=2, pady=10)

def show_tasks_notifications():
    clear_right_frame()

    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tasks_notifications")
    tasks_notifications = cursor.fetchall()
    cursor.close()

    columns = ("ID", "Тип", "Заголовок", "Описание", "Дата начала", "Дата окончания", "Постоянное", "ID агента", "Количество записей", "Выполнено")
    style = ttk.Style()
    style.configure("Treeview", font=("Helvetica", 12), rowheight=25)
    style.configure("Treeview.Heading", font=("Helvetica", 14, "bold"))
    style.map("Treeview", background=[("selected", "#347083")], foreground=[("selected", "white")])

    tree = ttk.Treeview(right_frame, columns=columns, show="headings")
    style_treeview(tree)

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="w", width=tk.font.Font().measure(col) + 20)

    for i, task_notification in enumerate(tasks_notifications):
        task_data = list(task_notification)

        # Convert 'Постоянное' to 'Да' or 'Нет'
        task_data[6] = "Да" if task_data[6] else "Нет"

        # Handle 'Количество записей' column to show '-' if null
        task_data[8] = task_data[8] if task_data[8] is not None else "-"

        agent_id = task_data[7]
        record_count = task_data[8] if task_notification[1] == 1 else None
        completed = check_task_completion(agent_id, record_count) if record_count else "-"

        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        tree.insert("", "end", values=tuple(task_data) + (completed,), tags=(tag,))

    tree.pack(expand=True, fill=tk.BOTH)

    button_frame = tk.Frame(right_frame, bg="white")
    button_frame.pack(fill=tk.X, pady=10)

    add_button = tk.Button(button_frame, text="Добавить задачу/уведомление", command=add_task_notification, **button_style)
    add_button.pack(side=tk.LEFT, padx=10, pady=5)

    edit_button = tk.Button(button_frame, text="Редактировать", command=lambda: edit_task_notification(tree), **button_style)
    edit_button.pack(side=tk.LEFT, padx=10, pady=5)

    delete_button = tk.Button(button_frame, text="Удалить", command=lambda: remove_task_notification(tree), **button_style)
    delete_button.pack(side=tk.LEFT, padx=10, pady=5)

def edit_task_notification(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите оповещение для редактирования")
        return

    item = tree.item(selected_item)["values"]
    task_id = item[0]

    def update_task_notification():
        task_type = entry_type.get()
        task_title = entry_title.get()
        task_description = entry_description.get()
        task_start_date = start_date_entry.get()
        task_end_date = end_date_entry.get()
        is_permanent = is_permanent_var.get()
        record_count = entry_record_count.get() if task_type == "1" else None

        cursor = conn.cursor()
        cursor.execute(
            "UPDATE tasks_notifications SET type = %s, title = %s, description = %s, start_date = %s, end_date = %s, is_permanent = %s, record_count = %s WHERE id = %s",
            (task_type, task_title, task_description, task_start_date, task_end_date, is_permanent, record_count, task_id)
        )
        conn.commit()
        cursor.close()

        messagebox.showinfo("Success", "Оповещение отредактировано успешно")
        task_window.destroy()
        show_tasks_notifications()

    task_window = tk.Toplevel()
    task_window.title("Редактировать задачу/уведомление")

    labels = ["Тип", "Заголовок", "Описание", "Дата начала", "Дата окончания", "Постоянное", "Количество записей"]
    entries = {}
    is_permanent_var = tk.BooleanVar(value=item[6])  

    for idx, label_text in enumerate(labels):
        label = tk.Label(task_window, text=label_text)
        label.grid(row=idx, column=0, padx=10, pady=5)
        if label_text == "Тип":
            entry = ttk.Combobox(task_window, values=["1", "2"])
            entry.set(item[1])  
        elif label_text == "Заголовок":
            entry = tk.Entry(task_window)
            entry.insert(0, item[2])  
        elif label_text == "Описание":
            entry = tk.Entry(task_window)
            entry.insert(0, item[3])  
        elif label_text == "Дата начала":
            entry = tk.Entry(task_window)
            entry.insert(0, item[4])  
        elif label_text == "Дата окончания":
            entry = tk.Entry(task_window)
            entry.insert(0, item[5])  
        elif label_text == "Постоянное":
            entry = tk.Checkbutton(task_window, variable=is_permanent_var)
        elif label_text == "Количество записей":
            entry = tk.Entry(task_window)
            entry.insert(0, item[8] if item[1] == "1" else "")  
        entry.grid(row=idx, column=1, padx=10, pady=5)
        entries[label_text] = entry

    entry_type = entries["Тип"]
    entry_title = entries["Заголовок"]
    entry_description = entries["Описание"]
    start_date_entry = entries["Дата начала"]
    end_date_entry = entries["Дата окончания"]
    entry_record_count = entries["Количество записей"]

    save_button = tk.Button(task_window, text="Сохранить", command=update_task_notification, bg="#4CAF50", fg="white")
    save_button.grid(row=len(labels), column=0, columnspan=2, pady=10)

def remove_task_notification(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Выберите оповещение для удаления")
        return

    item = tree.item(selected_item)["values"]
    task_id = item[0]

    if messagebox.askokcancel("Delete", "Вы уверены что хотите удалить это оповещение?"):
        cursor = conn.cursor()
        cursor.execute("DELETE FROM tasks_notifications WHERE id = %s", (task_id,))
        conn.commit()
        cursor.close()
        show_tasks_notifications()

def fetch_emails():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        result, data = mail.search(None, "ALL")
        email_ids = data[0].split()

        emails = []
        for email_id in email_ids[-20:]:  
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
        messagebox.showerror("Error", f"Ошибка получения писем: {str(e)}")
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

    send_email_button = tk.Button(right_frame, text="Отправить письмо", command=open_send_email_window, **button_style)
    send_email_button.pack(side="left", pady=10, padx=5)

    sent_button = tk.Button(right_frame, text="Отправленные", command=show_sent_emails, **button_style)
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
    
def open_send_email_window():
    send_email_window = tk.Toplevel(root)
    send_email_window.title("Отправить письмо")
    send_email_window.geometry("400x300")

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
            messagebox.showerror("Error", "Заполните все поля!")
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

            messagebox.showinfo("Success", "Письмо отправлено успешно!")
            send_email_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Ошибка отправки письма: {str(e)}")

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

def logout():
    if messagebox.askokcancel("Logout", "Вы уверены что хотите выйти?"):
        root.destroy()
        subprocess.Popen(["python", "login.py"])


def send_ad_email(recipient, subject, body):
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
    except Exception as e:
        messagebox.showerror("Error", f"Не удалось отправить письмо: {str(e)}")


def fetch_emails_from_db_tables(tables):
    emails = []
    try:
        cur = conn.cursor()
        for table in tables.split(','):
            cur.execute(f"SELECT email FROM {table.strip()}")
            rows = cur.fetchall()
            for row in rows:
                emails.append(row[0])
        cur.close()
        conn.close()
    except Exception as e:
        messagebox.showerror("Error", f"Не удалось получить адреса электронной почты: {str(e)}")
    return emails


def open_create_ad_window():
    create_ad_window = tk.Toplevel(root)
    create_ad_window.title("Создать рекламную рассылку")
    create_ad_window.geometry("600x600")

    tk.Label(create_ad_window, text="Выбор таблиц:", font=("Helvetica", 12)).pack(pady=5)
    tables_entry = tk.Entry(create_ad_window, width=50)
    tables_entry.pack(pady=5)

    tk.Label(create_ad_window, text="Тема:", font=("Helvetica", 12)).pack(pady=5)
    subject_entry = tk.Entry(create_ad_window, width=50)
    subject_entry.pack(pady=5)

    tk.Label(create_ad_window, text="Содержание:", font=("Helvetica", 12)).pack(pady=5)
    content_text = scrolledtext.ScrolledText(create_ad_window, wrap=tk.WORD, width=50, height=10)
    content_text.pack(pady=5)

    def save_ad_campaign():
        tables = tables_entry.get()
        subject = subject_entry.get()
        content = content_text.get("1.0", tk.END)

        if not tables or not subject or not content:
            messagebox.showerror("Error", "Все поля должны быть заполнены!")
            return

        try:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO ad_campaigns (subject, content, tables) VALUES (%s, %s, %s)",
                (subject, content, tables)
            )
            conn.commit()
            cur.close()
            conn.close()

            messagebox.showinfo("Success", "Рекламная рассылка создана успешно!")

            # Отправка писем
            recipients = fetch_emails_from_db_tables(tables)
            for recipient in recipients:
                send_ad_email(recipient, subject, content)

            create_ad_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Не удалось создать рассылку: {str(e)}")

    save_button = tk.Button(create_ad_window, text="Создать", command=save_ad_campaign, bg="lightgrey",
                            font=("Helvetica", 12), padx=10, pady=10)
    save_button.pack(pady=10)

def show_ad_campaigns():
    clear_right_frame()

    inner_frame = tk.Frame(right_frame, bg="white")
    inner_frame.pack(expand=True, fill="both")

    canvas = tk.Canvas(inner_frame, bg="white", bd=0, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    ad_campaigns_frame = tk.Frame(canvas, bg="white")
    canvas.create_window((0, 0), window=ad_campaigns_frame, anchor="nw")

    scrollbar = tk.Scrollbar(inner_frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas.bind("<Configure>", on_configure)


    # Добавление кнопки "Создать" в правую область
    create_ad_button = tk.Button(right_frame, text="Создать", command=open_create_ad_window, bg="#8EC6C5", fg="black",
                                 font=("Helvetica", 12), relief="flat", padx=10, pady=5, bd=0, highlightthickness=0,
                                 highlightbackground="white")
    create_ad_button.pack(side="top", pady=10, padx=5)



    def load_ad_campaigns():
        try:
            cur = conn.cursor()
            cur.execute("SELECT id, subject, content, tables FROM ad_campaigns")
            rows = cur.fetchall()
            cur.close()
            conn.close()

            for i, row in enumerate(rows):
                ad_label = tk.Label(ad_campaigns_frame, text=f"{row[1]} (таблицы: {row[3]})", bg="white", font=("Helvetica", 12), padx=20, pady=10, anchor="w")
                ad_label.grid(row=i, column=0, sticky="ew")

                if i % 2 == 0:
                    ad_label.config(bg="#f0f0f0")

                ad_label.bind("<Button-1>", lambda event, id=row[0], subject=row[1], content=row[2], tables=row[3]: show_ad_content(id, subject, content, tables))
        except Exception as e:
            messagebox.showerror("Error", f"Не удалось получить рассылки: {str(e)}")

    load_ad_campaigns()

# Функция для отображения содержания рекламной рассылки
def show_ad_content(id, subject, content, tables):
    clear_right_frame()

    inner_frame = tk.Frame(right_frame, bg="white")
    inner_frame.pack(expand=True, fill="both")

    subject_label = tk.Label(inner_frame, text=subject, bg="white", font=("Helvetica", 14, "bold"), padx=20, pady=10, anchor="w")
    subject_label.pack()

    tables_label = tk.Label(inner_frame, text=f"Таблицы: {tables}", bg="white", font=("Helvetica", 12), padx=20, pady=5, anchor="w")
    tables_label.pack()

    content_label = tk.Label(inner_frame, text=content, bg="white", font=("Helvetica", 12), padx=20, pady=10, anchor="w", justify="left", wraplength=500)
    content_label.pack(expand=True, fill="both")


# Функция для получения списка всех таблиц в базе данных
def get_all_tables():
    cursor = conn.cursor()
    cursor.execute("""
        SELECT table_name 
        FROM information_schema.tables 
        WHERE table_schema='public'
    """)
    tables = cursor.fetchall()
    cursor.close()
    return [table[0] for table in tables]


# Функция для проверки наличия нужных столбцов в таблице
def has_required_columns(table, columns):
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name = %s
    """, (table,))
    table_columns = [col[0] for col in cursor.fetchall()]
    cursor.close()
    return all(col in table_columns for col in columns)


# Функция для отображения списка агентов по отделам
def show_agents_by_department():
    clear_right_frame()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT отдел FROM агенты")
    departments = cursor.fetchall()
    cursor.close()

    for department in departments:
        dep_label = tk.Label(right_frame, text=f"Отдел: {department[0]}", font=("Helvetica", 14, "bold"), bg="white")
        dep_label.pack(pady=5)

        cursor = conn.cursor()
        cursor.execute("SELECT id, ФИО FROM агенты WHERE отдел=%s", (department[0],))
        agents = cursor.fetchall()
        cursor.close()

        for agent in agents:
            agent_button = tk.Button(right_frame, text=agent[1],
                                     command=lambda a_id=agent[0]: show_agent_analysis(a_id), bg="#4CAF50", fg="white")
            agent_button.pack(fill=tk.X, padx=10, pady=2)

    btn_show_all_performance = tk.Button(right_frame, text="Успеваемость всех агентов",
                                         command=show_all_agents_performance, bg="white")
    btn_show_all_performance.pack(fill=tk.X, pady=5)
# Функция для анализа агента
def show_agent_analysis(agent_id):
    # Очистка правой панели
    clear_right_frame()
    # Инициализация переменных для общего количества, успешных и неуспешных обработок
    total, successful, unsuccessful = 0, 0, 0
    tables = get_all_tables()
    monthly_data = {}  # Словарь для хранения данных по месяцам
    # Обход всех таблиц
    for table in tables:
        # Проверка наличия необходимых столбцов в таблице
        if has_required_columns(table, ['id_агента', 'обработано', 'дата']):
            cursor = conn.cursor()
            # Выполнение SQL-запроса для получения данных по агенту
            cursor.execute(f"""
                SELECT COUNT(*) AS total, 
                       SUM(CASE WHEN обработано = 'Да' THEN 1 ELSE 0 END) AS successful, 
                       SUM(CASE WHEN обработано = 'Нет' THEN 1 ELSE 0 END) AS unsuccessful,
                       DATE_TRUNC('month', дата) AS month
                FROM {table} 
                WHERE id_агента = %s
                GROUP BY month
                ORDER BY month
            """, (agent_id,))
            results = cursor.fetchall()
            cursor.close()
            # Обработка результатов запроса
            for result in results:
                month = result[3].strftime("%Y-%m")
                if month not in monthly_data:
                    monthly_data[month] = {'successful': 0, 'unsuccessful': 0}
                total += result[0]
                successful += result[1] if result[1] else 0
                unsuccessful += result[2] if result[2] else 0
                monthly_data[month]['successful'] += result[1] if result[1] else 0
                monthly_data[month]['unsuccessful'] += result[2] if result[2] else 0
    # Отображение общей информации
    tk.Label(right_frame, text=f"Всего обработано: {total}", font=("Helvetica", 12), bg="white").pack(pady=5)
    tk.Label(right_frame, text=f"Успешно обработано: {successful}", font=("Helvetica", 12), bg="white").pack(pady=5)
    tk.Label(right_frame, text=f"Не успешно обработано: {unsuccessful}", font=("Helvetica", 12), bg="white").pack(pady=5)
    # Построение графика
    if monthly_data:
        months = list(monthly_data.keys())
        successful_counts = [monthly_data[month]['successful'] for month in months]
        unsuccessful_counts = [monthly_data[month]['unsuccessful'] for month in months]
        fig, ax = plt.subplots()
        bar_width = 0.35
        index = range(len(months))
        ax.bar(index, successful_counts, bar_width, label='Успешно', color='green')
        ax.bar([i + bar_width for i in index], unsuccessful_counts, bar_width, label='Не успешно', color='red')
        ax.set_xlabel('Месяц')
        ax.set_ylabel('Количество')
        ax.set_title('Ежемесячный анализ успешно и неуспешно обработанных')
        ax.set_xticks([i + bar_width / 2 for i in index])
        ax.set_xticklabels(months)
        ax.legend()
        canvas = FigureCanvasTkAgg(fig, master=right_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10)
        
# Функция для анализа отдела
def show_department_analysis(department):
    clear_right_frame()

    total, successful, unsuccessful = 0, 0, 0
    tables = get_all_tables()

    monthly_data = {}  # Dictionary to store monthly data

    for table in tables:
        if has_required_columns(table, ['id_агента', 'обработано', 'дата']):
            cursor = conn.cursor()
            cursor.execute(f"""
                SELECT COUNT(*) AS total, 
                       SUM(CASE WHEN обработано = 'Да' THEN 1 ELSE 0 END) AS successful, 
                       SUM(CASE WHEN обработано = 'Нет' THEN 1 ELSE 0 END) AS unsuccessful,
                       DATE_TRUNC('month', дата) AS month
                FROM {table} 
                WHERE id_агента IN (SELECT id FROM агенты WHERE отдел = %s)
                GROUP BY month
                ORDER BY month
            """, (department,))
            results = cursor.fetchall()
            cursor.close()

            for result in results:
                month = result[3].strftime("%Y-%m")
                if month not in monthly_data:
                    monthly_data[month] = {'successful': 0, 'unsuccessful': 0}
                monthly_data[month]['successful'] += result[1] if result[1] else 0
                monthly_data[month]['unsuccessful'] += result[2] if result[2] else 0
                total += result[0]
                successful += result[1] if result[1] else 0
                unsuccessful += result[2] if result[2] else 0

    tk.Label(right_frame, text=f"Всего обработано: {total}", font=("Helvetica", 12), bg="white").pack(pady=5)
    tk.Label(right_frame, text=f"Успешно обработано: {successful}", font=("Helvetica", 12), bg="white").pack(pady=5)
    tk.Label(right_frame, text=f"Не успешно обработано: {unsuccessful}", font=("Helvetica", 12), bg="white").pack(pady=5)

    # Построение диаграммы
    if monthly_data:
        months = list(monthly_data.keys())
        successful_counts = [monthly_data[month]['successful'] for month in months]
        unsuccessful_counts = [monthly_data[month]['unsuccessful'] for month in months]

        fig, ax = plt.subplots()
        ax.bar(months, successful_counts, label='Успешно', color='green')
        ax.bar(months, unsuccessful_counts, bottom=successful_counts, label='Не успешно', color='red')

        ax.set_xlabel('Месяц')
        ax.set_ylabel('Количество')
        ax.set_title('Ежемесячный анализ успешно и неуспешно обработанных')
        ax.legend()

        canvas = FigureCanvasTkAgg(fig, master=right_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10)


# Функция для отображения отчета по всем агентам
def export_to_excel(data, filename="agents_performance.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Data exported to {filename}")

def show_all_agents_performance():
    clear_right_frame()

    # Create a canvas and a scrollbar
    canvas = tk.Canvas(right_frame, bg="white")
    scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas, style="My.TFrame")

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    export_data = []
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT отдел FROM агенты")
    departments = cursor.fetchall()
    cursor.close()

    grand_total, grand_successful, grand_unsuccessful = 0, 0, 0
    monthly_data = {}

    for department in departments:
        dep_total, dep_successful, dep_unsuccessful = 0, 0, 0
        dep_monthly_data = {}

        dep_frame = tk.Frame(scrollable_frame, bg="lightgrey", bd=2, relief=tk.GROOVE)
        dep_frame.pack(fill=tk.X, padx=10, pady=5)

        dep_label = tk.Label(dep_frame, text=f"Отдел: {department[0]}", font=("Helvetica", 14, "bold"), bg="lightgrey")
        dep_label.pack(pady=5)

        cursor = conn.cursor()
        cursor.execute("SELECT id, ФИО FROM агенты WHERE отдел=%s", (department[0],))
        agents = cursor.fetchall()
        cursor.close()

        for agent in agents:
            total, successful, unsuccessful = 0, 0, 0
            tables = get_all_tables()

            for table in tables:
                if has_required_columns(table, ['id_агента', 'обработано', 'дата']):
                    cursor = conn.cursor()
                    cursor.execute(f"""
                        SELECT COUNT(*) AS total, 
                               SUM(CASE WHEN обработано = 'Да' THEN 1 ELSE 0 END) AS successful, 
                               SUM(CASE WHEN обработано = 'Нет' THEN 1 ELSE 0 END) AS unsuccessful,
                               DATE_TRUNC('month', дата) AS month
                        FROM {table} 
                        WHERE id_агента = %s
                        GROUP BY month
                        ORDER BY month
                    """, (agent[0],))
                    results = cursor.fetchall()
                    cursor.close()

                    for result in results:
                        month = result[3].strftime("%Y-%m")
                        if month not in dep_monthly_data:
                            dep_monthly_data[month] = {'successful': 0, 'unsuccessful': 0}
                        dep_monthly_data[month]['successful'] += result[1] if result[1] else 0
                        dep_monthly_data[month]['unsuccessful'] += result[2] if result[2] else 0
                        total += result[0]
                        successful += result[1] if result[1] else 0
                        unsuccessful += result[2] if result[2] else 0

            dep_total += total
            dep_successful += successful
            dep_unsuccessful += unsuccessful

            agent_label = tk.Label(dep_frame,
                                   text=f"{agent[1]}: Всего: {total}, Успешно: {successful}, Не успешно: {unsuccessful}",
                                   font=("Helvetica", 12), bg="white")
            agent_label.pack(pady=2, padx=10)
            export_data.append({
                "Отдел": department[0],
                "Агент": agent[1],
                "Всего": total,
                "Успешно": successful,
                "Не успешно": unsuccessful,
            })

        dep_summary_label = tk.Label(dep_frame,
                                     text=f"Всего: {dep_total}, Успешно: {dep_successful}, Не успешно: {dep_unsuccessful}",
                                     font=("Helvetica", 12, "bold"), bg="lightgrey")
        dep_summary_label.pack(pady=5)

        grand_total += dep_total
        grand_successful += dep_successful
        grand_unsuccessful += dep_unsuccessful
        for month, counts in dep_monthly_data.items():
            if month not in monthly_data:
                monthly_data[month] = {'successful': 0, 'unsuccessful': 0}
            monthly_data[month]['successful'] += counts['successful']
            monthly_data[month]['unsuccessful'] += counts['unsuccessful']

    overall_frame = tk.Frame(scrollable_frame, bg="lightgrey", bd=2, relief=tk.GROOVE)
    overall_frame.pack(fill=tk.X, padx=10, pady=10)
    overall_label = tk.Label(overall_frame,
                             text=f"Общее: Всего: {grand_total}, Успешно: {grand_successful}, Не успешно: {grand_unsuccessful}",
                             font=("Helvetica", 14, "bold"), bg="lightgrey")
    overall_label.pack(pady=10)

    # Построение диаграммы
    if monthly_data:
        months = list(monthly_data.keys())
        successful_counts = [monthly_data[month]['successful'] for month in months]
        unsuccessful_counts = [monthly_data[month]['unsuccessful'] for month in months]

        fig, ax = plt.subplots()
        bar_width = 0.35
        index = range(len(months))

        ax.bar(index, successful_counts, bar_width, label='Успешно', color='green')
        ax.bar([i + bar_width for i in index], unsuccessful_counts, bar_width, label='Не успешно', color='red')

        ax.set_xlabel('Месяц')
        ax.set_ylabel('Количество')
        ax.set_title('Ежемесячный анализ успешно и неуспешно обработанных')
        ax.set_xticks([i + bar_width / 2 for i in index])
        ax.set_xticklabels(months)
        ax.legend()

        canvas_plot = FigureCanvasTkAgg(fig, master=scrollable_frame)
        canvas_plot.draw()
        canvas_plot.get_tk_widget().pack(pady=10)
    export_button = tk.Button(scrollable_frame, text="Выгрузить в Excel", command=lambda: export_to_excel(export_data))
    export_button.pack(pady=10)


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

# Инициализация главного окна
root = tk.Tk()
root.title("Интерфейс руководителя")

root.geometry("1550x1100")
root.configure(bg="white")

# Левый фрейм для кнопок
left_frame = tk.Frame(root, bg="#2C3E50", width=200)
left_frame.pack(side=tk.LEFT, fill=tk.Y)

# Правая область для отображения содержимого
right_frame = tk.Frame(root, bg="white")
right_frame.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH)

# Кнопка для отображения списка агентов
btn_agents = tk.Button(left_frame, text="Агенты", command=show_agents, bg="#34495E", fg="white", font=("Helvetica", 12), bd=0, relief=tk.FLAT)
btn_agents.pack(fill=tk.X, pady=5)

# Кнопка для отображения списка таблиц базы данных
btn_db_tables = tk.Button(left_frame, text="Таблицы", command=show_db_tables, bg="#34495E", fg="white", font=("Helvetica", 12), bd=0, relief=tk.FLAT)
btn_db_tables.pack(fill=tk.X, pady=5)
# Добавление кнопки Анализ в левую область
btn_ad_campaigns = tk.Button(left_frame, text="Отчет", command=show_agents_by_department, bg="#34495E", fg="white", font=("Helvetica", 12), bd=0, relief=tk.FLAT)
btn_ad_campaigns.pack(fill=tk.X, pady=5)
# Добавление кнопки Анализ в левую область
btn_analysis = tk.Button(left_frame, text="Почта", command=show_emails, bg="#34495E", fg="white", font=("Helvetica", 12), bd=0, relief=tk.FLAT)
btn_analysis.pack(fill=tk.X, pady=5)
btn_analysis = tk.Button(left_frame, text="Оповещения", command=show_tasks_notifications, bg="#34495E", fg="white", font=("Helvetica", 12), bd=0, relief=tk.FLAT)
btn_analysis.pack(fill=tk.X, pady=5)
# Добавление кнопки "Реклама" в левую область
btn_ad_campaigns = tk.Button(left_frame, text="Реклама", command=show_ad_campaigns, bg="#34495E", fg="white", font=("Helvetica", 12), bd=0, relief=tk.FLAT)
btn_ad_campaigns.pack(fill=tk.X, pady=5)
# Добавляем кнопку выхода из аккаунта
logout_button = tk.Button(left_frame, text="Выход", bg="#d9534f", fg="white", font=("Helvetica", 12, "bold"), bd=0, relief=tk.FLAT, command=logout)
logout_button.pack(fill="x", pady=2, side="bottom")
# Основной цикл
root.mainloop()

