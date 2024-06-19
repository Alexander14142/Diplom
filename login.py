import hashlib
import tkinter as tk
from tkinter import messagebox
import subprocess
import psycopg2

# Глобальная переменная для хранения id авторизованного пользователя
current_user_id = None


def login():
    global current_user_id
    username = entry_username.get()
    password = entry_password.get()

    # Хеширование введенного пароля
    hashed_password = hashlib.sha256(password.encode()).hexdigest()
    print(hashed_password)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM агенты WHERE login = %s AND password = %s", (username, hashed_password))
    agent = cursor.fetchone()

    cursor.execute("SELECT id FROM руководители WHERE login = %s AND password = %s", (username, hashed_password))
    manager = cursor.fetchone()
    cursor.close()

    if agent:
        current_user_id = agent[0]
        messagebox.showinfo("Login Info", "Welcome Agent!")
        root.destroy()  # Закрываем окно входа
        subprocess.Popen(
            ["python", "agent_interface.py", str(current_user_id)])  # Запускаем agent_interface.py с id агента
    elif manager:
        current_user_id = manager[0]
        messagebox.showinfo("Login Info", "Welcome Manager!")
        root.destroy()  # Закрываем окно входа
        subprocess.Popen(["python", "admin.py", str(current_user_id)])  # Запускаем admin.py с id руководителя
    else:
        messagebox.showerror("Login Error", "Invalid username or password")


# Создаем основное окно
root = tk.Tk()
root.title("Авторизация")
root.geometry("400x450")
root.configure(bg="#f0f0f0")

# Создаем и размещаем виджеты
frame = tk.Frame(root, bg="white", padx=20, pady=20)
frame.pack(pady=40)

label_title = tk.Label(frame, text="Авторизация", font=("Helvetica", 18, "bold"), bg="white")
label_title.pack(pady=10)

label_username = tk.Label(frame, text="Логин", font=("Helvetica", 12), bg="white")
label_username.pack(pady=5)

entry_username = tk.Entry(frame, font=("Helvetica", 12), width=30)
entry_username.pack(pady=5)

label_password = tk.Label(frame, text="Пароль", font=("Helvetica", 12), bg="white")
label_password.pack(pady=5)

entry_password = tk.Entry(frame, show="*", font=("Helvetica", 12), width=30)
entry_password.pack(pady=5)

button_login = tk.Button(frame, text="Войти", font=("Helvetica", 12, "bold"), bg="#4caf50", fg="white", command=login)
button_login.pack(pady=20)

# Запускаем главное событие
root.mainloop()
