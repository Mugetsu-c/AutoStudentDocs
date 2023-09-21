import tkinter as tk
import os
from tkinter import ttk
from utils import download_csv, read_filtered_data, create_certificate_jinja  # Импорт новой функции
from tkcalendar import DateEntry
from datetime import datetime

def create_export_folder():
    # Находим путь к рабочему столу
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

    # Название главной папки
    main_folder_name = 'AutoStudentDocs'

    # Полный путь к главной папке
    main_folder_path = os.path.join(desktop, main_folder_name)

    # Проверяем, существует ли уже такая папка. Если нет, создаем ее.
    if not os.path.exists(main_folder_path):
        os.makedirs(main_folder_path)

    # Создаем имя для вложенной папки с текущей датой и временем
    nested_folder_name = datetime.now().strftime("%d.%m.%Y %H-%M-%S")

    # Полный путь к вложенной папке
    nested_folder_path = os.path.join(main_folder_path, nested_folder_name)

    # Создаем вложенную папку
    os.makedirs(nested_folder_path)

    print(f"Main folder is located at: {main_folder_path}")
    print(f"Nested folder is located at: {nested_folder_path}")

    for i, record_str in enumerate(user_data):
        if checkbutton_vars[i].get():  # Проверяем, выбран ли студент
            record_list = record_str.split(", ")
            date_obj = datetime.strptime(record_list[0], '%d.%m.%Y')
            record = {
                "d": date_obj.day,
                "m": date_obj.month,
                "y": date_obj.year,
                "full_name": record_list[1],
                "group": record_list[2],
                "well": record_list[3],
                "destination": record_list[4],
                "quantity": int(record_list[5])  # Преобразование в int
            }
            
            # Выбираем шаблон справки
            if record['destination'] == 'в пенсионный фонд':
                template_name = 'в пенсионный фонд'
            else:
                template_name = 'по месту требования'
            
            date = f"{record['d']}.{record['m']}.{record['y']}"
            for i in range(record['quantity']):
                save_path = f"{nested_folder_path}/{record['full_name']} {date} {record['destination']} {i+1}.docx"
                
                # Создаем справку
                create_certificate_jinja(template_name, save_path, record)

    return nested_folder_path



checkbuttons = []
user_data = []
checkbutton_vars = []


def update_data():
    try:
        status_label.config(text="Обновление данных...")
        csv_url = 'https://docs.google.com/spreadsheets/d/1QzSCAb9mV7e_-OTzDwoWZuoWzC4V_KkAnS6HMQXPgQo/export?format=csv'
        csv_save_path = 'data/downloaded_data.csv'
        download_csv(csv_url, csv_save_path)
        status_label.config(text="Данные обновлены.")
    except Exception as e:
        print(f"An error occurred: {e}")


def load_data():
    global checkbuttons, user_data, checkbutton_vars  # Добавьте checkbutton_vars здесь
    for cb in checkbuttons:  # Удаление старых Checkbuttons
        cb.destroy()
    
    checkbuttons.clear()
    checkbutton_vars.clear()  # Очистите список переменных
    time_period = time_var.get()
    if time_period == "Произвольная дата":
        start_date = date_entry_from.get_date()
        end_date = date_entry_to.get_date()
    else:
        start_date, end_date = None, None
    
    csv_save_path = 'data/downloaded_data.csv'
    user_data = read_filtered_data(csv_save_path, time_period, start_date, end_date)
    
    for user in user_data:
        var = tk.BooleanVar(value=True)  # По умолчанию установлено значение True
        c = tk.Checkbutton(root, text=user, variable=var)
        c.pack()
        checkbuttons.append(c)  # Добавляем в список
        checkbutton_vars.append(var)  # Добавьте переменную в список
    print(user_data)


def toggle_date_entries(event):
    if time_var.get() == "Произвольная дата":
        date_entry_from.pack()
        date_entry_to.pack()
    else:
        date_entry_from.pack_forget()
        date_entry_to.pack_forget()

root = tk.Tk()
root.title('AutoStudentDocs')

user_data = []

update_button = tk.Button(root, text="Обновить данные", command=update_data)
update_button.pack()

load_button = tk.Button(root, text="Загрузить данные", command=load_data)
load_button.pack()

status_label = tk.Label(root, text="")
status_label.pack()

time_label = ttk.Label(root, text="Выберите временной период:")
time_label.pack()

time_options = ["Последние 24 часа", "Последние 48 часов", "Последние 72 часа", "Произвольная дата"]
time_var = tk.StringVar()
time_var.set(time_options[0])

time_combobox = ttk.Combobox(root, textvariable=time_var, values=time_options)
time_combobox.bind('<<ComboboxSelected>>', toggle_date_entries)
time_combobox.pack()

date_entry_from = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
date_entry_to = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)

export_button = tk.Button(root, text="Выгрузить справки", command=create_export_folder)
export_button.pack()

root.mainloop()
