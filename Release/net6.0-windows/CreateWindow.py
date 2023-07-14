# encoding: windows-1251

import tkinter as tk
from tkinter import filedialog
import tkinter.ttk
import pandas as pd
from CreateReport import create_report, prebuild


def select_csv():
    global csv_file
    global label_csv
    global csv_file_path
    global staff_exists
    csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_file_path:
        csv_file = csv_file_path.split("/")[-1]
        label_csv.config(text=csv_file)
    if csv_file and report_xlsx_file and staff_exists:
        run_button.config(state='active')


def check_employees_nonexistence(file) -> bool:
    df = pd.read_excel(file)
    return df.empty


def select_report_xlsx():
    global report_xlsx_file
    global label_report_xlsx
    global report_xlsx_file_path
    global staff_exists
    report_xlsx_file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
    if report_xlsx_file_path:
        report_xlsx_file = report_xlsx_file_path.split("/")[-1]
    if check_employees_nonexistence(report_xlsx_file_path):
        staff_exists = False
        label_report_xlsx.config(text="Нужно загрузить сотрудников в таблицу с отчётом.")
        run_button.config(state='disabled')
        button_staff_xlsx.config(state='active')
    else:
        staff_exists = True
        label_report_xlsx.config(text=report_xlsx_file)
        if csv_file and report_xlsx_file and staff_exists:
            run_button.config(state='active')
            button_staff_xlsx.config(state='disabled')


def select_staff_xlsx():
    staff_xlsx_file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
    prebuild(report_xlsx_file_path, staff_xlsx_file_path)
    label_report_xlsx.config(text=report_xlsx_file)
    run_button.config(state='active')
    button_staff_xlsx.config(state='disabled')


window = tk.Tk()
window.title("Генератор Отчёта")
window.geometry("500x550")
window.resizable(False, False)

label = tk.Label(window, text='Для работы программы нужены два файла. Один с расширением .csv (В названии файла необходимо указать дату в формате DD.MM.YYYY). Другой с расширением .xlsx(Здесь будут создаваться таблицы)', font=("Arial", 14), wraplength=450)
label.pack(pady=10)

frame = tk.Frame(window)
frame.pack()

staff_exists = False
csv_file = ""
csv_file_path = ""
report_xlsx_file = ""
report_xlsx_file_path = ""

label_csv = tk.Label(frame, text="Файл не выбран", font=("Arial", 14))
label_csv.pack(side=tk.TOP)
button_csv = tk.Button(frame, text="Выбрать таблицу для обработки", command=select_csv, width=16, height=3, wraplength=120)
button_csv.pack(side=tk.TOP, pady=10)

separator = tkinter.ttk.Separator(frame, orient='horizontal')
separator.pack(fill='x', pady=10)

label_report_xlsx = tk.Label(frame, text="Файл не выбран", font=("Arial", 14))
label_report_xlsx.pack(side=tk.TOP)
button_report_xlsx = tk.Button(frame, text="Выбрать файл для выгрузки отчёта", command=select_report_xlsx, width=16, height=3, wraplength=120)
button_report_xlsx.pack(side=tk.TOP, pady=10)

separator = tkinter.ttk.Separator(frame, orient='horizontal')
separator.pack(fill='x', pady=10)

button_staff_xlsx = tk.Button(frame, text="Выбрать файл для загрузки сотрудников", state='disabled', command=select_staff_xlsx, width=16, height=3, wraplength=120)
button_staff_xlsx.pack(side=tk.TOP, pady=10)

run_button = tk.Button(window, text="Создать отчёт", state='disabled', font=("Arial", 14), command=lambda: create_report(report_xlsx_file_path, csv_file_path))
run_button.pack(pady=20)

window.mainloop()