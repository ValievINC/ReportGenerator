# encoding: windows-1251

import threading
import tkinter as tk
from tkinter import filedialog, Checkbutton, BooleanVar
import tkinter.ttk
import pandas as pd
from datetime import date, datetime, timedelta
from tkcalendar import Calendar

from CreateReport import create_report, prebuild, reload_report


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
    else:
        staff_exists = True
        label_report_xlsx.config(text=report_xlsx_file)
        if report_xlsx_file and staff_exists:
            reload_button.config(state='active')
            if selected_date:
                run_button.config(state='active')
    button_staff_xlsx.config(state='active')


def select_staff_xlsx():
    staff_xlsx_file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
    prebuild(report_xlsx_file_path, staff_xlsx_file_path)
    label_report_xlsx.config(text=report_xlsx_file)
    run_button.config(state='active')
    reload_button.config(state='active')


def on_date_selected(event):
    global selected_date
    selected_date_str = cal.get_date()
    selected_date = datetime.strptime(selected_date_str, "%m/%d/%y").date()
    today = date.today()
    if selected_date <= today:
        checkbox.config(state='active')
        label_cal.config(text=str(selected_date))
        if report_xlsx_file and staff_exists:
            run_button.config(state='active')
    else:
        checkbox.config(state='disabled')
        run_button.config(state='disabled')
        label_cal.config(text="Выбрана недопустимая дата.")


def on_range_selected(event):
    global selected_date_range
    selected_date_range_str = range_cal.get_date()
    selected_date_range = datetime.strptime(selected_date_range_str, "%m/%d/%y").date()
    today = date.today()
    if today >= selected_date_range > selected_date:
        label_cal.config(text=f'{str(selected_date)} - {str(selected_date_range)}')
        if report_xlsx_file and staff_exists:
            run_button.config(state='active')
    else:
        run_button.config(state='disabled')
        label_cal.config(text="Выбрана недопустимая дата.")


def run_report_creation():
    if selected_date_range is None:
        run_button.config(state='disabled')
        t1 = threading.Thread(target=create_report(report_xlsx_file_path, selected_date))
        t1.start()
        t1.join()
        run_button.config(state='active')
    else:
        current_date = selected_date
        while current_date <= selected_date_range:
            run_button.config(state='disabled')
            t1 = threading.Thread(target=create_report(report_xlsx_file_path, current_date))
            t1.start()
            t1.join()
            run_button.config(state='active')
            current_date += timedelta(days=1)
        print('Цикл завершен')


def toggle_calendar():
    global selected_date_range
    if show_calendar.get():
        window.geometry("500x1050")
        run_button.config(state='disabled')
        range_cal.pack()
    else:
        window.geometry("500x850")
        selected_date_range = None
        label_cal.config(text=str(selected_date))
        range_cal.pack_forget()
        if report_xlsx_file and staff_exists:
            run_button.config(state='active')


window = tk.Tk()
window.title("Генератор Отчёта")
window.geometry("500x850")
window.resizable(False, False)

label = tk.Label(window,
                 text='Для работы программы нужно выбрать дату, по которой вы хотите сделать отчёт, и Excel-файл с расширением .xlsx, куда будет выгружаться отчёт. Если Excel-файл пустой или в отделе появились новые сотрудники, то необходимо внести базу сотрудников (Excel-документ)',
                 font=("Arial", 14),
                 wraplength=450)
label.pack(pady=10)

cal_frame = tk.Frame(window)
cal_frame.pack()

staff_exists = False
selected_date = None
selected_date_range = None
report_xlsx_file = ""
report_xlsx_file_path = ""

# Calendar
label_cal = tk.Label(cal_frame,
                     text="Выберите дату",
                     font=("Arial", 14))
label_cal.pack(side=tk.TOP)
cal = Calendar(cal_frame,
               selectmode="day",
               year=date.today().year,
               month=date.today().month,
               day=date.today().day)
cal.pack(side=tk.TOP,
         pady=10)
cal.bind("<<CalendarSelected>>", on_date_selected)

# Checkbox
show_calendar = BooleanVar()
checkbox = Checkbutton(cal_frame, state="disabled", text="Выбрать конечную дату", variable=show_calendar, command=toggle_calendar)
checkbox.pack(side=tk.TOP,
              pady=20)
range_cal = Calendar(cal_frame,
                     selectmode="day",
                     year=date.today().year,
                     month=date.today().month,
                     day=date.today().day)
range_cal.bind("<<CalendarSelected>>", on_range_selected)

# Separator
buttons_frame = tk.Frame(window)
buttons_frame.pack()


separator = tkinter.ttk.Separator(buttons_frame,
                                  orient='horizontal')
separator.pack(fill='x',
               pady=10)

# Excel Doc
label_report_xlsx = tk.Label(buttons_frame,
                             text="Файл не выбран",
                             font=("Arial", 14))
label_report_xlsx.pack(side=tk.TOP)
button_report_xlsx = tk.Button(buttons_frame,
                               text="Выбрать файл для выгрузки отчёта",
                               command=select_report_xlsx,
                               width=16,
                               height=3,
                               wraplength=120)
button_report_xlsx.pack(side=tk.TOP,
                        pady=10)

# Separator
separator = tkinter.ttk.Separator(buttons_frame,
                                  orient='horizontal')
separator.pack(fill='x',
               pady=10)

# Staff Prebuild
button_staff_xlsx = tk.Button(buttons_frame,
                              text="Выбрать файл для загрузки сотрудников",
                              state='disabled',
                              command=select_staff_xlsx,
                              width=16,
                              height=3,
                              wraplength=120)
button_staff_xlsx.pack(side=tk.TOP,
                       pady=10)

# Separator
separator = tkinter.ttk.Separator(buttons_frame,
                                  orient='horizontal')
separator.pack(fill='x',
               pady=10)

# Reload button
reload_button = tk.Button(window,
                       text="Обновить сводный лист",
                       state='disabled',
                       font=("Arial", 12),
                       command=lambda: reload_report(report_xlsx_file_path))
reload_button.pack(pady=10)

# Run button
run_button = tk.Button(window,
                       text="Создать отчёт",
                       state='disabled',
                       font=("Arial", 16),
                       command=run_report_creation)
run_button.pack(pady=20)

window.mainloop()
