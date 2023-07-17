# encoding: windows-1251

import threading
import tkinter as tk
from tkinter import filedialog
import tkinter.ttk
import pandas as pd
from datetime import date, datetime
from tkcalendar import Calendar

from CreateReport import create_report, prebuild


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
        label_report_xlsx.config(text="����� ��������� ����������� � ������� � �������.")
        run_button.config(state='disabled')
        button_staff_xlsx.config(state='active')
    else:
        staff_exists = True
        label_report_xlsx.config(text=report_xlsx_file)
        if selected_date and report_xlsx_file and staff_exists:
            run_button.config(state='active')
            button_staff_xlsx.config(state='disabled')


def select_staff_xlsx():
    staff_xlsx_file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
    prebuild(report_xlsx_file_path, staff_xlsx_file_path)
    label_report_xlsx.config(text=report_xlsx_file)
    run_button.config(state='active')
    button_staff_xlsx.config(state='disabled')


def on_date_selected(event):
    global selected_date
    selected_date_str = cal.get_date()
    selected_date = datetime.strptime(selected_date_str, "%m/%d/%y").date()
    today = date.today()
    if selected_date <= today:
        label_cal.config(text=str(selected_date))
    else:
        run_button.config(state='disabled')
        label_cal.config(text="������� ������������ ����.")


def run_report_creation():
    run_button.config(state='disabled')
    t1 = threading.Thread(target=create_report(report_xlsx_file_path, selected_date))
    t1.start()
    run_button.config(state='active')


window = tk.Tk()
window.title("��������� ������")
window.geometry("500x700")
window.resizable(False, False)

label = tk.Label(window,
                 text='��� ������ ��������� ������ ��� �����. ���� � ����������� .csv (� �������� ����� ���������� ������� ���� � ������� DD.MM.YYYY). ������ � ����������� .xlsx(����� ����� ����������� �������)',
                 font=("Arial", 14),
                 wraplength=450)
label.pack(pady=10)

frame = tk.Frame(window)
frame.pack()

staff_exists = False
selected_date = ""
report_xlsx_file = ""
report_xlsx_file_path = ""

label_cal = tk.Label(frame,
                     text="�������� ����",
                     font=("Arial", 14))
label_cal.pack(side=tk.TOP)
cal = Calendar(frame,
               selectmode="day",
               year=date.today().year,
               month=date.today().month,
               day=date.today().day)
cal.pack(side=tk.TOP,
         pady=20)
cal.bind("<<CalendarSelected>>", on_date_selected)

separator = tkinter.ttk.Separator(frame,
                                  orient='horizontal')
separator.pack(fill='x',
               pady=10)

label_report_xlsx = tk.Label(frame,
                             text="���� �� ������",
                             font=("Arial", 14))
label_report_xlsx.pack(side=tk.TOP)
button_report_xlsx = tk.Button(frame,
                               text="������� ���� ��� �������� ������",
                               command=select_report_xlsx,
                               width=16,
                               height=3,
                               wraplength=120)
button_report_xlsx.pack(side=tk.TOP,
                        pady=10)

separator = tkinter.ttk.Separator(frame,
                                  orient='horizontal')
separator.pack(fill='x',
               pady=10)

button_staff_xlsx = tk.Button(frame,
                              text="������� ���� ��� �������� �����������",
                              state='disabled',
                              command=select_staff_xlsx,
                              width=16,
                              height=3,
                              wraplength=120)
button_staff_xlsx.pack(side=tk.TOP,
                       pady=10)

run_button = tk.Button(window,
                       text="������� �����",
                       state='disabled',
                       font=("Arial", 14),
                       command=run_report_creation)
run_button.pack(pady=20)

window.mainloop()