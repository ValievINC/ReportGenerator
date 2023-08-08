import threading
import tkinter as tk
from tkinter import filedialog, Checkbutton, BooleanVar
import tkinter.ttk
from datetime import date, datetime, timedelta
from tkcalendar import Calendar

from CreateReport import create_report


def select_report_xlsx():
    global report_xlsx_file
    global report_xlsx_file_path
    report_xlsx_file_path = filedialog.askopenfilename(filetypes=[('XLSX files', '*.xlsx')])
    report_xlsx_file = report_xlsx_file_path.split("/")[-1]
    label_report_xlsx.config(text=report_xlsx_file)
    if report_xlsx_file_path and staff_xlsx_file_path:
        run_button.config(state='active')
    elif staff_xlsx_file_path is None:
        label_report_xlsx.config(text="Выберите файл с сотрудниками")


def on_date_selected(event):
    global selected_date
    selected_date_str = cal.get_date()
    selected_date = datetime.strptime(selected_date_str, "%m/%d/%y").date()
    today = date.today()
    if selected_date <= today:
        checkbox.config(state='active')
        label_cal.config(text=str(selected_date))
        if report_xlsx_file_path and staff_xlsx_file_path:
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
        if report_xlsx_file_path and staff_xlsx_file_path:
            run_button.config(state='active')
    else:
        run_button.config(state='disabled')
        label_cal.config(text="Выбрана недопустимая дата.")


def run_report_creation():
    if selected_date_range is None:
        run_button.config(state='disabled')
        t1 = threading.Thread(target=create_report(report_xlsx_file_path, selected_date, staff_xlsx_file_path))
        t1.start()
        t1.join()
        run_button.config(state='active')
    else:
        current_date = selected_date
        while current_date <= selected_date_range:
            run_button.config(state='disabled')
            t1 = threading.Thread(target=create_report(report_xlsx_file_path, current_date, staff_xlsx_file_path))
            t1.start()
            t1.join()
            run_button.config(state='active')
            current_date += timedelta(days=1)
        print('Цикл завершен')


def select_staff_xlsx():
    global staff_xlsx_file_path
    staff_xlsx_file_path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
    label_report_xlsx.config(text=report_xlsx_file)
    if report_xlsx_file_path and staff_xlsx_file_path:
        run_button.config(state='active')



def toggle_calendar():
    global selected_date_range
    if show_calendar.get():
        window.geometry("500x900")
        run_button.config(state='disabled')
        range_cal.pack()
    else:
        window.geometry("500x700")
        selected_date_range = None
        label_cal.config(text=str(selected_date))
        range_cal.pack_forget()
        if report_xlsx_file:
            run_button.config(state='active')


window = tk.Tk()
window.title("Генератор Отчёта")
window.geometry("500x700")
window.resizable(False, False)

selected_date = None
selected_date_range = None
report_xlsx_file = None
report_xlsx_file_path = None
staff_xlsx_file_path = None

# Label
label = tk.Label(window,
                 text='Для работы программы необходимо выбрать дату или диапазон дат и excel-файл, куда загружать отчёт, а также файл с сотрудниками.',
                 font=("Arial", 14),
                 wraplength=450)
label.pack(pady=10)

cal_frame = tk.Frame(window)
cal_frame.pack()

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
                              text="Выбрать файл с сотрудниками",
                              state='active',
                              command=select_staff_xlsx,
                              width=16,
                              height=3,
                              wraplength=120)
button_staff_xlsx.pack(side=tk.TOP,
                       pady=10)

# Run button
run_button = tk.Button(window,
                       text="Создать отчёт",
                       state='disabled',
                       font=("Arial", 16),
                       command=run_report_creation)
run_button.pack(pady=20)

window.mainloop()