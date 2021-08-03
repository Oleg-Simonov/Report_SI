import sys
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import messagebox
import main
import os


def foo():
    text = fd.askopenfilename()
    l_file_path.config(text=text)

    try:
        with open('path.txt', 'w') as fw:
            fw.write(text)
    except FileNotFoundError:
        messagebox.showerror(title='Error', message='File write error!')
        sys.exit()

    return text


def calc(file_path_):
    status = main.main(file_path_, days_reserve.get())
    if status == 'Report has been formed':
        messagebox.showinfo(title='Info message', message=status + '\nAddress: ' + os.getcwd() + '\\reports_SI')
    elif status == 'Everything is serviced!':
        messagebox.showinfo(title='Info message', message=status)
    elif type(status) == str:
        messagebox.showerror(title='Error', message=status)
    else:
        messagebox.showerror(title='Error', message='Unexpected error')


if __name__ == '__main__':

    try:
        with open('path.txt') as f:
            file_path = f.read()
    except FileNotFoundError:
        messagebox.showerror(title='Error', message='File error!')
        sys.exit()

    window = tk.Tk()
    window.title('reportSI')
    window.resizable(False, False)

    l_chosen_path = tk.Label(window, text='File:')
    l_file_path = tk.Label(window, text=file_path)
    l_day_reserve = tk.Label(window, text='Reserve, days:')

    days_reserve = tk.IntVar(window, value=15)
    spb_days = tk.Spinbox(window, from_=5, to=50, textvariable=days_reserve, state='readonly')

    btn_open = tk.Button(window, text='Click to open file', command=lambda: file_path == foo())
    btn_calc = tk.Button(window, text='Calculate', command=lambda: calc(l_file_path.cget('text')))

    l_chosen_path.grid(row=0, column=0, padx=15, pady=15, sticky='W')
    l_file_path.grid(row=0, column=1, padx=15, pady=15, sticky='W')
    l_day_reserve.grid(row=1, column=0, padx=15, pady=15, sticky='W')

    spb_days.grid(row=1, column=1, padx=15, pady=15, sticky='E')

    btn_open.grid(row=2, column=0, padx=15, pady=15, sticky='E')
    btn_calc.grid(row=2, column=1, padx=15, pady=15, sticky='E')

    window.mainloop()
