import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook


def upload_command():
    filename = filedialog.askopenfilename()
    global ws
    if filename.endswith('.xlsx'):
        ws = load_workbook(filename).active
    else:
        messagebox.showerror("Error", "Please select an .xlsx file.")


def generate_command():
    global ws
    if ws:
        print(ws['D6'].value)


root = tk.Tk()
ws = None
upload_button = tk.Button(root, text='Open', command=upload_command).pack()
generate_button = tk.Button(root, text='Generate', command=generate_command).pack()

root.mainloop()
