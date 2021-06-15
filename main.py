import time
import tkinter as tk
from tkinter import filedialog, messagebox

import genanki
from openpyxl import load_workbook


def upload():
    filename = filedialog.askopenfilename()
    global ws
    if filename.endswith('.xlsx'):
        ws = load_workbook(filename).active
    else:
        messagebox.showerror('Error', 'Please select an .xlsx file.')


def generate():
    global ws
    deck = genanki.Deck(int(time.time()), 'Test deck')
    if ws:
        deck.add_note(genanki.Note(model=anki_model, fields=[ws['D6'].value, ws['E6'].value]))
        genanki.Package(deck).write_to_file('output.apkg')
        messagebox.showinfo('Success', 'Successfully generated file')


root = tk.Tk()
ws = None
anki_model = genanki.Model(
    int(time.time()),
    'Basic Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Answer'}],
    templates=[{
        'name': 'Card',
        'qfmt': '{{Question}}',
        'afmt': '{{FrontSide}}<hr id="answer">{{Answer}}'}])
upload_button = tk.Button(root, text='Open', command=upload).pack()
generate_button = tk.Button(root, text='Generate', command=generate).pack()

root.mainloop()
