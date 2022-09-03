import tkinter as tk
import openpyxl
from barcode import Code128
from barcode.writer import SVGWriter
from barcode.writer import ImageWriter
from tkinter import *
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD

version = '0.1'
file_types = ('xlsx', 'xlsm', 'xltx', 'xltm')


class ChecklistProgram:
    def __init__(self):
        # Initialize main program
        self.root = TkinterDnD.Tk()
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.open_file)

        # Root window properties
        self.root.geometry('420x420')
        self.root.title('Checklist Generator')
        self.root.resizable(False, False)

        # Variables
        self.workbook = None
        self.sheet = None

        # Button
        self.open_file_button = tk.Button(self.root, text='Browse...', width=10, command=self.open_file, takefocus=0)
        self.open_file_button.pack()

        self.test_button = tk.Button(self.root, text='Test', width=10, command=self.test_barcode, takefocus=0)
        self.test_button.pack()

        # Version Label
        self.version_label = Label(self.root, text='ver ' + version)
        self.version_label.place(relx=1, rely=1.01, anchor='se')

        # Main Loop
        self.root.mainloop()

    def open_file(self, event=None):
        if event is None:
            file_path = filedialog.askopenfilename(filetypes=[('Excel File', file_types)])
            if not file_path:
                return
            self.process_file(file_path)
            return
        file_path = event.data.replace('{', '').replace('}', '')
        valid_file = False
        for file_type in file_types:
            if '.' + file_type in file_path:
                valid_file = True
                break
        if not valid_file:
            return
        self.process_file(file_path)

    def process_file(self, file_path: str):
        self.workbook = openpyxl.load_workbook(file_path)
        self.sheet = self.workbook.active
        for r in range(1, self.sheet.max_row + 1):
            for c in range(1, self.sheet.max_column + 1):
                cell_obj = self.sheet.cell(row=r, column=c)
                print(cell_obj.value)

    # noinspection PyMethodMayBeStatic
    def test_barcode(self):
        with open('test.png', "wb") as file:
            Code128('This is a test', writer=ImageWriter()).write(file, options={'write_text': False})
        with open('test.svg', "wb") as file:
            Code128('This is a test', writer=SVGWriter()).write(file, options={'write_text': False})


ChecklistProgram()
