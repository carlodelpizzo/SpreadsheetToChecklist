import os.path
import shutil
import tkinter as tk
import openpyxl
import openpyxl.styles
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
from openpyxl.drawing.xdr import XDRPositiveSize2D
from barcode import Code128  # Package called python-barcode
from barcode.writer import ImageWriter  # Requires Pillow
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
        self.cur_dir = os.getcwd() + '/'
        self.batch = None
        self.parts = {}
        self.parts_list = []

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
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        part_col = None
        type_col = None
        material_col = None
        break_col = None
        weld_col = None
        powder_col = None
        quantity_col = None
        batch_col = None
        read_col_labels = False
        all_col_read = False
        for r in range(1, sheet.max_row + 1):
            current_part = None
            for c in range(1, sheet.max_column + 1):
                cell_obj = sheet.cell(row=r, column=c)
                if read_col_labels:
                    if cell_obj.value == 'Part':
                        part_col = c
                    elif cell_obj.value == 'Type':
                        type_col = c
                    elif cell_obj.value == 'Material':
                        material_col = c
                    elif cell_obj.value == 'Break':
                        break_col = c
                    elif cell_obj.value == 'Weld':
                        weld_col = c
                    elif cell_obj.value == 'PowdCoat Y/N':
                        powder_col = c
                    elif cell_obj.value == 'Quantity':
                        quantity_col = c
                    elif cell_obj.value == 'Batch':
                        batch_col = c
                    else:
                        all_col_read = True
                        column_vars = [part_col, type_col, material_col, break_col, weld_col, powder_col, quantity_col,
                                       batch_col]
                        for column_var in column_vars:
                            if column_var is None:
                                all_col_read = False
                                break
                        if all_col_read:
                            read_col_labels = False
                elif all_col_read:
                    if c == part_col and cell_obj.value:
                        current_part = cell_obj.value
                        self.parts_list.append(current_part)
                        self.parts[current_part] = {}
                    elif current_part:
                        if c == type_col:
                            self.parts[current_part]['type'] = cell_obj.value
                        elif c == material_col:
                            self.parts[current_part]['material'] = cell_obj.value
                        elif c == break_col:
                            if cell_obj.value == 'N':
                                self.parts[current_part]['break'] = False
                            else:
                                self.parts[current_part]['break'] = True
                        elif c == weld_col:
                            if cell_obj.value == 'N':
                                self.parts[current_part]['weld'] = False
                            else:
                                self.parts[current_part]['weld'] = True
                        elif c == powder_col:
                            if cell_obj.value == 'N':
                                self.parts[current_part]['powder'] = False
                            else:
                                self.parts[current_part]['powder'] = True
                        elif c == quantity_col:
                            self.parts[current_part]['quantity'] = cell_obj.value
                        elif c == batch_col and self.batch is None:
                            self.batch = cell_obj.value
                if cell_obj.value == '#':
                    read_col_labels = True

        self.make_break_list()

    def make_break_list(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.worksheets[0]
        sheet.freeze_panes = sheet['A2']
        sheet.merge_cells('A1:B1')
        sheet['A1'].value = 'Batch: ' + str(self.batch)
        sheet.cell(row=1, column=1).alignment = openpyxl.styles.Alignment(horizontal='center')
        i = 2
        temp_dir = self.cur_dir + 'barcode_temp/'
        if os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir)
        os.mkdir(temp_dir)
        for part in self.parts_list:
            part_name = part + ' ['
            if self.parts[part]['break']:
                part_name += 'B'
            if self.parts[part]['weld']:
                if part_name[-1] != '[':
                    part_name += '-'
                part_name += 'W'
            if self.parts[part]['powder']:
                if part_name[-1] != '[':
                    part_name += '-'
                part_name += 'P'
            if part_name[-1] == '[':
                part_name = part_name[0:-2]
            else:
                part_name += ']'
            sheet.cell(row=i, column=2).alignment = openpyxl.styles.Alignment(vertical='top')
            sheet.cell(row=i, column=2).value = part_name
            with open(temp_dir + part + '.png', "wb") as file:
                Code128(part, writer=ImageWriter()).write(file, options={'write_text': False})
            img = openpyxl.drawing.image.Image(temp_dir + part + '.png')
            height = 18
            width = 210
            img.height = height
            img.width = width
            # Silly nonsense below
            row_offset = 0.25
            marker = AnchorMarker(col=0, row=i-1, rowOff=cm_to_EMU((row_offset * 49.77) / 99),
                                  colOff=cm_to_EMU((0.1 * 49.77) / 99))
            size = XDRPositiveSize2D(pixels_to_EMU(width), pixels_to_EMU(height))
            img.anchor = OneCellAnchor(_from=marker, ext=size)
            sheet.add_image(img)
            i += 1
        dims = {}
        for row in sheet.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
        for col, value in dims.items():
            sheet.column_dimensions[col].width = value + 2
        for row in range(1, sheet.max_row + 1):
            sheet.row_dimensions[row].height = 20
        sheet.column_dimensions['A'].width = 30
        workbook.save('test.xlsx')
        if os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir)
        print('DONE')

    # noinspection PyMethodMayBeStatic
    def test_barcode(self):
        with open('test.png', "wb") as file:
            Code128('This is a test', writer=ImageWriter()).write(file, options={'write_text': False})
        self.make_break_list()


ChecklistProgram()
