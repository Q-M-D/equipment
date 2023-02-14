import os
import tkinter.filedialog

import openpyxl
import datetime
from tkinter import *
import time

maxlen = 15
delete_lib = [' ', '/', '(', ')', '（', '）', '^']


class Format():
    def __init__(self, init_window):
        self.init_window = init_window

    def set_init_window(self):
        self.init_window.title("表格格式化工具")
        self.init_window.geometry('700x300+10+10')
        # self.init_window.geometry('320x160+10+10')
        self.init_title = Label(self.init_window, text="Format Function")
        self.init_title.grid(row=0, column=0)
        self.input_label = Label(self.init_window, text="Input path")
        self.input_label.grid(row=1, column=0)
        self.input_button = Button(self.init_window, text="upload", command=self.upload_file)
        self.input_button.grid(row=1, column=11)
        self.input = Entry(self.init_window, width=66)
        self.input.grid(row=1, column=1, columnspan=10)
        self.output_label = Label(self.init_window, text="Output")
        self.output_label.grid(row=2, column=0)
        self.output = Entry(self.init_window, width=66)
        self.output.grid(row=2, column=1, columnspan=10)
        self.status = Label(self.init_window, text="Status")
        self.status.grid(row=17, column=0)
        self.status = Entry(self.init_window, width=7)
        self.status.grid(row=17, column=1)

        self.format_button = Button(self.init_window, text="Format", command=self.format)
        self.format_button.grid(row=4, column=5)

        self.format_option_column_title = Label(self.init_window, text="column: ")
        self.format_option_column_title.grid(row=4, column=1)
        self.format_option_column = Entry(self.init_window, width=7)
        self.format_option_column.grid(row=4, column=2)

        self.format_option_length_title = Label(self.init_window, text="max-length: ")
        self.format_option_length_title.grid(row=4, column=3)
        self.format_option_length = Entry(self.init_window, width=7)
        self.format_option_length.grid(row=4, column=4)

        self.str_trans_to_md5_button = Button(self.init_window, text="Format all", bg="lightblue", width=10,
                                              command=self.format_all)  # 调用内部方法  加()为直接调用
        self.str_trans_to_md5_button.grid(row=12, columnspan=20)
        print(self.input.get())

    def format(self):
        if self.input.get() == '':
            print("No input")
            self.status.delete(0, 'end')
            self.status.insert(0, "Failed")
            self.status['bg'] = 'red'
            return

        workbook = openpyxl.load_workbook(self.input.get())
        sheet = workbook.active
        print("format model")
        column_num = int(self.format_option_column.get())
        print(column_num)
        max_length = int(self.format_option_length.get())
        print(max_length)
        for row in sheet:
            if (row[0].value == None):
                break
            cell = row[column_num-1]
            if cell.value == None:
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                else:
                    cell.value = "*"
            elif isinstance(cell.value, datetime.datetime):
                day = cell.value.day
                month = cell.value.month
                year = cell.value.year
                cell.value = str(year).zfill(2) + str(month).zfill(2)
            elif isinstance(cell.value, str):
                tmp = cell.value
                for item in delete_lib:
                    tmp = tmp.replace(item, '')
                tmpj = ""
                str_len = 0
                for i in range(len(tmp)):
                    # if tmp is digit, letter, '-', '*' or chinese, keep it
                    if '\u4e00' <= tmp[i] <= '\u9fa5':
                        if str_len < max_length - 1:
                            tmpj += tmp[i]
                            str_len += 2
                        else:
                            break
                    elif tmp[i].isdigit() or tmp[i].isalpha() or tmp[i] == '-' or tmp[i] == '*':
                        if str_len < max_length:
                            tmpj += tmp[i]
                            str_len += 1
                        else:
                            break
                if tmpj == "":
                    tmpj = "*"
                tmp = tmpj
                # if len(tmp) > max_length:
                #     tmp = tmp[:max_length]
                cell.value = tmp

        if self.output.get() == '':
            workbook.save(self.input.get())
        else:
            workbook.save(self.output.get())
        self.status.delete(0, 'end')
        self.status.insert(0, "Success")
        self.status['bg'] = 'green'
        return

    def format_all(self):
        if self.input.get() == '':
            print("No input")
            self.status.delete(0, 'end')
            self.status.insert(0, "Failed")
            self.status['bg'] = 'red'
            return

        workbook = openpyxl.load_workbook(self.input.get())
        sheet = workbook.active
        print("format all model")
        for row in sheet:
            if (row[0].value == None):
                break
            for cell in row:
                if cell.value == None:
                    if isinstance(cell, openpyxl.cell.cell.MergedCell):
                        continue
                    else:
                        cell.value = "*"
                elif isinstance(cell.value, datetime.datetime):
                    day = cell.value.day
                    month = cell.value.month
                    year = cell.value.year
                    cell.value = str(year).zfill(2) + str(month).zfill(2)
                elif isinstance(cell.value, str):
                    tmp = cell.value
                    for item in delete_lib:
                        tmp = tmp.replace(item, '')
                    tmpj = ""
                    str_len = 0
                    for i in range(len(tmp)):
                        # if tmp is digit, letter, '-', '*' or chinese, keep it
                        if ('\u4e00' <= tmp[i] <= '\u9fa5'):
                            if str_len < maxlen - 1:
                                tmpj += tmp[i]
                                str_len += 2
                            else:
                                break
                        elif (tmp[i].isdigit() or tmp[i].isalpha() or tmp[i] == '-' or tmp[i] == '*'):
                            if str_len < maxlen:
                                tmpj += tmp[i]
                                str_len += 1
                            else:
                                break
                    if tmpj == "":
                        tmpj = "*"
                    tmp = tmpj
                    # if len(tmp) > maxlen:
                    #     tmp = tmp[:maxlen]
                    cell.value = tmp
        if self.output.get() == '':
            workbook.save(self.input.get())
        else:
            workbook.save(self.output.get())
        self.status.delete(0, 'end')
        self.status.insert(0, "Success")
        self.status['bg'] = 'green'
        return


    def upload_file(self):
        self.input.delete(0, 'end')
        inputfile = tkinter.filedialog.askopenfilename()
        if inputfile[-4:] != 'xlsx':
            print('请上传excel表格')
        else:
            self.input.insert(0, inputfile)


def choose():
    # format_all()
    init_window = Tk()
    my_window = Format(init_window)
    my_window.set_init_window()
    init_window.mainloop()

choose()

