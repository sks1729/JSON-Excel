import tkinter as tk
from tkinter import font
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os
from pathlib import Path
import requests
import json
from pandas.io.json import json_normalize
import pandas as pd
import openpyxl
import xlrd


class Main:
    def __init__(self, master):
        self.master = master
        self.master.title("JSON ↔ Excel converter")
        self.master.iconbitmap("joker.ico")
        self.master.option_add('*Dialog.msg.font', 'Helvetica 12')
        self.master.resizable(False, False)

        style = ttk.Style()
        style.theme_create("TabStyle", parent="classic", settings={
            "TNotebook": {"configure": {"tabmargins": [5, 5, 10, 0]}},
            "TNotebook.Tab": {"configure": {"padding": [10, 5], "font": ("Helvetica", 15)},
                              "map": {"foreground": [("selected", "#DB1D1D")],
                                      "background": [("selected", "SystemButtonFace")]}}})
        style.configure(
            "TabStyle", focuscolor=style.configure(".")["background"])
        style.theme_use("TabStyle")

        self.download_folder = str(os.path.join(Path.home(), "Downloads"))

        self.widgets_font = font.Font(family="Helvetica", size=15)
        self.description = "Convert multiple JSON files into Excel files\nat once saved " \
                           "in your system or from URL\nusing offline converter"
        self.excel_description = "Convert multiple Excel files into JSON files at once\nsaved" \
                                 " in your system using offline converter"
        self.footer = "The Excel files save in the Downloads folder"

        self.excel_footer = "The JSON files save in the Download folder"
        self.json_files = {}
        self.excel_files = {}

        self.correct_col_name = ""
        self.create_widgets()

    def create_widgets(self):
        self.add_image = tk.PhotoImage(file="add_files.png")
        self.convert_image = tk.PhotoImage(file="convert.png")

        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack()

        self.tab1 = tk.Frame(self.notebook)
        self.tab2 = tk.Frame(self.notebook)

        self.tab1.pack()
        self.tab2.pack()

        self.notebook.add(self.tab1, text="JSON to Excel")
        self.notebook.add(self.tab2, text="Excel to JSON")

        self.master.bind("<Visibility>", self.selected)

        self.json_excel_widgets()
        self.excel_json_widgets()

    def selected(self, event):
        if self.notebook.index("current") == 0:
            self.master.geometry("825x445+275+100")
        else:
            self.master.geometry("865x400+275+100")

    def json_excel_widgets(self):
        self.json_first_frame = tk.Frame(self.tab1, highlightbackground="#272121", highlightthickness=3)
        self.json_first_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.N)

        self.json_heading_label = tk.Label(self.json_first_frame, text="JSON to Excel converter", fg="#ca3e47",
                                           font="Helvetica 20 bold")
        self.json_heading_label.grid(row=0, column=0, pady=(10, 0), columnspan=2)

        self.json_description_label = tk.Label(self.json_first_frame, text=self.description, font=self.widgets_font,
                                               fg="#414141",
                                               justify="center")
        self.json_description_label.grid(row=1, column=0, padx=10, pady=(5, 10), columnspan=2)

        self.json_second_frame = tk.Frame(self.tab1, highlightbackground="#0f4c75", highlightthickness=3)
        self.json_second_frame.grid(row=0, column=1, pady=10, padx=(0, 10), sticky=tk.N)

        self.json_files_label = tk.Label(self.json_second_frame, text="Selected JSON files", font="Helvetica 15 bold",
                                         fg="#ca3e47")
        self.json_files_label.grid(row=0, column=0, padx=10, pady=(15, 10), sticky=tk.W)

        self.json_listbox_frame = tk.Frame(self.json_second_frame)
        self.json_list_box = tk.Listbox(self.json_listbox_frame, width=30, height=8, selectmode=tk.EXTENDED,
                                        font=self.widgets_font, bg="#313131", fg="white", activestyle="none")
        self.json_listbox_frame.grid(row=1, column=0, columnspan=2)
        self.json_list_box.grid(row=1, column=0, pady=(0, 10), padx=10)

        self.json_clear_button = tk.Button(self.json_second_frame, text="Clear", font="Helvetica 12 bold", bg="#d8d3cd",
                                           command=self.clear_json)
        self.json_clear_button.grid(row=2, column=1, padx=11, pady=(0, 10), sticky=tk.E)

        self.json_reset_button = tk.Button(self.json_second_frame, text="Reset", font="Helvetica 12 bold", bg="#d8d3cd",
                                           command=self.reset_json)
        self.json_reset_button.grid(row=3, column=1, padx=11, sticky=tk.E)

        self.json_eg_button = tk.Button(self.json_second_frame, text="Show example", font="Helvetica 12 bold",
                                        bg="#d8d3cd",
                                        command=lambda: os.startfile("example.png"))
        self.json_eg_button.grid(row=2, column=0, padx=11, pady=(0, 10), sticky=tk.W)

        self.json_enter_col = tk.Button(self.json_second_frame, text="Enter column name", font="Helvetica 12 bold",
                                        bg="#d8d3cd",
                                        command=self.column, state=tk.DISABLED)
        self.json_enter_col.grid(row=3, column=0, padx=11, sticky=tk.W)

        self.json_col_entry = tk.Entry(self.json_second_frame, font="Helvetica 12 bold")
        self.json_col_entry.grid(row=4, column=0, padx=11, pady=(8, 13), sticky=tk.W)

        self.json_add_button = tk.Button(self.json_first_frame, image=self.add_image, bd=0, command=self.add_jsons)
        self.json_add_button.grid(row=2, column=0, sticky=tk.W, padx=10)

        self.json_convert_button = tk.Button(self.json_first_frame, image=self.convert_image, bd=0, state=tk.DISABLED,
                                             command=self.convert_json)
        self.json_convert_button.grid(row=2, column=1, padx=10, sticky=tk.E)

        self.json_url_label = tk.Label(self.json_first_frame, text="Use URL", font=self.widgets_font, bg="#d8d3cd")
        self.json_url_label.grid(row=3, column=0, padx=10, pady=(40, 5), sticky=tk.W)

        self.json_url_entry = tk.Entry(self.json_first_frame, font=self.widgets_font, relief=tk.RAISED,
                                       highlightbackground="#d8d3cd", highlightthickness=3)
        self.json_url_entry.grid(row=4, column=0, padx=10)

        self.json_url_entry_button = tk.Button(self.json_first_frame, text="Add URL", font=self.widgets_font,
                                               command=self.json_use_url, relief=tk.RAISED, bg="#d8d3cd")
        self.json_url_entry_button.grid(row=4, column=1, padx=10, sticky=tk.W)

        self.json_footer_label = tk.Label(self.json_first_frame, text=self.footer, font=self.widgets_font, bg="#d8d3cd")
        self.json_footer_label.grid(row=5, column=0, padx=10, pady=(20, 10), columnspan=2, sticky=tk.W)

    def excel_json_widgets(self):
        self.excel_first_frame = tk.Frame(self.tab2, highlightbackground="#272121", highlightthickness=3)
        self.excel_first_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.N)

        self.excel_heading_label = tk.Label(self.excel_first_frame, text="Excel to JSON converter", fg="#ca3e47",
                                            font="Helvetica 20 bold")
        self.excel_heading_label.grid(row=0, column=0, pady=(10, 0), columnspan=2)

        self.excel_description_label = tk.Label(self.excel_first_frame, text=self.excel_description,
                                                font=self.widgets_font,
                                                fg="#414141", justify="center")
        self.excel_description_label.grid(row=1, column=0, padx=10, pady=(5, 10), columnspan=2)

        self.excel_add_button = tk.Button(self.excel_first_frame, image=self.add_image, bd=0, command=self.add_excel)
        self.excel_add_button.grid(row=2, column=0, sticky=tk.W, padx=10)

        self.excel_convert_button = tk.Button(self.excel_first_frame, image=self.convert_image, bd=0, state=tk.DISABLED,
                                              command=self.convert_excel)
        self.excel_convert_button.grid(row=2, column=1, padx=10, sticky=tk.E)

        self.rows2json = tk.StringVar()
        self.row_json_check = tk.Checkbutton(self.excel_first_frame, font=self.widgets_font, text="Rows to JSON",
                                             variable=self.rows2json, onvalue="on", offvalue="off", bg="#d8d3cd")
        self.row_json_check.deselect()
        self.row_json_check.grid(row=3, column=0, pady=10, padx=10, sticky=tk.W)

        self.multi_sheet = tk.StringVar()
        self.sheets_check = tk.Checkbutton(self.excel_first_frame, font=self.widgets_font, text="Multiple sheets",
                                           variable=self.multi_sheet, onvalue="on", offvalue="off", bg="#d8d3cd")
        self.sheets_check.deselect()
        self.sheets_check.grid(row=4, column=0, padx=10, sticky=tk.W)

        self.pretty_json = tk.StringVar()
        self.pretty_check = tk.Checkbutton(self.excel_first_frame, font=self.widgets_font, text="Beautify JSON",
                                           variable=self.pretty_json, onvalue="on", offvalue="off", bg="#d8d3cd")
        self.pretty_check.deselect()
        self.pretty_check.grid(row=3, column=1, pady=10, padx=10, sticky=tk.E)

        self.excel_footer_label = tk.Label(self.excel_first_frame, text=self.excel_footer, font=self.widgets_font,
                                           bg="#d8d3cd")
        self.excel_footer_label.grid(row=6, column=0, padx=10, pady=(20, 10), columnspan=2, sticky=tk.W)

        self.excel_second_frame = tk.Frame(self.tab2, highlightbackground="#0f4c75", highlightthickness=3)
        self.excel_second_frame.grid(row=0, column=1, pady=10, padx=(0, 10), sticky=tk.N)

        self.excel_files_label = tk.Label(self.excel_second_frame, text="Selected Excel files",
                                          font="Helvetica 15 bold", fg="#ca3e47")
        self.excel_files_label.grid(row=0, column=0, padx=10, pady=(15, 10), sticky=tk.W)

        self.excel_listbox_frame = tk.Frame(self.excel_second_frame)
        self.excel_list_box = tk.Listbox(self.excel_listbox_frame, width=30, height=8, selectmode=tk.EXTENDED,
                                         font=self.widgets_font, bg="#313131", fg="white", activestyle="none")
        self.excel_listbox_frame.grid(row=1, column=0, columnspan=2)
        self.excel_list_box.grid(row=1, column=0, pady=(0, 10), padx=10)

        self.excel_clear_button = tk.Button(self.excel_second_frame, text="Clear", font="Helvetica 12 bold",
                                            bg="#d8d3cd",
                                            command=self.clear_excel)
        self.excel_clear_button.grid(row=2, column=0, padx=10, pady=(0, 40), sticky=tk.W)

        self.excel_reset_button = tk.Button(self.excel_second_frame, text="Reset", font="Helvetica 12 bold",
                                            bg="#d8d3cd",
                                            command=self.reset_excel)
        self.excel_reset_button.grid(row=2, column=0, padx=80, pady=(0, 40), sticky=tk.W)

    def clear_json(self):
        for index in reversed(self.json_list_box.curselection()):
            file_name = self.json_list_box.get(index)
            self.json_files.pop(file_name)
            self.json_list_box.delete(index)
        if len(self.json_files) == 0:
            self.reset_json()

    def clear_excel(self):
        for index in reversed(self.excel_list_box.curselection()):
            file_name = self.excel_list_box.get(index)
            self.excel_files.pop(file_name)
            self.excel_list_box.delete(index)
        if len(self.excel_files) == 0:
            self.reset_excel()

    def reset_json(self):
        self.json_list_box.delete(0, tk.END)
        self.json_convert_button["state"] = tk.DISABLED
        self.correct_col_name = ""
        self.json_col_entry.delete(0, tk.END)
        self.json_enter_col["state"] = tk.DISABLED
        self.json_files = {}

    def reset_excel(self):
        self.excel_list_box.delete(0, tk.END)
        self.excel_convert_button["state"] = tk.DISABLED
        self.rows2json.set("off")
        self.pretty_json.set("off")
        self.multi_sheet.set("off")
        self.excel_files = {}

    def json_use_url(self):
        try:
            url = self.json_url_entry.get()
            self.json_url_entry.delete(0, tk.END)
            resp = requests.get(url=url)
            df = json_normalize(resp.json())
            excel_file_name = self.download_folder + "\\" + url.split("/")[-1].replace(".json", ".xlsx")
            df.to_excel(excel_file_name, index=False)
            os.startfile(self.download_folder)
        except:
            messagebox.askretrycancel("Invalid URL", "Enter a valid URL")

    def column(self):
        self.entered_col_name = self.json_col_entry.get().strip()
        if self.entered_col_name == "":
            messagebox.askretrycancel("Empty column name", "Enter a valid column name")
            self.json_col_entry.delete(0, tk.END)
        else:
            try:
                for file in self.json_files.values():
                    data = json.load(open(file))
                    pd.DataFrame(data[self.entered_col_name])
            except:
                messagebox.askretrycancel("Invalid column name", "Enter a valid column name")
                self.correct_col_name = ""
            else:
                self.correct_col_name = self.entered_col_name
                self.json_col_entry.delete(0, tk.END)

    def add_jsons(self):
        self.files = filedialog.askopenfilenames(initialdir=os.getcwd(), title="Select JSON file(s)",
                                                 filetypes=[("JSON files", "*.json")])
        filedialog.geometry = "+250+100"
        for file in self.files:
            full_path = file
            file_name = full_path.split("/")[-1]
            self.json_files[file_name] = full_path
            self.json_list_box.insert(tk.END, file_name)
        if len(self.json_files) > 0:
            self.json_convert_button["state"] = tk.NORMAL
            self.json_enter_col["state"] = tk.NORMAL

    def add_excel(self):
        self.files = filedialog.askopenfilenames(initialdir=os.getcwd(), title="Select Excel file(s)",
                                                 filetypes=[("Excel files", ".xl*")])
        filedialog.geometry = "+250+100"
        for file in self.files:
            full_path = file
            file_name = full_path.split("/")[-1]
            self.excel_files[file_name] = full_path
            self.excel_list_box.insert(tk.END, file_name)
        if len(self.excel_files) > 0:
            self.excel_convert_button["state"] = tk.NORMAL

    def convert_json(self):
        try:
            if self.correct_col_name != "":
                for file in self.json_files.values():
                    full_file_path = self.download_folder + "\\" + str(file).split("/")[-1].replace(".json", ".xlsx")
                    data = json.load(open(file))
                    df = pd.DataFrame(data[str(self.correct_col_name)])
                    df.to_excel(full_file_path)
            else:
                for file in self.json_files.values():
                    full_file_path = self.download_folder + "\\" + str(file).split("/")[-1].replace(".json", ".xlsx")
                    data = json.load(open(file))
                    df = pd.DataFrame(data)
                    df.to_excel(full_file_path, index=False)
            os.startfile(self.download_folder)
        except:
            messagebox.askretrycancel("Conversion not possible", "Enter valid column name")

    def prettify_json(self, f_names):
        for f in f_names:
            file_object = open(f, "r+")
            json_object = json.load(file_object)
            formatted_json = json.dumps(json_object, indent=3)
            file_object.seek(0)
            file_object.write(formatted_json)

    def convert_excel(self):
        self.row_json_check["state"] = tk.DISABLED
        self.pretty_check["state"] = tk.DISABLED
        self.sheets_check["state"] = tk.DISABLED
        try:
            if self.rows2json.get() == "on":
                file_names = []
                for file in self.excel_files.values():
                    if self.multi_sheet.get() == "on":
                        for sheet in pd.ExcelFile(file).sheet_names:
                            df = pd.ExcelFile(file).parse(sheet)
                            for i in sorted(list(df.index), reverse=True):
                                extension_replace = str(file).split("/")[-1].replace(".xlsx", "").replace(".xls", "")
                                full_file_path = f"{self.download_folder}\\{extension_replace} {sheet} Record index ({i + 1}).json"
                                file_names.append(full_file_path)
                                data = df.iloc[i]
                                data.to_json(full_file_path)
                    else:
                        df = pd.read_excel(file)
                        for i in sorted(list(df.index), reverse=True):
                            extension_replace = str(file).split("/")[-1].replace(".xlsx", "").replace(".xls", "")
                            full_file_path = f"{self.download_folder}\\{extension_replace} Record index ({i + 1}).json"
                            file_names.append(full_file_path)
                            data = df.iloc[i]
                            data.to_json(full_file_path)
                if self.pretty_json.get() == "on":
                    self.prettify_json(file_names)
            else:
                file_names = []
                for file in self.excel_files.values():
                    if self.multi_sheet.get() == "on":
                        for sheet in pd.ExcelFile(file).sheet_names:
                            extension_replace = str(file).split("/")[-1].replace(".xlsx", "").replace(".xls", "")
                            full_file_path = f"{self.download_folder}\\{extension_replace} {sheet}.json"
                            file_names.append(full_file_path)
                            pd.ExcelFile(file).parse(sheet).to_json(full_file_path, orient="records")
                    else:
                        extension_replace = str(file).split("/")[-1].replace(".xlsx", "").replace(".xls", "")
                        full_file_path = f"{self.download_folder}\\{extension_replace}.json"
                        file_names.append(full_file_path)
                        pd.read_excel(file).to_json(full_file_path, orient="records")
                if self.pretty_json.get() == "on":
                    self.prettify_json(file_names)
            self.row_json_check["state"] = tk.NORMAL
            self.pretty_check["state"] = tk.NORMAL
            self.sheets_check["state"] = tk.NORMAL
            os.startfile(self.download_folder)
        except:
            messagebox.askretrycancel("Conversion not possible", "Try again")


if __name__ == '__main__':
    root = tk.Tk()
    Main(root)
    root.mainloop()
