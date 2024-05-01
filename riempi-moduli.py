# generate a python script that opens a window, where the user can enter pairs of keys and values, and the user can enter a list of files. The keys and values are stored in a hidden file, and are restored each time that the program is launched.

import os
import sys
import io
import glob
from pathlib import Path
from docxtpl import DocxTemplate
import openpyxl
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog


parametri_file_default = "parametri.xlsx"
template_dir_default = "./template/"
save_dir_default = "./filled/"


def fill():

    context = {}


    parametri_path = parametri_file.get()

    if os.path.exists(parametri_path):

        xlsxfile = openpyxl.load_workbook(parametri_path)

        sheet = xlsxfile.worksheets[0]

        for row in sheet.iter_rows():

            if row[0].value is None:
                break

            key = row[0].value
            value = row[1].value

            context[key] = value

            if context[key] == "":
                messagebox.showerror("Error", f"Campo {key} vuoto!")
                return

    else:
        messagebox.showerror("Error", f"File {parametri_path} non trovato!")
        return


    templates_path = template_dir.get()

    templates = []


    if os.path.exists(templates_path):

        for template in glob.glob(templates_path + "*.docx"):
            templates.append(template)
        
        if templates.__len__() == 0:
            messagebox.showerror("Error", f"Nessun template trovato in {templates_path}!")
            return
    else:
        messagebox.showerror("Error", f"Directory {templates_path} non trovata!")
        return



    save_path = save_dir.get()
    

    if not os.path.exists(save_path):

        os.makedirs(save_path) 


    for template_name in templates:
        
        doc_path = ""

        print(f"Processing {template_name}...")

        if "Suffisso_Nome_File" in context and context["Suffisso_Nome_File"] != "":
            doc_path = "{0}/{1}_{2}{3}".format(save_path, Path(template_name).stem, context["Suffisso_Nome_File"], Path(template_name).suffix)
        else:
            doc_path = save_path + os.path.basename(template_name)

        print(f"Saving to {doc_path}...")

        template = DocxTemplate(template_name)

        io_buffer = io.BytesIO()

        template.render(context)

        template.save(doc_path)




def select_xlsx():
    filetypes = [("Excel files", "*.xlsx")]
    filename = filedialog.askopenfilenames(filetypes=filetypes, multiple=False, initialdir=".")
    print(filename)
    savedir_value.set(filename)

def select_template_dir():
    dirname = filedialog.askdirectory(initialdir=template_dir_default)
    print(dirname)
    savedir_value.set(dirname)

def select_savedir():
    dirname = filedialog.askdirectory(initialdir=save_dir_default)
    print(dirname)
    savedir_value.set(dirname)


if __name__ == "__main__":

    root = tk.Tk()
    root.title("Riempitore di moduli")

    parametri_file = tk.StringVar(root,parametri_file_default)

    template_dir = tk.StringVar(root,template_dir_default)

    save_dir = tk.StringVar(root,save_dir_default)

    parametri_label = tk.Label(root, text="Parametri (XLSX)").grid(row=0, column=0)
    parametri_button = tk.Button(root, text="Seleziona file...", command=select_xlsx).grid(row=0, column=1)
    parametri_value = tk.Label(root, textvariable=parametri_file).grid(row=0, column=2)

    template_label = tk.Label(root, text="Directory template").grid(row=1, column=0)
    template_button = tk.Button(root, text="Seleziona directory...", command=select_template_dir).grid(row=1, column=1)
    template_value = tk.Label(root, textvariable=template_dir).grid(row=1, column=2)

    savedir_label = tk.Label(root, text="Directory salvataggio").grid(row=2, column=0)
    savedir_button = tk.Button(root, text="Seleziona directory...", command=select_savedir).grid(row=2, column=1)
    savedir_value = tk.Label(root, textvariable=save_dir).grid(row=2, column=2)

    filenames = tk.StringVar()


    fill_button = tk.Button(root, text="Genera", command=fill).grid(row=3, column=1)

    root.mainloop()

