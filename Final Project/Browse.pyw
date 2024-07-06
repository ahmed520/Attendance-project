from tkinter import *
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import customtkinter as ctk
from tkinter import messagebox
import os


text_file_path = "C:\Attendance\Code\Important.txt"


def is_excel_file(file_path):
    _, extension = os.path.splitext(file_path)
    return extension in ['.xlsx', '.xlsm', '.xlsb', '.xls']


def replace_line(text, filename, line_number):
    with open(filename, "r") as file:
        lines = file.readlines()

    # Replace the second line with the new text
    lines[line_number] = text + "\n"

    with open(filename, "w") as file:
        file.writelines(lines)


def my_func():
    urfile = filedialog.askdirectory(initialdir="C:/", title="save as")
    E_out.insert(0, urfile)


def my_funct():
    urfiles = filedialog.askdirectory(initialdir="C:/", title="choose")
    E_w.insert(0, urfiles)


def my_fun():
    myfile = filedialog.askopenfilename(
        initialdir="C:/", title="select a file")
    E_import.insert(0, myfile)


def get_input():
    user_in = E_import.get()
    user_inpu = E_out.get()
    user_inp = E_w.get()
    if os.path.exists(user_inpu):
        replace_line(user_inpu, text_file_path, 2)
    else:
        messagebox.showerror("Error", "File path not found.")
    if os.path.exists(user_inp):
        replace_line(user_inp, text_file_path, 0)
    else:
        messagebox.showerror("Error", "File path not found.")
    if os.path.isfile(user_in) and is_excel_file(user_in):
        replace_line(user_in, text_file_path, 1)
    else:
        if not is_excel_file(user_in):
            messagebox.showerror("Error", "Not excel file.")
        else:
            messagebox.showerror("Error", "File  not found.")
    if os.path.exists(user_inpu) and os.path.exists(user_inp) and os.path.isfile(user_in) and is_excel_file(user_in):
        root2.destroy()


ctk.set_appearance_mode("Dark")
root2 = ctk.CTk()
root2.geometry("350x200")
root2.title("Browse")
root2.resizable(False, False)

limport = ctk.CTkLabel(root2, text="Import Data Base", font=('Arial', 14))
# label1.place(x=-80, y=10, height=30, width=300)
limport.grid(row=0, column=0)

b_import = ctk.CTkButton(root2, text="Open", font=(
    'Arial', 10), width=30, command=lambda: (E_import.delete(0, tk.END), my_fun()))
b_import.grid(row=1, column=1, padx=5)

E_import = ctk.CTkEntry(root2, width=250)
# L_entry.place(x=10, y=60, height=20, width=300)
E_import.grid(row=1, column=0, padx=15)

l_out = ctk.CTkLabel(
    root2, text="Choose output destination", font=('Arial', 14))
# label1.place(x=-80, y=130, height=30, width=300)
l_out.grid(row=2, column=0)

b_out = ctk.CTkButton(root2, text="Save", font=(
    'Arial', 10), width=30, command=lambda: (E_out.delete(0, tk.END), my_func()))
# button11.place(x=320, y=160, height=20, width=50)
b_out.grid(row=3, column=1, padx=5)

E_out = ctk.CTkEntry(root2,  width=250)
# Le_entry.place(x=10, y=160, height=20, width=300)
E_out.grid(row=3, column=0)


l_w = ctk.CTkLabel(root2, text="Choose W.N.W file location", font=('Arial', 14))
# label10.place(x=-80, y=240, height=30, width=300)
l_w.grid(row=4, column=0)

b_w = ctk.CTkButton(root2, text="Open", font=(
    'Arial', 10), width=30, command=lambda: (E_w.delete(0, tk.END), my_funct()))
# button110.place(x=320, y=290, height=20, width=50)
b_w.grid(row=5, column=1, padx=5)

E_w = ctk.CTkEntry(root2, width=250)
# Le_entry1.place(x=10, y=290, height=20, width=300)
E_w.grid(row=5, column=0)

button222 = ctk.CTkButton(root2, text="Finish", font=(
    'Arial', 12), command=lambda: get_input())
button222.grid(row=6, column=0, padx=10)
root2.mainloop()
