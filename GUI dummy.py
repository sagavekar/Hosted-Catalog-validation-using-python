import tkinter as tk
from tkinter import filedialog
import pandas as pd

def browse_file1():
    file_path = filedialog.askopenfilename(filetypes=[("All files", ("*.xlsx", "*.xls","*.csv"))])
    file_label1.config(text=file_path)

def browse_file2():
    file_path = filedialog.askopenfilename(filetypes=[("All files", ("*.xlsx", "*.xls","*.csv"))])
    file_label2.config(text=file_path)

def load_files():
    file_path1 = file_label1.cget("text")
    file_path2 = file_label2.cget("text")
    if file_path1 and file_path2:
        df1 = None
        df2 = None
        if file_path1.endswith(".csv"):
            df1 = pd.read_csv(file_path1)
        else:
            df1 = pd.read_excel(file_path1)
        
        if file_path2.endswith(".csv"):
            df2 = pd.read_csv(file_path2)
        else:
            df2 = pd.read_excel(file_path2)
        
        #print("File 1:")
        #print(df1.head())
        #print("File 2:")
        #print(df2.head())
    else:
        print("Please select both files first.")

root = tk.Tk()
root.geometry("500x370")
root.title("ABM Catalog validator")

root.minsize(250,250)
root.maxsize(800,600)

label1 = tk.Label(root, text="Select system extract (CSV or Excel):")
label1.pack()

button1 = tk.Button(root, text="Browse", command=browse_file1)
button1.pack()

file_label1 = tk.Label(root, text="")
file_label1.pack()

label2 = tk.Label(root, text="Select Supplier template (CSV or Excel):")
label2.pack()

button2 = tk.Button(root, text="Browse", command=browse_file2)
button2.pack()

file_label2 = tk.Label(root, text="")
file_label2.pack()

submit_button = tk.Button(root, text="Submit", command=load_files)
submit_button.pack()

root.mainloop()
