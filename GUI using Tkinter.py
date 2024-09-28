import tkinter as tk
from PIL import Image , ImageTk # PIL  = python imaging library

def omkar():
    print("OKar")

root = tk.Tk() #root


#widthxheight
root.geometry("500x370")
root.title("ABM Catalog validator")

root.minsize(250,250)
root.maxsize(800,600)

label1 = tk.Label(text="ABM Catalog validator", fg = "white", bg ="Grey", font="Tahoma 11")
label1.pack(fill="x")

label3 = tk.Label(text="Designed and Developed by - Omkar Sagavekar ", fg = "white", bg ="Grey", font="Tahoma 10")
label3.pack(anchor="se",side="bottom",fill="x")

logo = tk.PhotoImage(file="GEP.png")
lable2 =  tk.Label(image=logo)
lable2.pack(anchor="nw", side="top")

f1 = tk.Frame(root, borderwidth=2, bg="red")
f1.pack(side="bottom")

b1 = tk.Button(f1,text="Run", command=omkar)
b1.pack(padx=2)

root.mainloop() #
