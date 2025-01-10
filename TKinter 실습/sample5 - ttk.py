from tkinter import *
from tkinter import ttk

root = Tk()
label = ttk.Label(root, text = "Hello")
label.pack()

label.config(text = 'The label is changed')

root.mainloop()
