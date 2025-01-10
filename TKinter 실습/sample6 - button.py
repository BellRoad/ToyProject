from tkinter import *
from tkinter import ttk

def clickCallback():
    print('clieked')

root = Tk()
button = ttk.Button(root, text = 'Click Me', command = clickCallback)
button.pack()

root.mainloop()
