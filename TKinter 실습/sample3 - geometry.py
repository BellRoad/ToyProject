from tkinter import *
root = Tk()

lbl = Label(root, text = "Name")
lbl.grid(row = 0, column = 0)

txt = Entry(root)
txt.grid(row = 1, column = 0)

btn = Button(root, text = "OK", width=15)
btn.grid(row = 2, column = 0)


cbtn = Checkbutton(root, text = "Check")
cbtn.grid(row = 3, column = 0)



root.mainloop()
