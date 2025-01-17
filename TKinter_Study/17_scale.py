import tkinter

window=tkinter.Tk()
window.title("YUN DAE HEE")
window.geometry("640x400+100+100")
window.resizable(False, False)

def select(value):
    label_text= "값 : {0}".format(str(value))
    label.config(text=label_text)

var=tkinter.IntVar()

scale=tkinter.Scale(window, variable=var, command=select, orient="horizontal", showvalue=False, tickinterval=50, to=500, length=300)
scale.pack()

label=tkinter.Label(window, text="값 : 0")
label.pack()

window.mainloop()