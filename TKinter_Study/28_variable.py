import tkinter

window=tkinter.Tk()
window.title("YUN DAE HEE")
window.geometry("640x400+100+100")
window.resizable(True, True)

def command_args(argument1, argument2, argument3):
    global arg1
    arg1 = argument1 * 2
    print(argument1, argument2, argument3)

arg1 = 1
arg2 = "alpha"
arg3 = "beta"

button = tkinter.Button(window, width=25, height=10, text="버튼", command=lambda: command_args(arg1, arg2, arg3))
button.pack(expand=True, anchor="center")

window.mainloop()