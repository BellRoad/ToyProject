import tkinter

window=tkinter.Tk()
window.title("BellRoad")
window.geometry("640x480+100+100")
window.resizable(0, 0)

count=0

def countUP():
	global count
	count +=1
	label.config(text=str(count))

label = tkinter.Label(window, text="0")
label.pack()

button = tkinter.Button(window, overrelief="solid", width=15, command=countUP, repeatdelay=1000, repeatinterval=100, text="PUSH!")
button.pack()

window.mainloop()
