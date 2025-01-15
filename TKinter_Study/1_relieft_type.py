import tkinter

window = tkinter.Tk()

window.title("BellRoad")
window.geometry("640x400+100+100")
window.resizable(False, False)


relief_options=["flat", "groove", "raised", "ridge", "solid", "sunken"]

for relief_type in relief_options:
	label=tkinter.Label(window, text=relief_type, width=10, height=3, fg="red", relief=relief_type, anchor="center", border=3, font="맑은고딕")
	label.pack(pady=5)

window.mainloop()
