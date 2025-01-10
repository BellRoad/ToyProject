from tkinter import *

class MyFrame(Frame):
    def __init__(self, master):
        img = PhotoImage(file='photo.png')
        lbl = Label(image=img)
        lbl.image = img
        lbl.place(x=10, y=10)

def main():
    root = Tk()
    root.title('View Image')
    root.geometry('800x600+300+300')
    myframe = MyFrame(root)
    root.mainloop()

if __name__ == '__main__':
    main()
