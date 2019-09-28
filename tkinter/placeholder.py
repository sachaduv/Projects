from tkinter import *
root=Tk()
e1=Entry(root)
e1.pack()
e2=Entry(root)
e2.pack()
def handleReturn(event):
    print("return: event.widget is",event.widget)
    print("focus is:",root.focus_get())

root.bind("<Return>",handleReturn)
root.mainloop()