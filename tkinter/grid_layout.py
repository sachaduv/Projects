from tkinter import *

root = Tk()

label1 = Label(root,text="Name")
label1.grid(row=0,column=0,sticky="e")
label2 = Label(root,text="Password")
label2.grid(row=1,column=0,sticky="E")
entryPassword = Entry()
entryPassword.grid(row=1,column=1)
entryName = Entry(root)
entryName.grid(row=0,column=1)
chkRemember = Checkbutton(root,text="Remember me")
chkRemember.grid(columnspan = 2)

root.mainloop()