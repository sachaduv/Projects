from tkinter import *

root = Tk()

topFrame = Frame(root)
topFrame.pack()

leftFrame = Frame(root)
leftFrame.pack(side=LEFT)

bottomFrame = Frame(root)
bottomFrame.pack(side=BOTTOM)

theLabel = Label(topFrame,text = "CWIS-ISOM")
theLabel.pack()

ord_ful_btn = Button(None,text="Order Fullfillment",fg="Green")
ord_ful_btn.pack(side=LEFT,fill = Y)
submit = Button(None,text="Submit",fg="Blue")
submit.pack(side = BOTTOM,fill = X)

root.mainloop()

