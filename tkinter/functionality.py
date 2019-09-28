from tkinter import *

root =Tk()

def leftClick(event):
    print("Left")
def rightClick(event):
    print("Right")
def scroll(event):
    print("Scroll")
def leftKey(event):
    print("Left key pressed")
def rightKey(event):
    print("Right key pressed")
def evaluateOnEnter(event):
    cal = entryCal.get()
    sol.configure(text="Solution is : "+str(eval(cal)))
label1 = Label(root,text="Enter your expression")
label1.pack()
entryCal = Entry()
entryCal.bind("<Return>",evaluateOnEnter)
entryCal.pack()

sol = Label(root)
sol.pack()

root.geometry("500x500")
root.bind("<Button-1>",leftClick)
root.bind("<Button-2>",rightClick)
root.bind("<Button-3>",scroll)
root.bind("<Left>",leftKey)
root.bind("<Right>",rightKey)
root.mainloop()