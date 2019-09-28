from tkinter import *

root = Tk()

equation = ""
expression = StringVar()

screen = Label(root,textvariable=expression)
screen.grid(columnspan=4)

def btnpress(num):
    global equation 
    equation+= str(num)
    expression.set(equation)

def clearScreen():
    global equation
    equation = ""
    expression.set("")

def evaluate():
    global equation
    sol = str(eval(equation))
    expression.set(sol)
    equation=""
    

btn1 = Button(root,text="1",command = lambda:btnpress(1))
btn1.grid(row=1,column=0,padx = 10,pady =10,sticky="ew")
btn2 = Button(root,text="2",command = lambda:btnpress(2))
btn2.grid(row=1,column=1,padx = 10,pady =10,sticky="ew")
btn3 = Button(root,text="3",command = lambda:btnpress(3))
btn3.grid(row=1,column=2,padx = 10,pady =10,sticky="ew")
btnplus = Button(root,text="+",command = lambda:btnpress("+"))
btnplus.grid(row=1,column=3,padx = 10,pady =10,sticky="ew")
btn4 = Button(root,text="4",command = lambda:btnpress(4))
btn4.grid(row=2,column=0,padx = 10,pady =10,sticky="ew")
btn5 = Button(root,text="5",command = lambda:btnpress(5))
btn5.grid(row=2,column=1,padx = 10,pady =10,sticky="ew")
btn6 = Button(root,text="6",command = lambda:btnpress(6))
btn6.grid(row=2,column=2,padx = 10,pady =10,sticky="ew")
btnsub = Button(root,text="-",command = lambda:btnpress("-"))
btnsub.grid(row=2,column=3,padx = 10,pady =10,sticky="ew")
btn7 = Button(root,text="7",command = lambda:btnpress(7))
btn7.grid(row=3,column=0,padx = 10,pady =10,sticky="ew")
btn8 = Button(root,text="8",command = lambda:btnpress(8))
btn8.grid(row=3,column=1,padx = 10,pady =10)
btn9 = Button(root,text="9",command = lambda:btnpress(9))
btn9.grid(row=3,column=2,padx = 10,pady =10,sticky="ew")
btnmul = Button(root,text="*",command = lambda:btnpress("*"))
btnmul.grid(row=3,column=3,padx = 10,pady =10,sticky="ew")
btnc = Button(root,text="C",command = clearScreen)
btnc.grid(row=4,column=0,padx = 10,pady =10,sticky="ew")
btn0 = Button(root,text="0",command = lambda:btnpress(0))
btn0.grid(row=4,column=1,padx = 10,pady =10,sticky="ew")
btneq = Button(root,text="=",command = evaluate)
btneq.grid(row=4,column=2,padx = 10,pady =10,sticky="ew")
btndiv = Button(root,text="/",command = lambda:btnpress("/"))
btndiv.grid(row=4,column=3,padx = 10,pady =10,sticky="ew")

root.mainloop()