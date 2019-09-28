from tkinter import *
import random

root = Tk()
canvas = Canvas(root,width=300,height=300)
canvas.pack()
canvas.create_rectangle(0,0,100,270)
canvas.create_line(0,0,300,300)
canvas.create_polygon(0,0,5,20,50,30,40,15)
canvas.create_arc(10,10,200,80,extent = 45,style = ARC)
canvas.create_arc(10,80,200,180,extent = -90,style = ARC)
canvas.create_text(150,150,text="Canvas module",font=('Times',10))



colors = ['blue','green','red','yellow']

def create_random_rects(num):
    for i in range(0,num):
        random.shuffle(colors)
        x1 = random.randrange(150)
        y1 = random.randrange(150)
        x2 = x1+random.randrange(150)
        y2 = y1+random.randrange(150)
        canvas.create_rectangle(x1,y1,x2,y2,fill=colors[0])
#create_random_rects(150)
root.mainloop()