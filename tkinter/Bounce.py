from tkinter import *
import time
import random

tk = Tk()
tk.title("Bounce")
tk.resizable(0,0)
tk.wm_attributes("-topmost",1)
canvas = Canvas(tk,width=500,height=500,bd=0,highlightthickness=0)
canvas.pack()
tk.update()
class Ball:

    def __init__(self,canvas,peddle,color):
        self.hit_bottom = False
        self.canvas = canvas
        self.peddle = peddle 
        self.id = canvas.create_oval(10,10,25,25,fill=color)
        self.canvas.move(self.id,245,100)
        start = [-3,-2,-1,0,1,2,4]
        random.shuffle(start)
        self.x = start[0]
        self.y = -3
        self.canvas_height = self.canvas.winfo_height()
        self.canvas_width = self.canvas.winfo_width()

    def hit_peddle(self,pos):
        peddle_pos = self.canvas.coords(self.peddle.id)
        if(pos[2]>=peddle_pos[0] and pos[0]<=peddle_pos[2]):
            if(pos[3]>=peddle_pos[1] and pos[3]<=peddle_pos[3]):
                return True
            return False

    def draw(self):
        self.canvas.move(self.id,self.x,self.y)
        pos = canvas.coords(self.id)
        if(pos[1]<=0):
            self.y=1
        if(pos[3]>=self.canvas_height):
            self.hit_bottom = True
            canvas.create_text(245,100,text="Game over")
        if(pos[2]>=self.canvas_width):
            self.x=-3
        if(pos[0]<=0):
            self.x=3
        if(self.hit_peddle(pos)==True):
            self.y=-3

        


class Peddle:

    def __init__(self,canvas,color):
        self.canvas = canvas
        self.id = canvas.create_rectangle(0,0,100,10,fill=color)
        self.canvas.move(self.id,200,300)
        self.x = 0
        self.canvas_width = self.canvas.winfo_width()
        self.canvas.bind_all("<KeyPress-Left>",self.turn_left)
        self.canvas.bind_all("<KeyPress-Right>",self.turn_right)

    def draw(self):
        self.canvas.move(self.id,self.x,0)
        pos=self.canvas.coords(self.id)
        if(pos[0]<=0):
            self.x=0
        if(pos[2]>=self.canvas_width):
            self.x=0

    def turn_left(self,event):
        self.x = -2

    def turn_right(self,event):
        self.x = 2

peddle = Peddle(canvas,"red")
ball = Ball(canvas,peddle,"blue")

while 1:
    if(ball.hit_bottom ==  False):
        peddle.draw()
        ball.draw()
    tk.update_idletasks()
    tk.update()
    time.sleep(0.01)
    