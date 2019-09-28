from tkinter import *
import time
import random

score_lft=0
score_ryt=0

tk=Tk()
tk.title("Pong")
tk.resizable(0,0)
tk.wm_attributes("-topmost",1)
canvas = Canvas(tk,width=500,height=400,bd=0,highlightthickness=0)
canvas.pack()
canvas.config(bg="black")
tk.update()
canvas.create_line(250,0,250,400,fill="white")

class Ball:

    def __init__(self,canvas,color,pdl_lft,pdl_ryt):
        self.canvas = canvas
        self.pdl_lft = pdl_lft
        self.pdl_ryt = pdl_ryt
        self.id = canvas.create_oval(10,10,25,25,fill=color)
        self.canvas.move(self.id,235,200)
        start_pos = [-3,3]
        random.shuffle(start_pos)
        self.x = start_pos[0]
        self.y = -3
        self.canvas_width = self.canvas.winfo_width()
        self.canvas_height = self.canvas.winfo_height()

    def draw(self):
        self.canvas.move(self.id,self.x,self.y)
        pos = self.canvas.coords(self.id)
        if(pos[0]<=0):
            self.x=3
            self.score(True)
        if(pos[2]>=self.canvas_width):
            self.x=-3
            self.score(False)
        if(pos[1]<=0):
            self.y=3
        if(pos[3]>=self.canvas_height):
            self.y=-3
        if(self.hit_paddle_lft(pos)==True):
            self.x=3
        if(self.hit_paddle_ryt(pos)==True):
            self.x=-3

    def hit_paddle_lft(self,pos):
        pdl_pos = self.canvas.coords(self.pdl_lft.id)
        if(pos[1]>=pdl_pos[1] and pos[1]<=pdl_pos[3]):
            if(pos[0]>=pdl_pos[0] and pos[0]<=pdl_pos[2]):
                return True
            return False
    
    def hit_paddle_ryt(self,pos):
        pdl_pos = self.canvas.coords(self.pdl_ryt.id)
        if(pos[1]>=pdl_pos[1] and pos[1]<=pdl_pos[3]):
            if(pos[2]>=pdl_pos[0] and pos[2]<=pdl_pos[2]):
                return True
            return False

    def score(self,val):
            global score_lft
            global score_ryt
            if(val==True):
                a=self.canvas.create_text(125,40,text=score_lft,font=('Arial',60),fill="white")
                self.canvas.itemconfig(a,fill="black")
                score_lft+=1
                a=self.canvas.create_text(125,40,text=score_lft,font = ('Arial',60),fill="white")

            if(val==False):
                a=self.canvas.create_text(375,40,text=score_ryt,font=('Arial',60),fill="white")
                self.canvas.itemconfig(a,fill="black")
                score_ryt+=1
                a=self.canvas.create_text(375,40,text=score_ryt,font = ('Arial',60),fill="white")

class Paddle:

    def __init__(self,canvas,color,x1,y1,x2,y2,key1,key2):
        self.canvas = canvas
        self.id = canvas.create_rectangle(x1,y1,x2,y2,fill=color)
        self.y=0
        self.canvas.height = self.canvas.winfo_height()
        self.canvas.width = self.canvas.winfo_width()
        self.canvas.bind_all(key1,self.pushtop)
        self.canvas.bind_all(key2,self.pushbottom)

    def draw(self):
        self.canvas.move(self.id,0,self.y)
        pos = self.canvas.coords(self.id)
        if(pos[1]<=0):
            self.y = 0
        if(pos[3]>=self.canvas.height):
            self.y = 0

    def pushtop(self,event):
        self.y=-2

    def pushbottom(self,event):
        self.y=2

paddle_left = Paddle(canvas,"blue",0,150,30,250,"a","d")
paddle_right = Paddle(canvas,"yellow",470,150,500,250,"<KeyPress-Left>","<KeyPress-Right>")
ball= Ball(canvas,"orange",paddle_left,paddle_right)

while 1:
    if(score_lft <10 and score_ryt<10):
        ball.draw()
        paddle_left.draw()
        paddle_right.draw()
    
    if(score_lft == 10):
        ball.x=235
        ball.y=200
        canvas.create_text(125,40,text="Player Yellow Won")
        
    if(score_ryt == 10):
        ball.x=235
        ball.y=200
        canvas.create_text(125,40,text="Player Blue Won")

    tk.update_idletasks()   
    tk.update()
    time.sleep(0.01)

