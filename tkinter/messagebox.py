from tkinter import *
import tkinter.messagebox
root = Tk()
tkinter.messagebox.showinfo("Windows Title","Hello there!...")
answer = tkinter.messagebox.askquestion("Question1","Are you human")
if answer == "yes":
    tkinter.messagebox.showinfo("Congrats","thank god it good to know another human is out there.")
else:
    tkinter.messagebox.showinfo("Alien","You are a bot")

root.mainloop()
