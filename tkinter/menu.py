from tkinter import *

root = Tk()

def random():
    print("sub menu")
main_menu = Menu(root)
root.configure(menu = main_menu)
sub_menu  = Menu(main_menu)
main_menu.add_cascade(label = "File",menu=sub_menu)
sub_menu.add_command(label = "Open File",command = random)
sub_menu.add_command(label = "Create File",command = random)
sub_menu.add_separator()
sub_menu.add_command(label = "Delete File",command = random)

root.mainloop()