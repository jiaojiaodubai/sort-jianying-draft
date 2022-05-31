from src import help
import tkinter as tk

t = tk.Tk()
m = help.Help(t)
m.pack(side=tk.LEFT)
t.geometry('520x150')
t.mainloop()
