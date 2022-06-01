import tkinter as tk
import tkinter.ttk as ttk


class MyProgram:

    def __init__(self):

        self.top2 = None
        self.root = root = tk.Tk()
        root.resizable(False, False)

        root.style = ttk.Style()
        # ('clam', 'alt', 'default', 'classic')
        root.style.theme_use("clam")

        menu = tk.Menu(root)
        root.config(menu=menu)

        fm = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Settings", menu=fm)
        fm.add_command(label="Preferences")

        hm = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="Help", menu=hm)
        hm.add_command(label="About", command=self.about_window)
        hm.add_command(label="Exit", command=root.quit)

    def about_window(self):
        self.top2.focus_set()


m = MyProgram()
