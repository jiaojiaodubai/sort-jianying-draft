import tkinter as tk
root = tk.Tk()
T1 = tk.Text(root)
T1.tag_configure("center", justify='center')
T1.insert(1.0, "啦啦啦")
T1.tag_add("center", "1.0", "end")
T1.pack()
root.mainloop()
