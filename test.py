from webbrowser import open as webopen
from tkinter import Tk, Text

# ...

root = Tk()  # 这里是思路，省略了一些代码
root.geometry('500x300')

text = Text(root)

text.tag_configure('link', foreground='blue', underline=True)
text.insert('end', '百度', 'link')
text.tag_bind('link', '<Button-1>', lambda event: webopen('www.baidu.com'))  # 被单击时调用浏览器打开网页
text.pack(fill='both')

root.mainloop()




