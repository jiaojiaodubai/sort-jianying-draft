from tkinter import DISABLED, scrolledtext
from tkinter.ttk import Frame, Label
from webbrowser import open as wb_open


class Help(Frame):
    def __init__(self, parent):
        super().__init__(parent, width=560, height=155)
        txt = scrolledtext.ScrolledText(self, width=54, height=7, padx=5, pady=5)

        txt.tag_configure('about author', justify='center', foreground='Violet')  # 设置tag属性
        txt.insert(1.0, '{:*^40}'.format('与我联系') + '\n')  # 插入第一行，第0个字符，行从1开始算，字符从0开始算
        txt.tag_add('about author', '1.0', '1.end')  # 将名为“about author”的tag的属性应用到第1行第0个字符直到该行结束

        txt.insert(2.0, 'CSDN：\n')
        txt.window_create('2.5',
                          window=LinkLabel(txt,
                                           link='https://blog.csdn.net/qq_43803536',
                                           text='jiaojiaodubai'))
        txt.insert(3.0, '博客园：\n')
        txt.window_create('3.4',
                          window=LinkLabel(txt,
                                           link='https://home.cnblogs.com/u/jiaojiaodubai',
                                           text='jiaojiaodubai23'))  # window_create用于插入其他控件

        txt.tag_configure('update', justify='center', foreground='Violet')
        txt.insert(4.0, '{:*^40}'.format('发布地址') + '\n')
        txt.tag_add('update', '4.0', '4.end')

        txt.insert(5.0, 'github：\n')
        txt.window_create('5.7',
                          window=LinkLabel(txt,
                                           link='',
                                           text='https://github.com/jiaojiaodubai/sort-jianying-draft'
                                           ))
        txt.insert(6.0, 'gitee：\n')
        txt.window_create('6.end',
                          window=LinkLabel(txt,
                                           link='',
                                           text='https://gitee.com/jiaojiaodubai/sort-jianying-draft'
                                           ))
        txt.tag_configure('note', justify='center', foreground='Violet')  # 设置tag属性
        txt.insert(7.0, '{:*^40}'.format('注意事项') + '\n')  # 插入第一行，第0个字符，行从1开始算，字符从0开始算
        txt.tag_add('note', '7.0', '7.end')
        txt.insert(8.0, '一山不容二虎，使用时一定不要打开剪映~\n\n')

        txt.config(state=DISABLED, font=('微软雅黑', 12))
        txt.pack(expand=True, fill='x')


class LinkLabel(Label):
    # https://blog.csdn.net/tinga_kilin/article/details/107121628
    link: str

    # LinkLabel可以显示超链接
    def __init__(self, master, link, text):
        # https://blog.csdn.net/yanyunfei0921/article/details/119256929
        super().__init__(master, text=text, background='#FFFFFF', font=('微软雅黑', 12))
        self.link = link
        self.bind('<Enter>', self.change_color)
        self.bind('<Leave>', self.change_cursor)
        self.bind('<Button-1>', self.go_link)
        self.is_click = False  # 未被点击

    def change_color(self, event):
        self['foreground'] = '#D52BC4'  # 鼠标进入，改变为紫色
        self['cursor'] = 'hand2'

    def change_cursor(self, event):
        if not self.is_click:  # 如果链接未被点击，显示会蓝色
            self['foreground'] = 'blue'
        self['cursor'] = 'xterm'

    def go_link(self, event):
        self.is_click = True  # 被链接点击后不再改变颜色
        wb_open(self.link)
