import configparser
from os import path
from threading import Thread
from tkinter import IntVar, messagebox, filedialog
from tkinter.ttk import Frame, LabelFrame, Label, Radiobutton, Combobox, Button

from win32com.client import Dispatch

from public import names2name


class Setting(Frame):
    module_name_display = '设置'
    message: Label
    version_val: IntVar
    d_comb_name = '选择草稿'
    p = configparser.ConfigParser
    draft_todo = []
    shells = Dispatch("WScript.Shell")

    def __init__(self, master, label: Label):
        super().__init__(master, width=560, height=155)
        self.message = label
        self.version_val = IntVar()
        self.columnconfigure(index=0, weight=1)
        version_fr = LabelFrame(self, text='选择你的剪映发行版本', width=560, padding=5)
        version_fr.grid(row=0, column=0, sticky='EW', padx=5)
        version_fr.columnconfigure(index=0, weight=1)
        version_fr.columnconfigure(index=1, weight=1)
        version_bt_cn = Radiobutton(master=version_fr, text='中文版', variable=self.version_val, value=0)
        version_bt_cn.grid(row=0, column=0)
        version_bt_it = Radiobutton(master=version_fr, text='国际版', variable=self.version_val, value=1)
        version_bt_it.grid(row=0, column=1)
        version_fr.grid(row=0, column=0, sticky='EW', padx=5)

        ini_updating = LabelFrame(self, text='旧版草稿文件转换', width=560, padding=5)
        ini_updating.grid(row=1, column=0, sticky='EW', padx=5)
        # 草稿行
        draft_label = Label(ini_updating, text=self.d_comb_name)
        draft_label.grid(row=0, column=0, pady=10, padx=5)
        self.draft_comb = Combobox(ini_updating, width=52, state='readonly')
        self.draft_comb.grid(row=0, column=1, columnspan=2, pady=10, padx=5, sticky='EW')
        draft_choose = Button(ini_updating, text='选取草稿', command=self.choose_draft, width=10)
        draft_choose.grid(row=0, column=3, pady=5, padx=5)
        self.main_fun_button = Button(ini_updating, text='一键转换',
                                      command=lambda: Thread(target=self.main_fun()).start(),
                                      width=20
                                      )
        self.main_fun_button.grid(row=2, column=1, columnspan=2)

    def choose_draft(self):
        # 每次点开选择按钮都刷新草稿列表
        self.p.collect_draft()
        links_select = filedialog.askopenfilenames(parent=self,
                                                   title='请选择你想要导出的草稿',
                                                   initialdir=r'.\draft-preview',
                                                   filetypes=[('快捷方式', '*.lnk *link')]
                                                   )
        batch = []  # 一次可能选中多个
        if links_select != '':
            for link in links_select:
                # 仅在draft-preview内选取的文件有效
                if path.abspath(path.dirname(link)) == path.abspath(r'.\draft-preview'):
                    batch.insert(0, self.shells.CreateShortCut(link).Targetpath)
                else:
                    messagebox.showwarning(title='操作有误', message='请在默认打开的位置内选取！')
                    return False
            self.draft_todo.insert(0, tuple(batch))
            self.draft_comb.config(values=names2name(self.draft_todo))
            self.draft_comb.current(0)
            self.message.config(text='草稿选取完毕！')
            return True
        else:
            messagebox.showwarning(title='操作有误', message='未选中任何草稿')
            return False

    def main_fun(self):
        pass
