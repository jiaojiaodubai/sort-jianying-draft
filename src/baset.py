from os import path
from threading import Thread
from tkinter import Frame, Label, Checkbutton, Button, filedialog, messagebox, BooleanVar
from tkinter.ttk import Combobox
# 在某些情况下，这个包会读取不成功，原因可能是虚拟环境没有配置好、路径没有加入环境变量
from win32com.client import Dispatch

from public import PathManager, names2name


class Template(Frame):
    """
    为了减少代码复用而为导出型的面板设计的模板类
    :parameter
        checks: CheckButton组成的list
        vals: CheckButton所对应的Boolean变量
        checks_names: CheckButton在config.ini中的str名字
        checks_names_display: CheckButton在面板上展示的tsr名字
        mould_name: 主按键在config.ini中的str名字
        mould_name_display: 主按键在面板上显示的str名字
    """
    message: Label
    p = PathManager()
    drafts_todo = []
    export_path = [p.DESKTOP, ]
    draft_comb: Combobox
    export_comb: Combobox
    # 这里不要把CheckButton实例化为空对象，否则它会带来一个空的根窗口
    checks: list[Checkbutton] = []
    vals: list[BooleanVar] = []
    checks_names: list[str]
    checks_names_display: list[str]
    mould_name: str
    mould_name_display: str
    box_frame: Frame
    export_button: Button
    shells = Dispatch("WScript.Shell")

    def __init__(self, parent, label: Label):
        super().__init__(parent, width=560, height=155)
        # if mode == 0:
        #     mode_name = '导出'
        # elif mode == 1:
        #     mode_name = '导入'
        # else:
        #     mode_name = '目标'
        draft_label = Label(self, text='已选草稿：')
        # export_label = Label(self, text=f'{mode_name}路径：')
        export_label = Label(self, text='导出路径：')
        draft_label.grid(row=0, column=0, pady=10, padx=5)
        export_label.grid(row=1, column=0, pady=10, padx=5)
        self.draft_comb = Combobox(self, width=52, state='readonly')
        self.draft_comb.grid(row=0, column=1, columnspan=2, pady=10, padx=5)
        self.export_comb = Combobox(self, width=52)
        self.export_comb.grid(row=1, column=1, columnspan=2, pady=10, padx=5)
        self.export_comb.config(values=self.export_path)
        self.export_comb.current(0)
        draft_choose = Button(self, text='选取草稿', command=self.choose_draft, width=10)
        draft_choose.grid(row=0, column=3, pady=5, padx=5)
        export_choose = Button(self, text='选择路径', command=self.choose_export, width=10)
        export_choose.grid(row=1, column=3, pady=5, padx=5)
        self.box_frame = Frame(self, width=520, height=20)
        # 启动guide的时候就已经检查过config.ini了，因此不再执行检查
        self.p.configer.read('config.ini', encoding='utf-8')
        for i in range(len(self.checks_names)):
            self.vals.append(BooleanVar())
            # 读取选项框预置
            if self.p.configer.has_section(f'{self.mould_name}_setting'):
                # 必须判断再get，否则没有配置项时将出错
                if self.p.configer.has_option(f'{self.mould_name}_setting', self.checks_names[i]):
                    self.vals[i].set(eval(self.p.configer.get(f'{self.mould_name}_setting', self.checks_names[i])))
            # 批量构造复选框并布局
            self.checks.append(Checkbutton(self.box_frame, text=self.checks_names_display[i], variable=self.vals[i]))
            self.checks[i].grid(row=0, column=i)
        # 读取导出预置
        if self.p.configer.has_option(f'{self.mould_name}_setting', 'export_path'):
            self.export_path.insert(0, self.p.configer.get(f'{self.mould_name}_setting', 'export_path'))
            self.export_comb.config(values=self.export_path)
            self.export_comb.current(0)
        self.box_frame.grid(row=2, column=0, columnspan=4)
        self.export_button = Button(self, text=f'一键{self.mould_name_display}', padx=40,
                                    command=lambda: Thread(target=self.export).start(),
                                    )
        self.export_button.grid(row=3, column=1, columnspan=2, padx=5, pady=5)
        self.message = label
        self.p.collect_draft()
        self.p.read_path()

    def choose_export(self):
        select_temp = filedialog.askdirectory(parent=self,
                                              title='请选择导出位置',
                                              initialdir=self.p.DESKTOP,
                                              )
        # 不要把工程导出到原工程路径，否则会引发困扰
        if path.isdir(select_temp) and select_temp not in self.p.paths[1]:
            self.export_path.insert(0, select_temp)
            self.export_comb.config(values=self.export_path)
            self.export_comb.current(0)
            self.message.config(text='导出路径选择完毕！')
        else:
            messagebox.showwarning(title='发生错误', message='请选择合适的文件夹！')

    def choose_draft(self):
        # 每次点开选择按钮都刷新草稿列表
        self.p.collect_draft()
        links_select = filedialog.askopenfilenames(parent=self,
                                                   title='请选择你想要导出的草稿',
                                                   initialdir=r'.\draft-preview',
                                                   filetypes=[('快捷方式', '*.lnk *link')]
                                                   )
        one_todo = []  # 一次可能选中多个
        for link in links_select:
            # 仅在draft-preview内选取的文件有效
            if path.abspath(path.dirname(link)) == path.abspath(r'.\draft-preview'):  # resolve返回绝对路径
                one_todo.insert(0, self.shells.CreateShortCut(link).Targetpath)
                self.drafts_todo.insert(0, tuple(one_todo))
                self.draft_comb.config(values=names2name(self.drafts_todo))
                self.draft_comb.current(0)
                self.message.config(text='草稿选取完毕！')
            else:
                messagebox.showwarning(title='操作有误', message='请在默认打开的位置内选取！')
                break

    def export(self):
        # isdir(self.export_comb.get())防止选择拿空，self.draft_comb.get() in self.todo_history防止不选直接按空
        if path.isdir(self.export_comb.get()) and self.draft_comb.get() in names2name(self.drafts_todo):
            self.p.read_path()
            self.message.config(text=f'正在{self.mould_name_display}...')
            # 写入设置项
            for i in range(len(self.checks)):
                self.p.configer.set(f'{self.mould_name}_setting', self.checks_names[i], str(self.vals[i].get()))
            self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
            self.main_fun()
            self.message.config(text=f'{self.mould_name_display}完毕！')
        else:
            messagebox.showwarning(title='路径无效', message='请检查路径是否存在！')

    def main_fun(self):
        pass
