from os import path
from threading import Thread
from tkinter import filedialog, messagebox, BooleanVar
from tkinter.ttk import Frame, Label, Button, Checkbutton, Combobox

# 在某些情况下，这个包会读取不成功，原因可能是虚拟环境没有配置好、路径没有加入环境变量
from win32com.client import Dispatch

from public import PathManager, names2name


class Template(Frame):
    """
    为了减少代码复用而为导出型的面板设计的模板类
    :parameter
        drafts_todo: 需要处理的草稿
        checks: CheckButton组成的list
        vals: CheckButton所对应的Boolean变量
        vals_names: CheckButton在config.ini中的str名字
        checks_names: CheckButton在面板上展示的tsr名字
        mould_name: 主按键在config.ini中的str名字
        mould_name_display: 主按键在面板上显示的str名字
    """
    # TODO: 如果变量不需要在构建方法以外的函数被调用，则不需要设为成员变量
    # 模块级属性
    message: Label
    mould_name: str
    mould_name_display: str
    # 草稿行
    p = PathManager()
    drafts_todo = [tuple()]
    draft_comb: Combobox
    d_comb_name = '已选草稿：'
    # 目标行
    target_path = []
    target_comb: Combobox
    t_comb_name = '目标路径：'
    # 复选行
    vals: list[BooleanVar] = []
    vals_names: list[str]
    checks_names: list[str]
    checks_frame: Frame
    # 按键行
    main_fun_button: Button
    shells = Dispatch("WScript.Shell")

    def __init__(self, parent, label: Label):
        super().__init__(parent, width=560, height=155)
        # 模块
        self.message = label
        # 部分组件创建时依赖预配置，因此先读取配置
        self.p.configer.read('config.ini', encoding='utf-8')  # 启动guide的时候就已经检查过config.ini了，因此不再执行检查
        self.p.collect_draft()
        self.p.read_path()

        # 草稿行
        draft_label = Label(self, text=self.d_comb_name)
        draft_label.grid(row=0, column=0, pady=10, padx=5)
        self.draft_comb = Combobox(self, width=52, state='readonly')
        self.draft_comb.grid(row=0, column=1, columnspan=2, pady=10, padx=5, sticky='EW')
        draft_choose = Button(self, text='选取草稿', command=self.choose_draft, width=10)
        draft_choose.grid(row=0, column=3, pady=5, padx=5)

        # 目标行
        target_label = Label(self, text=self.t_comb_name)
        target_label.grid(row=1, column=0, pady=10, padx=5)
        self.target_comb = Combobox(self, width=52)
        self.target_comb.grid(row=1, column=1, columnspan=2, pady=10, padx=5, sticky='EW')
        target_choose = Button(self, text='选择路径', command=self.choose_export, width=10)
        target_choose.grid(row=1, column=3, pady=5, padx=5)
        # 载入预记录的目标路径
        if self.p.configer.has_option(f'{self.mould_name}_setting', 'target_path'):
            self.target_path.insert(0, self.p.configer.get(f'{self.mould_name}_setting', 'target_path'))
            self.target_comb.config(values=self.target_path)
            self.target_comb.current(0)

        # 复选行
        checks: list[Checkbutton] = []
        self.checks_frame = Frame(self, width=520, height=20)
        for i in range(len(self.vals_names)):
            self.vals.append(BooleanVar())
            # 读取选项框预置
            if self.p.configer.has_section(f'{self.mould_name}_setting'):
                # 必须判断再get，否则没有配置项时将出错
                if self.p.configer.has_option(f'{self.mould_name}_setting', self.vals_names[i]):
                    self.vals[i].set(eval(self.p.configer.get(f'{self.mould_name}_setting', self.vals_names[i])))
            # 批量构造复选框并布局
            checks.append(Checkbutton(self.checks_frame, text=self.checks_names[i], variable=self.vals[i]))
            # 成员变量导致窗口定位失败
            # _tkinter.TclError: bad window path name
            # https://blog.csdn.net/xhw2021/article/details/121342515
            # https://stackoverflow.com/questions/29233029/python-tkinter-show-only-one-copy-of-window?noredirect=1
            # https://stackoverflow.com/questions/111155/how-do-i-handle-the-window-close-event-in-tkinter
            # https://stackoverflow.com/questions/3295270/overriding-tkinter-x-button-control-the-button-that-close-the-window
            checks[i].grid(row=0, column=i)
        self.checks_frame.grid(row=2, column=0, columnspan=4)

        # 按钮行
        self.main_fun_button = Button(self, text=f'一键{self.mould_name_display}',
                                      command=lambda: Thread(target=self.main_fun()).start(),
                                      width=20
                                      )
        self.main_fun_button.grid(row=3, column=1, columnspan=2, padx=5, pady=5)

        self.grid_columnconfigure(index=1, weight=1)
        self.grid_columnconfigure(index=2, weight=1)

    def choose_export(self):
        select_temp = filedialog.askdirectory(parent=self,
                                              title='请选择导出位置',
                                              initialdir=self.p.DESKTOP,
                                              )
        # 不要把工程导出到原工程路径，否则会引发困扰
        if path.isdir(select_temp) and select_temp not in self.p.paths[1]:
            self.target_path.insert(0, select_temp)
            self.target_comb.config(values=self.target_path)
            self.target_comb.current(0)
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
        batch = []  # 一次可能选中多个
        for link in links_select:
            # 仅在draft-preview内选取的文件有效
            if path.abspath(path.dirname(link)) == path.abspath(r'.\draft-preview'):
                batch.insert(0, self.shells.CreateShortCut(link).Targetpath)
                self.drafts_todo.insert(0, tuple(batch))
                self.draft_comb.config(values=names2name(self.drafts_todo))
                self.draft_comb.current(0)
                self.message.config(text='草稿选取完毕！')
            else:
                messagebox.showwarning(title='操作有误', message='请在默认打开的位置内选取！')
                break

    def main_fun(self):
        # 防止选择拿空
        if self.draft_comb.get() != '':
            self.p.read_path()
            self.message.config(text=f'正在{self.mould_name_display}...')

            # 导出预置项
            for i in range(len(self.checks_names)):
                self.p.configer.set(f'{self.mould_name}_setting', self.vals_names[i], str(self.vals[i].get()))
            self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))

            self.patch_fun()
            self.message.config(text=f'{self.mould_name_display}完毕！')
        else:
            messagebox.showwarning(title='路径无效', message='请检查路径是否存在！')

    def patch_fun(self):
        pass
