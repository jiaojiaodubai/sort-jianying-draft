"""
公共模块，提供了路径类PathX，管理.ini和获取路径的Initializer类，实例化各模块的抽象类Template。
"""
from abc import ABC, abstractmethod
from configparser import ConfigParser, NoSectionError, NoOptionError
from os import mkdir, listdir
from os.path import exists, isdir, expanduser, abspath, dirname
from shutil import rmtree
from threading import Thread
from tkinter import messagebox, BooleanVar, Misc, filedialog
from tkinter.ttk import Frame, Label, Combobox, Button, Checkbutton
from winreg import HKEY_CURRENT_USER

from win32com.client import Dispatch

from lib import get_key, names2name
from version import SetVersion


class PathX:
    """
    路径对象。

    Attributes:
        module: string，所属模块。
        name: string，路径在.ini中记录的名称。
        display: string，路径在GUI上显示的内容。
        content: list，所记录的路径内容。
    """
    module: str
    name: str
    display: str
    content: list[str]

    def __init__(self, module: str, name: str, display: str, content: list[str] = None):
        self.module = module
        self.name = name
        self.display = display
        if content is None:
            self.content = []
        else:
            self.content = content

    def __str__(self) -> str:
        return str(self.content)

    def now(self) -> str:
        return self.content[0]

    def add(self, c_i: str) -> None:
        self.content.insert(0, c_i)

    def clear(self) -> None:
        self.content.clear()

    def sort(self, as_length: bool = False) -> None:
        self.content = list(dict.fromkeys(self.content))
        if as_length:
            self.content.sort(key=lambda x: len(x))

    def load(self, parser: ConfigParser, destination: str = r'.\config.ini') -> None:
        if exists(destination):
            parser.read(destination, encoding='utf-8')
            self.content = parser[self.module][self.name].split(',')
        else:
            parser.add_section(self.module)
            parser.write(open(destination, 'w', encoding='utf-8'))

    def save(self, parser: ConfigParser, d: str = r'.\config.ini', update: bool = False):
        self.sort()
        parser.set(self.module, self.name, ','.join(self.content))
        if update:
            with open(d, 'w', encoding='utf-8') as f:
                parser.write(f)


class Initializer:
    install_path: PathX
    draft_path: PathX
    Data_path: PathX
    paths = []
    shells = Dispatch("WScript.Shell")
    configer = ConfigParser()
    version_exist = {'JianyingPro': False, 'CapCut': False}
    version_choose: str

    def __init__(self, first_time: bool = False):
        self.install_path = PathX('public', 'install_path', '安装路径')
        self.draft_path = PathX('public', 'draft_path', '草稿路径')
        self.Data_path = PathX('public', 'Data_path', 'Data路径')
        self.read_path()

    def batch_paths(self, appendix=None):
        if appendix is None:
            appendix = []
        self.paths = [self.install_path, self.draft_path, self.Data_path] + appendix

    def sync_paths(self):
        self.install_path = self.paths[0]
        self.draft_path = self.paths[1]
        self.Data_path = self.paths[2]

    def write_path(self, appendix=None):
        # 去重之后再写入ini文件
        self.batch_paths(appendix)
        for path in self.paths:
            path.sort()
            self.configer['public'][path.name] = ','.join(path.content)
        with open(r'.\config.ini', 'w', encoding='utf-8') as f:
            self.configer.write(f)

    def read_path(self) -> bool:
        """
        读取config.ini的默认配置，若存在则写入p.paths，若不存在则创建config.ini并增加相应节点。
        Returns:
            如果ini文件中的paths不少于3各则返回True，否则返回False。
        """
        self.batch_paths()
        # read对于文件不存在的情形不作处理，但不存在必定引起KeyError
        self.configer.read(r'.\config.ini', encoding='utf-8')
        try:
            for key in list(self.version_exist.keys()):
                self.version_exist[key] = self.configer['setting'][key]
            self.version_choose = self.configer['setting']['version_choose']
        except KeyError:
            self.reset_ini()
            return False
        for path in self.paths:
            try:
                path.content = (self.configer['public'][path.name]).split(',')
            except KeyError:
                self.reset_ini()
                return False
        self.sync_paths()
        return True

    def get_path(self, version: str):
        self.install_path.clear()
        self.draft_path.clear()
        self.Data_path.clear()
        self.install_path.add(get_key(HKEY_CURRENT_USER,
                                      fr'Software\Bytedance\{version}',
                                      'installDir').strip('\\')
                              )
        self.draft_path.add(get_key(HKEY_CURRENT_USER,
                                    fr'Software\Bytedance\{version}\GlobalSettings\History',
                                    'currentCustomDraftPath')
                            )
        # https://www.tutorialspoint.com/How-to-find-the-real-user-home-directory-using-Python
        self.Data_path.add(fr'{expanduser("~")}\AppData\Local\{version}\User Data')

    def reset_ini(self):
        for x in ['public', 'pack', 'unpack', 'ex_tittle', 'ex_voice', 'setting']:
            self.configer[x] = {}
        for key in list(self.version_exist.keys()):
            try:
                self.get_path(key)
                self.version_exist[key] = True
            except OSError:
                self.version_exist[key] = False
        for key in list(self.version_exist.keys()):
            self.configer['setting'][key] = str(self.version_exist[key])
        with open(r'.\config.ini', 'w', encoding='utf-8') as f:
            self.configer.write(f)
        if self.version_exist['JianyingPro'] or self.version_exist['CapCut']:
            if self.version_exist['JianyingPro'] and (not self.version_exist['CapCut']):
                self.version_choose = 'JianyingPro'
            elif (not self.version_exist['JianyingPro']) and self.version_exist['CapCut']:
                self.version_choose = 'CapCut'
            else:
                self.version_choose = SetVersion()
                # 已经包含异常处理
                self.get_path(self.version_choose)
            self.write_path()
            self.configer['setting']['version_choose'] = str(self.version_choose)
            with open(r'.\config.ini', 'w', encoding='utf-8') as f:
                self.configer.write(f)
        else:
            messagebox.showerror(title='获取路径失败',
                                 message='您未正确安装剪映，请尝试重新安装剪映再使用本程序。'
                                 )
            exit(1)

    def collect_draft(self):
        self.read_path()
        all_draft = []
        if not exists(r'.\draft-preview'):  # 判断文件夹是否存在
            mkdir(r'.\draft-preview')  # 不存在则新建文件夹
        else:
            # rmdir只能删除空文件夹，不空会报错，因此用shutil
            rmtree(r'.\draft-preview')
            mkdir(r'.\draft-preview')
        # 遍历所有名称满足条件草稿的路径
        for path in self.draft_path.content:
            for item in listdir(path):
                full_path = fr'{path}\{item}'
                if isdir(full_path):
                    shortcut = self.shells.CreateShortCut(fr'.\draft-preview\{item}.lnk')
                    shortcut.Targetpath = full_path
                    shortcut.save()
                    all_draft.append(full_path)
        return all_draft


class Template(Frame, ABC):
    """
    为了减少代码复用而为导出型的面板设计的模板类。

    Attributes:
        module_name: string，主按键在config.ini中的名字
        module_name_display: string，主按键在面板上显示的名字
        vals: tkinter.BoolVar，CheckButton所对应的Boolean变量
        vals_names: CheckButton在config.ini中的str名字
        checks_names: CheckButton在面板上展示的tsr名字

    """

    # TODO: 如果变量不需要在构建方法以外的函数被调用，则不需要设为成员变量
    # 模块级属性
    message: Label
    module_name: str
    module_name_display: str
    p: Initializer
    # 草稿行
    draft_todo = [tuple()]
    draft_comb: Combobox
    d_comb_name = '已选草稿：'
    # 目标行
    target_path: PathX
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

    def __init__(self, master: Misc, label: Label):
        super().__init__(master=master, width=560, height=155)
        # 模块
        self.message = label
        self.p = Initializer(first_time=True)
        # 部分组件创建时依赖预配置，因此先读取配置

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
        if self.p.configer.has_option(self.module_name, 'target_path'):
            self.target_path.add(self.p.configer.get(self.module_name, 'target_path'))
            self.target_comb.config(values=self.target_path.content)
            self.target_comb.current(0)

        # 复选行
        checks: list[Checkbutton] = []
        self.checks_frame = Frame(self, width=520, height=20)
        for i in range(len(self.vals_names)):
            self.vals.append(BooleanVar())
            # 读取选项框预置
            try:
                self.vals[i].set(eval(self.p.configer.get(self.module_name, self.vals_names[i])))
            except (NoSectionError, NoOptionError):
                pass
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
        self.main_fun_button = Button(self, text=f'一键{self.module_name_display}',
                                      command=lambda: Thread(target=self.main_fun).start(),
                                      width=20
                                      )
        self.main_fun_button.grid(row=3, column=1, columnspan=2, padx=5, pady=5)

        self.grid_columnconfigure(index=1, weight=1)
        self.grid_columnconfigure(index=2, weight=1)

    @abstractmethod
    def choose_export(self):
        select_temp = filedialog.askdirectory(parent=self,
                                              title='请选择导出位置',
                                              initialdir=self.target_path.now()
                                              )
        # 不要把工程导出到原工程路径，否则会引发困扰
        if isdir(select_temp) and select_temp not in self.p.draft_path.content:
            self.target_path.add(select_temp)
            self.target_comb.config(values=self.target_path.content)
            self.target_comb.current(0)
            self.message.config(text='导出路径选择完毕！')
            return True
        else:
            messagebox.showwarning(title='发生错误', message='请选择合适的文件夹！')
            return False

    @abstractmethod
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
                if abspath(dirname(link)) == abspath(r'.\draft-preview'):
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
            messagebox.showwarning(title='操作有误', message='未选中任何草稿！')
            return False

    def main_fun(self):
        # 防止选择拿空
        if self.draft_comb.get() != '':
            self.p.read_path()
            self.message.config(text=f'正在{self.module_name_display}...')
            # 导出预置项
            for i in range(len(self.checks_names)):
                self.p.configer.set(self.module_name, self.vals_names[i], str(self.vals[i].get()))
            self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
            self.patch_fun()
            self.message.config(text=f'{self.module_name_display}完毕！')
        else:
            messagebox.showwarning(title='路径无效', message='请检查路径是否存在！')

    @abstractmethod
    def patch_fun(self):
        pass
