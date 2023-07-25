from json import load
from os import startfile, remove
from os.path import exists, basename, split, abspath
from shutil import make_archive, rmtree
from time import strftime, localtime
from tkinter import BooleanVar, messagebox
from tkinter.ttk import Checkbutton

from public import names2name, win32_shell_copy, DESKTOP, PathX
from src import template


class PackDraft(template.Template):
    # 模块级
    module_name = 'pack'
    module_name_display = '导出草稿'
    components = []
    components_home: PathX
    target_path = PathX(m=module_name, n='target_path', d='导出路径', c=[DESKTOP])

    # 目标行
    t_comb_name = '导出路径：'

    # 复选行
    checks: list[Checkbutton] = []
    checks_names = ['仅导出索引', '打包为单文件', '完成后打开目录', '清除原草稿']
    # 必须重载这两个变量，否则用父类的构建方法访问的只能是父类的类属性，从而造成子类间的变量混淆
    vals: list[BooleanVar] = []
    vals_names = ['index_only', 'is_zip', 'is_open', 'is_clean']

    def __init__(self, parent, label):
        super().__init__(parent, label)
        self.components_home = PathX(m=self.module_name, n='components_home', d='素材目录')
        self.target_comb.config(values=self.target_path.content)
        self.target_comb.current(0)

    def analyse_meta(self, draft_path: str):
        # 每次执行新项目就要清空上次的记录，否则记录会叠加
        self.components.clear()
        self.components_home.clear()
        with open(fr'{draft_path}\draft_meta_info.json', 'r', encoding='utf-8') as f:
            data = load(f)
            meta_temp = data.pop('draft_materials')
            meta_dic = meta_temp[0]
            meta_list = meta_dic.pop('value')
            for item in meta_list:
                path = item.pop('file_Path')
                if exists(path):
                    # 这些路径仅用于替换，因此顺序不重要，直接用了append
                    self.components.append(path)
                    # https://docs.python.org/zh-cn/3.11/library/os.path.html#os.path.split
                    self.components_home.add(split(path)[0])
        # 去重
        self.components_home.sort()
        self.p.write_path()
        # 必须及时关闭，否则f时刻被占用，不可替换内部资源
        f.close()

    def patch_fun(self):
        # 实现从下拉框直接选
        for draft in self.draft_todo[names2name(self.draft_todo).index(self.draft_comb.get())]:
            # 不能使用冒号，否则OSError: [WinError 123] 文件名、目录名或卷标语法不正确
            suffix = strftime('%m.%d.%H-%M-%S', localtime())
            filepath = fr'{self.target_path.now()}\{basename(draft)}-收集的草稿-{suffix}'
            self.analyse_meta(draft)
            try:
                win32_shell_copy(draft, fr'{filepath}\{basename(draft)}')
                win32_shell_copy(abspath(r'.\config.ini'), fr'{filepath}\config.ini')
            except WindowsError as e:
                messagebox.showerror(title=f'复制过程中发生错误{e}')
                exit(1)
            # 未选中“仅导出索引，就导出素材”
            if self.vals[0].get() != 1:
                # 复制的是文件，则复制后也应当是文件
                for path in self.components:
                    win32_shell_copy(path, fr'{filepath}\meta\{basename(path)}')
            if self.vals[1].get() == 1:
                self.message.config(text='正在压缩单文件...')
                make_archive(filepath, 'zip', filepath)
                rmtree(filepath)
                self.message.config(text='草稿文件创建成功！')
            if self.vals[3].get() == 1:
                rmtree(draft)
                for path in self.components:
                    remove(path)
            if self.vals[2].get() == 1:
                startfile(self.target_path.now())
                # 写入导出路径
            self.target_path.save(self.p.configer)
            self.components_home.save(self.p.configer, update=True)
        self.message.config(text='草稿打包完毕！')
