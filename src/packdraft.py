from json import load
from os import startfile, remove
from os.path import exists, basename, split, abspath
from shutil import make_archive, rmtree
from time import strftime, localtime
from tkinter import Checkbutton, BooleanVar

import template
from public import names2name, win32_shell_copy


class PackDraft(template.Template):
    # 模块级
    mould_name = 'pack'
    mould_name_display = '导出草稿'

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
        self.target_path = [self.p.DESKTOP, ]
        self.target_comb.config(values=self.target_path)
        self.target_comb.current(0)

    def analyse_meta(self, draft_path: str):
        # 每次执行新项目就要清空上次的记录，否则记录会叠加
        self.p.paths[3].clear()
        self.p.paths[4].clear()
        with open(fr'{draft_path}\draft_meta_info.json', 'r', encoding='utf-8') as f:
            # TODO：路径统一写为生字符串
            data = load(f)
            meta_temp = data.pop('draft_materials')
            meta_dic = meta_temp[0]
            meta_list = meta_dic.pop('value')
            for item in meta_list:
                path = item.pop('file_Path')
                if exists(path):
                    # 这些路径仅用于替换，因此顺序不重要，直接用了append
                    self.p.paths[3].append(path)
                    self.p.paths[4].append(split(path)[0])
        # 去重
        self.p.paths[4] = list(set(self.p.paths[4]))
        self.p.write_path()
        # 必须及时关闭，否则f时刻被占用，不可替换内部资源
        f.close()

    def patch_fun(self):
        # 写入导出路径
        self.p.configer.set('pack_setting', 'target_path', ','.join(self.target_path))
        self.p.configer.write(open('config.ini', 'w', encoding='utf-8'))
        for draft in self.drafts_todo[names2name(self.drafts_todo).index(self.draft_comb.get())]:
            # 不能使用冒号，否则OSError: [WinError 123] 文件名、目录名或卷标语法不正确
            suffix = strftime('%m.%d.%H-%M-%S', localtime())
            filepath = '{}\\{}-收集的草稿-{}'.format(self.target_path[0], basename(draft), suffix)
            self.analyse_meta(draft)
            win32_shell_copy(draft, '{}\\{}'.format(filepath, basename(draft)))
            win32_shell_copy(abspath(r'.\config.ini'), '{}\\config.ini'.format(filepath))
            # 未选中“仅导出索引，就导出素材”
            if self.vals[0].get() != 1:
                # 复制的是文件，则复制后也应当是文件
                for path in self.p.paths[3]:
                    win32_shell_copy(path, '{}\\meta\\{}'.format(filepath, basename(path)))
            if self.vals[1].get() == 1:
                self.message.config(text='正在压缩单文件...')
                make_archive(filepath, 'zip', filepath)
                rmtree(filepath)
                self.message.config(text='草稿文件创建成功！')
            if self.vals[3].get() == 1:
                rmtree(draft)
                for path in self.p.paths[3]:
                    remove(path)
            if self.vals[2].get() == 1:
                startfile(self.target_path[0])
        self.message.config(text='草稿打包完毕！')
