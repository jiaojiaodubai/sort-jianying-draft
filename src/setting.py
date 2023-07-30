import configparser
from os import walk
from os.path import dirname
from threading import Thread
from tkinter import IntVar, messagebox, filedialog, DISABLED, NORMAL
from tkinter.ttk import Frame, LabelFrame, Label, Radiobutton, Combobox, Button

from win32com.client import Dispatch

from lib import names2name, DESKTOP
from public import Initializer, PathX


class Setting(Frame):
    module_name_display = '设置'
    message: Label
    version_bt_cn: Radiobutton
    version_bt_it: Radiobutton
    version: IntVar
    d_comb_name = '选择草稿'
    source_path = PathX(module='setting', name='source_path', display='', content=[DESKTOP])
    p = Initializer()
    draft_todo = []
    shells = Dispatch("WScript.Shell")

    def __init__(self, master, label: Label):
        super().__init__(master, width=560, height=155)
        self.message = label
        self.version = IntVar()
        self.version.set(('JianyingPro', 'CapCut').index(self.p.version_choose))
        self.columnconfigure(index=0, weight=1)
        version_fr = LabelFrame(self, text='选择你的剪映发行版本', width=560, padding=5)
        version_fr.grid(row=0, column=0, sticky='EW', padx=5)
        version_fr.columnconfigure(index=0, weight=1)
        version_fr.columnconfigure(index=1, weight=1)
        self.version_bt_cn = Radiobutton(master=version_fr, text='中文版',
                                         variable=self.version, value=0, command=self.set_ver)
        self.version_bt_cn.grid(row=0, column=0)
        self.version_bt_it = Radiobutton(master=version_fr, text='国际版',
                                         variable=self.version, value=1, command=self.set_ver)
        self.version_bt_it.grid(row=0, column=1)
        version_fr.grid(row=0, column=0, sticky='EW', padx=5)
        self.version_bt_cn.config(state=DISABLED)
        self.version_bt_it.config(state=DISABLED)
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
                                      command=lambda: Thread(target=self.main_fun).start(),
                                      width=20
                                      )
        self.main_fun_button.grid(row=2, column=1, columnspan=2)
        # print(self.p.version_exist)
        if self.p.version_exist['JianyingPro'] and self.p.version_exist['CapCut']:
            self.version_bt_cn.config(state=NORMAL)
            self.version_bt_it.config(state=NORMAL)

    def choose_draft(self):
        dir_select = filedialog.askdirectory(parent=self,
                                             title='请选择已打包的草稿或其父文件夹',
                                             initialdir=self.source_path.now()
                                             )
        batch = []
        if dir_select != '':
            for super_path, sub_path, sub_files in walk(dir_select):
                # 别把原来就有的给导进来了
                if 'draft_content.json' in sub_files and super_path not in self.p.draft_path.content:
                    batch.insert(0, super_path)
                    self.source_path.add(dirname(dirname(super_path)))  # 更新导入路径
            if len(batch) != 0:
                self.draft_todo.insert(0, tuple(batch))
                self.draft_comb.config(values=names2name(self.draft_todo))
                self.draft_comb.current(0)
                self.message.config(text='草稿选择完毕！')
                return True
            # 这里的逻辑与template选择草稿不同，这里遍历完才能确定是否有可用的草稿
            else:
                messagebox.showwarning(title='操作有误', message='未找到草稿文件！')
                return False
        else:
            messagebox.showwarning(title='操作有误', message='未选中任何目录！')
            return False

    def main_fun(self):
        for draft in self.draft_todo[names2name(self.draft_todo).index(self.draft_comb.get())]:
            paser_old = configparser.ConfigParser()
            paser_new = configparser.ConfigParser()
            paser_new['public'] = {}
            paser_old.read(fr'{dirname(draft)}\config.ini', encoding='utf-8')
            for path in self.p.paths:
                paser_new['public'][path.name] = paser_old['paths'][path.name]
            paser_new['pack'] = {}
            paser_new['pack']['components_home'] = paser_old['paths']['meta_path_simp']
            with open(fr'{dirname(draft)}\config.ini', 'w', encoding='utf-8') as f:
                paser_new.write(f)
        self.message.config(text='草稿转换完毕！')

    def set_ver(self):
        # self.p.version_choose =
        version_choose = ('JianyingPro', 'CapCut')[self.version.get()]
        # 初始化时已经进行异常处理
        self.p.get_path(version=version_choose)
        self.p.write_path()
        self.p.configer['public']['version_choose'] = version_choose
        with open(r'.\config.ini', 'w', encoding='utf-8') as f:
            self.p.configer.write(f)
