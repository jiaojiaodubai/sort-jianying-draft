from configparser import ConfigParser, NoOptionError
from os import walk, mkdir, listdir
from os.path import isdir, abspath, basename, dirname, exists
from pathlib import PurePath
from shutil import rmtree, copytree
from tkinter import filedialog, messagebox, BooleanVar
from tkinter.ttk import Checkbutton

from lib import DESKTOP, names2name, win32_shell_copy
from public import PathX, Template


class UnpackDraft(Template):
    # 模块级属性
    module_name = 'unpack'
    module_name_display = '导入草稿'

    # 草稿行
    source_path = PathX(module=module_name, name='source_path', display='草稿源', content=[DESKTOP])

    # 目标行
    t_comb_name = '素材路径：'
    target_path = PathX(module=module_name, name='target_path', display='素材目录')

    # 复选行
    checks: list[Checkbutton] = []
    checks_names = ['仅导入索引']
    vals: list[BooleanVar] = []
    vals_names = ['index_only']

    def __init__(self, parent, label):
        super().__init__(parent, label)
        self.target_path.add(fr'{self.p.Data_path.now()}\metas')
        try:
            mkdir(self.target_path.now())
        except FileExistsError:
            pass
        self.target_comb.config(values=self.target_path.content)
        self.target_comb.current(0)

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

    def choose_export(self):
        super().choose_export()

    def rewrite_json(self, draft_path: str) -> None:
        """
        对于待导入的草稿，将其中的json文件内记录的路径全部改写为新路径。
        Args:
            draft_path: 草稿json文件所在的文件夹
        Returns:
            None
        """
        jsons = ['draft_agency_config.json', 'draft_content.json', 'draft_meta_info.json']
        # 旧的素材目录名称需要自己读取、自己分割
        paths_old = [PathX(module='public', name='install_path', display='ipo'),
                     PathX(module='public', name='draft_path', display='dpo'),
                     PathX(module='public', name='Data_path', display='Dpo'),
                     PathX(module='pack', name='components_home', display='cho'),
                     ]
        parser = ConfigParser()
        for path in paths_old:
            try:
                path.load(parser, destination=fr'{abspath(dirname(draft_path))}\config.ini')
                path.sort(as_length=True)
            except NoOptionError:
                messagebox.showwarning(title='配置异常',
                                       message=f'草稿{basename(draft_path)}的配置文件有误\n请尝试在“设置”面板进行转换或重新导出。')
        paths_new = [self.p.install_path,
                     self.p.draft_path,
                     self.p.Data_path,
                     PathX(module='upack',
                           name='this_draft_components_home',
                           display='本草稿的素材目录',
                           content=[fr'{self.target_path.now()}\{basename(draft_path)}'])
                     ]
        # 遍历三个文件
        for json in jsons:
            with open(fr'{draft_path}\{json}', 'r', encoding='utf-8') as json_file:
                json_content = json_file.read()
                # 遍历四种路径，本次草稿位置、安装位置、全局草稿位置、缓存位置
                for path_old, path_new in zip(paths_old, paths_new):
                    # print(path_old, path_new)
                    # https://docs.python.org/zh-cn/3/library/pathlib.html#pathlib.PurePath.as_posix
                    json_content = json_content.replace(PurePath(path_old.now()).as_posix(),
                                                        PurePath(path_new.now()).as_posix())
            try:
                mkdir(fr'{draft_path}\temp')
            except FileExistsError:
                pass
            with open(fr'{draft_path}\temp\{json}', 'w', encoding='utf-8') as f:
                f.write(json_content)

    def patch_fun(self):
        # 因为素材路径可以自由选择，甚至手动录入，所以必须检查
        # os.path.isdir(self.export_comb.get())防止选择拿空，self.draft_comb.get() in self.todo_history防止不选直接按
        if isdir(self.target_comb.get()) and self.draft_comb.get() in names2name(self.draft_todo):
            self.message.config(text='正在导入...')
            self.source_path.save(parser=self.p.configer)
            self.target_path.add(self.target_comb.get())
            self.target_path.save(parser=self.p.configer, update=True)
            # 依据组合框的显示的值来确定操作哪一组草稿
            for draft in self.draft_todo[names2name(self.draft_todo).index(self.draft_comb.get())]:
                meta_path = fr'{abspath(dirname(draft))}\meta'
                self.rewrite_json(draft)
                try:
                    copy_done = win32_shell_copy(draft, fr'{self.p.draft_path.now()}')
                    copytree(fr'{draft}\temp', fr'{self.p.draft_path.now()}\{basename(draft)}', dirs_exist_ok=True)
                    rmtree(fr'{draft}\temp')
                    rmtree(fr'{self.p.draft_path.now()}\{basename(draft)}\temp')
                    if not copy_done:
                        messagebox.showwarning(title='遇到错误', message=f'复制草稿{basename(draft)}的过程中遇到错误。')
                except WindowsError:
                    messagebox.showwarning(title='遇到错误', message=f'复制草稿{basename(draft)}的过程中遇到错误。')
                if self.vals[0].get() != 1:
                    if exists(meta_path):
                        try:
                            mkdir(fr'{self.target_path.now()}\{basename(draft)}')
                        except FileExistsError:
                            pass
                        copy_done = win32_shell_copy([fr'{meta_path}\{x}' for x in listdir(meta_path)],
                                                     fr'{self.target_path.now()}\{basename(draft)}')
                        if not copy_done:
                            messagebox.showwarning(title='遇到错误',
                                                   message=f'复制草稿{basename(draft)}的过程中遇到错误。')

                    else:
                        messagebox.showinfo('缺少文件', fr'草稿{basename(draft)}没有媒体文件\n请在后续步骤中链接素材！')
            self.message.config(text='导入完成！')
        else:
            messagebox.showwarning(title='路径无效', message='请检查路径是否存在！')
