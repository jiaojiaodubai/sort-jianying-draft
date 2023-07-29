from configparser import ConfigParser, NoOptionError
from os import mkdir, listdir
from os.path import exists, isdir, expanduser
from shutil import rmtree
from tkinter import messagebox
from winreg import HKEY_CURRENT_USER

from win32com.client import Dispatch

from lib import get_key
from version import SetVersion


class PathX:
    module: str
    name: str
    display: str
    content: list[str]

    def __init__(self, m: str, n: str, d: str, c: list[str] = None):
        self.module = m
        self.name = n
        self.display = d
        if c is None:
            self.content = []
        else:
            self.content = c

    def __str__(self):
        return str(self.content)

    def now(self):
        return self.content[0]

    def add(self, c_i: str):
        self.content.insert(0, c_i)

    def clear(self):
        self.content.clear()

    def sort(self, as_length: bool = False):
        self.content = list(dict.fromkeys(self.content))
        if as_length:
            self.content.sort(key=lambda x: len(x))

    def load(self, parser: ConfigParser, d: str = r'.\config.ini'):
        if exists(d):
            parser.read(d, encoding='utf-8')
            if parser.has_option(self.module, self.name):
                self.content = parser.get(self.module, self.name).split(',')
            else:
                raise NoOptionError
        else:
            parser.add_section(self.module)
            parser.write(open(d, 'w', encoding='utf-8'))

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

    def __init__(self):
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
        if exists(r'.\config.ini'):
            self.configer.read(r'.\config.ini', encoding='utf-8')
            try:
                self.version_choose = self.configer['public']['version_choose']
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
        else:
            self.reset_ini()
            return False

    def get_path(self, version: str):
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
        for x in ['public', 'pack', 'unpack', 'ex_tittle', 'ex_voice']:
            self.configer[x] = {}
        for key in list(self.version_exist.keys()):
            try:
                self.get_path(key)
                self.version_exist[key] = True
            except OSError:
                self.version_exist[key] = False
        if self.version_exist['JianyingPro'] or self.version_exist['CapCut']:
            if self.version_exist['JianyingPro'] and (not self.version_exist['CapCut']):
                self.version_choose = 'JianyingPro'
            elif (not self.version_exist['JianyingPro']) and self.version_exist['CapCut']:
                self.version_choose = 'CapCut'
            else:
                self.version_choose = SetVersion()
                self.install_path.clear()
                self.draft_path.clear()
                # https://www.tutorialspoint.com/How-to-find-the-real-user-home-directory-using-Python
                self.Data_path.clear()
                self.get_path(self.version_choose)
            self.write_path()
            self.configer['public']['version_choose'] = str(self.version_choose)
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
