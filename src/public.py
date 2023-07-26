from configparser import ConfigParser
from os import mkdir, listdir
from os.path import exists, isdir, abspath, basename, expanduser
from re import match
from shutil import rmtree
from tkinter import messagebox
from winreg import QueryValueEx, OpenKey, HKEY_CURRENT_USER

from past.types import basestring
from win32com.client import Dispatch
from win32comext.shell.shell import SHFileOperation
from win32comext.shell.shellcon import FO_COPY, FOF_NOCONFIRMMKDIR

# noinspection SpellCheckingInspection
img = '''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'''

rule = [
    [r'.*\\JianyingPro$', r'.*\\JianyingPro\\Apps$'],
    # [r'.*\\JianyingPro\\(\d|[.])*$', r'.*\\JianyingPro\\Apps\\(\d|[.])*$'],
    [r'.*\\JianyingPro Drafts$', r'.*\\com[.]lveditor[.]draft$', ],
    [r'.*\\JianyingPro\\User Data$', ],
]


def get_key(main_key: int, sub_key: str, value_name: str):
    value = QueryValueEx(OpenKey(main_key, sub_key),
                         value_name
                         )[0]
    return value


DESKTOP = get_key(HKEY_CURRENT_USER,
                  r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',
                  "Desktop"
                  )


# TODO：有望用更轻量的模块来实现
#  https://pypi.org/project/ifileoperation/
#  https://learn.microsoft.com/zh-cn/windows/win32/api/shobjidl_core/nf-shobjidl_core-ifileoperation-copyitems
def win32_shell_copy(src, dest) -> bool:
    """
    使用Windows Shell复制目录或文件。
    https://docs.microsoft.com/zh-cn/windows/win32/api/shellapi/ns-shellapi-shfileopstructa?redirectedfrom=MSDN
    https://stackoverflow.com/questions/16867615/copy-using-the-windows-copy-dialog
    https://stackoverflow.com/questions/5768403/shfileoperation-doesnt-move-all-of-a-folders-contentsh
    https://stackoverflow.com/questions/4611237/double-null-terminated-string
    Args:
        src: Path or a list of paths to copy.
            Filename portion of a path (but not directory portion) can contain wildcards "*" and "?".
        dest: destination directory.
    Returns:
        True: if the operation completed successfully,
        False: if it was aborted by user (completed partially).
    Raises:
        WindowsError: if anything went wrong. Typically, when source file was not found.
    """
    if isinstance(src, basestring):  # in Py3 replace basestring with str
        src = abspath(src)
    else:  # iterable
        src = '\0'.join(abspath(path) for path in src)

    result, aborted = SHFileOperation((
        0,
        FO_COPY,
        src,
        abspath(dest),
        FOF_NOCONFIRMMKDIR,  # flags
        None,
        None))
    if not aborted and result != 0:
        # Note: raising a WindowsError with correct error code is quite
        # difficult due to SHFileOperation historical idiosyncrasies.
        # Therefore, we simply pass a message.
        raise WindowsError('SHFileOperation failed: 0x%08x' % result)

    return not aborted


def names2name(names: list):
    name = []
    for group in names:
        # 每一组内都要将temp重置，否则会保留上一个group的记录
        temp_group = []
        for one_name in group:
            temp_group.append(basename(one_name))  # append确保新生成的名称表的顺序与原来一致
        name.append(','.join(temp_group))
    return name


def is_match(path_str: str, position: int):
    is_find = False
    for key in rule[position]:
        if match(key, path_str):
            is_find = True
    return is_find


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
            self.configer.set('public', path.name, ','.join(path.content))
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
            for i in range(3):
                if self.configer.has_option('public', self.paths[i].name):
                    self.paths[i].content = self.configer.get('public', self.paths[i].name).split(',')
                else:
                    self.reset_ini()
                    return False
            self.sync_paths()
            return True
        else:
            self.reset_ini()
            return False

    def reset_ini(self):
        self.configer.add_section('public')
        try:
            self.install_path.add(get_key(HKEY_CURRENT_USER,
                                          r'Software\Bytedance\JianyingPro',
                                          'installDir').strip('\\')
                                  )
            self.draft_path.add(get_key(HKEY_CURRENT_USER,
                                        r'Software\Bytedance\JianyingPro\GlobalSettings\History',
                                        'currentCustomDraftPath')
                                )
            # https://www.tutorialspoint.com/How-to-find-the-real-user-home-directory-using-Python
            self.Data_path.add(fr'{expanduser("~")}\AppData\Local\JianyingPro\User Data')
            self.write_path()
        except OSError:
            messagebox.showerror(title='获取路径失败',
                                 message='您可能未正确安装剪映，请尝试重新安装剪映'
                                 )
            exit(1)
        self.configer.add_section('pack')
        self.configer.add_section('unpack')
        self.configer.add_section('ex_tittle')
        self.configer.add_section('ex_voice')
        self.configer.write(open(r'.\config.ini', 'w', encoding='utf-8'))

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
            ls = listdir(path)
            for item in ls:
                full_path = fr'{path}\{item}'
                if isdir(full_path):
                    shortcut = self.shells.CreateShortCut(fr'.\draft-preview\{item}.lnk')
                    shortcut.Targetpath = full_path
                    shortcut.save()
                    all_draft.append(full_path)
        return all_draft
