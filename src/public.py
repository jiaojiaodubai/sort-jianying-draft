import configparser
import os
import shutil
import winreg
import win32com.client
from past.types import basestring
from win32comext.shell import shellcon, shell

# noinspection SpellCheckingInspection
img = '''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
/AiEyk2L6pbvLAAAAAElFTkSuQmCC '''


class PathManager:
    path_names = ['draft_path', 'cache_path', 'equip_path', 'meta_path', 'meta_path_simp']
    draft_path = []
    cache_path = []
    equip_path = []
    meta_path = []
    meta_path_simp = []
    paths = [draft_path, cache_path, equip_path, meta_path, meta_path_simp]
    shells = win32com.client.Dispatch("WScript.Shell")

    DESKTOP = winreg.QueryValueEx(winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                                 r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'),
                                  "Desktop"
                                  )[0]

    def __init__(self):
        self.read_path()

    configer = configparser.ConfigParser()

    def write_path(self):
        for i in range(5):
            self.configer.set('paths', self.path_names[i], ','.join(self.paths[i]))
        with open('config.ini', 'w', encoding='utf-8') as f:
            self.configer.write(f)

    def read_path(self):
        if os.path.exists('config.ini'):
            self.configer.read('config.ini', encoding='utf-8')
            for i in range(4):
                if self.configer.has_option('paths', self.path_names[i]):
                    self.paths[i] = self.configer.get('paths', self.path_names[i]).split(',')
                elif i < 3:
                    return False
            return True
        else:
            self.configer.add_section('paths')
            self.configer.add_section('unpack_setting')
            self.configer.add_section('pack_setting')
            self.configer.write(open('config.ini', 'w', encoding='utf-8'))
            return False

    def collect_draft(self):
        self.read_path()
        all_draft = []
        if not os.path.exists('../draft-preview'):  # 判断文件夹是否存在
            os.mkdir('../draft-preview')  # 不存在则新建文件夹
        else:
            # os.rmdir只能删除空文件夹，不空会报错，因此用shutil
            shutil.rmtree('../draft-preview')
            os.mkdir('../draft-preview')
        for path in self.paths[0]:
            ls = os.listdir(path)
            for item in ls:
                full_path = r'{}\{}'.format(path, item)
                if os.path.isdir(full_path):
                    # 以下代码会造成程序体积飞涨
                    # icon = Image.open(r'{}\draft_cover.jpg'.format(full_path))
                    # icon = icon.crop(cal_corner(icon.size[0], icon.size[1]))
                    # icon.save(r'{}\draft_cover.ico'.format(full_path), sizes=[(250, 250)])
                    # icon.save(r'{}\draft_cover.png'.format(full_path), sizes=[(250, 250)])
                    # winshell.CreateShortcut(
                    #     # 必须使用绝对路径
                    #     Path=os.path.join(os.path.abspath('draft-preview'), '{}.lnk'.format(item)),
                    #     Target=full_path,
                    #     Icon=('{}\\draft_cover.ico'.format(full_path), 0),
                    #     Description='单击选中该草稿'
                    # )
                    shortcut = self.shells.CreateShortCut(r'draft-preview\{}.lnk'.format(item))
                    shortcut.Targetpath = full_path
                    shortcut.save()
                    all_draft.append(full_path)
        return all_draft


def win32_shell_copy(src, dest):
    """
    Copy files and directories using Windows shell.

    :param src: Path or a list of paths to copy. Filename portion of a path
       (but not directory portion) can contain wildcards ``*`` and
       ``?``.
    :param dest: destination directory.
    :returns: ``True`` if the operation completed successfully,
       ``False`` if it was aborted by user (completed partially).
    :raises: ``WindowsError`` if anything went wrong. Typically, when source
      file was not found.

    .. see also:
     `SHFileOperation on MSDN <https://msdn.microsoft.com/en-us/library/windows/desktop/bb762164(v=vs.85).aspx>`
    """
    if isinstance(src, basestring):  # in Py3 replace basestring with str
        src = os.path.abspath(src)
    else:  # iterable
        src = '\0'.join(os.path.abspath(path) for path in src)

    result, aborted = shell.SHFileOperation((
        0,
        shellcon.FO_COPY,
        src,
        os.path.abspath(dest),
        shellcon.FOF_NOCONFIRMMKDIR,  # flags
        None,
        None))
    if not aborted and result != 0:
        # Note: raising a WindowsError with correct error code is quite
        # difficult due to SHFileOperation historical idiosyncrasies.
        # Therefore, we simply pass a message.
        raise WindowsError('SHFileOperation failed: 0x%08x' % result)

    return not aborted


# 以下方法只能创建文件夹，不能复制内部的内容，原因有待考究
# noinspection SpellCheckingInspection
# https: // docs.microsoft.com / zh - cn / windows / win32 / api / shellapi / ns - shellapi - shfileopstructa?redirectedfrom = MSDN
#
# https: // stackoverflow.com / questions / 16867615 / copy - using - the - windows - copy - dialog
#
# https: // stackoverflow.com / questions / 5768403 / shfileoperation - doesnt - move - all - of - a - folders - contents
# https: // stackoverflow.com / questions / 4611237 / double - null - terminated - string
# shell.SHFileOperation((0,
#                        shellcon.FO_COPY,
#                        '{}\0'.format(draft),
#                        '{}\\{}\0'.format(filepath, os.path.basename(draft)),
#                        shellcon.FOF_NOCONFIRMATION | shellcon.FOF_NOERRORUI,
#                        None,
#                        None))

def names2name(names: list):
    name = []
    for group in names:
        # 每一组内都要将temp重置，否则会保留上一个group的记录
        temp_group = []
        for one_name in group:
            temp_group.append(os.path.basename(one_name))  # append确保新生成的名称表的顺序与原来一致
        name.append(','.join(temp_group))
    return name
