import pathlib
import hashlib
import os
import shutil
try:
    import win32api
    import win32con
    WIN32API = True
except ImportError:
    WIN32API = False


class Path(pathlib.Path):
    def __new__(cls, *args, **kwargs):
        if cls is Path:
            cls = WindowsPath if os.name == 'nt' else pathlib.PosixPath

        self = cls._from_parts(args, init=False)
        if not self._flavour.is_supported:
            raise NotImplementedError("cannot instantiate %r on your system"
                                      % (cls.__name__,))
        self._init()
        return self

    if WIN32API:
        def select(self):
            """
            Select the position.
            """
            win32api.ShellExecute(
                None, "open", "explorer.exe", '/select,{}'.format(self.path), '', win32con.SW_SHOW)

    def execute(self):
        """
        Run by system.
        """
        os.startfile(self)

    def md5(self):
        """
        Get file MD5 value.
        """
        with self.open("rb") as f:
            return hashlib.new("md5", f.read())

    def rmdir(self, recursive: bool):
        """
        Remove Directory empty or recursive.
        """
        if recursive:
            shutil.rmtree(self)
        else:
            super().rmdir()

    def empty(self):
        """
        Check folder is empty.
        """
        count = 0
        for _ in self.iterdir():
            count += 1
        return count == 0

    def copy_to(self, dest):
        if self.is_file():
            shutil.copy2(self, dest)
        elif self.is_dir():
            shutil.copytree(self, dest)

    def walk_dir(self):
        for root, dirs, files in os.walk(self):
            for d in dirs:
                yield Path(os.path.join(root, d))
            
    def walk_file(self):
        for root, dirs, files in os.walk(self):
            for f in files:
                if '~$' in f:
                    continue
                else:
                    yield Path(os.path.join(root, f))

    def info(self):
        if self.is_file():
            yield self.stat()
        elif self.is_dir():
            for p in self.walk_file():
                yield p.stat()



class PureWindowsPath(pathlib.PureWindowsPath):
    pass


class WindowsPath(Path, PureWindowsPath):
    pass

