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

    def select(self):
        """
        Select the position.
        """

        if WIN32API:
            win32api.ShellExecute(
                None, "open", "explorer.exe", '/select,{}'.format(self.path), '', win32con.SW_SHOW)
        else:
            raise ImportError()

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

    def walk_dir(self, *skip):
        for root, dirs, files in os.walk(self):
            if any([Path(root).is_relative_to(k) for k in skip]):
                continue
            for d in dirs:
                yield Path(os.path.join(root, d))

    def walk_file(self, *skip):
        for root, dirs, files in os.walk(self):
            if any([Path(root).is_relative_to(k) for k in skip]):
                continue
            for f in files:
                if '~$' in f:
                    continue
                else:
                    yield Path(os.path.join(root, f))

    def stat_dir(self):
        for p in self.walk_file():
            yield p.stat()

    def is_relative_to(self, path):
        try:
            rel = bool(self.relative_to(path))
        except ValueError as e:
            rel = False
        return rel


class PureWindowsPath(pathlib.PureWindowsPath):
    pass


class WindowsPath(Path, PureWindowsPath):
    pass
