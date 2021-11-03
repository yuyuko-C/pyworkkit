import win32api
import win32gui
import win32con

from on_script.simulatekm import SimulateKM
from typing import Union


def active(method):
    def inner(self: Window, *args, **kwargs):
        if self.forground_only and not self.is_forground():
            raise ValueError('window is not active')
        else:
            return method(self, *args, **kwargs)
    return inner


class BaseWindow:
    """
    docstring
    """

    def __init__(self, hwnd):
        self.hwnd = hwnd
        self._update_geometry()
        self.sleep_ms = 100

    def hide(self):
        win32gui.ShowWindow(self.hwnd, False)

    def show(self):
        win32gui.ShowWindow(self.hwnd, True)
        self._update_geometry()

    def is_forground(self):
        return self.hwnd == win32gui.GetForegroundWindow()

    def foreground_window(self):
        self.show()
        win32gui.SetForegroundWindow(self.hwnd)

    def _update_geometry(self):
        x, y, rx, ry = win32gui.GetWindowRect(self.hwnd)
        self.x, self.y = x, y
        self.w, self.h = rx-x, ry-y

    def _set_geometry(self, lx, ly, rx, ry):
        win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOP,
                              lx, ly, rx, ry, win32con.SWP_SHOWWINDOW)
        self._update_geometry()

    def get_size(self):
        self._update_geometry()
        return (self.w, self.h)

    def get_position(self):
        self._update_geometry()
        return (self.x, self.y)

    def get_geometry(self):
        return (self.x, self.y, self.x+self.w, self.y+self.h)

    def set_size(self, w, h):
        self._set_geometry(self.x, self.y,
                           self.x+w, self.y+h)

    def set_position(self, x, y):
        self._set_geometry(x, y, self.w, self.h)

    def _in_range(self, x, y):
        if any([x < self.x, y < self.y,
                x > self.w+self.x, y > self.h+self.y]):
            return False
        return True

    def global_to_local(self, point: tuple):
        if isinstance(point, tuple):
            return (point[0]-self.x, point[1]-self.y)
        else:
            raise ValueError('point is NoneType')

    def local_to_global(self, point: tuple):
        if isinstance(point, tuple):
            return (point[0]+self.x, point[1]+self.y)
        else:
            raise ValueError('point is NoneType')


class Window(BaseWindow, SimulateKM):
    """
    docstring
    """

    def __init__(self, hwnd):
        super().__init__(hwnd)
        self.forground_only = True

    def __range_check(self, x, y):
        if not self._in_range(x, y):
            raise ValueError('point is out of window')

    @active
    def move_to(self, x, y):
        self.__range_check((x, y))
        return super().move_to(x, y)

    @active
    def click(self, x, y):
        self.__range_check((x, y))
        return super().click(x, y)

    @active
    def scroll(self, clicks: Union[int, float]):
        return super().scroll(clicks)

    @active
    def get_color(self, x, y):
        self.__range_check(x, y)
        return super().get_color(x, y)

    @active
    def locate_image_center(self, img_path: str, region=None, confidence=0.9):
        return super().locate_image_center(img_path, region=region, confidence=confidence)

    @active
    def cursor_info(self):
        self.foreground_window()
        g_point = win32gui.GetCursorPos()
        self.__range_check(g_point[0], g_point[1])
        lx, ly = self.global_to_local(g_point)
        print('local pos:', (lx, ly))
        print('color:', self.get_color(g_point))

    def clear(self):
        """
        清空
        """
        self.hotkey('ctrl', 'a')
        self.press('backspace')

    def select_all(self):
        """
        全选
        """
        self.hotkey('ctrl', 'a')

    def copycls(self):
        """
        复制
        """
        self.hotkey('ctrl', 'c')

    def stick(self):
        """
        粘贴
        """
        self.hotkey('ctrl', 'v')

    pass


def findwindow(window_name, w_classname=None):
    hwnd = win32gui.FindWindow(w_classname, window_name)
    return hwnd


def findwindow_ex(parent, after, lpclass, lpname):
    hwnd = win32gui.FindWindowEx(parent, after, lpclass, lpname)
    return hwnd


def setKeyboardLayout_en(inner):

    def wrapper(*args, **kwargs):
        if win32api.LoadKeyboardLayout('0x0409409', win32con.KLF_ACTIVATE) == None:
            return Exception('加载键盘失败')
        # 语言代码
        # https://msdn.microsoft.com/en-us/library/cc233982.aspx
        LID = {0x8040804: "Chinese (Simplified) (People's Republic of China)",
               0x0409409: 'English (United States)'}

        # 获取前景窗口句柄
        hwnd = win32gui.GetForegroundWindow()

        # 获取前景窗口标题
        title = win32gui.GetWindowText(hwnd)
        # 获取键盘布局列表
        im_list = win32api.GetKeyboardLayoutList()
        im_list = list(map(hex, im_list))
        oldKey = hex(win32api.GetKeyboardLayout())
        # 设置键盘布局为英文
        result = win32api.SendMessage(
            hwnd,
            win32con.WM_INPUTLANGCHANGEREQUEST,
            0,
            0x4090409)
        if result == 0:
            pass
            print('设置英文键盘成功！')

        ret = inner(*args, **kwargs)

        result = win32api.SendMessage(
            hwnd,
            win32con. WM_INPUTLANGCHANGEREQUEST,
            0,
            oldKey)
        if result == 0:
            print('还原键盘成功！')
        return ret

    return wrapper
