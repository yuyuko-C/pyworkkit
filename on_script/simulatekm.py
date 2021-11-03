
from pynput.keyboard import Controller
import pyautogui
from typing import Union

__all__ = ['SimulateKM']


class SimulateKM:

    def __init__(self) -> None:
        self.keyboard = Controller()
        self.pg = pyautogui

    def press(self, key: str):
        self.pg.press(key)

    def move_to(self, x, y):
        self.pg.moveTo(x, y)

    def click(self, x, y):
        re_x,re_y = self.pg.position()
        self.pg.click(x, y)
        self.move_to(re_x,re_y)

    def scroll(self, clicks: Union[int, float]):
        self.pg.scroll(clicks)

    def hotkey(self, key_A, key_B):
        self.pg.hotkey(key_A, key_B)

    def typewriter(self, msg: str):
        self.keyboard.type(msg)

    # 得到某点的颜色

    def get_color(self, x, y):
        return self.pg.screenshot().getpixel((x, y))

    # 定位图像在屏幕的box信息

    def locate_image(self, img_path: str, confidence=0.9):
        return self.pg.locateOnScreen(img_path, confidence=confidence)

    # 定位box的中间点信息

    def locate_image_center(self, img_path: str, region=None, confidence=0.9):
        if region == None:
            ret = self.locate_image(img_path, confidence)
        else:
            sc_size = self.pg.size()
            region = list(region)
            region[0] = 0 if region[0] < 0 else region[0]
            region[1] = 0 if region[1] < 0 else region[1]
            region[2] = sc_size[0] if region[2] > sc_size[0] else region[2]
            region[3] = sc_size[1] if region[3] > sc_size[1] else region[3]
            region = tuple(region)
            ret = self.pg.locateOnScreen(
                img_path, region=region, confidence=confidence)

        return ret != None and self.pg.center(ret) or None
pass
