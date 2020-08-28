import win32gui, win32api, win32con
import collections
import time
from keycodes import Key


Point = collections.namedtuple("Point", "x y")
Color = collections.namedtuple("Color", "r g b")


class Window:
    def __init__(self, hwnd=None):
        if hwnd is None:
            hwnd = win32gui.GetForegroundWindow()
        elif type(hwnd) is str:
            hwnd = win32gui.FindWindow(None, hwnd)
        self.hwnd = hwnd

    def pixel(self, x, y=None):
        if y is None:
            y = x[1]
            x = x[0]
        try:
            dc = win32gui.GetWindowDC(self.hwnd)
            colorref = win32gui.GetPixel(dc, x + 8, y + 8)
        except:
            raise(Exception, "Could not get pixel")
        win32gui.DeleteDC(dc)
        r, g, b = rgbint2rgbtuple(colorref)
        return Color(r, g, b)

    def click(self, x=None, y=None, button=Key.VK_LBUTTON, duration=.3):
        if x:
            if not y:
                y = x[1]
                x = x[0]
        else:
            x, y = win32api.GetCursorPos()

        l_param = win32api.MAKELONG(x, y)
        # Not sure why, but when the window is in foreground it doesn't seem to work well with post message
        if win32gui.GetForegroundWindow() == self.hwnd:
            win32api.SetCursorPos((x, y))
        else:
            win32gui.PostMessage(self.hwnd, win32con.WM_MOUSEMOVE, 0, l_param)
        self.mouse_down(button, l_param)
        time.sleep(duration)
        self.mouse_up(button, l_param)

    def mouse_down(self, button=Key.VK_LBUTTON, l_param=None):
        # May need to call WM_MOUSEMOVE before hand depending on usage if not using through click()
        params = {
            Key.VK_LBUTTON: [
                win32con.WM_LBUTTONDOWN,
                win32con.MK_LBUTTON
            ],
            Key.VK_RBUTTON: [
                win32con.WM_RBUTTONDOWN,
                win32con.MK_RBUTTON
            ],
            Key.VK_MBUTTON: [
                win32con.WM_MBUTTONDOWN,
                win32con.MK_MBUTTON
            ]
        }

        u_int, w_param = params[button]
        win32api.PostMessage(self.hwnd, u_int, w_param, l_param)

    def mouse_up(self, button=Key.VK_LBUTTON, l_param=None):
        params = {
            Key.VK_LBUTTON: [
                win32con.WM_LBUTTONUP,
                None
            ],
            Key.VK_RBUTTON: [
                win32con.WM_RBUTTONUP,
                None
            ],
            Key.VK_MBUTTON: [
                win32con.WM_MBUTTONUP,
                None
            ]
        }

        u_int, w_param = params[button]
        win32api.PostMessage(self.hwnd, u_int, w_param, l_param)

    def bring_to_front(self):
        front = win32gui.GetForegroundWindow()
        while win32gui.GetWindowText(front) != win32gui.GetWindowText(self.hwnd):
            win32gui.ShowWindow(front, win32con.SW_FORCEMINIMIZE)
            try:
                win32gui.SetForegroundWindow(self.hwnd)
            except:
                pass
            front = win32gui.GetForegroundWindow()

    def press(self, key, duration=.6):
        self.key_down(key)
        time.sleep(duration)
        self.key_up(key)

    def key_down(self, key):
        win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, key, 0)

    def key_up(self, key):
        win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, key, 0)


def points_to_region(p1, p2=None):
    if p2 is None:
        p2, p1 = p1
    xi = min(p1[0], p2[0])
    xf = max(p1[0], p2[0])
    yi = min(p1[1], p2[1])
    yf = max(p1[1], p2[1])
    return [xi, yi, xf - xi, yf - yi]


def rgbint2rgbtuple(rgb_int):
    blue = rgb_int & 255
    green = (rgb_int >> 8) & 255
    red = (rgb_int >> 16) & 255
    return red, green, blue



