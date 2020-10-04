import win32gui, win32api, win32con, win32ui
import collections
import time
from PIL import Image


Point = collections.namedtuple("Point", "x y")
Color = collections.namedtuple("Color", "r g b")


class Key:
    VK_LBUTTON = 0x01  # Left mouse button
    VK_RBUTTON = 0x02  # Right mouse button
    VK_CANCEL = 0x03  # Control-break processing
    VK_MBUTTON = 0x04  # Middle mouse button (three-button mouse)
    VK_XBUTTON1 = 0x05  # X1 mouse button
    VK_XBUTTON2 = 0x06  # X2 mouse button
    VK_BACK = 0x08  # BACKSPACE key
    VK_TAB = 0x09  # TAB key
    VK_CLEAR = 0x0C  # CLEAR key
    VK_RETURN = 0x0D  # ENTER key
    VK_SHIFT = 0x10  # SHIFT key
    VK_CONTROL = 0x11  # CTRL key
    VK_MENU = 0x12  # ALT key
    VK_PAUSE = 0x13  # PAUSE key
    VK_CAPITAL = 0x14  # CAPS LOCK key
    VK_KANA = 0x15  # IME Kana mode
    VK_HANGUEL = 0x15  # IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
    VK_HANGUL = 0x15  # IME Hangul mode
    VK_JUNJA = 0x17  # IME Junja mode
    VK_FINAL = 0x18  # IME final mode
    VK_HANJA = 0x19  # IME Hanja mode
    VK_KANJI = 0x19  # IME Kanji mode
    VK_ESCAPE = 0x1B  # ESC key
    VK_CONVERT = 0x1C  # IME convert
    VK_NONCONVERT = 0x1D  # IME nonconvert
    VK_ACCEPT = 0x1E  # IME accept
    VK_MODECHANGE = 0x1F  # IME mode change request
    VK_SPACE = 0x20  # SPACEBAR
    VK_PRIOR = 0x21  # PAGE UP key
    VK_NEXT = 0x22  # PAGE DOWN key
    VK_END = 0x23  # END key
    VK_HOME = 0x24  # HOME key
    VK_LEFT = 0x25  # LEFT ARROW key
    VK_UP = 0x26  # UP ARROW key
    VK_RIGHT = 0x27  # RIGHT ARROW key
    VK_DOWN = 0x28  # DOWN ARROW key
    VK_SELECT = 0x29  # SELECT key
    VK_PRINT = 0x2A  # PRINT key
    VK_EXECUTE = 0x2B  # EXECUTE key
    VK_SNAPSHOT = 0x2C  # PRINT SCREEN key
    VK_INSERT = 0x2D  # INS key
    VK_DELETE = 0x2E  # DEL key
    VK_HELP = 0x2F  # HELP key
    VK_0 = 0x30  # 0 key
    VK_1 = 0x31  # 1 key
    VK_2 = 0x32  # 2 key
    VK_3 = 0x33  # 3 key
    VK_4 = 0x34  # 4 key
    VK_5 = 0x35  # 5 key
    VK_6 = 0x36  # 6 key
    VK_7 = 0x37  # 7 key
    VK_8 = 0x38  # 8 key
    VK_9 = 0x39  # 9 key
    VK_A = 0x41
    VK_B = 0x42
    VK_C = 0x43
    VK_D = 0x44
    VK_E = 0x45
    VK_F = 0x46
    VK_G = 0x47
    VK_H = 0x48
    VK_I = 0x49
    VK_J = 0x4A
    VK_K = 0x4B
    VK_L = 0x4C
    VK_M = 0x4D
    VK_N = 0x4E
    VK_O = 0x4F
    VK_P = 0x50
    VK_Q = 0x51
    VK_R = 0x52
    VK_S = 0x53
    VK_T = 0x54
    VK_U = 0x55
    VK_V = 0x56
    VK_W = 0x57
    VK_X = 0x58
    VK_Y = 0x59
    VK_Z = 0x5A
    VK_LWIN = 0x5B  # Left Windows key (Natural keyboard)
    VK_RWIN = 0x5C  # Right Windows key (Natural keyboard)
    VK_APPS = 0x5D  # Applications key (Natural keyboard)
    VK_SLEEP = 0x5F  # Computer Sleep key
    VK_NUMPAD0 = 0x60  # Numeric keypad 0 key
    VK_NUMPAD1 = 0x61  # Numeric keypad 1 key
    VK_NUMPAD2 = 0x62  # Numeric keypad 2 key
    VK_NUMPAD3 = 0x63  # Numeric keypad 3 key
    VK_NUMPAD4 = 0x64  # Numeric keypad 4 key
    VK_NUMPAD5 = 0x65  # Numeric keypad 5 key
    VK_NUMPAD6 = 0x66  # Numeric keypad 6 key
    VK_NUMPAD7 = 0x67  # Numeric keypad 7 key
    VK_NUMPAD8 = 0x68  # Numeric keypad 8 key
    VK_NUMPAD9 = 0x69  # Numeric keypad 9 key
    VK_MULTIPLY = 0x6A  # Multiply key
    VK_ADD = 0x6B  # Add key
    VK_SEPARATOR = 0x6C  # Separator key
    VK_SUBTRACT = 0x6D  # Subtract key
    VK_DECIMAL = 0x6E  # Decimal key
    VK_DIVIDE = 0x6F  # Divide key
    VK_F1 = 0x70  # F1 key
    VK_F2 = 0x71  # F2 key
    VK_F3 = 0x72  # F3 key
    VK_F4 = 0x73  # F4 key
    VK_F5 = 0x74  # F5 key
    VK_F6 = 0x75  # F6 key
    VK_F7 = 0x76  # F7 key
    VK_F8 = 0x77  # F8 key
    VK_F9 = 0x78  # F9 key
    VK_F10 = 0x79  # F10 key
    VK_F11 = 0x7A  # F11 key
    VK_F12 = 0x7B  # F12 key
    VK_F13 = 0x7C  # F13 key
    VK_F14 = 0x7D  # F14 key
    VK_F15 = 0x7E  # F15 key
    VK_F16 = 0x7F  # F16 key
    VK_F17 = 0x80  # F17 key
    VK_F18 = 0x81  # F18 key
    VK_F19 = 0x82  # F19 key
    VK_F20 = 0x83  # F20 key
    VK_F21 = 0x84  # F21 key
    VK_F22 = 0x85  # F22 key
    VK_F23 = 0x86  # F23 key
    VK_F24 = 0x87  # F24 key
    VK_NUMLOCK = 0x90  # NUM LOCK key
    VK_SCROLL = 0x91  # SCROLL LOCK key
    VK_LSHIFT = 0xA0  # Left SHIFT key
    VK_RSHIFT = 0xA1  # Right SHIFT key
    VK_LCONTROL = 0xA2  # Left CONTROL key
    VK_RCONTROL = 0xA3  # Right CONTROL key
    VK_LMENU = 0xA4  # Left MENU key
    VK_RMENU = 0xA5  # Right MENU key
    VK_BROWSER_BACK = 0xA6  # Browser Back key
    VK_BROWSER_FORWARD = 0xA7  # Browser Forward key
    VK_BROWSER_REFRESH = 0xA8  # Browser Refresh key
    VK_BROWSER_STOP = 0xA9  # Browser Stop key
    VK_BROWSER_SEARCH = 0xAA  # Browser Search key
    VK_BROWSER_FAVORITES = 0xAB  # Browser Favorites key
    VK_BROWSER_HOME = 0xAC  # Browser Start and Home key
    VK_VOLUME_MUTE = 0xAD  # Volume Mute key
    VK_VOLUME_DOWN = 0xAE  # Volume Down key
    VK_VOLUME_UP = 0xAF  # Volume Up key
    VK_MEDIA_NEXT_TRACK = 0xB0  # Next Track key
    VK_MEDIA_PREV_TRACK = 0xB1  # Previous Track key
    VK_MEDIA_STOP = 0xB2  # Stop Media key
    VK_MEDIA_PLAY_PAUSE = 0xB3  # Play/Pause Media key
    VK_LAUNCH_MAIL = 0xB4  # Start Mail key
    VK_LAUNCH_MEDIA_SELECT = 0xB5  # Select Media key
    VK_LAUNCH_APP1 = 0xB6  # Start Application 1 key
    VK_LAUNCH_APP2 = 0xB7  # Start Application 2 key
    VK_OEM_1 = 0xBA  # Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ';:' key
    VK_OEM_PLUS = 0xBB  # For any country/region, the '+' key
    VK_OEM_COMMA = 0xBC  # For any country/region, the ',' key
    VK_OEM_MINUS = 0xBD  # For any country/region, the '-' key
    VK_OEM_PERIOD = 0xBE  # For any country/region, the '.' key
    VK_OEM_2 = 0xBF  # Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '/?' key =
    VK_OEM_3 = 0xC0  # Used for miscellaneous characters; it can vary by keyboard. = For the US standard keyboard, the '`~' key
    VK_OEM_4 = 0xDB  # Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '[{' key
    VK_OEM_5 = 0xDC  # Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '\|' key
    VK_OEM_6 = 0xDD  # Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ']}' key
    VK_OEM_7 = 0xDE  # Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the 'single-quote/double-quote' key
    VK_OEM_8 = 0xDF  # Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_102 = 0xE2  # Either the angle bracket key or the backslash key on the RT 102-key keyboard 0xE3-E4 OEM specific
    VK_PROCESSKEY = 0xE5  # IME PROCESS key 0xE6 = OEM specific #
    VK_PACKET = 0xE7  # Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP #-
    VK_ATTN = 0xF6  # Attn key
    VK_CRSEL = 0xF7  # CrSel key
    VK_EXSEL = 0xF8  # ExSel key
    VK_EREOF = 0xF9  # Erase EOF key
    VK_PLAY = 0xFA  # Play key
    VK_ZOOM = 0xFB  # Zoom key
    VK_NONAME = 0xFC  # Reserved
    VK_PA1 = 0xFD  # PA1 key
    VK_OEM_CLEAR = 0xFE  #
    dict_keycodes = {
        'VK_LBUTTON': 0x01,  # Left mouse button
        'VK_RBUTTON': 0x02,  # Right mouse button
        'VK_CANCEL': 0x03,  # Control-break processing
        'VK_MBUTTON': 0x04,  # Middle mouse button (three-button mouse)
        'VK_XBUTTON1': 0x05,  # X1 mouse button
        'VK_XBUTTON2': 0x06,  # X2 mouse button
        'VK_BACK': 0x08,  # BACKSPACE key
        'VK_TAB': 0x09,  # TAB key
        'VK_CLEAR': 0x0C,  # CLEAR key
        'VK_RETURN': 0x0D,  # ENTER key
        'VK_SHIFT': 0x10,  # SHIFT key
        'VK_CONTROL': 0x11,  # CTRL key
        'VK_MENU': 0x12,  # ALT key
        'VK_PAUSE': 0x13,  # PAUSE key
        'VK_CAPITAL': 0x14,  # CAPS LOCK key
        'VK_KANA': 0x15,  # IME Kana mode
        'VK_HANGUEL': 0x15,  # IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
        'VK_HANGUL': 0x15,  # IME Hangul mode
        'VK_JUNJA': 0x17,  # IME Junja mode
        'VK_FINAL': 0x18,  # IME final mode
        'VK_HANJA': 0x19,  # IME Hanja mode
        'VK_KANJI': 0x19,  # IME Kanji mode
        'VK_ESCAPE': 0x1B,  # ESC key
        'VK_CONVERT': 0x1C,  # IME convert
        'VK_NONCONVERT': 0x1D,  # IME nonconvert
        'VK_ACCEPT': 0x1E,  # IME accept
        'VK_MODECHANGE': 0x1F,  # IME mode change request
        'VK_SPACE': 0x20,  # SPACEBAR
        'VK_PRIOR': 0x21,  # PAGE UP key
        'VK_NEXT': 0x22,  # PAGE DOWN key
        'VK_END': 0x23,  # END key
        'VK_HOME': 0x24,  # HOME key
        'VK_LEFT': 0x25,  # LEFT ARROW key
        'VK_UP': 0x26,  # UP ARROW key
        'VK_RIGHT': 0x27,  # RIGHT ARROW key
        'VK_DOWN': 0x28,  # DOWN ARROW key
        'VK_SELECT': 0x29,  # SELECT key
        'VK_PRINT': 0x2A,  # PRINT key
        'VK_EXECUTE': 0x2B,  # EXECUTE key
        'VK_SNAPSHOT': 0x2C,  # PRINT SCREEN key
        'VK_INSERT': 0x2D,  # INS key
        'VK_DELETE': 0x2E,  # DEL key
        'VK_HELP': 0x2F,  # HELP key
        'VK_0': 0x30,  # 0 key
        'VK_1': 0x31,  # 1 key
        'VK_2': 0x32,  # 2 key
        'VK_3': 0x33,  # 3 key
        'VK_4': 0x34,  # 4 key
        'VK_5': 0x35,  # 5 key
        'VK_6': 0x36,  # 6 key
        'VK_7': 0x37,  # 7 key
        'VK_8': 0x38,  # 8 key
        'VK_9': 0x39,  # 9 key
        'VK_A': 0x41,
        'VK_B': 0x42,
        'VK_C': 0x43,
        'VK_D': 0x44,
        'VK_E': 0x45,
        'VK_F': 0x46,
        'VK_G': 0x47,
        'VK_H': 0x48,
        'VK_I': 0x49,
        'VK_J': 0x4A,
        'VK_K': 0x4B,
        'VK_L': 0x4C,
        'VK_M': 0x4D,
        'VK_N': 0x4E,
        'VK_O': 0x4F,
        'VK_P': 0x50,
        'VK_Q': 0x51,
        'VK_R': 0x52,
        'VK_S': 0x53,
        'VK_T': 0x54,
        'VK_U': 0x55,
        'VK_V': 0x56,
        'VK_W': 0x57,
        'VK_X': 0x58,
        'VK_Y': 0x59,
        'VK_Z': 0x5A,
        'VK_LWIN': 0x5B,  # Left Windows key (Natural keyboard)
        'VK_RWIN': 0x5C,  # Right Windows key (Natural keyboard)
        'VK_APPS': 0x5D,  # Applications key (Natural keyboard)
        'VK_SLEEP': 0x5F,  # Computer Sleep key
        'VK_NUMPAD0': 0x60,  # Numeric keypad 0 key
        'VK_NUMPAD1': 0x61,  # Numeric keypad 1 key
        'VK_NUMPAD2': 0x62,  # Numeric keypad 2 key
        'VK_NUMPAD3': 0x63,  # Numeric keypad 3 key
        'VK_NUMPAD4': 0x64,  # Numeric keypad 4 key
        'VK_NUMPAD5': 0x65,  # Numeric keypad 5 key
        'VK_NUMPAD6': 0x66,  # Numeric keypad 6 key
        'VK_NUMPAD7': 0x67,  # Numeric keypad 7 key
        'VK_NUMPAD8': 0x68,  # Numeric keypad 8 key
        'VK_NUMPAD9': 0x69,  # Numeric keypad 9 key
        'VK_MULTIPLY': 0x6A,  # Multiply key
        'VK_ADD': 0x6B,  # Add key
        'VK_SEPARATOR': 0x6C,  # Separator key
        'VK_SUBTRACT': 0x6D,  # Subtract key
        'VK_DECIMAL': 0x6E,  # Decimal key
        'VK_DIVIDE': 0x6F,  # Divide key
        'VK_F1': 0x70,  # F1 key
        'VK_F2': 0x71,  # F2 key
        'VK_F3': 0x72,  # F3 key
        'VK_F4': 0x73,  # F4 key
        'VK_F5': 0x74,  # F5 key
        'VK_F6': 0x75,  # F6 key
        'VK_F7': 0x76,  # F7 key
        'VK_F8': 0x77,  # F8 key
        'VK_F9': 0x78,  # F9 key
        'VK_F10': 0x79,  # F10 key
        'VK_F11': 0x7A,  # F11 key
        'VK_F12': 0x7B,  # F12 key
        'VK_F13': 0x7C,  # F13 key
        'VK_F14': 0x7D,  # F14 key
        'VK_F15': 0x7E,  # F15 key
        'VK_F16': 0x7F,  # F16 key
        'VK_F17': 0x80,  # F17 key
        'VK_F18': 0x81,  # F18 key
        'VK_F19': 0x82,  # F19 key
        'VK_F20': 0x83,  # F20 key
        'VK_F21': 0x84,  # F21 key
        'VK_F22': 0x85,  # F22 key
        'VK_F23': 0x86,  # F23 key
        'VK_F24': 0x87,  # F24 key
        'VK_NUMLOCK': 0x90,  # NUM LOCK key
        'VK_SCROLL': 0x91,  # SCROLL LOCK key
        'VK_LSHIFT': 0xA0,  # Left SHIFT key
        'VK_RSHIFT': 0xA1,  # Right SHIFT key
        'VK_LCONTROL': 0xA2,  # Left CONTROL key
        'VK_RCONTROL': 0xA3,  # Right CONTROL key
        'VK_LMENU': 0xA4,  # Left MENU key
        'VK_RMENU': 0xA5,  # Right MENU key
        'VK_BROWSER_BACK': 0xA6,  # Browser Back key
        'VK_BROWSER_FORWARD': 0xA7,  # Browser Forward key
        'VK_BROWSER_REFRESH': 0xA8,  # Browser Refresh key
        'VK_BROWSER_STOP': 0xA9,  # Browser Stop key
        'VK_BROWSER_SEARCH': 0xAA,  # Browser Search key
        'VK_BROWSER_FAVORITES': 0xAB,  # Browser Favorites key
        'VK_BROWSER_HOME': 0xAC,  # Browser Start and Home key
        'VK_VOLUME_MUTE': 0xAD,  # Volume Mute key
        'VK_VOLUME_DOWN': 0xAE,  # Volume Down key
        'VK_VOLUME_UP': 0xAF,  # Volume Up key
        'VK_MEDIA_NEXT_TRACK': 0xB0,  # Next Track key
        'VK_MEDIA_PREV_TRACK': 0xB1,  # Previous Track key
        'VK_MEDIA_STOP': 0xB2,  # Stop Media key
        'VK_MEDIA_PLAY_PAUSE': 0xB3,  # Play/Pause Media key
        'VK_LAUNCH_MAIL': 0xB4,  # Start Mail key
        'VK_LAUNCH_MEDIA_SELECT': 0xB5,  # Select Media key
        'VK_LAUNCH_APP1': 0xB6,  # Start Application 1 key
        'VK_LAUNCH_APP2': 0xB7,  # Start Application 2 key
        'VK_OEM_1': 0xBA,  # Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ';:' key
        'VK_OEM_PLUS': 0xBB,  # For any country/region, the '+' key
        'VK_OEM_COMMA': 0xBC,  # For any country/region, the ',' key
        'VK_OEM_MINUS': 0xBD,  # For any country/region, the '-' key
        'VK_OEM_PERIOD': 0xBE,  # For any country/region, the '.' key
        'VK_OEM_2': 0xBF,  # Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '/?' key
        'VK_OEM_3': 0xC0,  # Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '`~' key
        'VK_OEM_4': 0xDB,  # Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '[{' key
        'VK_OEM_5': 0xDC,  # Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '\|' key
        'VK_OEM_6': 0xDD,  # Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ']}' key
        'VK_OEM_7': 0xDE,  # Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the 'single-quote/double-quote' key
        'VK_OEM_8': 0xDF,  # Used for miscellaneous characters; it can vary by keyboard.
        'VK_OEM_102': 0xE2,  # Either the angle bracket key or the backslash key on the RT 102-key keyboard 0xE3-E4 OEM specific
        'VK_PROCESSKEY': 0xE5,  # IME PROCESS key 0xE6 OEM specific #
        'VK_PACKET': 0xE7,  # Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP #-
        'VK_ATTN': 0xF6,  # Attn key
        'VK_CRSEL': 0xF7,  # CrSel key
        'VK_EXSEL': 0xF8,  # ExSel key
        'VK_EREOF': 0xF9,  # Erase EOF key
        'VK_PLAY': 0xFA,  # Play key
        'VK_ZOOM': 0xFB,  # Zoom key
        'VK_NONAME': 0xFC,  # Reserved
        'VK_PA1': 0xFD,  # PA1 key
        'VK_OEM_CLEAR': 0xFE  #
    }


class Window:
    def __init__(self, hwnd=None):
        if hwnd is None:
            hwnd = win32gui.GetForegroundWindow()
        elif type(hwnd) is str:
            hwnd = win32gui.FindWindow(None, hwnd)
        self.hwnd = hwnd
        self.text = win32gui.GetWindowText(self.hwnd)

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

    def capture(self, region):
        w = region[1][0] - region[0][0]
        h = region[1][1] - region[0][1]
        wDC = win32gui.GetWindowDC(self.hwnd)
        dcObj = win32ui.CreateDCFromHandle(wDC)
        cDC = dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, w, h)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0, 0), (w, h), dcObj, region[0], win32con.SRCCOPY)
        bmpinfo = dataBitMap.GetInfo()
        bmpstr = dataBitMap.GetBitmapBits(True)
        return bmpstr, bmpinfo

    def bitmap_to_image(self, bmpstr, bmpinfo):
        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        return im

    def click(self, x=None, y=None, button=Key.VK_LBUTTON, duration=.2):
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
            win32gui.PostMessage(self.hwnd, win32con.WM_MOUSEMOVE, None, l_param)
        self.mouse_down(button, l_param)
        time.sleep(duration)
        self.mouse_up(button, l_param)

    def mouse_down(self, button=Key.VK_LBUTTON, l_param=None):
        # May need to call WM_MOUSEMOVE before hand depending on usage if not using through click()
        params = {
            Key.VK_LBUTTON: [
                win32con.WM_LBUTTONDOWN,
                None
            ],
            Key.VK_RBUTTON: [
                win32con.WM_RBUTTONDOWN,
                None
            ],
            Key.VK_MBUTTON: [
                win32con.WM_MBUTTONDOWN,
                None
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
        front = Window()
        while front.hwnd != self.hwnd:
            try:
                win32gui.SetForegroundWindow(self.hwnd)
            except:
                front.minimize()
                while Window().hwnd == front.hwnd:
                    pass
            front = Window()

    def minimize(self):
        win32gui.ShowWindow(self.hwnd, win32con.SW_FORCEMINIMIZE)

    def press(self, key, duration=.1):
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
    red = rgb_int & 255
    green = (rgb_int >> 8) & 255
    blue = (rgb_int >> 16) & 255
    return red, green, blue



