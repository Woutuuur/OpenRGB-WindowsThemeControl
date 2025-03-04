from functools import partial
from typing import Callable
import openrgb
from openrgb.utils import RGBColor

import win32api
import win32con
import win32gui
import winreg

WM_THEMECHANGED = 0x031A
THEME_COLOR_CHANGED_SETTING_NAME = "ImmersiveColorSet"

STATIC_MODE_NAMES = {"direct", "static", "fixed", "solid"}

current_color = None

class Color():
    def __init__(self, r, g, b):
        self.r = r
        self.g = g
        self.b = b
    
    def __repr__(self):
        return f"Color({self.r}, {self.g}, {self.b})"
    
    def __eq__(self, other):
        if other is None:
            return False

        return self.r == other.r and self.g == other.g and self.b == other.b

class WindowsApiHelper:
    def get_accent_color():
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\DWM") as key:
                accent_color_reg = winreg.QueryValueEx(key, "AccentColor")[0]
                return Color(
                    (accent_color_reg & 0xFF),
                    ((accent_color_reg >> 8) & 0xFF),
                    ((accent_color_reg >> 16) & 0xFF)
                )
        except FileNotFoundError:
            return None

class OpenRGBClient(openrgb.OpenRGBClient):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        for device in self.devices:
            for mode in device.modes:
                if mode.name in STATIC_MODE_NAMES:
                    device.set_mode(mode.name)
                    break

    def set_color_to_all_devices(self, color: RGBColor):
        for device in self.devices:
            device.set_color(color)

class ThemeChangeListener:
    def __init__(self, on_theme_change: Callable[[Color], None] = None):
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = self.wnd_proc
        wc.lpszClassName = "ThemeChangeListener"
        wc.hInstance = win32api.GetModuleHandle(None)

        self.class_atom = win32gui.RegisterClass(wc)
        self.hwnd = win32gui.CreateWindow(
            self.class_atom,
            "Theme Change Listener",
            0,
            0, 0, 0, 0,
            0, 0, wc.hInstance, None
        )
        self.on_theme_change = on_theme_change

    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_SETTINGCHANGE:
            setting_name = win32gui.PyGetString(lparam)

            if setting_name == THEME_COLOR_CHANGED_SETTING_NAME:
                if self.on_theme_change:
                    self.on_theme_change(WindowsApiHelper.get_accent_color())

        return 0

    def run(self):
        win32gui.PumpMessages()

def apply_color_to_openrgb(client: OpenRGBClient, color: Color):
    global current_color

    if current_color == color:
        return

    current_color = color

    print(f"Setting new color: {color}")
    client.set_color_to_all_devices(RGBColor(color.r, color.g, color.b))

if __name__ == "__main__":
    openrgb_client = OpenRGBClient()

    apply_color_to_openrgb(openrgb_client, WindowsApiHelper.get_accent_color())

    listener = ThemeChangeListener(partial(apply_color_to_openrgb, openrgb_client))
    listener.run()
