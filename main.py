import ctypes
from ctypes import wintypes

from winshell import recycle_bin
from PIL import Image
from pystray import Icon, Menu, MenuItem
from pystray._util import win32
from threading import Thread
from time import sleep
from os import system, walk, getlogin
from os.path import getsize, islink, join
import win32security
from env import BIN0, BIN1, BIN2, BIN3

class DDIcon(Icon):
    def _on_notify(self, wparam, lparam):
        """Handles ``WM_NOTIFY``.

        If this is a left button click, this icon will be activated. If a menu
        is registered and this is a right button click, the popup menu will be
        displayed.
        """
        if lparam == 0x0203: # left double-click
            clear_bin()

        elif self._menu_handle and lparam == win32.WM_RBUTTONUP:
            # TrackPopupMenuEx does not behave unless our systray window is the
            # foreground window
            win32.SetForegroundWindow(self._hwnd)

            # Get the cursor position to determine where to display the menu
            point = wintypes.POINT()
            win32.GetCursorPos(ctypes.byref(point))

            # Display the menu and get the menu item identifier; the identifier
            # is the menu item index
            hmenu, descriptors = self._menu_handle
            index = win32.TrackPopupMenuEx(
                hmenu,
                win32.TPM_RIGHTALIGN | win32.TPM_BOTTOMALIGN
                | win32.TPM_RETURNCMD,
                point.x,
                point.y,
                self._menu_hwnd,
                None)
            if index > 0:
                descriptors[index - 1](self)

def clear_bin():
    try:
        recycle_bin().empty(confirm=False, show_progress=False, sound=False)
        icon.icon = Image.open(BIN0)
    except Exception as e:
        return(f"ERROR: {e}")

def open_bin():
    system("start shell:RecycleBinFolder")

def get_actual_img():
    size = get_bin_size() # KiB

    if size == 0:
        return Image.open(BIN0)
    elif size < 102400:
        return Image.open(BIN1)
    elif size < 1048576:
        return Image.open(BIN2)
    else:
        return Image.open(BIN3)

def get_bin_size():
    size = 0 # in bytes
    for bin_path, dirs, files in walk(BIN_PATH):
        for file in files:
            file_path = join(bin_path, file)
            if not islink(file_path):
                try:
                    size += getsize(file_path)
                except OSError as e:
                    print(f"OSError: {e}")
                    size += 1 # to user know there are some files
                    continue
    return size // 1024 # KiB

def update_icon():
    while True:
        BIN_IMG = get_actual_img()
        icon.icon = BIN_IMG
        sleep(3)

def create_icon():
    menu = Menu(
        MenuItem("Clear bin", clear_bin),
        MenuItem("Open", open_bin),
        MenuItem("Exit", lambda: icon.stop())
    )

    return DDIcon(name="Mini Bin", icon=BIN_IMG, title="Mini Bin", menu=menu)

def get_bin_path():
    path = r"C:/$Recycle.Bin/"

    username = getlogin()

    sid_binary, _, _ = win32security.LookupAccountName(None, username)
    SID = win32security.ConvertSidToStringSid(sid_binary)

    path += SID

    return path

BIN_PATH = get_bin_path()
icon = None
BIN_IMG = get_actual_img()

if __name__ == "__main__":
    icon = create_icon()
    icon_pull = Thread(target=icon.run, daemon=True)
    icon_pull.start()

    Thread(target=update_icon, daemon=True).start()

    icon_pull.join()