import os
import sys
import time

import comtypes.client as cc
# cc.GetModule(os.path.join(sys.path[0], 'lib/tl.tlb'))
import comtypes.gen.TaskbarLib as tbl
import win32api
import win32gui
from plyer import notification

taskbar = None
hWnd = None



def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = sys.path[0]
    return os.path.join(base_path, relative_path)


def getConsoleHwnd():
    global hWnd
    if not hWnd:
        old_title = win32api.GetConsoleTitle()
        new_window_title = f'{win32api.GetTickCount()}/{win32api.GetCurrentProcess()}'
        win32api.SetConsoleTitle(new_window_title)
        time.sleep(0.04)
        hWnd = win32gui.FindWindow(None, new_window_title)
        win32api.SetConsoleTitle(old_title)
    return hWnd


def update_taskbar_progress(progress):
    global taskbar
    hWnd = getConsoleHwnd()
    if not hWnd:
        return
    # TBPF_NOPROGRESS = 0
    # TBPF_INDETERMINATE = 0x1
    # TBPF_NORMAL = 0x2
    # TBPF_ERROR = 0x4
    # TBPF_PAUSED = 0x8
    if not taskbar:
        cc.GetModule(resource_path(r'lib\tl.tlb'))
        taskbar = cc.CreateObject('{56FDF344-FD6D-11d0-958A-006097C9A090}', interface=tbl.ITaskbarList3)
        taskbar.HrInit()
    if progress == 0:
        taskbar.SetProgressState(hWnd, 0)
    else:
        taskbar.SetProgressValue(hWnd, int(progress * 100), 100)
        taskbar.SetProgressState(hWnd, 0x2)
