
""" Recommended use: 
from robee import windows as win"""

import win32gui
import win32con
import time


def top(win_name):
    '''Moves specified window to the foreground.
    '''
    handle = win32gui.FindWindow(0, win_name)
    win32gui.SetForegroundWindow(handle)
    return True


def close(win_name):
    '''Closes specified window.
    '''
    handle = win32gui.FindWindow(0, win_name)
    win32gui.PostMessage(handle, win32con.WM_CLOSE,0,0)
    return True


def wait(win_name , timeout = -1):
    ''' [>] Waits for specified window - searches for the window every second.
    [>] timeout is the limit of seconds for performing this action.
    [>] If timeout is not specified by the user or is -1: Action is performed infinitely.
    '''

    infinite_loop = False
    if timeout == -1:
        timeout = 2
        infinite_loop = True
    i = 0
    is_visible=False
    while not is_visible:
        i += 1
        if infinite_loop:
            Timeout +=1
        handle = win32gui.FindWindow(0, WinName)
        is_visible = win32gui.IsWindowVisible(handle)
        if is_visible or i >= timeout:
            return is_visible
        if not infinite_loop:
            time.sleep(1)


def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))


def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst


def find(win_name_part, return_single = True):
    ''' Finds and prints a list of windows with names that contain the (win_name_part)
    '''
    appwindows = get_app_list()
    list_win_name = []
    print(f'[>] Windows with name "{win_name_part}":')
    for i in appwindows:
        if win_name_part.lower() in i[1].lower():
            print (i)
            if return_single:
                return i
            else:
                list_win_name.append(i)
    if not return_single:
        return list_win_name


def get_win_info():
    handle = win32gui.GetForegroundWindow()
    print("[>] Foreground window number: " + str(handle))
    win_name = str(win32gui.GetWindowText(handle))
    print(f'Foreground window text: {win_name}')