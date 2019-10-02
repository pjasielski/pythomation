# pythomation (Work In Progress)

# RPA - Robotic Process Automation with Python
__version__ = '2019.10.01'

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
#from selenium.common.exceptions import timeoutException

import logging
import datetime
import time

import win32gui
import win32con
import keyboard
import pyautogui
import os

import re


#====================[ UTILITY: ACTION MANAGEMENT ]========================#

def until_done(action,
               result_checker,
               checker_condition,
               interval=0,
               retry_number=-1):
    '''
    [action] - function to perform 
    [result_checker] - function to check if action was successfull
    [checker_condition] - value result_checker returns if action is successfull
    [interval = 0] - time interval between action and result_checker
    [retry_number = -1] - how many times action will be repeated
    >if not specified - it will be performed until successful.
    '''
    infinite_loop = False
    if retry_number == -1:
        retry_number = 2
        infinite_loop = True
    i = 0
    while i < retry_number:
        i+=1
        action
        #print(str(action))
        time.sleep(interval)
        result_checker
        #print(str(result_checker))
        if result_checker == checker_condition:
            return True
        if infinite_loop:
            retry_number += 1


#====================[ UTILITY: WINDOWS ]========================#
# TODO: Selenium/Winium desktop app handling

def wintop(win_name):
    '''Moves specified window to the foreground.
    '''
    handle = win32gui.FindWindow(0, win_name)
    win32gui.SetForegroundWindow(handle)
    return True


def winclose(win_name):
    '''Closes specified window.
    '''
    handle = win32gui.FindWindow(0, win_name)
    win32gui.PostMessage(handle, win32con.WM_CLOSE,0,0)
    return True


def winwait(win_name , timeout = -1):
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


def winmove_mainmax(win_name): # subject to further improvement
    '''Moves specified window to the main screen and maximizes it.
    '''
    handle = win32gui.FindWindow(0, win_name)
    win32gui.SetForegroundWindow(handle)
    win32gui.ShowWindow(handle,1)
    win32gui.MoveWindow(handle, 100,100, 200,200, 1)   
    win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)


def winmove_main_citrix(win_name): # subject to further improvement
    '''[For Citrix]:
    Moves specified window to the main screen and maximizes it.
    '''
    handle = win32gui.FindWindow(0, win_name)
    win32gui.ShowWindow(handle, 3)
    win32gui.SetForegroundWindow(handle)
    keyboard.press('alt')
    keyboard.send('space')
    keyboard.release('alt')
    time.sleep(1)
    keyboard.send('r')
    win32gui.SetForegroundWindow(handle)
    keyboard.press('alt')
    keyboard.send('space')
    keyboard.release('alt')
    time.sleep(1)
    keyboard.send('m')
    time.sleep(1)
    pyautogui.FAILSAFE = False
    time.sleep(1)
    pyautogui.dragTo(100, 100,1, button='left')
    time.sleep(1)
    pyautogui.doubleClick()
    pyautogui.FAILSAFE = True


#====================[ UTILITY: IMAGES ]========================#

def picclick(picture_path, confidence_level=0.99 ,is_grayscale=False , offset_x=None ,offset_y=None): # + region? picfind_area
    '''[>] Clicks on specified image (picture_path) on the screen.
    [>] confidence_level specifies the percentage of pixels that have to match the source image.
    [>] If is_grayscale = True: image is searched in grayscale (colors are ignored)
    [>] The click can be offset by setting 2 additional parameters: offset_x and offset_y
    offset_x specifies how many pixels the cursor moves from the top left corner of the image HORIZONTALLY
    offset_y specifies how many pixels the cursor moves from the top left corner of the image VERTICALLY
    '''
    image = pyautogui.locateOnScreen(picture_path, confidence = confidence_level, grayscale = is_grayscale)
    if offset_x == None and offset_y == None:
        pyautogui.click(image)
    else:
        pyautogui.click(image[0] + offset_x , image[1]+ offset_y)


def picwait(picture_path, confidence_level=0.99 , is_grayscale=False , timeout=-1):
    '''[>] Waits for specified image (picture_path) - searches for the window every second.
    [>] timeout is the limit of seconds for performing this action.
    If timeout is not specified by the user or is -1: 
    [>] Action is performed infinitely.
    '''
    image = None
    infinite_loop = False
    if timeout == -1:
        timeout = 2
        infinite_loop = True
    i = 0
    while i < timeout:
        i += 1
        if Infinite_Loop:
            timeout +=1
        image = pyautogui.locateOnScreen(picture_path, confidence=confidence_level, grayscale=is_grayscale)
        if image is not None:
            return image
        if not infinite_loop:
            time.sleep(1)

#TODO: relational secondaary image search (finishing touches)

#====================[ UTILITY: WEB ]========================#

def webopen_chrome(chrome_driver, website_address): # subject to further improvement
    ''' [>] Opens specified website (WebsiteAddress) with Chrome Browser.
    [>] ChromeDriver - path to the Chrome driver that is required for the version of Chrome that is utilized.
    '''
    options = Options()
    options.add_argument("--disable-extensions")
    options.add_argument("--start-maximized")
    options.add_experimental_option("useAutomationExtension", False)
    prefs = {'download.prompt_for_download': True,}
    options.add_experimental_option('prefs', prefs)
    #ChooseDriver
    browser = webdriver.Chrome(chrome_driver ,options=options)
    #Open Website
    browser.get(website_address)

    
# TODO: Add web module here

#====================[ UTILITY: LOG ]========================#

def botcomment(comment_text, date_format='%Y-%m-%d %H:%M:%S',  script_name=os.path.basename(__file__)):
    '''[>] Prints a comment in specified format
    '''
    t = datetime.datetime.now()
    current_time = str(t.strftime(date_format))
    botcomment = "[" + current_time + "]" + " " + script_name + ": " + comment_text
    print(botcomment)
    return botcomment


def datenow(date_format):
    '''[>] Returns current time and date in specified format (date_format)
    '''
    t = datetime.datetime.now()
    current_time = str(t.strftime(date_format))
    return current_time


def botlog(comment,
            log_file_name,
            script_name =  str(os.path.basename(__file__)),
            file_mode = 'a',
            comment_format = '%(message)s'):
    '''[>] Creates or appends a log file with bot comments
    '''
    logging.basicConfig(filename = log_file_name,
                            level = logging.DEBUG,
                            filemode = 'a',
                            format = '%(message)s',
                            datefmt = '%H:%M:%S')
    logging.info(str(botcomment(comment, script_name = script_name)))
    
    
def screenshot(folder_path, file_name, file_format='.png', clean=False, replacement = ''):
    '''
    folder_path - path to folder in which the screenshot is supposed to be saved
    file_name - screenshot file name
    file_format - file format extension ('.png' is default)
    clean - if set to True, replaces special characters
    replacement - every special character in file name will be replaced with this
    '''   
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    screen_name = os.path.join(folder_path, file_name)
    if clean:
        screen_name = re.sub('[^a-zA-Z0-9 \n\.]', replacement, screen_name)
    pyautogui.screenshot(screen_name + file_format)


#====================[ UTILITY: FINDERS ]========================#

def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))


def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst


def find_windows(win_name_part, return_single = True):
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
    print("Find Windows Ex: " + str(win32gui.FindWindowEx(handle,0,None,None)))
    #print(win32api.GetHandleInformation(handle))
    print("Parrent: " + str(win32gui.GetParent(handle)))
    return handle, win_name


def win_find_by_handle(handle):
    win_name = str(win32gui.GetWindowText(handle))
    print(f'[>] For handle {handle} the window name is: {win_name}')
    return win_name
