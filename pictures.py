
""" Recommended use: 
from robee import pictures as pic"""

import pyautogui
import time

def click(picture_path, confidence_level = 0.99 ,is_grayscale = False , offset_x = None ,offset_y = None): # + region? picfind_area
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


def wait(picture_path, confidence_level=0.99 , is_grayscale=False , timeout=-1):
    '''[>] Waits for specified image (picture_path) - searches for the window every second.
    [>] timeout is the limit of seconds for performing this action.
    [>] If timeout is not specified by the user or is -1: 
    Action is performed infinitely.
    '''
    image = None
    infinite_loop = False
    if timeout == -1:
        timeout = 2
        infinite_loop = True
    i = 0
    while i < timeout:
        i += 1
        if infinite_loop:
            timeout +=1
        image = pyautogui.locateOnScreen(picture_path, confidence = confidence_level, grayscale = is_grayscale)
        if image is not None:
            return image
        if not infinite_loop:
            time.sleep(1)


def find(picture_path,confidence_level, is_grayscale):
    """Returns the location of specified image on screen."""
    image = pyautogui.locateOnScreen(picture_path, confidence = confidence_level, grayscale = is_grayscale)
    return image