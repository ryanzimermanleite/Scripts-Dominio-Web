import time
from time import sleep
from os import path

import PIL.Image
from pyautogui import locateOnScreen, locateCenterOnScreen, press, hotkey, click


if __name__ == '__main__':
    while True:
        sleep(1)
        
        """while locateOnScreen(path.join('imgs', 'Screenshot_1.png'), confidence=0.8) or \
                                        locateOnScreen(path.join('imgs', 'Screenshot_4.png'), confidence=0.8) or \
                                        locateOnScreen(path.join('imgs', 'Screenshot_5.png'), confidence=0.8) or \
                                        locateOnScreen(path.join('imgs', 'Screenshot_6.png'), confidence=0.8):"""
            
        """if locateOnScreen(path.join('imgs', 'Screenshot_1.png'), confidence=0.8):
            click(locateCenterOnScreen(path.join('imgs', 'Screenshot_1.png'), confidence=0.8))
            press('x')
        if locateOnScreen(path.join('imgs', 'Screenshot_4.png'), confidence=0.8):
            click(locateCenterOnScreen(path.join('imgs', 'Screenshot_4.png'), confidence=0.8))
            press('x')
        if locateOnScreen(path.join('imgs', 'Screenshot_5.png'), confidence=0.8):
            click(locateCenterOnScreen(path.join('imgs', 'Screenshot_5.png'), confidence=0.8))
            press('x')"""
        if locateOnScreen(path.join('imgs', 'Screenshot_6.png'), confidence=0.8):
            click(locateCenterOnScreen(path.join('imgs', 'Screenshot_6.png'), confidence=0.8))
            press('r')
            time.sleep(1)
            press('y')
        
        """click()
                                sleep(1)
                                press('pgdn')"""
