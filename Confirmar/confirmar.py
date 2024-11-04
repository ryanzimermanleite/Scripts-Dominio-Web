from time import sleep
from os import path
from pyautogui import locateOnScreen, press, hotkey


if __name__ == '__main__':
    while 1 > 0:
        sleep(1)
        
        if locateOnScreen(path.join('imgs', 'empresa_nao_esta_pronto_para_provisao.png'), confidence=0.8):
            press('enter')
        if locateOnScreen(path.join('imgs', 'lancamentos_regerados.png'), confidence=0.8):
            hotkey('alt', 'y')
        if locateOnScreen(path.join('imgs', 'erros_avisos.png'), confidence=0.8):
            hotkey('alt', 'n')
            sleep(2)
            press('esc', presses=100)
