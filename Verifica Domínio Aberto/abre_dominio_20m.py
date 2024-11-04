# -*- coding: utf-8 -*-
import pyautogui as p
import time
from sys import path

path.append(r'..\..\_comum')
from comum_comum import _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login


def cronometro(window, tempo_total_minutos):
    tempo_total_segundos = tempo_total_minutos * 60
    inicio = time.perf_counter()

    for segundos_restantes in range(tempo_total_segundos, -1, -1):
        minutos, segundos = divmod(segundos_restantes, 60)
        while True:
            try:
                window['-Mensagens-'].update(f"{minutos:02d}:{segundos:02d}")
                window.refresh()
                break
            except:
                pass
        time_restante = max(0, 1 - (time.perf_counter() - inicio))
        time.sleep(time_restante)
        inicio = time.perf_counter()


@_barra_de_status
def run(window):
    while True:
        p.hotkey('win', 'm')
        while True:
            try:
                window['-Mensagens-'].update(f"Abrindo domínio")
                window.refresh()
                break
            except:
                pass
        _login_web(force_open=True)
        _abrir_modulo('escrita_fiscal', usuario='ROBO', senha='Rb#0086*')
        cronometro(window, 20)


if __name__ == '__main__':
    run()
