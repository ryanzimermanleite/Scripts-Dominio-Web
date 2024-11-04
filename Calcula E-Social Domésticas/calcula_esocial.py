# -*- coding: utf-8 -*-
import datetime, pyperclip, time, os, shutil, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login, _salvar_pdf, _login_web, _abrir_modulo


def calcula(empresa, comp, andamento):
    cod, cnpj, nome = empresa
    
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'p')
    time.sleep(0.5)
    # Cálculos
    p.press('c')
    time.sleep(0.5)
    
    while not _find_img('tela_calculo.png', conf=0.9):
        time.sleep(1)
   
    p.write(comp)
    time.sleep(1)
    
    # calcular
    p.hotkey('alt', 'r')
    
    while not _find_img('calculou.png', conf=0.9):
        time.sleep(1)
    
    p.press('esc', presses=4)
    time.sleep(2)
    
    
@_time_execution
@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('folha')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)
    
        if not _login(empresa, andamentos):
            continue
            
        calcula(empresa, comp, andamentos)


if __name__ == '__main__':
    empresas = _open_lista_dados()
    andamentos = 'Cálculo eSocial'
    comp = _get_comp(printable='mm/aaaa', strptime='%m/%Y')
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
