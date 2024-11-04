#!/usr/bin/env python
import sys
import PySimpleGUI as sg
import subprocess
import os
import pyperclip, time, os, subprocess, pyautogui as p
from sys import path
from threading import Thread


path.append(r'..\..\_comum')

from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _salvar_pdf, _encerra_dominio


def run_script_thread_parametro():
    parametro()
def run_script_thread_patrimonio():
    patrimonio()

def patrimonio():

    while not _find_img('patrimoniais.png', conf=0.9):
        _click_img('dominio.png', conf=0.9)
        time.sleep(0.5)
        p.hotkey('alt', 'u')
        time.sleep(0.5)
        p.press('r')
        time.sleep(0.5)
        p.press('r')
        p.press('enter')
        time.sleep(0.5)
        p.press('p')
        time.sleep(1)

    _click_img('clientes.png', conf=0.9)
    time.sleep(0.5)
    p.press('l')
    time.sleep(1)

    if _find_img('aviso_cadastro.png', conf=0.9):
        time.sleep(0.5)
        p.press('enter')
        window['output'].update('Cliente OK!')
        cliente = 'ok'
    else:
        p.press('t')
        time.sleep(0.5)
        p.press('g')
        time.sleep(2)
        p.press('y')
        window['output'].update('Cliente OK!')
        cliente = 'ok'


    _click_img('fornecedores.png', conf=0.9)
    time.sleep(0.5)
    p.press('l')
    time.sleep(1)

    if _find_img('aviso_cadastro.png', conf=0.9):
        time.sleep(0.5)
        p.press('enter')
        window['output'].update('Fornecedor OK!')
        fornecedor = 'ok'
    else:
        p.press('t')
        time.sleep(2)
        p.press('g')
        time.sleep(2)
        p.press('y')
        time.sleep(2)
        window['output'].update('Fornecedor OK!')
        fornecedor = 'ok'

    if cliente == 'ok' and fornecedor == 'ok':
        p.press('t')
        time.sleep(0.5)
        p.press('r')
        time.sleep(0.5)
        p.press('y')
        time.sleep(0.5)
        p.press('enter')
        window['output'].update('Gerado com sucesso!')




def parametro():
    cod, cnpj = empresa

    while not _find_img('parametro.png', conf=0.9):
        p.hotkey('alt', 'c')
        p.press('p')
        time.sleep(3)


    if _find_img('parametro.png', conf=0.9):
        _click_img('contabilidade.png', conf=0.9)
        time.sleep(0.5)

    p.press('tab', presses=25)

    time.sleep(1)

    p.hotkey('ctrl', 'c')
    time.sleep(1)
    p.hotkey('ctrl', 'c')
    parametro_fornecedor = pyperclip.paste()
    fornecedor = parametro_fornecedor.split('.')


    p.press('tab', presses=4)
    time.sleep(1)

    p.hotkey('ctrl', 'c')
    time.sleep(1)
    p.hotkey('ctrl', 'c')
    parametro_cliente = pyperclip.paste()
    cliente = parametro_cliente.split('.')


    if cliente[2] == '3' or fornecedor[2] == '3':
        window['output'].update('Parametro Errado!')
        return 'erro'
    else:
        window['output'].update('Parametro Ok!')
        return 'ok'


if __name__ == '__main__':
    sg.theme('Dark')

    sg.set_options(element_padding=(0, 0),
                   button_element_size=(12, 1), auto_size_buttons=False)

    layout = [[sg.Input(size=(10, 10), key='-DATA_INICIO-'),
               sg.Input(size=(10, 10), key='-DATA_FINAL-'),
               sg.Button('Parâmetro'),
               sg.Button('Patrimonio'),
               sg.Button('Regerar', button_color=('white', '#35008B')),
               sg.Button('EXIT', button_color=('white', 'firebrick3'))],
              [sg.Text('', text_color='white', size=(50, 1), key='output')]]

    window = sg.Window('Floating Toolbar',
                       layout,
                       no_titlebar=True,
                       grab_anywhere=True,
                       keep_on_top=True)

    # ---===--- Loop taking in user input and executing appropriate program --- #
    while True:
        event, values = window.read()
        if event == 'EXIT' or event == sg.WIN_CLOSED:
            break  # exit button clicked
        if event == 'Parâmetro':
            script_thread = Thread(target=run_script_thread_parametro)
            script_thread.start()
        elif event == 'Patrimonio':
            script_thread = Thread(target=run_script_thread_patrimonio)
            script_thread.start()

        elif event == 'Regerar':
            print('Regerar')

    window.close()
