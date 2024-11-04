# -*- coding: utf-8 -*-
import datetime, pyperclip, time, os, subprocess, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf, _encerra_dominio


def atualiza(empresa, andamentos):
    cod, cnpj, nome = empresa

    # Aguarda aparecer o menun de utilitarios
    _wait_img('utilitarios.png', conf=0.9, timeout=-1)

    # Aperta o atalho para entrar no portal do empregado
    p.hotkey('alt', 'u')
    time.sleep(0.5)
    p.press('p')
    time.sleep(3)

    # Se após o s 3 segundos ele abrir a tela do portal do emprego é porque está habilitado e faz o processo normal
    if _find_img('config_portal.png', conf=0.9):
        # Habilita a opção todos
        _click_img('todos.png', conf=0.9)
        time.sleep(0.5)
        # Aperta para Lista
        p.press('l')
        time.sleep(0.5)
        # Seleciona Todos
        p.press('t')
        time.sleep(0.5)
        # Clicka para Atualizar
        p.press('a')

        # Enquanto estiver atualizando e mensagem de processando ele espera 1 segundo
        while _find_img('processando.png', conf=0.9):
            time.sleep(1)

        # Apos a Mensagem processando sumir ele fecha a tela e escreve no excel
        time.sleep(1)
        p.press('f')
        print('✔ Empregados Atualizados')
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empregados Atualizados']), nome=andamentos)
        return 'ok'

    # Se apos os 3 segundos ele não encontrou a imagem do portal do empregado
    # Quer dizer que o botão esta desabilitado ou seja não entrou no portal
    # Aperta esc e escreve no excel
    else:
        p.press('esc')
        print('✔ Portal Desabilitado')
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Portal Desabilitado']), nome=andamentos)
        return 'ok'
    
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
        
        while True:
            if not _login(empresa, andamentos):
                break
            else:
                resultado = atualiza(empresa, andamentos)
                
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'ok':
                    break
    _escreve_header_csv('CÓDIGO;CNPJ/CPF;NOME;STATUS', nome=andamentos)
    _encerra_dominio()

if __name__ == '__main__':
    empresas = _open_lista_dados()
    andamentos = 'Atualiza Portal Empregado'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
    