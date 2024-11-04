# -*- coding: utf-8 -*-
import datetime, re, shutil, time, os, pyautogui as p
import pyperclip
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf


def relatorio_comparativo(inicio, final, empresa, andamentos):
    cod, cnpj, nome = empresa

    _click_img('dominio.png', conf=0.9)
    
    time.sleep(2)

    _click_img('conteudo.png', conf=0.9)
    time.sleep(2)

    _click_img('icone_comparativo.png', conf=0.9)
    
    while not _find_img('titulo_comparativo.png', conf=0.9):
        time.sleep(1)
        if _find_img('filial.png', conf=0.9):
            time.sleep(1)
            p.press("enter")
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Empresa Filial;', nome=andamentos)
            return 'ok'
            
    time.sleep(1)
    p.write(inicio)
    time.sleep(0.5)
    p.press("tab")
    time.sleep(0.5)
    p.write(final)
    time.sleep(2)
    if _find_img('simples_nada.png', conf=0.9):
        p.hotkey('alt', 'c')
    else:
        p.press("tab")
        time.sleep(0.5)
        p.press("tab")
        time.sleep(0.5)
        p.press("r")
        time.sleep(1)
        p.hotkey('alt', 'c')


    while not _find_img('versao_impressao.png', conf=0.9):
        if _find_img('modulo.png', conf=0.9):
            time.sleep(1)
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Sem Parametros Escrita Fiscal;', nome=andamentos)
            return 'ok'
        time.sleep(1)
        _click_img('seta_baixo.png', conf=0.9)
        time.sleep(1)
        if _find_img('comparar.png', conf=0.9):
            try:
                _click_img('comparar.png', conf=0.9)
            except:
                pass

    try:
        _click_img('versao_impressao.png', conf=0.9)
    except:
        pass
    
    while not _find_img('img1.png', conf=0.9):
        time.sleep(1)
        if _find_img('aviso1.png', conf=0.9):
            p.press("enter")
        elif _find_img('aviso3.png', conf=0.9):
            p.press("enter")
        elif _find_img('aviso_super_simples.png', conf=0.9):
            p.press('enter')
        print('X1')
        
    p.press("tab")
    time.sleep(0.5)
    p.press("enter")
    time.sleep(1)
    print('2')
    
    while not _find_img('print.png', conf=0.9):
        print('oiiii')
        time.sleep(1)
        
    while not _find_img('preview_pdf.png', conf=0.9):
    
        time.sleep(1)
        _click_img('seta_lado.png', conf=0.9)
        time.sleep(1)
        
    _click_img('preview_pdf.png', conf=0.9)
    

    time.sleep(1)
    print('abcDDD')
    p.hotkey('alt', 'p')

    while not _find_img('preview_pdf2.png', conf=0.9):
        time.sleep(1)
        _click_img('adobe.png', conf=0.9)
        time.sleep(2)

    _click_img('preview_pdf2.png', conf=0.9)

    p.hotkey('ctrl', 'shift', 's')

    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)


    while not _find_img('salvar_como2.png', conf=0.9):
        time.sleep(1)
        _click_img('pasta_diferente.png', conf=0.9)

    salvar(cod, cnpj, nome, andamentos)

    p.press("esc", presses=5, interval=0.2)
    _click_img('print.png', conf=0.9)
    p.press("esc", presses=5, interval=0.2)

    return 'ok'

def salvar(cod, cnpj, nome, andamentos):
    os.makedirs(r'{}\{}'.format(e_dir, 'Boletos'), exist_ok=True)
    nome = nome.replace("/", "-")
    p.write(cnpj + ' - ' + nome)
    # Selecionar local

    p.press('tab', presses=6)

    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Domínio\Gera Relatório Comparativo Regime Tributário\{}\{}'.format(e_dir, 'Boletos'))
    p.hotkey('ctrl', 'v')
    ##
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Guia gerada;', nome=andamentos)
    p.hotkey('alt', 'f4')
    print('✔ Guia gerada')
    time.sleep(1)
    p.hotkey('ctrl', 'w')

@_time_execution
def run():
    _login_web()
    _abrir_modulo('escrita_fiscal')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
        
        while True:
            if not _login(empresa, andamentos):
                break
            else:
                resultado = relatorio_comparativo(inicio, final, empresa, andamentos)
                
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'ok':
                    break

if __name__ == '__main__':
    inicio = p.prompt(text='Qual a competência inicial?', title='Script incrível', default='00/0000')
    final = p.prompt(text='Qual a competência final?', title='Script incrível', default='00/0000')
    empresas = _open_lista_dados()
    andamentos = 'Relatório Comparativo Regime Tributario'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
        