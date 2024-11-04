# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Envia Patrimonio                                           #
# Arquivo:  envia_patrimonio.py                                        #
# Versão:   1.0.0                                                       #
# Modulo:   Dominio                                                     #
# Objetivo: Verifica Parametro Empresa                                  #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     07/11/2023                                                  #
# ----------------------------------------------------------------------#
import datetime, pyperclip, time, os, subprocess, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _salvar_pdf, _encerra_dominio


def verifica_empresa(cod):
    erro = 'sim'
    while erro == 'sim':
        try:
            # Faz um clique no canto superior direito para verificar a empresa
            p.click(1258, 82)

            while True:
                try:
                    # Copia o nome
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    cnpj_codigo = pyperclip.paste()
                    break
                except:
                    pass

            time.sleep(0.5)
            codigo = cnpj_codigo.split('-')
            codigo = str(codigo[1])
            codigo = codigo.replace(' ', '')
            # Compara o codigo digitado com o codigo da empresa que entrou se for diferente a empresa não existe
            if codigo != cod:
                print(f'Código da empresa: {codigo}')
                print(f'Código encontrado no Domínio: {cod}')
                return False
            else:
                return True
        except:
            erro = 'sim'


def login(empresa, andamentos):
    cod, cnpj, status = empresa
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    # Faz um clique no meio da tale pra dar um focus no dominio
    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta='imgs_c', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(cod)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        if _find_img('sem_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Parametro não cadastrado para esta empresa']), nome=andamentos)
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not _find_img('parametros.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        # Se der a mensagem que nao tem parametros escreve no relatorio e pula a empresa
        if _find_img('nao_existe_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Não existe parametro cadastrado para esta empresa']),
                                   nome=andamentos)
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        # Se a empresa não usar o sistema escreve no relatorio e pula a empresa
        if _find_img('empresa_nao_usa_sistema.png', pasta='imgs_c', conf=0.9) or _find_img(
                'empresa_nao_usa_sistema_2.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa não está marcada para usar este sistema']),
                                   nome=andamentos)
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('fase_dois_do_cadastro.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)

        if _find_img('aviso_regime.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)

        if _find_img('aviso.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'o')
            time.sleep(1)

        if _find_img('erro_troca_empresa.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5, interval=1)
            login(empresa, andamentos)

    if not verifica_empresa(cod):
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa não encontrada']), nome=andamentos)
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False

    p.press('esc', presses=5)
    time.sleep(1)

    return True


def envia_eventos(ano, empresa, andamentos):
    cod, cnpj, status = empresa

    while not _find_img('envio.png', conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)

        p.press('n')
        time.sleep(0.5)
        p.press('f')
        time.sleep(0.5)
        p.press('n')
        time.sleep(0.5)
        p.press('n')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.press('e')
        time.sleep(2)

    p.write(ano)
    time.sleep(0.5)
    p.press('tab', presses=2)
    time.sleep(0.5)
    p.press('up', presses=6)
    time.sleep(0.5)
    #p.press('v')
    time.sleep(20)


    _escreve_relatorio_csv(';'.join([cod, cnpj, 'REABERTURA COMPLETA']), nome=andamentos)
    return 'ok'


@_barra_de_status
def run(window):
    # abre o Domínio Web e o módulo, no caso será o módulo Folha
    _login_web()
    _abrir_modulo('escrita_fiscal')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        while True:
            # abre a empresa no domínio
            if not login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = envia_eventos(str(ano), empresa, andamentos)

            if resultado == 'dominio fechou':
                _login_web()
                _abrir_modulo('escrita_fiscal')

            if resultado == 'modulo fechou':
                _abrir_modulo('escrita_fiscal')

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('CÓDIGO;CNPJ;STATUS', nome=andamentos)
    _encerra_dominio()


if __name__ == '__main__':
    # Captura o ano base que vai ser usado para apuração e envio de reinf
    ano = p.prompt(text='Qual data base?', title='Script incrível', default='00/0000')
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado_REINF'

    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        run()
