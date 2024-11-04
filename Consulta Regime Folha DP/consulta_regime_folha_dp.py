import datetime, pyperclip, shutil, os, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login


def consulta_regime_folha(empresa, andamentos):
    cod, cnpj, nome = empresa

    while not _find_img('titulo_parametro.png', conf=0.9):
        p.hotkey('alt', 'c')
        time.sleep(0.5)
        p.press('p')
        time.sleep(3)

    while not _find_img('personaliza.png', conf=0.9):
        time.sleep(1)
        _click_img('geral.png', conf=0.9)

    while not _find_img('data_pagamento.png', conf=0.9):
        time.sleep(1)
        _click_img('personaliza.png', conf=0.9)

    while not _find_img('empregados.png', conf=0.9):
        time.sleep(1)
        _click_img('data_pagamento.png', conf=0.9)

    _click_img('empregados.png', conf=0.9)

    p.press('tab', presses=4, interval=0.2)
    time.sleep(0.5)
    p.hotkey('ctrl', 'c')
    time.sleep(0.5)
    p.hotkey('ctrl', 'c')
    time.sleep(0.5)
    p.hotkey('ctrl', 'c')
    time.sleep(0.5)
    total_dias = pyperclip.paste()
    p.press('esc', presses=5)
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, str(total_dias)]),
                           nome=andamentos)
    return 'ok'
@_barra_de_status
def run(window):
    # abre o Domínio Web e o módulo, no caso será o módulo Folha
    _login_web()
    _abrir_modulo('folha')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        while True:
            # abre a empresa no domínio
            if not _login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = consulta_regime_folha(empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Regime Folha DP'
    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()