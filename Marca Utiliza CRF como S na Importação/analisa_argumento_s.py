import datetime, shutil, os, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login


def marca_s_importacao(empresa, andamentos):
    cod, cnpj, nome = empresa

    _wait_img('utilitarios.png', conf=0.9, timeout=-1)

    print('>>> Gravando Argumento CRF -> S')


    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('importacao_arquivo.png', conf=0.9):
        # Relatórios
        p.hotkey('alt', 'u')
        time.sleep(0.5)
        p.press('i')
        time.sleep(0.5)
        p.press('i')
        time.sleep(0.5)
        p.press('i')
        time.sleep(2)
    time.sleep(1)
    p.press('up')
    p.press('up')
    p.press('up')
    while not _find_img('txt.png', conf=0.9):
        time.sleep(0.5)
        p.press('up')


    time.sleep(0.5)
    p.press('tab', presses=2, interval=0.2)
    time.sleep(0.5)
    p.press('2')
    time.sleep(0.5)
    _click_img('argumentos.png', conf=0.9)
    time.sleep(1)
    p.press('tab', presses=4, interval=0.2)
    time.sleep(0.5)###
    p.press('S')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('esc', presses=2, interval=0.1)
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Argumento CRF para S!']), nome=andamentos)
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
            if not _login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = marca_s_importacao(empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)


if __name__ == '__main__':
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Marca S Importação'

    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
