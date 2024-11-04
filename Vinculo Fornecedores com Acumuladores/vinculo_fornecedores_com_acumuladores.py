import datetime
import time
import pyautogui as p
import pyperclip
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login

def vinculo_fornecedor_acumulador():
    while not _find_img('titulo_importacao_padrao.png', conf=0.9):
        p.hotkey('alt', 'u')
        time.sleep(0.5)
        p.press('i')
        time.sleep(0.5)
        p.press('p')
        time.sleep(0.5)
        p.press('n')
        time.sleep(3)

    p.hotkey('alt', 'c')

    while not _find_img('titulo_configuracao_importacao.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)
    while not _find_img('acumuladores.png', conf=0.9):

        _click_img('entradas.png', conf=0.9)
        time.sleep(2)
    time.sleep(1)
    while not _find_img('titulo_acumuladores.png', conf=0.9):
        _click_img('acumuladores.png', conf=0.9)
        time.sleep(2)
    time.sleep(1)


def acessa_acumulador(acumulador):
    _click_img('seta_esq.png', conf=0.9, clicks=30)
    time.sleep(2)
    _click_position_img('acumulador_2.png', '+', pixels_y=20, conf=0.9)
    time.sleep(1)
    for i in range(50):
        p.press('pgdn')
    time.sleep(0.5)
    while True:
        try:
            p.hotkey('ctrl', 'c')
            p.hotkey('ctrl', 'c')
            codigo = pyperclip.paste()
            break
        except:
            pass
    verifica = ''
    while str(codigo) != str(acumulador):
        if _find_img('linha_1.png'):
            verifica = '1'
            break

        _click_img('seta_esq.png', conf=0.9, clicks=3)
        time.sleep(1)
        p.press('up')
        time.sleep(1)
        while True:
            try:
                p.hotkey('ctrl', 'c')
                p.hotkey('ctrl', 'c')
                codigo = pyperclip.paste()
                break
            except:
                pass
        time.sleep(1)

    if verifica == '1':
        p.hotkey('alt', 'i')
        time.sleep(1)
        p.write(acumulador)
        time.sleep(1)
    time.sleep(0.5)
    p.press('right')
    time.sleep(0.5)
    p.press('tab', presses=4, interval=0.1)
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    _click_img('setinha.png', conf=0.9, clicks=10)
    time.sleep(1)#
    _click_img('selecionar.png', conf=0.9, clicks=10)

    while not _find_img('titulo_inclusao_fornecedores.png', conf=0.9):
        time.sleep(1)
        try:
            _click_position_img('selecionar.png', '+', pixels_x=307, conf=0.9)
            time.sleep(2)
        except:
            pass
        time.sleep(1)

    time.sleep(1)
def seta_acumulador(fornecedor):
    p.hotkey('alt', 'i')
    p.press('f2')
    while not _find_img('buscar_fornecedores.png', conf=0.9):
        time.sleep(0.5)
    p.write(fornecedor)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'g')
    time.sleep(2)
    _click_position_img('selecionar.png', '+', pixels_x=307, conf=0.9)
    while not _find_img('titulo_inclusao_fornecedores.png', conf=0.9):
        time.sleep(0.5)

@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('escrita_fiscal')

    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    ultimo_acumulador = None
    ultimo_cod = None
    for count, empresa in enumerate(empresas[index:], start=1):
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)
        cod, cnpj, nome, acumulador, fornecedor = empresa

        if cod != ultimo_cod:
            print('Codigo Diferente')
            time.sleep(1)
            if _find_img('titulo_inclusao_fornecedores.png', conf=0.9):
                p.hotkey('alt', 'f')
                while not _find_img('titulo_2.png', conf=0.9):

                    time.sleep(1)
                time.sleep(1)
                p.hotkey('alt', 'g')
                while not _find_img('gravar_ok.png', conf=0.9):

                    time.sleep(1)
                time.sleep(1)

                p.press('esc', presses=5)
            try:
                if _login(empresa, andamentos):
                    vinculo_fornecedor_acumulador()
                    acessa_acumulador(acumulador)
                    seta_acumulador(fornecedor)
                    resultado = 'ok'
                else:
                    resultado = 'erro'
            except Exception as e:
                p.press('esc', presses=5)
                print(f"Erro ao vincular fornecedor para a empresa {nome}: {e}")
                resultado = 'erro'
        else:
            print('Codigo Igual')
            try:
                if ultimo_acumulador == acumulador:
                    print('Acumulador Igual')
                    seta_acumulador(fornecedor)
                else:
                    print('Acumulador Diferente')
                    if _find_img('titulo_inclusao_fornecedores.png', conf=0.9):
                        p.hotkey('alt', 'f')
                        print('TESTEEEE')
                        while not _find_img('titulo_2.png', conf=0.9):
                            time.sleep(1)
                    time.sleep(1)

                    _click_img('seta_esq.png', conf=0.9, clicks=30)
                    time.sleep(2)
                    _click_position_img('acumulador_2.png', '+', pixels_y=20, conf=0.9)
                    time.sleep(1)
                    for i in range(50):
                        p.press('pgdn')
                    time.sleep(0.5)
                    acessa_acumulador(acumulador)
                    seta_acumulador(fornecedor)
                resultado = 'ok'
            except Exception as e:
                print(f"Erro ao acessar ou setar acumulador para a empresa {nome}: {e}")
                resultado = 'erro'

        if resultado == 'ok':
            ultimo_acumulador = acumulador
            ultimo_cod = cod
        _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{acumulador};{fornecedor};{resultado}', nome=andamentos)

    _escreve_header_csv('COD;CNPJ;NOME;ACUMULADOR;FORNECEDOR;STATUS', nome=andamentos)

if __name__ == '__main__':
    empresas = _open_lista_dados()
    andamentos = 'Resultado Vinculo Fornecedores com Acumuladores'
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
