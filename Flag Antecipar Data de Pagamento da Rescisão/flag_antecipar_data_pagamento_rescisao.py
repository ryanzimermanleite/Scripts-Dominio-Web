import datetime
import time
import os
import pyautogui as p
from _comum.pyautogui_comum import _find_img, _click_img, _click_position_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _get_host_name, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _escreve_header_csv
from _comum.dominio_comum import _login_web, _abrir_modulo, _login

def flag_antecipar_data_pagamento_rescisao(cod, cnpj, nome):
    while not _find_img('titulo_parametros.png', conf=0.9):
        time.sleep(1)
        p.hotkey('alt', 'c')
        time.sleep(1)
        p.press('p')
        time.sleep(5)
    time.sleep(1)

    while not _find_img('personaliza.png', conf=0.9):
        time.sleep(1)
        _click_position_img('regime.png', '-', pixels_x=45, conf=0.9)
        time.sleep(1)
    time.sleep(1)

    while not _find_img('opcao.png', conf=0.9):
        time.sleep(1)
        _click_img('personaliza.png', conf=0.9)
        time.sleep(1)
    time.sleep(1)

    while not _find_img('rescisao.png', conf=0.9):
        time.sleep(1)
        _click_img('opcao.png', conf=0.9)
        time.sleep(1)

    while not _find_img('data_pagamento.png', conf=0.9):
        time.sleep(1)
        _click_img('rescisao.png', conf=0.9)
        time.sleep(1)
    time.sleep(1)

    while not _find_img('antecipar_data_pagamento.png', conf=0.9):
        time.sleep(1)
        _click_position_img('salario.png', '+', pixels_y=33, conf=0.9)
        time.sleep(1)
    time.sleep(1)

    while not _find_img('antecipar.png', conf=1):
        time.sleep(1)
        _click_img('antecipar_data_pagamento.png', conf=0.9)
        time.sleep(1)
        _click_position_img('salario.png', '+', pixels_y=33, conf=0.9)
        time.sleep(3)
    time.sleep(1)

    p.hotkey('alt', 'g')

    while not _find_img('img.png', conf=0.9):
        time.sleep(1)
        if _find_img('calculo.png', conf=0.9):
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(2)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Existem calculos apos a data de vigencia']), nome=andamentos)
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'
        if _find_img('banco.png', conf=0.9):
            time.sleep(1)
            p.press('esc')
            time.sleep(2)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'A empresa informada já iniciou o envio das informaçoes em outro banco de dados']),
                                   nome=andamentos)
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'

        if _find_img('e_social_erro.png', conf=0.9):
            time.sleep(1)
            p.press('enter')
            time.sleep(3)
        if _find_img('previdencia.png', conf=0.9):
            time.sleep(1)
            p.hotkey('alt', 'y')
            time.sleep(3)
            if _find_img('calculo.png', conf=0.9):
                time.sleep(1)
                p.hotkey('alt', 'n')
                time.sleep(2)
                _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Existem calculos apos a data de vigencia']),
                                       nome=andamentos)
                time.sleep(2)
                p.press('esc', presses=5)
                return 'ok'

        if _find_img('provisao.png', conf=0.9):
            time.sleep(1)
            p.hotkey('alt', 'o')
            time.sleep(2)
            p.press('esc')
            time.sleep(2)
            _escreve_relatorio_csv(';'.join(
                [cod, cnpj, nome, 'É necessário configurar o grupo Configuração do relatório de provisão das férias calculadas no mês.']),
                                   nome=andamentos)
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'

        if _find_img('conforme.png', conf=0.9):
            time.sleep(1)
            p.hotkey('alt', 'y')
            time.sleep(3)
            if _find_img('banco.png', conf=0.9):
                time.sleep(1)
                p.press('esc')
                time.sleep(2)
                _escreve_relatorio_csv(';'.join([cod, cnpj, nome,
                                                 'A empresa informada já iniciou o envio das informaçoes em outro banco de dados']),
                                       nome=andamentos)
                time.sleep(2)
                p.press('esc', presses=5)
                return 'ok'
        if _find_img('centralizada.png', conf=0.9):
            p.press('enter')
            time.sleep(2)
            p.press('esc')
            time.sleep(2)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome,
                                             'A empresa centralizadora para envio dos eventos Empregador, Reabertura e Fechamento eSocial deve possuir o mesmo responsavel legal']),
                                   nome=andamentos)
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'

    time.sleep(1)


    p.press('esc', presses=5)
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'OK']), nome=andamentos)
    return 'ok'

@_time_execution_monitor_db
def run(controle):

    empresas = _open_lista_dados()
    if not empresas:
        return False

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _login_web()
    _abrir_modulo('folha')


    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao, controle=controle, usando_bd=True, nome_rotina=andamentos +  f' - {_get_host_name()}', planilha=os.path.join('execução', andamentos + '.csv'))

        cod, cnpj, nome = empresa

        while True:

            if not _login(empresa, andamentos):
                break
            # abre a empresa no domínio
            resultado = flag_antecipar_data_pagamento_rescisao(cod, cnpj, nome)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    andamentos = 'Flag Antecipar Data de Pagamento da Rescição'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)
