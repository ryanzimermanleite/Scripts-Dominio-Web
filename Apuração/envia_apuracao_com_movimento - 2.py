# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Gera apuração                                               #
# Arquivo:  envia_apuracao.py                                           #
# Versão:   1.0.0                                                       #
# Modulo:   Dominio                                                     #
# Objetivo: Atualizar periodo de apuração das empresas                  #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     30/10/2023                                                  #
# ----------------------------------------------------------------------#
import datetime
import time
import pyautogui as p
import os
from _comum.pyautogui_comum import _find_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _get_host_name, _open_lista_dados, \
    _escreve_relatorio_csv, _where_to_start, _escreve_header_csv
from _comum.dominio_comum import _login_web, _abrir_modulo, _login


def apurar(mes_ano, empresa, andamentos):
    cod, cnpj = empresa

    # Se o mes escolhido for 12 ele def07/2024ine como 1 para a imagem
    mes = mes_ano.split('/')
    if mes[0] == '12':
        mes = '01'
    # Se não for mes 12 ele soma +1 ao mes escolhe para a imagem
    else:
        mes = str(int(mes[0]) + 1)

    # Se o mês da empresa estiver 1 mês na frente do mês escolhido ele pula a apuração
    while not _find_img('mes' + mes + '_.png', conf=0.9):

        if _find_img('empresas_2020.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração 2020']), nome=andamentos)
            return 'ok'
        if _find_img('empresas_2021.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração 2021']), nome=andamentos)
            return 'ok'
        if _find_img('empresas_2022.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração 2022']), nome=andamentos)
            return 'ok'
        # espera o botão de movimentos do domínio aparecer na tela
        # _wait_img('movimentos.png', conf=0.9, timeout=-1)

        print('>>> Apurando Empresa')
        # Enquanto não achar a janela de apuraçãp. Tenta abrir a janela de apuração
        while not _find_img('apuracao.png', conf=0.9):
            # Relatórios
            p.hotkey('alt', 'm')
            time.sleep(0.5)
            # gerador de relatórios
            p.press('r')
            time.sleep(0.5)
            if _find_img('apuracao_2.png', conf=0.9):
                break

        # Escreve o mes_ano que foi escolhi e aperta o comando para gerar novo periodo
        time.sleep(0.5)
        p.write(mes_ano)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.write(mes_ano)
        # gera novo periodo
        time.sleep(0.5)
        p.hotkey('alt', 'g')
        time.sleep(1)
        if _find_img('bloco_k.png', conf=0.9):
            p.hotkey('alt', 'y')
        # Equanto não achar imagem do processo de apuração espera 1 segundorrr
        while not _find_img('progresso_apuracao.png', conf=0.9):
            time.sleep(1)
            if _find_img('fora_horario.png', conf=0.9):
                p.press('enter')
                time.sleep(1)
                p.press('esc', presses=5)
                _escreve_relatorio_csv(';'.join([cod, cnpj, 'Fora de Periodo']), nome=andamentos)
                return 'ok'
            if _find_img('bloco_k.png', conf=0.9):
                p.hotkey('alt', 'y')

            if _find_img('fechamento.png', conf=0.9):
                p.press('enter')
                time.sleep(1)
                p.press('esc', presses=5)
                _escreve_relatorio_csv(';'.join([cod, cnpj, 'Periodo menor que o periodo de fechamento']),
                                       nome=andamentos)
                return 'ok'
            if _find_img('progresso_apuracao_2.png', conf=0.9):
                break

        # Enquanto não achar a imagem de fim de apuração ele procura pela imagem de avisos apuração
        while not _find_img('fim_apuracao.png', conf=0.9):
            time.sleep(1)
            if _find_img('bloco_k.png', conf=0.9):
                p.hotkey('alt', 'y')
            # Se encontrar avisos ele fecha
            if _find_img('avisos_apuracao.png', conf=0.9):
                p.press('f')
            elif _find_img('atencao.png', conf=0.9):
                p.click(833, 384)
                time.sleep(1)
                p.press('y')
            if _find_img('fim_apuracao_2.png', conf=0.9):
                break

        # Fecha a janela de apuração
        time.sleep(1)
        p.press('n')
        time.sleep(0.5)
        p.press('esc')
        p.press('n')

        print('✔ Apuração Concluida')
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração com Sucesso']), nome=andamentos)
        # fechar qualquer possível tela aberta
        p.press('esc', presses=5)
        time.sleep(1)
        return 'ok'
    _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa já apurada!']), nome=andamentos)
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
    _abrir_modulo('escrita_fiscal')

    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index=index, tempos=tempos,
                                         tempo_execucao=tempo_execucao, controle=controle, usando_bd=True,
                                         nome_rotina=andamentos + f' - {_get_host_name()}',
                                         planilha=os.path.join('execução', andamentos + '.csv'))

        while True:
            # abre a empresa no domínio
            if not _login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = apurar(str(ano), empresa, andamentos)

            if resultado == 'dominio fechou':
                _login_web()
                _abrir_modulo('escrita_fiscal')

            if resultado == 'modulo fechou':
                _abrir_modulo('escrita_fiscal')

            if resultado == 'ok':
                break
    _escreve_header_csv('CÓDIGO;CNPJ', nome=andamentos)


if __name__ == '__main__':
    ano = p.prompt(text='Qual periodo base?', title='Script incrível', default='00/0000')
    andamentos = 'Apuração DAIANI Com Movimento Invertida'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)

