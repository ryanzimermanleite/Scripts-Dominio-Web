import datetime
import time
import cv2
import pyautogui as p
from _comum.pyautogui_comum import _find_img, _click_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _get_host_name, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _escreve_header_csv
from _comum.dominio_comum import _login_web, _abrir_modulo, _login


def abre_tela_parametros_empresa():
    while not _find_img('parametros_empresa.png', conf=0.9):
        p.hotkey('alt', 'c')
        time.sleep(0.5)
        p.press('p')
        time.sleep(3)

    while not _find_img('geral_opcoes.png', conf=0.9):
        time.sleep(1)
        _click_img('geral.png', conf=0.9)
        time.sleep(1)
        _click_img('federal.png', conf=0.9)
        time.sleep(1)

    while not _find_img('dctf.png', conf=0.9):
        time.sleep(1)
        _click_img('opcao.png', conf=0.9)
        time.sleep(1)

def flag_dctfweb(cod, cnpj, nome):
    p.hotkey('alt', 'n')
    time.sleep(1)
    p.write(ano)
    time.sleep(1)

    while not _find_img('dctf_flagado.png', conf=1):
        time.sleep(1)
        _click_img('dctf.png', conf=0.9)
        time.sleep(1)
        _click_img('opcao.png', conf=0.9)
        time.sleep(1)
        if _find_img('dctf_bloqueado.png', conf=1):
            time.sleep(1)
            p.press('esc', presses=5)
            time.sleep(3)
            p.hotkey('alt', 'n')
            time.sleep(3)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Flag Bloqueada']), nome=andamentos)
            return 'ok'

    p.hotkey('alt', 'g')
    time.sleep(3)
    while not _find_img('gravar.png', conf=0.9):
        time.sleep(1)
        if _find_img('opcao_x.png', conf=0.9):
            p.hotkey('alt', 's')
        if _find_img('menor.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('maior.png', conf=0.9):
            p.press('enter')
        if _find_img('gerador.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('gerador_2.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('vigencia.png', conf=0.9):
            p.press('enter')
        if _find_img('vigencia_2.png', conf=0.9):
            p.press('enter')
        if _find_img('vigencia_4.png', conf=0.9):
            p.press('enter')
        if _find_img('icms.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('acumuladores.png', conf=0.9):
            p.hotkey('alt', 'y')#
        if _find_img('acumuladores_2.png', conf=0.9):
            p.press('enter')
        if _find_img('modulo.png', conf=0.9):
            p.press('enter')
        if _find_img('lancado.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('regime.png', conf=0.9):
            p.hotkey('alt', 's')
        if _find_img('simples.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('apuracao.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('rateio.png', conf=0.9):
            p.press('enter')
        if _find_img('periodo.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('certificado.png', conf=0.9):
            p.press('enter')
        if _find_img('filial.png', conf=0.9):
            p.press('enter')
        if _find_img('bloco_k.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('cupom.png', conf=0.9):
            p.press('enter')#

        if _find_img('informativo.png', conf=0.9):
            p.press('enter')
            time.sleep(3)
            p.press('esc')
            time.sleep(3)
            p.hotkey('alt', 'n')
            time.sleep(3)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Selecionar o perfil correspondente no enquadramento estadual da empresa.']),
                                   nome=andamentos)
            return 'ok'

        if _find_img('vigencia_3.png', conf=0.9):
            p.press('enter')
            time.sleep(3)
            p.press('esc')
            time.sleep(3)
            p.hotkey('alt', 'n')
            time.sleep(3)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Período de vigência inferior a data de Fechamento.']),
                                   nome=andamentos)
            return 'ok'




        if _find_img('uf.png', conf=0.9):
            p.press('enter')
            time.sleep(3)
            p.press('esc')
            time.sleep(3)
            p.hotkey('alt', 'n')
            time.sleep(3)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'O cliente deve ser da mesma UF da empresa.']), nome=andamentos)
            return 'ok'




        if _find_img('cadastrado.png', conf=0.9):
            p.press('enter')
            time.sleep(3)
            p.press('esc')
            time.sleep(3)
            p.hotkey('alt', 'n')
            break
        if _find_img('cadastrado_2.png', conf=0.9):
            p.press('enter')
            time.sleep(3)
            p.press('esc')
            time.sleep(3)
            p.hotkey('alt', 'n')
            break



    time.sleep(2)
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
    _abrir_modulo('escrita-fiscal')


    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao, controle=controle, usando_bd=True, nome_rotina=andamentos +  f' - {_get_host_name()}')

        cod, cnpj, nome = empresa

        while True:

            if not _login(empresa, andamentos):
                break
            # abre a empresa no domínio
            abre_tela_parametros_empresa()
            resultado = flag_dctfweb(cod, cnpj, nome)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    ano = p.prompt(text='Qual periodo base?', title='Script incrível', default='00/0000')
    andamentos = 'Flag Cálculo Tributos Federais DCTFWeb Fiscal'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)


