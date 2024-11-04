import datetime
import time
import pyautogui as p
from _comum.pyautogui_comum import _find_img, _click_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _get_host_name, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _escreve_header_csv
from _comum.dominio_comum import _login_web, _abrir_modulo, _login


def abre_tela_parametros_empresa():
    while not _find_img('parametros.png', conf=0.9):
        p.hotkey('alt', 'c')
        time.sleep(0.5)
        p.press('p')
        time.sleep(3)

    while not _find_img('lancamento.png', conf=0.9):
        time.sleep(1)
        _click_img('geral.png', conf=0.9)
        time.sleep(1)


    while not _find_img('localizador.png', conf=0.9):
        time.sleep(1)
        _click_img('lancamento.png', conf=0.9)
        time.sleep(1)

def flag_localizador(cod, nome, cnpj):
    while not _find_img('informar_localizador_ok.png', conf=1):
        time.sleep(1)
        _click_img('informar_localizador.png', conf=0.9)
        time.sleep(1)
        _click_img('lancamento.png', conf=0.9)
        time.sleep(1)
    p.hotkey('alt', 'g')
    time.sleep(3)
    if _find_img('somente_um_signatario.png', conf=0.9):
        p.hotkey('alt', 'o')
        time.sleep(1)
        p.hotkey('alt', 'f')
        time.sleep(1)
        if _find_img('deseja_gravar_dados.png', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc', presses=5)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, r"Somente um dos signatários poderá ter 'Sim' na coluna Responsável assinatura - ECD."]), nome=andamentos)
            return 'ok'

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
    _abrir_modulo('contabil')


    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao, controle=controle, usando_bd=True, nome_rotina=andamentos +  f' - {_get_host_name()}')

        cod, nome, cnpj = empresa

        while True:

            if not _login(empresa, andamentos):
                break
            # abre a empresa no domínio
            abre_tela_parametros_empresa()
            resultado = flag_localizador(cod, nome, cnpj)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    andamentos = 'Flag Informar Localizador na Tela de Lançamentos Contabil'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)


