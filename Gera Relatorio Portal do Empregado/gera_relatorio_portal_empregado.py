import datetime, time, pyautogui as p
import os
import shutil
from _comum.pyautogui_comum import _find_img, _click_img, _wait_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _escreve_header_csv, _get_host_name
from _comum.dominio_comum import _login_web, _abrir_modulo, _login


def gera_relatorio_portal_empregado(empresa, andamentos):
    cod, cnpj, nome = empresa
    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('gerenciador_de_relatorios.png', conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('i', presses=2)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(2)

    time.sleep(0.5)
    p.press('down', presses=20)
    while not _find_img('relacao_empregados.png', conf=0.9):
        p.press('up')
        time.sleep(1)

    while not _find_img('portal_empregado.png', conf=0.9):
        time.sleep(1)
        _click_img('relacao_empregados.png', conf=0.9, clicks=2)
        time.sleep(2)

    while not _find_img('empregado_todos.png', conf=0.9):
        time.sleep(1)
        _click_img('portal_empregado.png', conf=0.9, clicks=2)
        time.sleep(2)

    time.sleep(1)
    p.hotkey('alt', 'e')
    while not _find_img('codigo_nome.png', conf=0.9):
        time.sleep(1)
    salvar_pdf(cod)
    mover_arquivo((cod) + ' - Portal Empregado' + '.xls')
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatorio Gerado']), nome=andamentos)
    return 'ok'


def salvar_pdf(cod):
    p.click(833, 384)
    time.sleep(0.5)

    _click_img('salvar.png', conf=0.9)
    timer = 0

    while not _find_img('arquivo_relat_2.png', conf=0.9):
        time.sleep(1)
    while not _find_img('planilha_cabecalho.png', conf=0.9):
        _click_img('botao_2.png', conf=0.9)
        time.sleep(2)

    _click_img('planilha_cabecalho.png', conf=0.9)
    time.sleep(0.5)
    _click_img('3pontos.png', conf=0.9)
    time.sleep(0.5)

    while not _find_img('selecione_arquivo.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False

    time.sleep(1)
    p.write(str(cod) + ' - Portal Empregado' + '.xls')
    time.sleep(0.5)

    if not _find_img('cliente_c_selecionado.png', pasta='imgs', conf=0.9):
        while not _find_img('cliente_c.png', pasta='imgs', conf=0.9) or _find_img('cliente_m.png', pasta='imgs',
                                                                                    conf=0.9):
            _click_img('botao.png', pasta='imgs', conf=0.9)
            time.sleep(3)

        #_click_img('cliente_m.png', pasta='imgs', conf=0.9, timeout=1)
        _click_img('cliente_c.png', pasta='imgs', conf=0.9, timeout=1)
        time.sleep(5)

    time.sleep(1)
    p.hotkey('alt', 's')
    _wait_img('salvar_fechar.png', conf=0.9)
    time.sleep(1)
    p.hotkey('alt', 's')
    while not _find_img('img.png', conf=0.9):
        time.sleep(5)
        p.press('esc')
        if _find_img('gravar_dados.png', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
    time.sleep(1)
    p.press('esc', presses=5, interval=0.2)
    return True

def mover_arquivo(nome_arquivo):
    pasta_origem = 'C:\\'
    pasta_destino = f'V:\Setor Robô\Scripts Python\Domínio\Gera Relatorio Portal do Empregado\execução\Relatorios\\'
    p.press('esc', presses=5)
    time.sleep(1)

    try:
        shutil.move(os.path.join(pasta_origem, nome_arquivo), os.path.join(pasta_destino, nome_arquivo))
    except:
        pass

    caminho = pasta_destino + nome_arquivo

    return caminho

@_time_execution_monitor_db
def run(window):

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
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao, controle=controle, usando_bd=True, nome_rotina=andamentos +  f' - {_get_host_name()}')
#
        while True:
            # abre a empresa no domínio
            if not _login(empresa, andamentos, ignora_sem_parametro=True):
                break
            # Chama a função de apurar
            resultado = gera_relatorio_portal_empregado(empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    andamentos = 'Gera Relatorio Portal do Empregado'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)
