import datetime, time, pyautogui as p
from sys import path
path.append(r'..\..\_comum')
from comum_comum import _indice, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _barra_de_status, _escreve_header_csv
from dominio_comum import _login, _login_web, _abrir_modulo
from pyautogui_comum import _find_img, _click_img, _wait_img

def plano_referencial(empresa, andamentos):
    cod, cnpj, nome = empresa
    retorno = abre_tela_vinculo_plano_referencial(cod, cnpj, nome)
    if retorno == 'ok':
        marca_as_duas_checkbox()
        time.sleep(5)
        status = seleciona_sped_contabil(cod, cnpj, nome)
        #status = 'ok'
        if status == 'ok':
            cont_cliente = conta_cliente()
            cont_fornecedor = conta_fornecedor()
            time.sleep(1)
            if _find_img('vazio.png', conf=0.9):
                status = 'vazio'
            else:
                status = 'simples/conta'
            p.press('esc', presses=5)
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{cont_cliente};{cont_fornecedor};{status}', nome=andamentos)
    return 'ok'

def abre_tela_vinculo_plano_referencial(cod, cnpj, nome):
    p.hotkey('alt', 'u')
    time.sleep(2)
    if _find_img('titulo_clique.png', conf=0.9):
        p.press('v')
    else:
        cont_cliente = 0
        cont_fornecedor = 0
        status = 'Alterar Cadastro não possui opção Vinculo Plano Referencial'
        _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{cont_cliente};{cont_fornecedor};{status}', nome=andamentos)
        p.press('esc')
        time.sleep(1)
        return 'erro'

    while not _find_img('titulo_vincular_plano_referencial.png', conf=0.9):
        time.sleep(1)
    time.sleep(3)
    return 'ok'

def marca_as_duas_checkbox():
    while _find_img('checkbox_vazia.png', conf=0.9):
        _click_img('checkbox_vazia.png', conf=0.9)
        time.sleep(1)

def seleciona_sped_contabil(cod, cnpj, nome):
    _click_img('seta_baixo_abrir_dropdown.png', conf=0.9)
    time.sleep(1)
    if _find_img('sped_contabil.png', conf=0.9):
        _click_img('sped_contabil.png', conf=0.9)
    else:
        cont_cliente = 0
        cont_fornecedor = 0
        status = 'Sped Contabil PJ em Geral não encontrado'
        _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{cont_cliente};{cont_fornecedor};{status}', nome=andamentos)
        p.press('esc', presses=5)
        return 'erro'

    while not _find_img('carregou.png', conf=0.9):
        time.sleep(0.5)
    return 'ok'

def conta_cliente():
    cont_cliente = 0
    while _find_img('conta_cliente.png', conf=0.95):
        _click_img('conta_cliente.png', conf=0.95)
        p.press('tab', presses=9, interval=0.1)
        #p.press('tab', presses=8, interval=0.1)
        p.write('1.01.02.02.01')
        _click_img('plano_cliente.png', conf=0.9)
        gravar_vinculo()
        cont_cliente += 1
    time.sleep(1)
    return cont_cliente

def conta_fornecedor():###
    cont_fornecedor = 0#
    while _find_img('conta_fornecedor.png', conf=0.95):
        _click_img('conta_fornecedor.png', conf=0.95)
        p.press('tab', presses=9, interval=0.1)
        #p.press('tab', presses=8, interval=0.1)
        p.write('2.01.01.03.01')
        _click_img('plano_fornecedor.png', conf=0.9)
        gravar_vinculo()
        cont_fornecedor += 1
        time.sleep(1)
    return cont_fornecedor
def gravar_vinculo():
    p.hotkey('alt', 'g')

@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('contabilidade')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        while True:

            if not _login(empresa, andamentos):
                break

            resultado = plano_referencial(empresa, andamentos)

            if resultado == 'ok':
                break
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;CLIENTE;FORNECEDOR;STATUS', nome=andamentos)

if __name__ == '__main__':

    empresas = _open_lista_dados()

    andamentos = 'Resultado Analisa Plano Referencial'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()