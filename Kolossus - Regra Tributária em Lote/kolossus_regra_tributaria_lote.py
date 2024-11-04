import datetime, shutil, os, time, pyautogui as p
from xlrd import open_workbook
from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login
# Adicione aqui os imports que estão faltando

def gerar(empresa, andamentos, tipo):
    cod, cnpj, nome = empresa

    _abrir_gerador_relatorios()

    # Seleciona o tipo de relatório
    if tipo == 'entrada':
        _selecionar_relatorio_entrada()
    elif tipo == 'saida':
        _selecionar_relatorio_saida()

    # Define o período
    _definir_periodo(ano, ano2)

    # Executa o relatório
    if not executar_relatorio(cod, tipo):
        return 'Erro na execução do relatório'

    # Move o arquivo para a pasta de destino
    arquivo = mover_arquivo(cod, tipo)

    # Processa o arquivo Excel
    processar_arquivo_excel(arquivo, tipo, nome)

    # Escreve no relatório de andamentos
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Gravação com sucesso!']), nome=andamentos)

    return 'ok'

def _abrir_gerador_relatorios():
    while not _find_img('gerenciador_relatorios.png', conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('r')
        time.sleep(0.5)
    p.press('up', presses=20, interval=0.1)

def _selecionar_relatorio_entrada():
    _click_img('conferencia_entradas.png', conf=0.9, clicks=2)
    time.sleep(1)
    _click_img('movimento_produto.png', conf=0.9, clicks=2)

def _selecionar_relatorio_saida():
    _click_img('conferencia_saida.png', conf=0.9, clicks=2)
    time.sleep(1)
    _click_img('saida.png', conf=0.9, clicks=2)

def _definir_periodo(ano, ano2):
    p.press('tab', presses=2, interval=0.2)
    time.sleep(0.5)
    p.write(ano)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=10)
    time.sleep(0.5)
    p.write(ano2)
    time.sleep(0.5)
    p.hotkey('alt', 'e')

def executar_relatorio(cod, tipo):
    if tipo == 'entrada':
        while not _find_img('titulo_entrada.png', conf=0.9):
            time.sleep(1)
            if _find_img('sem_dados.png', conf=0.9):
                p.press('enter')
                time.sleep(3)
                p.press('esc', presses=5)
                return False

    if tipo == 'saida':
        while not _find_img('titulo_saida.png', conf=0.9):
            time.sleep(1)
            if _find_img('sem_dados.png', conf=0.9):
                p.press('enter')
                time.sleep(3)
                p.press('esc', presses=5)
                return False

    salvar_arquivo(cod, tipo)
    p.press('esc', presses=5)
    return True

def salvar_arquivo(cod, tipo):
    p.click(833, 384)
    time.sleep(0.5)

    _click_img('salvar.png', conf=0.9)
    timer = 0
    _wait_img('salvar_relat.png', conf=0.9)
    time.sleep(0.5)
    _click_img('botao.png', conf=0.9)
    time.sleep(0.5)
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
    p.write(str(cod) + ' - ' + tipo + '.xls')
    time.sleep(0.5)

    if not _find_img('cliente_c_selecionado.png', pasta='imgs_c', conf=0.9):
        while not _find_img('cliente_c.png', pasta='imgs_c', conf=0.9) or _find_img('cliente_m.png', pasta='imgs_c',
                                                                                    conf=0.9):
            _click_img('botao.png', pasta='imgs_c', conf=0.9)
            time.sleep(3)

        _click_img('cliente_m.png', pasta='imgs_c', conf=0.9, timeout=1)
        _click_img('cliente_c.png', pasta='imgs_c', conf=0.9, timeout=1)
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

def mover_arquivo(cod, tipo):
    pasta_origem = 'C:\\'
    pasta_destino = 'V:\Setor Robô\Scripts Python\Domínio\\Kolossus - Regra Tributária em Lote\\execução\\' + tipo + '\\'
    nome_arquivo = str(cod) + ' - ' + tipo + '.xls'
    p.press('esc', presses=5)
    time.sleep(1)
    try:
        shutil.move(os.path.join(pasta_origem, nome_arquivo), os.path.join(pasta_destino, nome_arquivo))
    except:
        pass
    arquivo = pasta_destino + nome_arquivo
    return arquivo

def processar_arquivo_excel(caminho_arquivo, tipo, nome):
    workbook = open_lista_dados(caminho_arquivo)
    if not workbook:
        return  # Retornar se o arquivo não puder ser aberto

    for count1 in range(1, workbook.nrows):  # Começar a partir da segunda linha
        if tipo == 'entrada':
            CODIGO_PRODUTO = workbook.cell_value(count1, 12)
            DESCRICAO = workbook.cell_value(count1, 13)
            NCM = workbook.cell_value(count1, 14)
            CST_ORIGEM = str(workbook.cell_value(count1, 22)).split('.')[0]  # Convertendo para string e removendo decimais
            CST_ORIGEM = CST_ORIGEM.zfill(3)  # Adicionando zeros à esquerda
            CEST = ''
            NCM = NCM.zfill(8)
            print(NCM)

        elif tipo == 'saida':
            CODIGO_PRODUTO = workbook.cell_value(count1, 10)
            NCM = workbook.cell_value(count1, 12)
            CEST = workbook.cell_value(count1, 13)
            DESCRICAO = workbook.cell_value(count1, 14)
            CST_ORIGEM = str(workbook.cell_value(count1, 24)).split('.')[0]  # Convertendo para string e removendo decimais
            CST_ORIGEM = CST_ORIGEM.zfill(3)  # Adicionando zeros à esquerda
            NCM = NCM.zfill(8)


        EX = ''
        _escreve_relatorio_csv(f'{CST_ORIGEM[0]};{str(CODIGO_PRODUTO)};{str(DESCRICAO)};{str(NCM)};{str(EX)};{str(CEST)}', nome=nome)


def open_lista_dados(input_excel):
    workbook = ''
    file = input_excel

    if not file:
        return False

    if file.endswith('.xls') or file.endswith('.XLS'):
        workbook = open_workbook(file).sheet_by_index(0)

    return workbook


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
            if not _login(empresa, andamentos):#######
                break
            cod, cnpj, nome = empresa
            status_entrada = gerar(empresa, andamentos, 'entrada')
            status_saida = gerar(empresa, andamentos, 'saida')

            print('fimmmmmmm')
            _escreve_header_csv('CST_ORIGEM;'
                                'CODIGO_PRODUTO;'
                                'DESCRICAO;'
                                'NCM;'
                                'EX;'
                                'CEST;',
                                nome=nome)
            if status_entrada and status_saida == 'ok':
                break

    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    empresas = _open_lista_dados()  # Implemente a lógica para abrir a lista de empresas

    andamentos = 'Relatorio Geração Arquivos'
    ano = p.prompt(text='Qual periodo inicio?', title='Script incrível', default='00/00/0000')
    ano2 = p.prompt(text='Qual periodo final?', title='Script incrível', default='00/00/0000')

    index = _where_to_start(tuple(i[0] for i in empresas))  # Implemente a lógica para definir o índice inicial

    if index is not None:
        run()  # Executa o script
