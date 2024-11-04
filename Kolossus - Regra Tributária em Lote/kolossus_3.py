import shutil
import os
import time
import pandas as pd
import pyperclip
import pyautogui as p
import PySimpleGUI as sg
import traceback
from datetime import datetime
from functools import wraps
from threading import Thread
from xlrd import open_workbook
from sys import path
from pathlib import Path
from pyautogui import alert
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img

diretorio_importar = Path('execução/Relatorios Importar Kolossus')

def escreve_header_csv(texto, nome='resumo', local=diretorio_importar, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    with open(os.path.join(local, f"{nome}.csv"), 'r', encoding=encode) as f:
        conteudo = f.read()
    
    with open(os.path.join(local, f"{nome}.csv"), 'w', encoding=encode) as f:
        f.write(texto + '\n' + conteudo)

def escreve_relatorio_csv(texto, nome='resumo', local=diretorio_importar, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()


def login(empresa, window):
    window['-Mensagens-'].update('Fazendo login na empresa!')
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta='imgs', conf=0.9):
        time.sleep(1)

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por códig
    if _find_img('codigo.png', pasta='imgs', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs\codigo.png', confidence=0.9))
    p.write(empresa)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs', conf=0.9):
        time.sleep(1)
        if _find_img('sem_parametro.png', pasta='imgs', conf=0.9):
            window['-Mensagens-'].update('Parametro não cadastrado para esta empresa!')
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not _find_img('parametros.png', pasta='imgs', conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if _find_img('nao_existe_parametro.png', pasta='imgs', conf=0.9):
            window['-Mensagens-'].update('Não existe parametro cadastrado para esta empresa!')
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta='imgs', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('empresa_nao_usa_sistema.png', pasta='imgs', conf=0.9) or _find_img(
                'empresa_nao_usa_sistema_2.png', pasta='imgs', conf=0.9):
            window['-Mensagens-'].update('Empresa não está marcada para usar este sistema!')
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            while _find_img('trocar_empresa.png', pasta='imgs', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('fase_dois_do_cadastro.png', pasta='imgs', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta='imgs', conf=0.9):
            p.press('enter')
            time.sleep(1)

        if _find_img('aviso_regime.png', pasta='imgs', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)

        if _find_img('aviso.png', pasta='imgs', conf=0.9):
            p.hotkey('alt', 'o')
            time.sleep(1)

        if _find_img('erro_troca_empresa.png', pasta='imgs', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5, interval=1)
            login(empresa, window)

    if not verifica_empresa(empresa, window):
        window['-Mensagens-'].update('Empresa não encontrada!')
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False

    p.press('esc', presses=5)
    time.sleep(1)
    return True

def verifica_empresa(empresa, window):
    window['-Mensagens-'].update('Verificando empresa!')
    erro = 'sim'
    while erro == 'sim':
        try:
            p.click(1258, 82)

            while True:
                try:
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    cnpj_codigo = pyperclip.paste()
                    break
                except:
                    pass

            time.sleep(0.5)
            codigo = cnpj_codigo.split('-')
            codigo = str(codigo[1])
            codigo = codigo.replace(' ', '')

            if codigo != empresa:
                print(f'Código da empresa: {codigo}')
                print(f'Código encontrado no Domínio: {empresa}')
                return False
            else:
                return True
        except:
            erro = 'sim'

def barra_de_status(func):
    @wraps(func)
    def wrapper():
        sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
        layout = [
            [sg.Text('Codigo:'),
             sg.Input(key='-codigo-', size=(4, 1)),
             sg.Text('Inicio:'),
             sg.Input(key='-data_inicio-', size=(9, 1)),
             sg.Text('Final:'),
             sg.Input(key='-data_final-', size=(9, 1)),
             sg.Radio("JOÃO", "pessoa", key="-joao-", font=("Helvetica", 9)),
             sg.Radio("VINICIUS", "pessoa", key="-vinicius-", font=("Helvetica", 9)),
             sg.Button('RUN', key='-iniciar-', border_width=0, button_color='green1'),
             sg.Button('STOP', key='-stop-', border_width=0, button_color='yellow'),
             sg.Button('EXIT', key='-exit-', border_width=0, button_color='red'),
             sg.Text('', key='-Mensagens-', size=100)],
        ]

        # guarda a janela na variável para manipula-la
        screen_width, screen_height = sg.Window.get_screen_size()
        window = sg.Window('', layout, no_titlebar=True, location=(0, 0), size=(screen_width, 35), keep_on_top=True)

        def run_script_thread():
            try:
                func(window, values)
            except Exception as e:
                traceback_str = traceback.format_exc()
                alert(f'Traceback: {traceback_str}\n\n'
                      f'Erro: {e}')
                print(f'Traceback: {traceback_str}\n\n'
                      f'Erro: {e}')

        while True:
            # captura o evento e os valores armazenados na interface
            event, values = window.read()

            if event == sg.WIN_CLOSED or event == '-exit-':
                break

            elif event == '-iniciar-':
                # Cria uma nova thread para executar o script
                script_thread = Thread(target=run_script_thread)
                script_thread.start()

        window.close()
    return wrapper


@barra_de_status
def run(window, values):
    codigo = values['-codigo-']
    data_inicio = values['-data_inicio-']
    data_final = values['-data_final-']
    login(codigo, window)
    abrir_gerador_relatorios(window)
    selecionar_relatorio_saida()
    definir_periodo(data_inicio, data_final)
    tipo = 'JOÃO' if values['-joao-'] else 'VINICIUS'

    if not executar_relatorio(codigo, window):
        return 'Erro na execução do relatório'

    caminho_saida = mover_arquivo(codigo)
    nome_arquivo_importar = processar_arquivo_excel(caminho_saida, codigo, window, tipo)
    abrir_kolossus_auditor(window)
    importar_arquivo_kolossus(window, nome_arquivo_importar)
    verifica_mensagens_kolossus()
    insere_informacoes(tipo)
    atualiza_status()
    caminho_kolossus = baixar_arquivo_kolossus(codigo)
    time.sleep(5)
    saida = formata_planilha(codigo, caminho_saida)
    output_path = f'V:\Setor Robô\Scripts Python\Domínio\Kolossus - Regra Tributária em Lote\execução\Relatorios Processados\\{codigo} - Processado.xlsx'
    window['-Mensagens-'].update('Processando Planilha Final!')

    adiciona_colunas_fixamente_2040(output_path, saida, caminho_kolossus)
    adiciona_colunas_fixamente_2046(output_path, caminho_kolossus)
    adiciona_colunas_fixamente_3093(output_path, caminho_kolossus)


    window['-Mensagens-'].update('Finalizado!')


def adiciona_colunas_fixamente_2040(output_path, saida, caminho_kolossus):
    try:
        while True:
            try:
                # Passo 1: Ler a planilha ' - Kolossus' na aba 2040, ignorando as 12 primeiras linhas.
                kolossus_df = pd.read_excel(caminho_kolossus, sheet_name='2040', skiprows=12)
                break
            except:
                print('2040')
                pass
    
        # Passo 2: Capturar os valores das colunas especificadas e adicioná-los a uma lista.
        kolossus_data = []
        for index, row in kolossus_df.iterrows():
            kolossus_data.append(
                [row['CODIGO_PRODUTO'], row['ALIQ_ICMS_INTRA'], row['CONDICIONAL_REDUCAO'], row['CARGA_TRIB_BC_REDUZIDA'],
                 row['ALERTA_SOBRE_REDUCAO']])
    
        # Passo 3: Remover elementos duplicados da lista.
        kolossus_data_unique = [list(x) for x in set(tuple(x) for x in kolossus_data)]
    
        # Passo 4: Ler a planilha Processada.
        processada_df = pd.read_excel(saida)
    
        # Adicionando colunas vazias inicialmente
        for i in range(1, 9):
            processada_df[f'2040-Aliquota Interna {i}'] = None
    
        # Passo 5-7: Verificar e retornar os elementos correspondentes na planilha Processada.
        for index, row in processada_df.iterrows():
            codigo = row['codi_pdi']
    
            matching_rows = [item for item in kolossus_data_unique if item[0] == codigo]
            if matching_rows:
    
                for i, match in enumerate(matching_rows[:2]):  # Assume no more than 2 matches
                    base_idx = 1 + 4 * i
                    processada_df.at[index, f'2040-Aliquota Interna {base_idx}'] = match[1]
                    processada_df.at[index, f'2040-Aliquota Interna {base_idx + 1}'] = match[2]
                    processada_df.at[index, f'2040-Aliquota Interna {base_idx + 2}'] = match[3]
                    processada_df.at[index, f'2040-Aliquota Interna {base_idx + 3}'] = match[4]
    
    except ValueError as e:
        if "Worksheet named '2040' not found" in str(e):
            processada_df = pd.read_excel(output_path)
            for i in range(1, 9):
                processada_df[f'2040-Aliquota Interna {i}'] = "Não tem aba"
        
        # Salvar as alterações na planilha Processada
    processada_df.to_excel(output_path, index=False)

def adiciona_colunas_fixamente_2046(output_path, caminho_kolossus):
    try:
        kolossus_df = pd.read_excel(caminho_kolossus, sheet_name='2046', skiprows=12)
        kolossus_data = [[row['COD_PRODUTO'], row['CST']] for index, row in kolossus_df.iterrows()]
        kolossus_data_unique = [list(x) for x in set(tuple(x) for x in kolossus_data)]
    
        processada_df = pd.read_excel(output_path)
        for i in range(1, 3):
            processada_df[f'2046 - Aliquota de PIS e COFINS {i}'] = None
    
        for index, row in processada_df.iterrows():
            codigo = row['codi_pdi']
            matching_rows = [item for item in kolossus_data_unique if item[0] == codigo]
            for i, match in enumerate(matching_rows[:2]):  # Assume no more than 2 matches
                processada_df.at[index, f'2046 - Aliquota de PIS e COFINS {i+1}'] = match[1]
       
    except ValueError as e:
        if "Worksheet named '2046' not found" in str(e):
            processada_df = pd.read_excel(output_path)
            for i in range(1, 3):
                processada_df[f'2046 - Aliquota de PIS e COFINS {i}'] = "Não tem aba"
    
    processada_df.to_excel(output_path, index=False)
    
def adiciona_colunas_fixamente_3093(output_path, caminho_kolossus):
    try:
        kolossus_df = pd.read_excel(caminho_kolossus, sheet_name='3093', skiprows=12)
        kolossus_data = [[row['CODIGO_PRODUTO'], f"{row['FUNDAMENTO']} - {row['ALERTA']}"] for index, row in kolossus_df.iterrows()]
        kolossus_data_unique = [list(x) for x in set(tuple(x) for x in kolossus_data)]
    
        processada_df = pd.read_excel(output_path)
        for i in range(1, 5):
            processada_df[f'3093 - Isenção {i}'] = None
    
        for index, row in processada_df.iterrows():
            codigo = row['codi_pdi']
            matching_rows = [item for item in kolossus_data_unique if item[0] == codigo]
            for i, match in enumerate(matching_rows[:4]):  # Assume no more than 4 matches
                processada_df.at[index, f'3093 - Isenção {i+1}'] = match[1]
                
    except ValueError as e:
        if "Worksheet named '3093' not found" in str(e):
            processada_df = pd.read_excel(output_path)
            for i in range(1, 5):
                processada_df[f'3093 - Isenção {i}'] = "Não tem aba"
    
    processada_df.to_excel(output_path, index=False)


def formata_planilha(codigo, caminho_saida):
    caminho_saida_atualizado = f'V:\Setor Robô\Scripts Python\Domínio\Kolossus - Regra Tributária em Lote\execução\Relatorios Saida Dominio\\{codigo} - Saida Atualizado.xlsx'
    dados = pd.read_excel(caminho_saida, engine='xlrd')
    # Salvar como .xlsx usando o openpyxl como engine
    dados.to_excel(caminho_saida_atualizado, index=False, engine='openpyxl')

    # Carregar a planilha
    dados = pd.read_excel(caminho_saida_atualizado, engine='openpyxl')

    # Remover espaços em branco da coluna 'codi_pdi'
    dados['codi_pdi'] = dados['codi_pdi'].astype(str).str.strip()

    # Salvar como .xlsx usando o openpyxl como engine
    dados.to_excel(caminho_saida_atualizado, index=False, engine='openpyxl')
#
    return caminho_saida_atualizado

def abrir_gerador_relatorios(window):
    window['-Mensagens-'].update('Abrindo gerenciador de relatórios!')
    while not _find_img('gerenciador_relatorios.png', conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('r')
        time.sleep(0.5)
    p.press('up', presses=20, interval=0.1)

def selecionar_relatorio_saida():
    _click_img('conferencia_saida.png', conf=0.9, clicks=2)
    time.sleep(1)
    _click_img('saida.png', conf=0.9, clicks=2)

def definir_periodo(inicio, final):
    p.press('tab')
    time.sleep(0.5)
    p.write(inicio)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write(final)
    time.sleep(0.5)
    p.hotkey('alt', 'e')

def executar_relatorio(codigo, window):
    while not _find_img('titulo_saida.png', conf=0.9):
        time.sleep(1)
        if _find_img('sem_dados.png', conf=0.9):
            p.press('enter')
            time.sleep(3)
            p.press('esc', presses=5)
            return False

    salvar_arquivo(codigo, window)
    p.press('esc', presses=5)
    return True

def salvar_arquivo(codigo, window):
    window['-Mensagens-'].update('Salvando relatório!')
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
    p.write(str(codigo) + ' - Saida.xls')
    time.sleep(0.5)

    if not _find_img('cliente_c_selecionado.png', pasta='imgs', conf=0.9):
        while not _find_img('cliente_c.png', pasta='imgs', conf=0.9) or _find_img('cliente_m.png', pasta='imgs',
                                                                                    conf=0.9):
            _click_img('botao.png', pasta='imgs', conf=0.9)
            time.sleep(3)

        _click_img('cliente_m.png', pasta='imgs', conf=0.9, timeout=1)
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

def mover_arquivo(codigo):
    pasta_origem = 'C:\\'
    pasta_destino = 'V:\Setor Robô\Scripts Python\Domínio\\Kolossus - Regra Tributária em Lote\\execução\\Relatorios Saida Dominio\\'
    nome_arquivo = str(codigo) + ' - Saida.xls'
    p.press('esc', presses=5)
    time.sleep(1)
    
    try:
        shutil.move(os.path.join(pasta_origem, nome_arquivo), os.path.join(pasta_destino, nome_arquivo))
    except:
        pass
    caminho_saida = pasta_destino + nome_arquivo

    return caminho_saida

def processar_arquivo_excel(caminho_saida, codigo, window, tipo):

    window['-Mensagens-'].update('Processando arquivo excel!')
    workbook = open_lista_dados(caminho_saida)

    nome_arquivo_importar = f'{str(codigo)} - Importar'
    if not workbook:
        return  # Retornar se o arquivo não puder ser aberto

    dados_para_escrever = []
    for count1 in range(1, workbook.nrows):  # Começar a partir da segunda linha

        codigo_produto = workbook.cell_value(count1, 8).strip()
        ncm = workbook.cell_value(count1, 9).strip()
        descricao = workbook.cell_value(count1, 10).strip()
        descricao_formatada = descricao.replace(';', '')
        ncm = ncm.zfill(8)
        cest = ''
        ex = ''
        if tipo == 'JOÃO':
            cst_origem = '0'
            linha = f'{str(cst_origem)};{str(codigo_produto)};{str(descricao_formatada)};{str(ncm)};{str(ex)};{str(cest)}'
        else:
            cst_origem = str(workbook.cell_value(count1, 19)).split('.')[0]  # Convertendo para string e removendo decimais
            cst_origem = cst_origem.zfill(3)  # Adicionando zeros à esquerda
            linha = f'{str(cst_origem[0])};{str(codigo_produto)};{str(descricao_formatada)};{str(ncm)};{str(ex)};{str(cest)}'
            
        if linha not in dados_para_escrever:
            dados_para_escrever.append(linha)
    for linha in dados_para_escrever:

        escreve_relatorio_csv(f'{str(linha)}', nome=nome_arquivo_importar)

    escreve_header_csv('CST_ORIGEM;'
                        'CODIGO_PRODUTO;'
                        'DESCRICAO;'
                        'NCM;'
                        'EX;'
                        'CEST',
                        nome=nome_arquivo_importar)
    window['-Mensagens-'].update('Planilha gerada com sucesso!')

    return nome_arquivo_importar

def open_lista_dados(input_excel):

    workbook = ''
    file = input_excel
    if not file:
        return False
    if file.endswith('.xls') or file.endswith('.XLS'):
        workbook = open_workbook(file).sheet_by_index(0)
    return workbook

def abrir_kolossus_auditor(window):
    window['-Mensagens-'].update('Abrindo Kolossus Auditor!')
    _click_img('dominio.png', conf=0.9)
    while not _find_img('kolossus_auditor.png', conf=0.9):
        time.sleep(1)

    _click_img('kolossus_auditor.png', conf=0.9)
    while not _find_img('titulo_kolossus.png', conf=0.9):
        print('123')
        time.sleep(1)
        #_click_img('zoom.png', conf=0.9)
    time.sleep(3)
def importar_arquivo_kolossus(window, nome_arquivo_importar):

    window['-Mensagens-'].update('Importando arquivo gerado!')

    _click_img('regra_tributaria.png', conf=0.9)
    time.sleep(1)

    while _find_img('dicas_uso.png', conf=0.9):
        p.click(801, 183)
        time.sleep(1)

    #_click_img('x.png', conf=0.9)

    time.sleep(1)
    _click_img('meus_processos.png', conf=0.9)
    time.sleep(0.5)
    _click_img('meus_processos.png', conf=0.9)
    time.sleep(3)
    try:
        _click_img('upload.png', conf=0.9)
    except:
        pass
    
    while not _find_img('upload2.png', conf=0.9):
        time.sleep(1)
   
    while not _find_img('open.png', conf=0.9):
        time.sleep(1)
        try:
            _click_img('upload2.png', conf=0.9)
            time.sleep(1)
        except:
            pass
        
    time.sleep(1)
    p.write(nome_arquivo_importar)
    p.press('tab', presses=4, interval=0.1)
    time.sleep(0.5)
    p.press('enter')
    caminho2 = 'V:\Setor Robô\Scripts Python\Domínio\Kolossus - Regra Tributária em Lote\execução\Relatorios Importar Kolossus'
    pyperclip.copy(caminho2)
    time.sleep(0.5)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(2)
    p.hotkey('alt', 'o')
    time.sleep(1)
    p.hotkey('alt', 'o')
    time.sleep(1)
    p.hotkey('alt', 'o')
    
def verifica_mensagens_kolossus():
    while not _find_img('protocolo.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)

    while not _find_img('filtrar.png', conf=0.9):
        try:
            _click_img('voltar.png', conf=0.9)
        except:
            pass
        time.sleep(1)
        if _find_img('dicas_uso.png', conf=0.9):
            p.click(801, 183)
    time.sleep(1)
    _click_img('atualizar.png', conf=0.9)
    time.sleep(1)
    if _find_img('dicas_uso.png', conf=0.9):
        p.click(801, 183)

    while not _find_img('quadrado_verde.png', conf=0.9):
        if _find_img('dicas_uso.png', conf=0.9):
            p.click(801, 183)
        p.click(339, 379)
        time.sleep(2)

    while not _find_img('processar_tributos.png', conf=0.9):
        if _find_img('dicas_uso.png', conf=0.9):
            p.click(801, 183)
        _click_img('atualizar.png', conf=0.9)
        time.sleep(2)
        p.click(339, 379)
        time.sleep(3)

    _click_img('processar_tributos.png', conf=0.9)

def insere_informacoes(tipo):
    mes = datetime.now().month
    ano = datetime.now().year
    
    while not _find_img('ipi_sim.png', conf=0.9):
        _click_img('ipi_nao.png', conf=0.9)
        time.sleep(1)
    while not _find_img('pis_cofins_sim.png', conf=0.9):
        _click_img('pis_cofins_nao.png', conf=0.9)
        time.sleep(1)
    while not _find_img('imposto_sim.png', conf=0.9):
        _click_img('imposto_nao.png', conf=0.9)
        time.sleep(1)
    while not _find_img('cprb_sim.png', conf=0.9):
        _click_img('cprb_nao.png', conf=0.9)
        time.sleep(1)
    while not _find_img('icms_interestadual_sim.png', conf=0.9):
        _click_img('icms_interestadual_nao.png', conf=0.9)
        time.sleep(1)
    while not _find_img('icms_interno_sim.png', conf=0.9):
        _click_img('icms_interno_nao.png', conf=0.9)#
        time.sleep(1)
    while not _find_img('cest_sim.png', conf=0.9):
        _click_img('cest_nao.png', conf=0.9)
        time.sleep(1)
    while not _find_img('icms_produtos_sim.png', conf=0.9):
        _click_img('icms_produtos_nao.png', conf=0.9)
        time.sleep(1)
        
    _click_img('data.png', conf=0.9)
    time.sleep(1)
    p.write(str(mes))
    time.sleep(1)
    p.press('right')
    time.sleep(1)
    p.write(str(ano))
    _click_img('estados_origem.png', conf=0.9)
    time.sleep(1)
    p.write('SP - ')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    _click_img('estados_destino.png', conf=0.9)
    time.sleep(1)
    p.write('Todos')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    _click_img('atividades.png', conf=0.9)
    time.sleep(1)
    if tipo == 'JOÃO':
        p.write('VAREJISTA (ATIVIDADE DE COMERCIO)')
    else:
        p.write('Todos')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    _click_img('regime.png', conf=0.9)
    time.sleep(1)
    p.write('Todos')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    _click_img('estados_internos.png', conf=0.9)
    time.sleep(1)
    p.write('SP - ')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.press('tab')
    time.sleep(1)
    _click_img('aplicar.png', conf=0.9)
    
def atualiza_status():

    while not _find_img('meus_processos3.png', conf=0.9):
        time.sleep(1)
        _click_img('meus_processos2.png', conf=0.9)
        time.sleep(1)

    while not _find_img('exportar3.png', conf=0.9):
        if _find_img('dicas_uso.png', conf=0.9):
            p.click(801, 183)
            time.sleep(1)
        while not _find_img('quadrado_verde.png', conf=0.9):
            if _find_img('dicas_uso.png', conf=0.9):
                p.click(801, 183)
                time.sleep(1)
            p.click(339, 379)
            time.sleep(1)
        time.sleep(3)
        if _find_img('exportar3.png', conf=0.9):
            break
        else:
            _click_img('atualizar.png', conf=0.9)
            time.sleep(1)

    _click_img('exportar3.png', conf=0.9)


    while not _find_img('download3.png', conf=0.9):
        if _find_img('dicas_uso.png', conf=0.9):
            p.click(801, 183)
            time.sleep(1)
        while not _find_img('quadrado_verde.png', conf=0.9):
            if _find_img('dicas_uso.png', conf=0.9):
                p.click(801, 183)
                time.sleep(1)
            p.click(339, 379)
            time.sleep(1)
        time.sleep(3)


        if _find_img('download3.png', conf=0.9):
            break
        else:
            _click_img('atualizar.png', conf=0.9)
            time.sleep(1)

    _click_img('download3.png', conf=0.9)

    while not _find_img('save_as.png', conf=0.9):
        time.sleep(1)
    
def baixar_arquivo_kolossus(codigo):
    time.sleep(3)
    _click_img('pasta.png', conf=0.9)
    time.sleep(1)
    caminho3 = 'V:\Setor Robô\Scripts Python\Domínio\Kolossus - Regra Tributária em Lote\execução\Relatorios Exportado Kolossus\\'
    pyperclip.copy(caminho3)
    time.sleep(1)
    pyperclip.copy(caminho3)
    time.sleep(1)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.press('enter')
    time.sleep(3)
    #p.press('tab', presses=6, interval=0.2)
    _click_img('regra_lote.png', conf=0.9)  #
    time.sleep(1)
    p.press('backspace')
    time.sleep(1)
    p.write(str(codigo) + ' - Kolossus')
    time.sleep(2)
    p.hotkey('alt', 's')
    caminho_kolossus = str(caminho3) + (str(codigo) + ' - Kolossus.xlsx')
    time.sleep(2)
    return caminho_kolossus

if __name__ == '__main__':
   
    run()

   