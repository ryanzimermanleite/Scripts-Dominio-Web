# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
import PySimpleGUI as sg
import shutil
import pandas as pd
from functools import wraps
from pyautogui import alert
import traceback
from threading import Thread
from sys import path
from xlrd import open_workbook
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv

data_anterior = []

def login(codigo, window):
    # espera a tela inicial do domínio#
    while not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por códig
    if _find_img('codigo.png', pasta='imgs_c', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(codigo)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        if _find_img('sem_parametro.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Parametro não cadastrado para esta empresa!')
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not _find_img('parametros.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if _find_img('nao_existe_parametro.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Não existe parametro cadastrado para esta empresa!')
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('empresa_nao_usa_sistema.png', pasta='imgs_c', conf=0.9) or _find_img(
                'empresa_nao_usa_sistema_2.png', pasta='imgs_c', conf=0.9):
            window['-Mensagens-'].update('Empresa não está marcada para usar este sistema!')
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('fase_dois_do_cadastro.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)

        if _find_img('aviso_regime.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)

        if _find_img('aviso.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'o')
            time.sleep(1)

        if _find_img('erro_troca_empresa.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5, interval=1)
            login(codigo, window)

    if not verifica_empresa(codigo):
        window['-Mensagens-'].update('Empresa não encontrada!')
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False

    p.press('esc', presses=5)
    time.sleep(1)
    return True

def verifica_empresa(empresa):
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
    sg.LOOK_AND_FEEL_TABLE['tema'] = {'BACKGROUND': '#3b4042',
                                      'TEXT': '#ffffff',
                                      'INPUT': '#f0f0f0',
                                      'TEXT_INPUT': '#000000',
                                      'SCROLL': '#ffffff',
                                      'BUTTON': ('#ffffff', '#313638'),
                                      'PROGRESS': ('#ffffff', '#ffffff'),
                                      'BORDER': 0,
                                      'SLIDER_DEPTH': 0,
                                      'PROGRESS_DEPTH': 0}
    @wraps(func)
    def wrapper():
        sg.theme('tema')  # Define o tema do PySimpleGUI
        bancos2 = ['Banco Neutro - 0', 'Banco Caixa - 5', 'Banco Itaú - 21', 'Santander - 1090', 'Banco - 386',
                   'Banco - 1060']
        modelo = ['Planilha modelo padrão - 1', 'Casa Bela Revestimentos Limitada']
        layout = [
            [sg.Text('Codigo:'), sg.InputText(key="-codigo-", size=(5, 5)),
             sg.FileBrowse('Arquivo', target='-input_excel_cliente-', key='-Abrir-', file_types=(('Planilhas Excel', '*.xlsx *.xls'),)),
             sg.InputText(key='-input_excel_cliente-',  size=20, disabled=True),
             sg.Text('Tipo:   '), sg.Radio("Pagar", "tipo", enable_events=True, key="-pagar-", default=True),
              sg.Radio("Receber", "tipo", enable_events=True, key="-receber-"),
             sg.Text('|'),
             sg.Text('Bancos:'),
              sg.Combo(bancos2, default_value='Banco Neutro - 0', font=("Helvetica", 10),size=(10,1), expand_x=True, enable_events=True, readonly=False, key='-COMBO-'),
             sg.Text('Modelos:'),
              sg.Combo(modelo, default_value='Planilha modelo padrão - 1', font=("Helvetica", 10),size=(30,1), expand_x=True, enable_events=True, readonly=False,
                       key='-COMBO-'),
             sg.Text('Status: '), sg.InputText("", key="texto", disabled=True, size=(10, 1)),
             sg.Button("Gerar", key='-gerar-', expand_x=True),
              sg.Button("Sair", expand_x=True)]

        ]

        # guarda a janela na variável para manipula-la
        screen_width, screen_height = sg.Window.get_screen_size()
        window = sg.Window('', layout, no_titlebar=True, location=(0, 0), size=(screen_width, 35), keep_on_top=True)

        def run_script_thread():
            try:
                func(window, values)
            except:
                pass

        while True:
            # captura o evento e os valores armazenados na interface
            event, values = window.read()

            if event == sg.WIN_CLOSED or event == 'Sair':
                break

            elif event == '-gerar-':
                # Cria uma nova thread para executar o script
                script_thread = Thread(target=run_script_thread)
                script_thread.start()

        window.close()
    return wrapper

@barra_de_status
def run(window, values):
    try:
        input_excel_cliente = values['-input_excel_cliente-']
        codigo = values["-codigo-"]
        tipo_pagar = values['-pagar-']
        tipo_receber = values['-receber-']
        banco = values['-COMBO-']
        numero_banco = banco.split('- ')
    except:
        input_excel_cliente = 'Desktop'

    if not all([input_excel_cliente, codigo]):
        alert(text=f'Por favor selecione uma planilha do Cliente.' + '\n\n'
                                                                     f'Digite o código do domínio referente a Empresa.')
    else:
        try:
            principal(input_excel_cliente, codigo, tipo_pagar, tipo_receber, banco, numero_banco, window)
        except Exception as e:
            # Obtém a pilha de chamadas de volta como uma string
            traceback_str = traceback.format_exc()
            alert(f'Traceback: {traceback_str}\n\n'
                  f'Erro: {e}')
            print(f'Traceback: {traceback_str}\n\n'
                  f'Erro: {e}')


def principal(input_excel_cliente, codigo, tipo_pagar, tipo_receber, banco, numero_banco, window):
    window["texto"].update('Rodando!')
    if not login(codigo, window):
        return
    if tipo_pagar is True:
         # RELATORIO A PAGAR
         titulo_tela = 'titulo_contas_a_pagar.png'
         tecla = 'p'
         imagem = 'aberta_paga_parcial.png'
         titulo_relatorio = 'contas_a_pagar_fornecedor.png'
         tipo = 'A Pagar'
         caminho = 'Contas a Pagar'
         nome_relatorio = 'Relatorio Contas a Pagar'
         # RELATORIO FORNECEDORES
         titulo_tela2 = 'titulo_fornecedores.png'
         tecla2 = 'f'
         titulo_relatorio2 = 'relacao_fornecedores.png'
         tipo2 = 'Fornecedores'
         caminho2 = 'Relação Fornecedores'
         nome_relatorio2 = 'Relatorio Fornecedores'
         # EXTRAI DADOS EXCEL
         tipo3 = 'pagar_fornecedor'

    if tipo_receber is True:
        # RELATORIO A RECEBER
        titulo_tela = 'titulo_contas_a_receber.png'
        tecla = 'r'
        imagem = 'aberta_recebida_parcial.png'
        titulo_relatorio = 'contas_a_receber_por_cliente.png'
        tipo = 'A Receber'
        caminho = 'Contas a Receber'
        nome_relatorio = 'Relatorio Contas a Receber'
        # RELATORIO CLIENTES
        titulo_tela2 = 'titulo_clientes.png'
        tecla2 = 'c'
        titulo_relatorio2 = 'relacao_clientes.png'
        tipo2 = 'Clientes'
        caminho2 = 'Relação Clientes'
        nome_relatorio2 = 'Relatorio Clientes'
        # EXTRAI DADOS EXCEL
        tipo3 = 'receber_cliente'

    arquivo_pagar_receber = gera_relatorio_pagar_receber(codigo, titulo_tela, tecla, imagem, titulo_relatorio, tipo,caminho, nome_relatorio)
    arquivo_fornecedor_cliente = gera_relatorio_fornecedor_cliente(codigo, titulo_tela2, tecla2, titulo_relatorio2,tipo2, caminho2, nome_relatorio2)
    dados_pagar_receber = open_lista_dados(arquivo_pagar_receber)
    dados_fornecedor_cliente = open_lista_dados(arquivo_fornecedor_cliente)
    extrai_dados_excel(codigo, dados_pagar_receber, dados_fornecedor_cliente, tipo3, input_excel_cliente, numero_banco, window)
    window["texto"].update('Finalizado!')
def gera_relatorio_pagar_receber(codigo, titulo_tela, tecla, imagem, titulo_relatorio, tipo, caminho, nome_relatorio):
    tela_contas_pagar_receber(titulo_tela, tecla, imagem)
    time.sleep(1)
    p.hotkey('alt', 'o')
    status = verifica_possui_relatorio(titulo_relatorio)
    if status == 'ok':
        salvar_pdf(codigo, tipo)
        arquivo = mover_arquivo(codigo, caminho, tipo, nome_relatorio)
    return arquivo

def tela_contas_pagar_receber(titulo, tecla, aberta_parcial):
    while not _find_img(titulo, conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('g')
        time.sleep(0.5)
        p.press(tecla)
        time.sleep(2)
    _click_img(aberta_parcial, conf=0.9)

def verifica_possui_relatorio(titulo_relatorio):
    while not _find_img(titulo_relatorio, conf=0.9):
        time.sleep(1)
        if _find_img('sem_dados.png', conf=0.9):
            print('❌ Sem dados para emitir')
            p.press('enter')
            time.sleep(0.5)
            p.press('esc')
            time.sleep(0.5)
            return 'sem_dados'
    return 'ok'

def salvar_pdf(cod, tipo):
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

def mover_arquivo(cod, caminho, tipo, nome_relatorio):
    pasta_origem = 'C:\\'
    pasta_destino = 'V:\Setor Robô\Scripts Python\Domínio\Baixa de Pagamento e Recebimento Contabil\execução\\' + caminho + '\\'
    nome_arquivo = str(cod) + ' - ' + tipo + '.xls'
    p.press('esc', presses=5)
    time.sleep(1)
    try:
        shutil.move(os.path.join(pasta_origem, nome_arquivo), os.path.join(pasta_destino, nome_arquivo))
    except:
        pass
    arquivo = pasta_destino + nome_arquivo
    return arquivo

def gera_relatorio_fornecedor_cliente(codigo, titulo_tela2, tecla2, titulo_relatorio2, tipo2, caminho2, nome_relatorio2):
    tela_clientes_fornecedor(titulo_tela2, tecla2)
    time.sleep(1)
    p.hotkey('alt', 'o')
    status = verifica_possui_relatorio(titulo_relatorio2)
    if status == 'ok':
        salvar_pdf(codigo, tipo2)
        arquivo = mover_arquivo(codigo, caminho2, tipo2, nome_relatorio2)
    return arquivo

def tela_clientes_fornecedor(titulo, tecla):
    while not _find_img(titulo, conf=0.9):
        p.hotkey('alt', 'a')
        time.sleep(0.5)
        p.press(tecla)
        time.sleep(5)
    time.sleep(1)
    p.press('l')
    while not _find_img('codigo_nome.png', conf=0.9):
        time.sleep(1)
    p.hotkey('alt', 'r')
    while not _find_img('listagem_relatorio.png', conf=0.9):
        time.sleep(1)
    _click_img('adicionar.png', conf=0.9)

def extrai_dados_excel(codigo, dados_pagar_receber, dados_fornecedor_cliente, tipo3, input_excel_cliente, numero_banco, window):
    dados_cliente, status = open_lista_dados_cliente(input_excel_cliente)
    global data_anterior
    if status == 'ok':
        nota_cliente = []

        for linha_cliente in dados_cliente.itertuples():
            numero_nota = linha_cliente[1]
            data_pagamento = linha_cliente[2]
            format = '%d/%m/%Y'
            data_pagamento = data_pagamento.strftime(format)
            valor_pagamento = linha_cliente[3]
            juros = linha_cliente[4]
            multa = linha_cliente[5]
            desconto = linha_cliente[6]
            nota_cliente.append([str(numero_nota), str(data_pagamento), str(valor_pagamento), str(juros), str(multa), str(desconto)])

            if tipo3 == 'pagar_fornecedor':
                nome_excel_encontrados = 'A Pagar Fornecedor'


            elif tipo3 == 'receber_cliente':
                nome_excel_encontrados = 'A Receber Cliente'

            gera_importação_dominio(dados_pagar_receber, dados_fornecedor_cliente, numero_nota, nome_excel_encontrados, data_pagamento, valor_pagamento, juros, multa, desconto, codigo, numero_banco, window)

        escreve_header(codigo, nome_excel_encontrados)
        gera_planilha_erros(codigo, nome_excel_encontrados, input_excel_cliente, numero_banco, window)

def open_lista_dados_cliente(input_excel_cliente):
    arquivo_cliente = ''
    try:
        arquivo_cliente = pd.read_excel(input_excel_cliente)
        return arquivo_cliente, 'ok'

    except:
        alert(text=f'Você selecionou uma planilha fora do padrão.')
        return arquivo_cliente, 'erro'

def gera_importação_dominio(dados_pagar_receber, dados_fornecedor_cliente, numero_nota, nome_excel, data_pagamento, valor_pagamento, juros, multa, desconto, codigo, numero_banco, window):
    notas_pagar_receber = []

    global data_anterior

    for count1, linha1 in enumerate(range(dados_fornecedor_cliente.nrows), start=1):
        codigo1 = dados_fornecedor_cliente.cell_value(linha1, 0)
        cnpj = dados_fornecedor_cliente.cell_value(linha1, 10)
        codigo1 = (str(codigo1).split('.'))
        for count, linha2 in enumerate(range(dados_pagar_receber.nrows), start=1):
            data_atual = []
            documento = dados_pagar_receber.cell_value(linha2, 16)
            #window["texto"].update(str(documento).strip())
            data_vencimento = dados_pagar_receber.cell_value(linha2, 19)
            codigo2 = dados_pagar_receber.cell_value(linha2, 31).split(' ')
            notas_pagar_receber.append(str(documento))

            if str(numero_nota) == str(documento):
                if codigo2[0] == codigo1[0]:

                    data_pagamento = str(data_pagamento)
                    data_atual.append([numero_nota, data_vencimento])
                    valor_pagamento = str(valor_pagamento).replace(".", ",")
                    juros = str(juros).replace(".", ",")
                    multa = str(multa).replace(".", ",")
                    desconto = str(desconto).replace(".", ",")
                    cnpj_formatado = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

                    for elemento_x in data_atual:

                        #window["texto"].update(str(elemento_x).strip())
                        if elemento_x not in data_anterior:
                            try:
                                _escreve_relatorio_csv(f'{numero_nota};{cnpj_formatado};{data_vencimento};{data_pagamento};{valor_pagamento};{juros};{multa};{desconto};{numero_banco[1]}', nome=codigo + ' - ' + nome_excel)
                            except Exception as erro:
                                print(erro)

                            data_anterior.append(elemento_x)
                            return
                        else:
                            continue

def escreve_header(codigo, nome_excel):
    try:
        _escreve_header_csv('Número da Nota;'
                            'CPF/CNPJ do fornecedor;'
                            'Data de Vencimento da parcela;'
                            'Data da baixa da parcela;'
                            'Valor pago;'
                            'Valor juros;'
                            'Valor multa;'
                            'Valor Desconto;'
                            'Código da conta banco;', nome=codigo + ' - ' + nome_excel)
    except:
        alert(text=f'NENHUMA NOTA FOI ENCONTRADA!')


def gera_planilha_erros(codigo, nome_excel_encontrados, input_excel_cliente, numero_banco, window):
    dados_resultado = []
    nota_data_valor = []

    dados_cliente, status = open_lista_dados_cliente(input_excel_cliente)

    for linha_cliente in dados_cliente.itertuples():
        numero_nota = linha_cliente[1]
        #window["texto"].update[str(numero_nota).strip()]
        data_pagamento = linha_cliente[2]
        format = '%d/%m/%Y'
        data_pagamento = data_pagamento.strftime(format)  # MEXI AQUI
        valor_pagamento = linha_cliente[3]
        juros = linha_cliente[4]
        multa = linha_cliente[5]
        desconto = linha_cliente[6]
        nota_data_valor.append([str(numero_nota), str(data_pagamento), str(valor_pagamento).replace(".", ","),
                                str(juros).replace(".", ","), str(multa).replace(".", ","),
                                str(desconto).replace(".", ",")])

    empresas = abre_arquivo_resultado(codigo, nome_excel_encontrados)
    if empresas is False:
        for linha_cliente in dados_cliente.itertuples():
            numero_nota = linha_cliente[1]
            #window["texto"].update[str(numero_nota).strip()]
            data_pagamento = linha_cliente[2]
            format = '%d/%m/%Y'
            data_pagamento = data_pagamento.strftime(format)  # MEXI AQU
            valor_pagamento = linha_cliente[3]
            juros = linha_cliente[4]
            multa = linha_cliente[5]
            desconto = linha_cliente[6]
            cnpj = 'Não Encontrado'
            data_venc = 'Não Encontrado'
            cnpj_formatado = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

            try:
                _escreve_relatorio_csv(
                    f'{str(numero_nota)};{cnpj_formatado};{data_venc};{str(data_pagamento)};{str(valor_pagamento).replace(".", ",")};{str(juros).replace(".", ",")};{str(multa).replace(".", ",")};{str(desconto).replace(".", ",")};{numero_banco[1]}',
                    nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')
            except Exception as erro:
                print(erro)
        try:
            _escreve_header_csv('Número da Nota;'
                                'CPF/CNPJ do fornecedor;'
                                'Data de Vencimento da parcela;'
                                'Data da baixa da parcela;'
                                'Valor pago;'
                                'Valor juros;'
                                'Valor multa;'
                                'Valor Desconto;'
                                'Código da conta banco;',
                                nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')
        except:
            alert(text=f'TODAS AS NOTAS FORAM ENCONTRADAS!.')
    else:
        for count, empresa in enumerate(empresas, start=1):
            nota_resultado = empresa[0]
            #window["texto"].update[str(nota_resultado).strip()]
            data_baixa = empresa[3]
            valor_resultado = empresa[4]
            juros_resultado = empresa[5]
            multa_resultado = empresa[6]
            desconto_resultado = empresa[7]
            dados_resultado.append(
                [str(nota_resultado), str(data_baixa), str(valor_resultado), str(juros_resultado), str(multa_resultado),
                 str(desconto_resultado)])
        #
        for elemento_y in nota_data_valor:
            #window["texto"].update[str(elemento_y).strip()]
            if elemento_y not in dados_resultado:
                cnpj = 'Não Encontrado'
                data_venc = 'Não Encontrado'
                cnpj_formatado = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

                try:
                    _escreve_relatorio_csv(
                        f'{elemento_y[0]};{cnpj_formatado};{data_venc};{elemento_y[1]};{elemento_y[2]};{elemento_y[3]};{elemento_y[4]};{elemento_y[5]};{numero_banco[1]}',
                        nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')
                except Exception as erro:
                    print(erro)
        try:
            _escreve_header_csv('Número da Nota;'
                                'CPF/CNPJ do fornecedor;'
                                'Data de Vencimento da parcela;'
                                'Data da baixa da parcela;'
                                'Valor pago;'
                                'Valor juros;'
                                'Valor multa;'
                                'Valor Desconto;'
                                'Código da conta banco;',
                                nome=codigo + ' - ' + nome_excel_encontrados + ' Não Encontrados')
        except:
            alert(text=f'TODAS AS NOTAS FORAM ENCONTRADAS!.')

def abre_arquivo_resultado(codigo, nome_excel_encontrados):
    file = 'V:\Setor Robô\Scripts Python\Domínio\Baixa de Pagamento e Recebimento Contabil\execução\\' + codigo + ' - ' + nome_excel_encontrados + '.csv'
    if not file:
        return False

    try:
        with open(file, 'r', encoding='latin-1') as f:
            dados = f.readlines()
    except Exception as e:
        return False

    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))

def open_lista_dados(input_excel):
    workbook = ''
    file = input_excel

    if not file:
        return False

    if file.endswith('.xls') or file.endswith('.XLS'):
        workbook = open_workbook(file)
        workbook = workbook.sheet_by_index(0)

    return workbook

if __name__ == '__main__':
    run()
