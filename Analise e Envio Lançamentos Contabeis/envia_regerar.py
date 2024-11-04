# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Envia Regerar                                               #
# Arquivo:  envia_regerar.py                                            #
# Versão:   1.0.0                                                       #
# Modulo:   Dominio                                                     #
# Objetivo: Faz envio regerar                                           #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     07/11/2023                                                  #x
# ----------------------------------------------------------------------#
import datetime, pyperclip, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo

def salvar_pdf(cod):

    p.click(833, 384)
    time.sleep(0.5)
    p.hotkey('ctrl', 'd')
    time.sleep(1)
    timer = 0
    print('1')
    while not _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
        print('2')
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False

    time.sleep(1)
    print('3')
    p.write(str(cod) + ' - Erro Regerar')
    time.sleep(0.5)


    if not _find_img('cliente_v_selecionado.png', pasta='imgs_c', conf=0.9):
        print('4')
        while not _find_img('cliente_v.png', pasta='imgs_c', conf=0.9):
            print('5')
            _click_img('botao.png', pasta='imgs_c', conf=0.9)
            time.sleep(3)

    _click_img('cliente_v.png', pasta='imgs_c', conf=0.9, timeout=1)
    time.sleep(5)

    _click_img('pasta.png', conf=0.9, )
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 's')

    timer = 0
    while not _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        print('6')
        if _find_img('sera_finalizada.png', pasta='imgs_c', conf=0.9):
            p.press('esc')
            time.sleep(2)
            return False

        if _find_img('erro_pdf.png', pasta='imgs_c', conf=0.9) or _find_img('erro_pdf_2.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            p.hotkey('alt', 'f4')

        if _find_img('substituir.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('adobe.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
        time.sleep(1)
        timer += 1
        if timer > 30:
            p.hotkey('ctrl', 'd')
            while not _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
                print('7')
                time.sleep(1)

            if not _find_img('cliente_v_selecionado.png', pasta='imgs_c', conf=0.9):
                while not _find_img('cliente_v.png', pasta='imgs_c', conf=0.9):
                    _click_img('botao.png', pasta='imgs_c', conf=0.9)
                    time.sleep(3)
                _click_img('cliente_v.png', pasta='imgs_c', conf=0.9)
                time.sleep(5)

            p.press('enter')
            timer = 0

    while _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        print('8')
        p.hotkey('alt', 'f4')
        time.sleep(3)

    while _find_img('sera_finalizada.png', pasta='imgs_c', conf=0.9):
        print('9')
        p.press('esc')
        time.sleep(2)
    p.press('esc', presses=5)

    return True
def verifica_empresa(cod):
    erro = 'sim'
    while erro == 'sim':
        try:
            # Faz um clique no canto superior direito para verificar a empresa
            p.click(1258, 82)

            while True:
                try:
                    # Copia o nome
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
            # Compara o codigo digitado com o codigo da empresa que entrou se for diferente a empresa não existe
            if codigo != cod:
                print(f'Código da empresa: {codigo}')
                print(f'Código encontrado no Domínio: {cod}')
                return False
            else:
                return True
        except:
            erro = 'sim'


def login(empresa, andamentos):
    cod, cnpj, status = empresa
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    # Faz um clique no meio da tale pra dar um focus no dominio
    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta='imgs_c', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(cod)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        if _find_img('sem_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Parametro não cadastrado para esta empresa']), nome=andamentos)
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not _find_img('parametros.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        # Se der a mensagem que nao tem parametros escreve no relatorio e pula a empresa
        if _find_img('nao_existe_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Não existe parametro cadastrado para esta empresa']),
                                   nome=andamentos)
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

        # Se a empresa não usar o sistema escreve no relatorio e pula a empresa
        if _find_img('empresa_nao_usa_sistema.png', pasta='imgs_c', conf=0.9) or _find_img(
                'empresa_nao_usa_sistema_2.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa não está marcada para usar este sistema']),
                                   nome=andamentos)
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
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

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
            login(empresa, andamentos)

    if not verifica_empresa(cod):
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa não encontrada']), nome=andamentos)
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False

    p.press('esc', presses=5)
    time.sleep(1)

    return True


def regerar(ano, ano2, empresa, andamentos):
    cod, cnpj, status = empresa

    if status == 'Erro Imune - Associação' or status == 'Parametro Errado' or status == 'Patrimonio Errado':
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'REGERAR PULADO!']), nome=andamentos)
        return 'ok'
    else:
        while not _find_img('regerar.png', conf=0.9):
            p.hotkey('alt', 'u')
            time.sleep(0.5)
            p.press('a')
            time.sleep(0.5)
            p.press('a')
            time.sleep(0.5)
            p.press('enter')
            time.sleep(0.5)
            p.press('c')
            time.sleep(3)

        time.sleep(0.5)
        p.write(ano)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.write(ano2)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('+')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('+')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('+')
        time.sleep(0.5)

        p.press('r')
        print('15')
        while not _find_img('final.png', conf=0.9):
            print('11')
            if _find_img('erro_regerar.png', conf=0.9):
                p.press('enter')
                time.sleep(2)
                print('12')
                salvar_pdf(cod)
                print('13')
                p.press('esc', presses=5)
                p.press('n')
                _escreve_relatorio_csv(';'.join([cod, cnpj, 'REGERAR ERRO']), nome=andamentos)
                return 'ok'
            elif _find_img('ja_existe.png', conf=0.9):
                print('17')
                p.press('enter')
        time.sleep(2)

        if _find_img('sucesso.png', conf=0.9):
            print('18')
            p.press('enter')
            time.sleep(1)

            # GRAVAR
            p.press('g')

            _wait_img('aviso.png', conf=0.9, timeout=-1)

            if _find_img('gravar_sucesso.png', conf=0.9):
                p.press('enter')
                time.sleep(0.5)
                p.press('esc', presses=5)
                _escreve_relatorio_csv(';'.join([cod, cnpj, 'Regerar OK']), nome=andamentos)
                return 'ok'
            else:
                p.press('enter')
                p.press('esc', presses=5)
                _escreve_relatorio_csv(';'.join([cod, cnpj, 'Regerar Errado']), nome=andamentos)
                return 'ok'
        else:
            p.press('enter')
            p.press('esc', presses=5)
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Regerar Errado']), nome=andamentos)
            return 'ok'


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
            if not login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = regerar(str(ano), str(ano2), empresa, andamentos)

            if resultado == 'dominio fechou':
                _login_web()
                _abrir_modulo('escrita_fiscal')

            if resultado == 'modulo fechou':
                _abrir_modulo('escrita_fiscal')

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('CÓDIGO;CNPJ;STATUS', nome=andamentos)



if __name__ == '__main__':
    # Captura o ano base que vai ser usado para apuração e envio de reinf
    ano = p.prompt(text='Qual data inicio?', title='Script incrível', default='00/00/0000')
    ano2 = p.prompt(text='Qual data final?', title='Script incrível', default='00/00/0000')
    
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()
    
    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado_Regerar'
    
    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
