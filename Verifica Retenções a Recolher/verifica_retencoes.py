import datetime, pyperclip, shutil, os, fitz, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _encerra_dominio, _login


def salvar_pdf(cnpj):

    p.click(833, 384)
    time.sleep(0.5)
    p.hotkey('ctrl', 'd')
    time.sleep(1)
    timer = 0
    while not _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False

    time.sleep(1)
    p.write(str(cnpj))
    time.sleep(0.5)


    if not _find_img('cliente_t_selecionado.png', pasta='imgs_c', conf=0.9):
        while not _find_img('cliente_t.png', pasta='imgs_c', conf=0.9):
            _click_img('botao.png', pasta='imgs_c', conf=0.9)
            time.sleep(3)

    _click_img('cliente_t.png', pasta='imgs_c', conf=0.9, timeout=1)
    time.sleep(5)

    _click_img('pasta.png', conf=0.9, clicks=2, timeout=1)
    time.sleep(1)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 's')

    timer = 0
    while not _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
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
                time.sleep(1)

            if not _find_img('cliente_t_selecionado.png', pasta='imgs_c', conf=0.9):
                while not _find_img('cliente_t.png', pasta='imgs_c', conf=0.9):
                    _click_img('botao.png', pasta='imgs_c', conf=0.9)
                    time.sleep(3)
                _click_img('cliente_t.png', pasta='imgs_c', conf=0.9)
                time.sleep(5)

            p.press('enter')
            timer = 0

    while _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        p.hotkey('alt', 'f4')
        time.sleep(3)

    while _find_img('sera_finalizada.png', pasta='imgs_c', conf=0.9):
        p.press('esc')
        time.sleep(2)
    p.press('esc', presses=5)

    return True


def gera_relat(ano, ano2, empresa, andamentos):
    cod, cnpj, nome = empresa

    _wait_img('relatorios.png', conf=0.9, timeout=-1)

    print('>>> Gerando o relatório')

    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('retencoes_recolher.png', conf=0.9):
        # Relatórios
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('i')
        time.sleep(0.5)
        p.press('r')
        time.sleep(2)

    # Insere codigo da empresa
    time.sleep(1)
    p.write(ano)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write(ano2)
    time.sleep(0.5)
    p.hotkey('alt', 'o')
    time.sleep(0.5)

    # enquanto o relatório não é gerado, verifica se aparece a mensagem dizendo que não possuí dados para emitir
    while not _find_img('relatorio.png', conf=0.9):
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para emitir']), nome=andamentos)
            print('❌ Sem dados para emitir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok'

    salvar_pdf(cnpj)
    # fechar qualquer possível tela aberta
    p.press('esc', presses=5)
    time.sleep(3)
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Guia Gerada']), nome=andamentos)
    return 'ok'


@_barra_de_status
def run(window):
    # Captura o ano base que vai ser usado para apuração e envio de reinf
    ano = p.prompt(text='Qual data inicio?', title='Script incrível', default='00/00/0000')
    ano2 = p.prompt(text='Qual data final?', title='Script incrível', default='00/00/0000')

    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Retenções a Recolher 1'

    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return andamentos

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
            if not _login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = gera_relat(str(ano), str(ano2), empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)
    _encerra_dominio()

if __name__ == '__main__':
    run()