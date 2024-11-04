import datetime, shutil, os, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login

def apurar(mes_ano, empresa, andamentos):
    cod, cnpj, nome = empresa

    # Se o mes escolhido for 12 ele define como 1 para a imagem
    mes = mes_ano.split('/')
    if mes[0] == '12':
        mes = '01'
    # Se não for mes 12 ele soma +1 ao mes escolhe para a imagem
    else:
        mes = str(int(mes[0]) + 1)

    # Se o mês da empresa estiver 1 mês na frente do mês escolhido ele pula a apuração
    while not _find_img('mes' + mes + '_.png', conf=0.9):

        if _find_img('empresas_2020.png', conf=0.9):
            return 'ok'
        if _find_img('empresas_2021.png', conf=0.9):
            return 'ok'
        if _find_img('empresas_2022.png', conf=0.9):
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

        # Equanto não achar imagem do processo de apuração espera 1 segundo
        while not _find_img('progresso_apuracao.png', conf=0.9):
            time.sleep(1)
            if _find_img('progresso_apuracao_2.png', conf=0.9):
                break

        # Enquanto não achar a imagem de fim de apuração ele procura pela imagem de avisos apuração
        while not _find_img('fim_apuracao.png', conf=0.9):
            time.sleep(1)
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
        # fechar qualquer possível tela aberta
        p.press('esc', presses=5)
        time.sleep(1)
        return 'ok'
    return 'ok'

def valor_folha_sn(empresa, andamentos):
    cod, cnpj, nome = empresa

    _wait_img('movimentos.png', conf=0.9, timeout=-1)

    print('>>> Importando e Gravando')

    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('valor_da_folha.png', conf=0.9):
        # Relatórios
        p.hotkey('alt', 'm')
        time.sleep(0.5)
        p.press('o')
        time.sleep(0.5)
        p.press('n')
        time.sleep(0.5)
        p.press('f')
        time.sleep(2)
        if _find_img('valor_da_folha_2.png', conf=0.9):
            break

    time.sleep(0.5)
    p.hotkey('alt', 'i')

    while not _find_img('botao_ativado.png', conf=0.9):
        time.sleep(1)

    p.hotkey('alt', 'g')

    while not _find_img('botao_desativado.png', conf=0.9):
        time.sleep(1)

    p.press('esc', presses=5)
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Gravação com sucesso!']), nome=andamentos)
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
            if not _login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado2 = apurar(str(ano), empresa, andamentos)
            if resultado2 == 'ok':
                resultado = valor_folha_sn(empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)


if __name__ == '__main__':
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Valor da Folha SN'
    ano = p.prompt(text='Qual periodo base?', title='Script incrível', default='00/0000')
    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
        