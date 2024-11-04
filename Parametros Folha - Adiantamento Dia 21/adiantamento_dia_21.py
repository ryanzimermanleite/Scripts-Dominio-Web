import datetime, shutil, os, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login

def parametro_adiantamento(empresa, andamentos):
    cod, cnpj, nome = empresa
    while not _find_img('parametros.png', conf=0.9):
        p.hotkey('alt', 'c')
        time.sleep(1)
        p.press('p')
        time.sleep(1)
    
    while not _find_img('proporcionalidade.png', conf=0.9):
        _click_img('adiantamento.png', conf=0.9)
        time.sleep(1)
        
    while not _find_img('calcular_adiantamento.png', conf=0.9):
        _click_img('proporcionalidade.png', conf=0.9)
        time.sleep(1)
        
    time.sleep(2)
    
    if _find_img('calcular_nao.png', conf=1):
        _click_img('titulo_calcular.png', conf=0.9)
    elif _find_img('calcular_ok.png', conf=1):
        _click_img('titulo_calcular.png', conf=0.9, clicks=2)
        
    time.sleep(1)
    
    p.press('tab', presses=3)
    time.sleep(0.5)
    p.press('space')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write('21')
    time.sleep(0.5)
    p.hotkey('alt', 'g')
    while not _find_img('e_social.png', conf=0.9):
        time.sleep(1)
        if _find_img('calcula_vigencia.png', conf=1):
            p.hotkey('alt', 'y')
        if _find_img('provisao.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'É necessário configurar o grupo "Configuração do relatório de provisão das férias calculadas no mês."']), nome=andamentos)
            p.hotkey('alt', 'o')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'
        if _find_img('centralizadora.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Informe apenas uma empresa como centralizadora ou configure todas as empresas como centralizada indicando que a centralizadora consta em outro banco de dados.']), nome=andamentos)
            p.press('enter')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'
        if _find_img('banco_dados.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'A empresa informada já iniciou o envio das informações em outro banco de dados. Para que seja possível continuar o envio das informações nesse banco de dados, deverá ser efetuada a Importação de Dados do eSocial ou a Conversão de Dados do eSocial. É necessário selecionar o processo desejado.']), nome=andamentos)
            p.press('esc')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'
        if _find_img('conforme.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Conforme Portaria Conjunta SEPRT/RFB/ME nº71, de 29 de junho de 2021 as datas de envio do eSocial são 1º Grupo: Tabelas 01/01/2018,Não periódicos: 01/03/2018,Periódicos: 01/05/2018. 2º Grupo: Tabelas 16/07/2018;Não periódicos: 10/10/2018, Periódicos: 01/01/2019, 3º Grupo: Tabelas 10/01/2019;Não periódicos: 10/04/2019,Periódicos:PJ:01/05/2021 e PF 01/07/2021 e Segurado Especial 01/10/2021 e o 4º Grupo 21/07/2021.']), nome=andamentos)
            p.hotkey('alt', 'n')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'
        if _find_img('codigo_gov.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'O código de acesso foi descontinuado pelo eSocial em 06/2023. Para gerar a guia DAE, é necessário configurar nos parâmetros da empresa o Tipo de acesso para a Conta gov.br']), nome=andamentos)
            p.hotkey('alt', 'n')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'
        if _find_img('s1070.png', conf=1):
            p.press('enter')
        if _find_img('apuracao_previdenciaria.png', conf=1):
            p.hotkey('alt', 'y')
        if _find_img('mesmo_cnpj.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Já existe empresa com este mesmo CNPJ configurada para envio ao ambiente oficial do eSocial.']), nome=andamentos)
            p.press('enter')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'
        if _find_img('s1000.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'A empresa centralizadora para envio dos eventos S-1000 Empregador, S-1298 Reabertura e S-1299 Fechamento ao eSocial deve possuir o mesmo responsável legal da empresa centralizada.']),
                                   nome=andamentos)
            p.press('enter')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'

        if _find_img('s1000_2.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'A empresa centralizadora para envio dos eventos S-1000 Empregador, S-1298 Reabertura e S-1299 Fechamento ao eSocial não foi informada. É obrigatória sua informação.']),
                                   nome=andamentos)
            p.press('enter')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'

        if _find_img('minimo.png', conf=1):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome,
                                             'O mínimo de dias trabalhados para cálculo do adiantamento não foi informado.']),
                                   nome=andamentos)
            p.press('enter')
            time.sleep(2)
            p.press('esc', presses=5)
            return 'ok'

    p.press('esc', presses=5)
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, '21 Dias Setado!']), nome=andamentos)
    return 'ok'


@_barra_de_status
def run(window):

    # abre o Domínio Web e o módulo, no caso será o módulo Folha
    _login_web()
    _abrir_modulo('folha', usuario, senha)
    
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
            
            resultado = parametro_adiantamento(empresa, andamentos)
            
            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)


if __name__ == '__main__':
    usuario = p.prompt(text='Qual o usuario do dominio?', title='Script incrível', default='')
    senha = p.password(text='Qual a senha do dominio?', title='Script incrível', default='')
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Folha Parametro Adiantamento Dia 21'
    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
#