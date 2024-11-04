import datetime, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login

def robo_novo(empresa, andamentos):
    cod, cnpj, nome = empresa

    while not _find_img('modulos.png', conf=0.9):
        p.hotkey('alt', 'c')
        time.sleep(0.5)
        p.press('e')
        time.sleep(5)

    _click_img('modulos.png', conf=0.9)

    while not _find_img('titulo_modulos.png', conf=0.9):
        time.sleep(1)

    _click_img('onvio.png', conf=0.9)

    while not _find_img('modulos_utilizados.png', conf=0.9):
        time.sleep(1)
    time.sleep(2)

    '''while not _find_img('referencia.png', conf=0.9):
        _click_img('empregado.png', conf=0.9)
        time.sleep(1)
        if _find_img('aviso.png', conf=0.9):
            p.press('enter')
        time.sleep(1)
        _click_img('cliente.png', conf=0.9)
        time.sleep(2)



    p.hotkey('alt', 'o')
    time.sleep(1)
    p.hotkey('alt', 'g')
    time.sleep(3)

    if _find_img('cliente_desde.png', conf=0.9):
        p.press('enter')
        time.sleep(1)
        p.press('esc')
        time.sleep(1)
        if _find_img('gravar.png', conf=0.9):
            p.hotkey('alt', 'n')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Você precisar informar o campo Cliente desde.']), nome=andamentos)
            p.press('esc', presses=5)
            return 'ok'

    if _find_img('e_social.png', conf=0.9):
        p.press('enter')
        time.sleep(1)
        p.press('esc')
        time.sleep(1)
        if _find_img('gravar.png', conf=0.9):
            p.hotkey('alt', 'n')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Alteração não permitida! A empresa já possui eventos validados no eSocial e a alteração da inscrição ocasionará erros no envio dos eventos.']), nome=andamentos)
            p.press('esc', presses=5)
            return 'ok'

    if _find_img('aviso2.png', conf=0.9):
        p.press('enter')
        time.sleep(1)
        p.press('esc')
        time.sleep(1)
        if _find_img('gravar.png', conf=0.9):
            p.hotkey('alt', 'n')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'A data do campo Início das atividades da empresa não pode ser maior que a data do campo Cliente desde.']), nome=andamentos)
            p.press('esc', presses=5)
            return 'ok'


    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Flag Onvio Setado!']), nome=andamentos)'''

    if _find_img('referencia.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Flag Onvio Ativado!']), nome=andamentos)
    else:
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Flag Onvio Desativado!']), nome=andamentos)

    p.press('esc', presses=5)
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

            resultado = robo_novo(empresa, andamentos)

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
    andamentos = 'Resultado Flag Modulo Onvio'
    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
