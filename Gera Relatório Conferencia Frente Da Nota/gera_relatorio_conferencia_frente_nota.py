import datetime, shutil, os, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login

def salvar_pdf(cod):
    p.click(833, 384)
    time.sleep(0.5)
    
    _click_img('salvar.png', conf=0.9)
    timer = 0
    _wait_img('salvar_relat.png', conf=0.9)
    time.sleep(0.5)
    _click_img('botao.png', conf=0.9)
    time.sleep(0.5)
    _click_img('tabela_html.png', conf=0.9)
    time.sleep(0.5)
    _click_img('3pontos.png', conf=0.9)
    time.sleep(0.5)

    while not _find_img('selecione_arquivo.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False
    
    time.sleep(1)
    p.write(str(cod) + '.html')
    time.sleep(0.5)
    
    if not _find_img('cliente_c_selecionado.png', pasta='imgs_c', conf=0.9):
        while not _find_img('cliente_c.png', pasta='imgs_c', conf=0.9) or _find_img('cliente_m.png', pasta='imgs_c', conf=0.9):
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
    time.sleep(5)
    p.press('esc', presses=5)
    time.sleep(1)

    return True

def gera_relat(ano, ano2, empresa, andamentos):
    cod, cnpj, nome = empresa
    
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    
    print('>>> Gerando o relatório')
    
    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('gerenciador_de_relatorios.png', conf=0.9):
        print('x')
        # Relatórios
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        # gerador de relatórios
        p.press('r')
        time.sleep(2)

    time.sleep(0.5)

    print('>>> Buscando relatório')
    p.press('pgup', presses=10)
    while not _find_img('conferencias_entrada.png', conf=0.9):
        print('y')
        p.press('pgdn')
        time.sleep(0.5)

    _click_img('conferencias_entrada.png', conf=0.9, clicks=2, timeout=1)
    time.sleep(1)
    _click_img('conferencias_entrada_2.png', conf=0.9, clicks=2, timeout=1)
    
    # Insere codigo da empresa
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(cod)

    # Insere data inicio
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=10)
    time.sleep(0.5)
    p.write(ano)

    # Insere data final
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=10)
    time.sleep(0.5)
    p.write(ano2)

    # executa
    time.sleep(0.5)
    p.hotkey('alt', 'e')
    
    # enquanto o relatório não é gerado, verifica se aparece a mensagem dizendo que não possuí dados para emitir
    while not _find_img('notas_entradas.png', conf=0.9) or _find_img('notas_entradas2.png'):
        print('z')
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para emitir']), nome=andamentos)
            print('❌ Sem dados para emitir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok'
    
    salvar_pdf(cod)
    # fechar qualquer possível tela aberta
    p.press('esc', presses=5)
    time.sleep(2)
    pasta_origem = 'C:\\'
    pasta_destino = 'V:\Setor Robô\Scripts Python\Domínio\Gera Relatório Conferencia Frente Da Nota\execução\\notas_html\\'
    nome_arquivo = str(cod) + '.html'
    p.press('esc', presses=5)
    time.sleep(1)
    try:
        shutil.move(os.path.join(pasta_origem, nome_arquivo), os.path.join(pasta_destino, nome_arquivo))
    except:
        pass

    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatorio Gerado']), nome=andamentos)
    
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
            resultado = gera_relat(str(ano), str(ano2), empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;STATUS', nome=andamentos)
    
    
if __name__ == '__main__':
    # Captura o ano base que vai ser usado para apuração e envio de reinf
    ano = p.prompt(text='Qual data inicio?', title='Script incrível', default='00/00/0000')
    ano2 = p.prompt(text='Qual data final?', title='Script incrível', default='00/00/0000')

    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Frente Nota_2'

    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
