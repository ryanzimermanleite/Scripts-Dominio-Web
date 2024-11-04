import datetime
import time
import pyautogui as p
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By

from _comum.chrome_comum import _initialize_chrome, _send_input, _find_by_id, _find_by_path
from _comum.pyautogui_comum import _find_img, _click_img, _click_position_img
from _comum.comum_comum import _indice, _escreve_relatorio_xlsx, _envia_email, _get_host_name, _concatena, _configura_dados, _time_execution_monitor_db
from _comum.dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf

def gera_relatorio_pat(cnpj, nome, pasta_final):
    while not _find_img('gerenciador_de_relatorios.png', conf=0.9):
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('i', presses=2)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(2)

    time.sleep(0.5)
    p.press('pgup', presses=20)
    time.sleep(2)
    while not _find_img('diversos.png', conf=0.9):
        time.sleep(1)
        p.press('down')
    time.sleep(0.5)

    _click_img('diversos.png', conf=0.9, clicks=2)
    time.sleep(0.5)

    while not _find_img('vale.png', conf=0.9):
        p.press('down')
        time.sleep(1)
    time.sleep(0.5)  #

    _click_img('vale.png', conf=0.9, clicks=2)

    while not _find_img('referencia.png', conf=0.9):
        time.sleep(1)
        _click_img('vale.png', conf=0.9, clicks=2)
    time.sleep(0.5)

    p.press('tab', presses=3)
    time.sleep(1)

    # Data atual
    data_atual = datetime.datetime.now()

    # Calcular o mês anterior
    mes_anterior = data_atual.month - 1
    ano_anterior = data_atual.year

    # Se o mês atual for janeiro (1), então o mês anterior será dezembro (12) do ano anterior
    if mes_anterior == 0:
        mes_anterior = 12
        ano_anterior -= 1

    mes_ano_anterior = f"{mes_anterior:02d}/{ano_anterior}"
    print(mes_ano_anterior)
    p.write(mes_ano_anterior)
    time.sleep(1)
    p.hotkey('alt', 'e')

    while not _find_img('titulo_relatorio.png', conf=0.9):
        time.sleep(1)
        if _find_img('sem_dados.png', conf=0.9):
            p.press('enter')
            time.sleep(3)
            p.press('esc', presses=5)
            return 'Sem dados'

    _salvar_pdf()

    time.sleep(3)

    while not _find_img('e_social.png', conf=0.9):
        time.sleep(3)
        p.press('esc', presses=5)
        time.sleep(3)

    diretorio_relatorios = mover_pdf(pasta_final, cnpj, nome)

    return 'Relatório Gerado', diretorio_relatorios

def mover_pdf(pasta_final, cnpj, nome):
    nome_arquivo_origem = 'R.POSTAL - Vale Alimentação.pdf'
    nome_arquivo_final = str(cnpj) + '-' + str(nome) + '.pdf'

    pasta_origem = 'C:\\'
    pasta_origem_final = str(pasta_final) + '\\' + 'Relatórios'

    p.press('esc', presses=5)
    time.sleep(1)

    try:
        shutil.move(os.path.join(pasta_origem, nome_arquivo_origem), os.path.join(pasta_origem_final, nome_arquivo_final))
    except:
        pass

    return pasta_origem_final

def importa_relatorios_onvio(diretorio_relatorios):
    driver = login()
    while not _find_img('upload_3.png', conf=0.9):
        time.sleep(1)
        print('aaa')
        try:
            _click_img('botao_upload.png', conf=0.9)
            time.sleep(3)
        
        except:
            pass
    print('b')
    time.sleep(2)
    _click_position_img('seta.png', '+', pixels_x=25, conf=0.9)
    time.sleep(1)
    p.write(diretorio_relatorios)
    time.sleep(1)
    p.press('enter')
    time.sleep(5)

    while not _find_img('pdf2.png', conf=0.9):
        time.sleep(1)
        _click_img('pdf.png', conf=0.9)
        time.sleep(1)
        p.hotkey('ctrl', 'a')
        time.sleep(2)
    time.sleep(1)

    try:
        _click_img('abrir.png', conf=0.9)
    except:
        pass

    time.sleep(99)
    p.hotkey('alt', 'f4')


def login():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument('--disable-blink-features=AutomationControlled')
    # Percorre a planilha de dados atribuindo seus valores nas variaveis

    status, driver = _initialize_chrome(options)
    driver.set_page_load_timeout(60)
    print('>>> Acessando site')
    timer = 0
    timer_erro = 0
    try:
        # Tenta acessar o site
        driver.get('https://onvio.com.br/#/')
    except:
        return driver, 'erro'

    # ENQUANTO NÃO ENCONTRAR BOTAO DE ENTRAR
    while not _find_by_id('trauth-continue-signin-btn', driver):
        print('🕤')
        time.sleep(1)
        timer += 1
        timer_erro += 1

        if timer > 10:
            timer = 0
            try:
                # Tenta acessar o site
                driver.get('https://onvio.com.br/#/')
            except:
                return driver, 'erro'

        if timer_erro > 60:
            return driver, 'erro'

    # APERTA NO BOTAO ENTRAR
    while not _find_by_id('trauth-continue-signin-btn', driver):
        time.sleep(1)
    time.sleep(2)

    try:
        driver.find_element(by=By.ID, value="trauth-continue-signin-btn").click()
    except:
        pass

    # ESPERA APARECER O CAMPO USERNAME
    while not _find_by_id('username', driver):
        time.sleep(1)

    # DIGITA NO CAMPO DE USERNAME
    _send_input('username', 'daiani@veigaepostal.com.br', driver)
    time.sleep(1)

    # APERTA EM ENTRAR
    while not _find_by_path('/html/body/div/div/div/div[2]/div/main/section/div/div/div/div[1]/div/form/div[2]/button',
                            driver):
        time.sleep(1)
    driver.find_element(by=By.XPATH,
                        value="/html/body/div/div/div/div[2]/div/main/section/div/div/div/div[1]/div/form/div[2]/button").click()

    # ESPERA O CAMPO DE SENHA
    while not _find_by_id('password', driver):
        time.sleep(1)

    # DIGITA NO CAMPO A SENHA
    _send_input('password', 'abacate2024*//', driver)

    # BOTAO ENTRAR
    while not _find_by_path('/html/body/div/div/div/div[2]/div/main/section/div/div/div/form/div[2]/button', driver):
        time.sleep(1)
    driver.find_element(by=By.XPATH,
                        value="/html/body/div/div/div/div[2]/div/main/section/div/div/div/form/div[2]/button").click()

    # ESPERA APARECER O MENU
    while not _find_by_id('bm-header-app-menu-toggle', driver):
        time.sleep(2)
    time.sleep(2)
    # APERTA NO MENU LATERAL
    while True:
        try:
            driver.find_element(by=By.ID, value="bm-header-app-menu-toggle").click()
            break
        except:
            pass

    # CLICKA EM PROCESSOS
    while not _find_by_path('/html/body/bm-optional-header/bm-staff-custom-header/bm-header/div[1]/ul/li[1]/a/span',
                            driver):
        time.sleep(2)
    time.sleep(2)

    while True:
        try:
            driver.find_element(by=By.XPATH,
                                value="/html/body/bm-optional-header/bm-staff-custom-header/bm-header/div[1]/ul/li[1]/a/span").click()
            break
        except:
            pass
    time.sleep(1)
    while True:
        try:
            driver.switch_to.window(driver.window_handles[1])
            break
        except:
            time.sleep(1)

    # ESPERA O MENU DE PROCESSOS APARECER
    while not _find_by_id('gestta-menu', driver):
        time.sleep(1)

    driver.get(f'https://app.gestta.com.br/express/#/upload')
    while not _find_img('botao_upload.png', conf=0.9):
        time.sleep(1)

    time.sleep(3)

    return driver


@_time_execution_monitor_db
def run(controle):
    planilha = 'V:\Setor Robô\Scripts Python\Domínio\Gera Relatórico PAT na Folha de Pagamento\ignore\dados.xlsx'
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = _configura_dados(pasta_final_, andamentos, nova_planilha=False, planilha_dados=planilha)

    if not total_empresas:
        return False, pasta_final
    caminho = os.path.join(pasta_final, 'Relatórios')
    os.makedirs(os.path.join(pasta_final, 'Relatórios'), exist_ok=True)

    _login_web()
    _abrir_modulo('folha')

    tempos = [datetime.datetime.now()]
    tempo_execucao = []

    for count, empresa in enumerate(df_empresas[index:].values.tolist(), start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index=index, tempos=tempos,
                                         tempo_execucao=tempo_execucao, controle=controle, usando_bd=True, nome_rotina=andamentos +  f' - {_get_host_name()}')

        codigo, cnpj, nome = empresa
        cnpj = str(cnpj)
        cnpj = _concatena(cnpj, 14, 'antes', 0)
        nome = nome.replace('/', ' ')

        while True:

            if not _login(empresa, andamentos):
                break

            resultado, diretorio_relatorios = gera_relatorio_pat(cnpj, nome, pasta_final)


            if resultado != 'erro':
                break
        
        _escreve_relatorio_xlsx({'CODIGO': codigo, 'CNPJ': cnpj, 'NOME': nome, 'RESULTADO': resultado}, nome=andamentos,
                                local=pasta_final)
    importa_relatorios_onvio(pasta_final)
    
    #_envia_email('guilherme@veigaepostal.com.br, stephani@veigaepostal.com.br', andamentos, pasta_final, False, 'joao@veigaepostal.com.br, willian.rocha@veigaepostal.com.br')
    _envia_email('ryan.leite@veigaepostal.com.br, willian.rocha@veigaepostal.com.br', andamentos, pasta_final, False, 'joao@veigaepostal.com.br, daiani@veigaepostal.com.br')


if __name__ == '__main__':
    pasta_final_ = r'\\vpsrv03\Arq_Robo\Domínio'
    andamentos = 'Gera Relatório PAT Folha de Pagamento'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)
    # ONVIO NAO RECONHECE OS RELATORIO NO MOMENTO NAO IMPORTAR NO ONVIO ENVIAR O RESULTADO PARA STEPHANY E GUILHERME
    #importa_relatorios_onvio(r'\\vpsrv03\Arq_Robo\Domínio\Gera Relatório PAT Folha de Pagamento\2024\09-2024\Execução (1)\Relatórios')