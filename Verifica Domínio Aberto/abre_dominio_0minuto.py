# -*- coding: utf-8 -*-
import pyautogui as p
import time
import chromedriver_autoinstaller
import shutil
import os
import pyperclip
import pyautogui as a
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import SessionNotCreatedException, WebDriverException
from selenium.webdriver.common.by import By
from pathlib import Path
#16-08-2024#
e_dir = Path('execução')
imagens = "V:\\_imagens_comum_python\\imgs_comum_dominio"

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Domínio.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.readline()
user = user.split('/')


def verifica_empresa(cod):
    p.click(1000, 500)
    time.sleep(1)
    while True:
        p.click(1258, 82)
        while True:
            try:
                p.hotkey('ctrl', 'c')
                p.hotkey('ctrl', 'c')
                cnpj_codigo = pyperclip.paste()
                break
            except:
                pass

        codigo = cnpj_codigo.split('-')
        codigo = str(codigo[-1].strip())
        codigo = codigo.replace(' ', '')
        if codigo != '':
            break

    if codigo != cod:
        print(f'Código da empresa: {cod}')
        print(f'Código encontrado no Domínio: {codigo}')
        return False
    else:
        return True


_verifica_empresa = verifica_empresa


def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()


_escreve_relatorio_csv = escreve_relatorio_csv


def _login(empresa, andamentos, retorna_erro_parametro=False):
    regime = ''
    try:
        cod, cnpj, nome, regime, movimento = empresa
        regime += ';'
    except:
        try:
            cod, cnpj, nome, regime = empresa
            regime += ';'
        except:
            try:
                cod, cnpj, nome = empresa
            except:
                cod, cnpj = empresa
                nome = ''

    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta=imagens, conf=0.9):
        if _find_img('inicial_2.png', pasta=imagens, conf=0.9):
            break
        time.sleep(1)

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta=imagens, conf=0.9):
        p.press('f8')
        if _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
            break

    time.sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta=imagens, conf=0.9):
        _click_img('codigo.png', pasta=imagens, conf=0.9)
    p.write(cod)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
        time.sleep(1)
        if _find_img('sem_parametro.png', pasta=imagens, conf=0.9):
            print('❌ Parametro não cadastrado para esta empresa')
            if retorna_erro_parametro:
                return 'Sem parâmetros'
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{regime}Parametro não cadastrado para esta empresa',
                                   nome=andamentos)
            p.press('enter')
            time.sleep(1)
            while not _find_img('parametros.png', pasta=imagens, conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if _find_img('nao_existe_parametro.png', pasta=imagens, conf=0.9) or _find_img('nao_existe_parametro_2.png',
                                                                                       pasta=imagens, conf=0.9):
            if retorna_erro_parametro:
                return 'Sem parâmetros'
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{regime}Não existe parametro cadastrado para esta empresa',
                                   nome=andamentos)
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta=imagens, conf=0.9):
                time.sleep(1)
            while _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
                time.sleep(1)
            return False

        if (_find_img('empresa_nao_usa_sistema.png', pasta=imagens, conf=0.9) or
                _find_img('empresa_nao_usa_sistema_2.png', pasta=imagens, conf=0.9) or
                _find_img('empresa_nao_usa_sistema_3.png', pasta=imagens, conf=0.9)):
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{regime}Empresa não está marcada para usar este sistema',
                                   nome=andamentos)
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            while _find_img('trocar_empresa.png', pasta=imagens, conf=0.9):
                time.sleep(1)
            while _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
                time.sleep(1)
            return False

        if _find_img('fase_dois_do_cadastro.png', pasta=imagens, conf=0.9) or _find_img('fase_dois_do_cadastro_2.png',
                                                                                        pasta=imagens, conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta=imagens, conf=0.9) or _find_img('conforme_modulo_2.png',
                                                                                  pasta=imagens, conf=0.9):
            p.press('enter')
            time.sleep(1)

        if _find_img('aviso_regime.png', pasta=imagens, conf=0.9) or _find_img('aviso_regime_2.png', pasta=imagens,
                                                                               conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)

        if _find_img('aviso.png', pasta=imagens, conf=0.9) or _find_img('aviso_2.png', pasta=imagens, conf=0.9):
            p.hotkey('alt', 'o')
            time.sleep(1)

        if _find_img('erro_troca_empresa.png', pasta=imagens, conf=0.9) or _find_img('erro_troca_empresa_2.png',
                                                                                     pasta=imagens, conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5, interval=1)
            _login(empresa, andamentos)

    if not verifica_empresa(cod):
        _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{regime}Empresa não encontrada', nome=andamentos)
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False

    p.press('esc', presses=5)
    time.sleep(1)

    return True

def click_position_img(img, operacao, pixels_x=0, pixels_y=0, pasta='imgs', conf=1.0, clicks=1):
    try:
        if img.endswith('.png'):
            img = os.path.join(pasta, img)
            a.moveTo(a.locateCenterOnScreen(img, confidence=conf))
            local_mouse = a.position()
        else:
            local_mouse = img

        if operacao == '+':
            a.click(int(local_mouse[0] + int(pixels_x)), int(local_mouse[1] + int(pixels_y)), clicks=clicks)
            return True
        if operacao == '-':
            a.click(int(local_mouse[0] - int(pixels_x)), int(local_mouse[1] - int(pixels_y)), clicks=clicks)
            return True
        if operacao == '+x-y':
            a.click(int(local_mouse[0] + int(pixels_x)), int(local_mouse[1] - int(pixels_y)), clicks=clicks)
            return True
        if operacao == '-x+y':
            a.click(int(local_mouse[0] - int(pixels_x)), int(local_mouse[1] + int(pixels_y)), clicks=clicks)
            return True
    except:
        return False
_click_position_img = click_position_img


def _abrir_modulo(modulo, usuario=user[2], senha=user[3]):
    if _find_img('inicial.png', pasta=imagens, conf=0.9) or _find_img('inicial_2.png', pasta=imagens, conf=0.9):
        return True

    print(f'>>> Abrindo modulo {modulo.capitalize()}\n')
    while not _find_img('modulos.png', pasta=imagens, conf=0.9):
        time.sleep(1)
        try:
            p.getWindowsWithTitle('Lista de Programas')[0].activate()
        except:
            pass
    time.sleep(1)
    _click_img('modulo_' + modulo + '.png', pasta=imagens, conf=0.9, button='left', clicks=2)

    timer = 0
    contador = 1
    while not _find_img('login_modulo.png', pasta=imagens, conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            with p.hold('alt'):
                if contador == 1:
                    p.press('tab')
                    time.sleep(1)
                if contador == 2:
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(1)
                if contador == 3:
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(1)
                if contador == 4:
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(1)
                    contador = 0
            contador += 1
            time.sleep(1)

        if timer > 60:
            while not _find_img('tela_modulos.png', pasta=imagens, conf=0.9):
                contador = 1
                with p.hold('alt'):
                    if contador == 1:
                        p.press('tab')
                        time.sleep(1)
                    if contador == 2:
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(1)
                    if contador == 3:
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(1)
                    if contador == 4:
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(1)
                        contador = 0
                contador += 1
                time.sleep(1)

            _click_img('tela_modulos.png', pasta=imagens, conf=0.9)
            time.sleep(1)
            _click_img('modulo_' + modulo + '.png', pasta=imagens, conf=0.9, button='left', clicks=2)
            timer = 0

    _click_position_img('insere_usuario.png', '+', pixels_x=120, pasta=imagens, conf=0.9, clicks=2)

    time.sleep(0.5)
    p.press('del', presses=10)
    p.write(usuario)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=10)
    p.write(senha)
    time.sleep(0.5)
    p.hotkey('alt', 'o')
    while not _find_img('onvio.png', pasta=imagens, conf=0.9):
        time.sleep(1)

    time.sleep(5)

    if _find_img('aviso.png', pasta=imagens, conf=0.9):
        p.hotkey('alt', 'o')
    return True

def send_input_xpath(elem_path, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_path)
            elem.send_keys(data)
            break
        except:
            pass
_send_input_xpath = send_input_xpath


def click_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, button='left', clicks=1):
    img = os.path.join(pasta, img)
    try:
        aux = 0
        while True:
            box = a.locateCenterOnScreen(img, confidence=conf)
            if box:
                a.click(a.locateCenterOnScreen(img, confidence=conf), button=button, clicks=clicks)
                return True
            time.sleep(delay)
            if timeout < 0:
                continue
            if timeout == aux:
                break
            aux += 1
        else:
            return False
    except:
        return False
_click_img = click_img


def find_img(img, pasta='imgs', conf=1):
    try:
        path = os.path.join(pasta, img)
        if conf != 1:
            return a.locateOnScreen(path, confidence=conf)
        else:
            return a.locateOnScreen(path)
    except:
        return False
_find_img = find_img


def initialize_chrome(options=webdriver.ChromeOptions(), timeout=90):
    pasta_driver = 'V:\Setor Robô\Scripts Python\_comum\Chrome driver'

    service = Service(pasta_driver)
    while True:
        for pasta_atual, subpastas, arquivos in os.walk(pasta_driver):
            # Agora você pode processar os arquivos na pasta atual normalmente
            for file in arquivos:
                caminho_completo = os.path.join(pasta_atual, file)
                if caminho_completo == 'V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.exe':
                    continue
                service = Service(caminho_completo)
                print('Chrome driver selecionado:', caminho_completo)

        if not options:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")

        options.add_argument("--ignore-certificate-errors")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        # retorna o chromedriver aberto
        try:
            print('>>> Inicializando Chromedriver...')
            driver = webdriver.Chrome(options=options, service=service)
            driver.set_page_load_timeout(timeout)
            break
        except SessionNotCreatedException:
            print('>>> Atualizando Chromedriver...')
            shutil.rmtree(pasta_driver)
            time.sleep(1)
            os.makedirs(pasta_driver, exist_ok=True)
            # biblioteca para baixar o chromedriver atualizado
            chromedriver_autoinstaller.install(path=pasta_driver)
        except WebDriverException:
            print('>>> Baixando Chromedriver...')
            os.makedirs(pasta_driver, exist_ok=True)
            # biblioteca para baixar o chromedriver atualizado
            chromedriver_autoinstaller.install(path=pasta_driver)

    return True, driver


_initialize_chrome = initialize_chrome

def find_by_id(item, driver):
    try:
        driver.find_element(by=By.ID, value=item)
        return True
    except:
        return False
_find_by_id = find_by_id


def find_by_path(item, driver):
    try:
        driver.find_element(by=By.XPATH, value=item)
        return True
    except:
        return False
_find_by_path = find_by_path

def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.ID, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass
_send_input = send_input


def send_input_xpath(elem_path, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_path)
            elem.send_keys(data)
            break
        except:
            pass
_send_input_xpath = send_input_xpath

def _login_web(usuario=user[0], senha=user[1], force_open=False):
    if not _find_img('app_controler.png', pasta=imagens, conf=0.99) or not _find_img('app_controler_desfocado.png',
                                                                                     pasta=imagens,
                                                                                     conf=0.99) or force_open:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")

        status, driver = _initialize_chrome(options)

        driver.get('https://www.dominioweb.com.br/')

        # clica para começar o login
        while not _find_by_id('enterButton', driver):
            time.sleep(1)
        driver.find_element(by=By.ID, value='enterButton').click()

        # insere o e-mail
        while not _find_by_id('username', driver):
            time.sleep(1)
        _send_input('username', usuario, driver)
        time.sleep(0.5)

        # confirma o e-mail
        driver.find_element(by=By.XPATH,
                            value='/html/body/div/div/div/div[2]/div/main/section/div/div/div/div[1]/div/form/div[2]/button').click()
        time.sleep(1)

        # insere a senha
        while not _find_by_id('password', driver):
            time.sleep(1)
        _send_input('password', senha, driver)
        time.sleep(0.5)

        # clica em entrar
        driver.find_element(by=By.XPATH,
                            value='/html/body/div/div/div/div[2]/div/main/section/div/div/div/form/div[2]/button').click()

        abrir_apps = ['configurar_depois.png', 'abrir_app.png', 'abrir_app_2.png', 'abrir_app_3.png', 'abrir_app_4.png']

        print('>>> Aguardando modulos')
        while not _find_img('modulos.png', pasta=imagens, conf=0.9):
            time.sleep(1)
            for abrir_app in abrir_apps:
                if _find_img(abrir_app, pasta=imagens, conf=0.9):
                    _click_img(abrir_app, pasta=imagens, conf=0.9)

        driver.quit()
        return True
    else:
        if _find_img('app_controler_desfocado.png', pasta=imagens, conf=0.99):
            _click_img('app_controler_desfocado.png', pasta=imagens, conf=0.99, timeout=1)
        else:
            _click_img('app_controler.png', pasta=imagens, conf=0.99, timeout=1)
        time.sleep(2)
        if _find_img('lista_de_programas.png', pasta=imagens, conf=0.9):
            p.press('right', presses=2, interval=0.5)
            time.sleep(1)
            p.press('enter')

def run():
    p.hotkey('win', 'm')
    _login_web(force_open=True)
    _abrir_modulo('escrita_fiscal', usuario='ROBO', senha='Rb#0086*')

if __name__ == '__main__':
    run()
