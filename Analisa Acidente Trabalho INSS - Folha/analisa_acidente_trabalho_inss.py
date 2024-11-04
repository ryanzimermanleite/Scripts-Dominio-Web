import datetime
import pytesseract
import time
import cv2
import os
import pyautogui as p
from _comum.pyautogui_comum import _find_img, _click_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _get_host_name, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _escreve_header_csv
from _comum.dominio_comum import _login_web, _abrir_modulo, _login
from PIL import Image
pytesseract.pytesseract.tesseract_cmd = r'V:\Setor Robô\Tesseract-OCR\tesseract.exe'


def analisa_acidente_trabalho(cod, nome, cnpj):
    while not _find_img('titulo_servicos.png', conf=0.9):
        time.sleep(1)
        p.hotkey('alt', 'a')
        time.sleep(1)
        p.press('v')
        time.sleep(1)
        p.press('enter')
        time.sleep(5)
    time.sleep(1)

    while not _find_img('acid_trabalho.png', conf=0.9):
        time.sleep(1)
        _click_img('inss.png', conf=0.9)
        time.sleep(3)
    time.sleep(1)

    # Defina a região de interesse na tela (coordenadas x, y, largura, altura)
    # Você pode usar pyautogui.displayMousePosition() para encontrar essas coordenadas
    x, y, largura, altura = 505, 565, 783, 583

    # Captura a tela e salva como uma imagem
    imagem_screenshot = p.screenshot()
    imagem_screenshot.save('recorte.png')

    # Abre a imagem e faz o recorte
    im = Image.open('recorte.png')
    im = im.crop((int(x), int(y), int(largura), int(altura)))
    im.save('recorte.png')

    # Carregar a imagem pré-recortada
    imagem = cv2.imread('recorte.png')

    # Redimensionar a imagem para uma resolução maior
    imagem = cv2.resize(imagem, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)

    # Converter para escala de cinza
    imagem_cinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)

    # Aplicar um filtro de suavização (filtro Gaussiano)
    imagem_suavizada = cv2.GaussianBlur(imagem_cinza, (5, 5), 0)

    # Aplicar a binarização (limiar adaptativo)
    _, imagem_binarizada = cv2.threshold(imagem_suavizada, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Ajustar o contraste
    imagem_ajustada = cv2.equalizeHist(imagem_binarizada)

    # Salvar a imagem pré-processada
    cv2.imwrite('recorte_preprocessado.png', imagem_ajustada)



    # Usa o pytesseract para extrair texto da imagem pré-processada
    texto_extraido = pytesseract.image_to_string('recorte_preprocessado.png')
    texto_extraido = texto_extraido.replace('\n', '')
    texto_extraido = texto_extraido.replace('.', '')
    texto_extraido = texto_extraido.replace(',', '')

    print(str(texto_extraido))
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, texto_extraido]), nome=andamentos)
    p.press('esc', presses=5)
    return 'ok'

@_time_execution_monitor_db
def run(controle):

    empresas = _open_lista_dados()
    if not empresas:
        return False

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _login_web()
    _abrir_modulo('folha')


    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao, controle=controle, usando_bd=True, nome_rotina=andamentos +  f' - {_get_host_name()}', planilha=os.path.join('execução', andamentos + '.csv'))

        cod, nome, cnpj = empresa

        while True:

            if not _login(empresa, andamentos):
                break

            resultado = analisa_acidente_trabalho(cod, nome, cnpj)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    andamentos = 'Analisa Acidente Trabalho INSS - Dominio Folha'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)



