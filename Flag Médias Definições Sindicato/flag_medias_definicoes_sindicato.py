import datetime
import time
import cv2
import pytesseract
import pyautogui as p
from PIL import Image
from _comum.pyautogui_comum import _find_img, _click_img
from _comum.comum_comum import _escreve_relatorio_csv, _escreve_header_csv, _time_execution_monitor_db, _get_host_name
from _comum.dominio_comum import _login_web, _abrir_modulo

pytesseract.pytesseract.tesseract_cmd = r'V:\Setor Robô\Tesseract-OCR\tesseract.exe'
def captura_codigo_empresa_ocr():
    time.sleep(3)

    # Defina a região de interesse na tela (coordenadas x, y, largura, altura)
    # Você pode usar pyautogui.displayMousePosition() para encontrar essas coordenadas
    x, y, largura, altura = 343, 159, 405, 176

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

    # Configurações para extrair apenas números
    custom_config = r'--psm 6 outputbase digits'

    # Usa o pytesseract para extrair texto da imagem pré-processada
    texto_extraido = pytesseract.image_to_string('recorte_preprocessado.png', config=custom_config)

    return texto_extraido

def abre_tela_sindicato_empregados():
    while not _find_img('titulo_sindicatos_dos_empregados.png', conf=0.9):
        p.hotkey('alt', 'a')
        time.sleep(0.5)
        p.press('s')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.press('e')
        time.sleep(30)
    time.sleep(1)
    _click_img('primeiro.png', conf=0.9)
    time.sleep(3)
    while not _find_img('definicoes_aba.png', conf=0.9):
        time.sleep(1)
        _click_img('medias.png', conf=0.9)
        time.sleep(1)
    time.sleep(1)
    _click_img('definicoes.png', conf=0.9)

def consulta_definicoes(andamentos, texto_extraido):
    # Função para verificar e clicar nas imagens
    time.sleep(1)
    if _find_img(f'calcular_ok.png', conf=1):
        status2 = 'OK'
    else:
        while not _find_img(f'calcular_ok.png', conf=1):
            _click_img(f'calcular.png', conf=0.9)
            time.sleep(5)
            _click_img('definicoes.png', conf=0.9)
            time.sleep(1)
        time.sleep(3)
        p.hotkey('alt', 'g')
        status2 = 'OK'#
        while not _find_img('img2.png', conf=0.9):
            time.sleep(1)
            if _find_img('vinculado.png', conf=0.9):
                p.hotkey('alt', 's')
            if _find_img('informadas.png', conf=0.9):
                p.press('enter')
                status2 = 'Não foram informadas as configurações de médias para.'
            if _find_img('informadas_2.png', conf=0.9):
                p.press('enter')
                status2 = 'Não foram informadas as configurações de médias para Aviso Prévio.'
            if _find_img('estabilidade.png', conf=0.9):
                p.press('enter')
                time.sleep(5)
                p.hotkey('alt', 'c')
                time.sleep(10)
                status2 = 'Os motivos de rescisão não foram informados para cálculo de estabilidade.'
            
            if _find_img('entidade.png', conf=0.9):
                p.press('enter')
                time.sleep(5)
                p.hotkey('alt', 'c')
                time.sleep(10)
                status2 = 'O campo Código da Entidade deve ser informado!.'
    
    time.sleep(1)
    _escreve_relatorio_csv(';'.join([texto_extraido.strip(), status2]), nome=andamentos)

    p.click(481, 165)
    time.sleep(5)
    while not _find_img('definicoes_aba.png', conf=0.9):
        time.sleep(1)
        _click_img('medias.png', conf=0.9)
        time.sleep(1)
    time.sleep(1)
    _click_img('definicoes.png', conf=0.9)

    if str(texto_extraido).strip() == '275':
        return 'fim'
    return 'ok'

@_time_execution_monitor_db
def run(controle):
    _login_web()
    _abrir_modulo('folha')
    abre_tela_sindicato_empregados()#

    while True:
        texto_extraido = captura_codigo_empresa_ocr()
        resultado = consulta_definicoes(andamentos, texto_extraido)

        if resultado == 'ok':
            continue
        if resultado == 'fim':
            break

    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)

if __name__ == '__main__':
    andamentos = 'Flag Médias Definições Sindicato'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)
