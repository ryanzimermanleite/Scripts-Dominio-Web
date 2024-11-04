import datetime, time, cv2, pyautogui, pytesseract, pyautogui as p
import traceback
from PIL import Image
from sys import path
from pyautogui import alert

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from dominio_comum import _login_web, _abrir_modulo, _login
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

pytesseract.pytesseract.tesseract_cmd = r'V:\Setor Robô\Tesseract-OCR\tesseract.exe'

def abre_tela_sindicato():
    while not _find_img('sindicato_empregado.png', conf=0.9):
        p.hotkey('alt', 'a')
        time.sleep(0.5)
        p.press('s')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.press('e')
        time.sleep(10)
        
    while not _find_img('adicional.png', conf=0.9):
        time.sleep(1)
        _click_img('calculo.png', conf=0.9)
    
    while not _find_img('regra.png', conf=0.9):
        time.sleep(1)
        _click_img('adicional.png', conf=0.9)
    
def verificar_presenca_e_atribuir(nome_arquivo, valor_atribuido):
    return valor_atribuido if _find_img(nome_arquivo, conf=0.9) else None

def ocr():
    time.sleep(3)
  
    # Defina a região de interesse na tela (coordenadas x, y, largura, altura)
    # Você pode usar pyautogui.displayMousePosition() para encontrar essas coordenadas
    x, y, largura, altura = 343, 159, 405, 176
    
    # Captura a tela e salva como uma imagem
    imagem_screenshot = pyautogui.screenshot()
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
    
    status_qtd, status_maior = verifica_imagens()
   
    _escreve_relatorio_csv(';'.join([texto_extraido.strip(), status_qtd, status_maior]), nome=andamentos)
    
    if str(texto_extraido).strip() == '257':
        return 'fim'
    
    p.click(481, 165)
    time.sleep(5)
    
    return 'ok'

def verifica_imagens():
    # Verificar presença de imagens e atribuir valores
    maior_que_branco = verificar_presenca_e_atribuir('maior_que_branco.png', '1')
    maior_igual_branco = verificar_presenca_e_atribuir('maior_igual_branco.png', '2')
    maior_que_marrom = verificar_presenca_e_atribuir('maior_que_marrom.png', '1')
    maior_igual_marrom = verificar_presenca_e_atribuir('maior_igual_marrom.png', '2')
    
    
    menor_que_branco = verificar_presenca_e_atribuir('menor_que_branco.png', '1')
    menor_igual_branco = verificar_presenca_e_atribuir('menor_igual_branco.png', '2')
    menor_que_marrom = verificar_presenca_e_atribuir('menor_que_marrom.png', '1')
    menor_igual_marrom = verificar_presenca_e_atribuir('menor_igual_marrom.png', '2')
    
    # Determinar status com base nos valores encontrados
    status_maior = 'Maior que ou Maior igual' if maior_que_branco or maior_igual_branco or maior_que_marrom or maior_igual_marrom else ''
    status_menor = 'Menor que ou Menor igual' if menor_que_branco or menor_igual_branco or menor_que_marrom or menor_igual_marrom else ''
    
    return status_maior, status_menor

def run():
   
    #_login_web()
    #_abrir_modulo('folha')
    #abre_tela_sindicato()
    time.sleep(3)
    while True:
        resultado = ocr()
        
        if resultado == 'fim':
            break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('CODIGO;MAIOR;MENOR', nome=andamentos)


if __name__ == '__main__':
  
    andamentos = 'Resultado Sindicato Empregado'
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    try:
        run()
    except Exception as e:
        traceback_str = traceback.format_exc()
        alert(f'Traceback: {traceback_str}\n\n'
              f'Erro: {e}')
        print(f'Traceback: {traceback_str}\n\n'
              f'Erro: {e}')
 