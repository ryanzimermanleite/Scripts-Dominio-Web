import datetime, time, cv2, pyautogui, pytesseract, pyautogui as p
import os.path

from PIL import Image

from _comum.pyautogui_comum import _find_img, _click_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _get_host_name, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _escreve_header_csv
from _comum.dominio_comum import _login_web, _abrir_modulo

pytesseract.pytesseract.tesseract_cmd = r'V:\Setor Robô\Tesseract-OCR\tesseract.exe'

def abre_tela_saldo_recolher():
    while not _find_img('apuracao_tributos_federais.png', conf=0.9):
        p.hotkey('alt', 'p')
        time.sleep(0.5)
        p.press('t')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(1)

    p.hotkey('alt', 'c')

    while not _find_img('consulta_apuracao_tributos_federais.png', conf=0.9):
        time.sleep(1)
def consulta_saldo_recolher(empresa, andamentos):
    cod, cnpj, nome = empresa

    p.write(ano)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(3)
    p.press('del', presses=5)
    p.write(cod)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(3)

    if _find_img('apuracao_nao_calculada.png', conf=0.9):
        p.press('enter')
        time.sleep(0.5)
        p.press('backspace', presses=6)
        time.sleep(0.5)
        p.hotkey('alt', 'f')
        while not _find_img('apuracao_tributos_federais.png', conf=0.9):
            time.sleep(1)
        p.hotkey('alt', 'c')
        while not _find_img('consulta_apuracao_tributos_federais.png', conf=0.9):
            time.sleep(1)
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Tributos Federais não foi calculado!', 'Não é possivel ver se tem checkbox marcada']), nome=andamentos)
        return 'ok'
    _click_img('data.png', conf=0.9)

    # Defina a região de interesse na tela (coordenadas x, y, largura, altura)
    # Você pode usar pyautogui.displayMousePosition() para encontrar essas coordenadas
    x, y, largura, altura = 1026, 568, 1157, 588

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
    texto_extraido = texto_extraido.replace('\n', '')
    texto_extraido = texto_extraido.replace('.', '')
    texto_extraido = texto_extraido.replace(',', '')
    numero_formatado = '{:,.2f}'.format(int(texto_extraido) / 100)
    numero_formatado = numero_formatado.replace(',', '#')  # Substitui vírgula por um caractere temporário, '#' neste caso
    numero_formatado = numero_formatado.replace('.', ',')  # Substitui ponto por vírgula
    numero_formatado = numero_formatado.replace('#', '.')  # Substitui o caractere temporário ('#') por ponto
    # Substituindo o ponto pela vírgula e vice-versa

    if _find_img('checkbox.png', conf=0.9):
        checkbox = 'CheckBox Marcada'
    else:
        checkbox = 'CheckBox Não Marcada'

    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, numero_formatado, checkbox]), nome=andamentos)
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
    abre_tela_saldo_recolher()
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao, controle=controle, usando_bd=True,
                                         nome_rotina=andamentos +  f' - {_get_host_name()}', planilha=os.path.join('execução', andamentos + '.csv'))

        while True:
            # abre a empresa no domínio

            resultado = consulta_saldo_recolher(empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;VALOR;CHECKBOX', nome=andamentos)


if __name__ == '__main__':
    ano = p.prompt(text='Qual periodo base?', title='Script incrível', default='00/0000')
    andamentos = 'Consulta Saldo a Recolher Folha OCR'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    run(controle)