from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QPushButton,
    QLineEdit, QFileDialog, QMessageBox, QSpacerItem, QSizePolicy
)
from PySide6.QtCore import Qt, QThread, Signal
import pandas as pd
import sys
import pyautogui as p
import os
import time
import pyperclip
from pathlib import Path

# 1.1

andamentos = 'Importação Plano de Contas'
e_dir = Path('execução')
class WorkerThread(QThread):
    # Sinais para atualizar os dois inputs de exibição
    update_signal_exibicao1 = Signal(str)
    update_signal_exibicao2 = Signal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.running = True

    def run(self):
        self.update_signal_exibicao2.emit('Processando!')
        df = pd.read_excel(self.file_path, usecols=[0])
        for cod in df.iloc[:, 0]:
            if not self.running:
                break
            # Atualiza o input_exibicao1 com o valor atual
            self.update_signal_exibicao1.emit(str(cod))
            # Chama a função para processar o valor e atualizar o input_exibicao2
            while True:
                if not self._login(cod):
                    break
                # abre a empresa no domínio
                resultado = self.importacao_plano_contas(cod)

                if resultado == 'ok':
                    break
        self.update_signal_exibicao2.emit('Finalizado!')

    def importacao_plano_contas(self, cod):

        while not self._find_img('titulo_empresas.png', conf=0.9):
            time.sleep(1)
            p.hotkey('alt', 'c')
            time.sleep(1)
            p.press('e')
            time.sleep(5)
        time.sleep(1)

        while not self._find_img('titulo_empresas_importacao.png', conf=0.9):
            time.sleep(1)
            self._click_img('botao_importacao.png', conf=0.9)
            time.sleep(3)
        time.sleep(1)

        while not self._find_img('conforme.png', conf=0.9):
            time.sleep(1)
            p.write('conf')
            time.sleep(3)
        time.sleep(1)

        p.press('tab')
        time.sleep(1)
        p.write('20')
        time.sleep(1)
        p.press('tab')
        time.sleep(1)

        p.hotkey('alt', 'i')

        while not self._find_img('confirma_copia.png', conf=0.9):
            time.sleep(1)
        time.sleep(1)

        p.hotkey('alt', 'y')

        while not self._find_img('progresso_importacao.png', conf=0.9):
            time.sleep(1)
        time.sleep(1)

        while not self._find_img('estrutura_dre.png', conf=0.9):
            time.sleep(1)
        time.sleep(1)

        posicao = p.locateCenterOnScreen('imgs/progresso_importacao.png', confidence=0.9)

        if posicao:
            # Move o cursor até a imagem e arrasta para baixo
            p.moveTo(posicao)
            p.mouseDown()
            time.sleep(0.2)  # tempo de segurança para assegurar o clique
            p.moveRel(0, 250)  # arrasta 300 pixels para baixo
            p.mouseUp()


        while not self._find_img('cadastro_dre.png', conf=0.9):
            time.sleep(1)
            if self._find_img('resumo_dados.png', conf=0.9):
                break
        time.sleep(1)

        if self._find_img('cadastro_dre.png', conf=0.9):
            self._click_img('cadastro_dre.png', conf=0.9)
            time.sleep(1)
            p.hotkey('alt', 'i')

        while not self._find_img('estrutura_dlpa.png', conf=0.9):
            time.sleep(1)
            if self._find_img('resumo_dados.png', conf=0.9):
                break
        time.sleep(1)

        if self._find_img('estrutura_dlpa.png', conf=0.9):
            self._click_img('estrutura_dlpa.png', conf=0.9)
            time.sleep(1)
            p.hotkey('alt', 'i')

        while not self._find_img('estrutura_dmpl.png', conf=0.9):
            time.sleep(1)
            if self._find_img('resumo_dados.png', conf=0.9):
                break
        time.sleep(1)

        if self._find_img('estrutura_dmpl.png', conf=0.9):
            self._click_img('estrutura_dmpl.png', conf=0.9)
            time.sleep(1)
            p.hotkey('alt', 'i')

        while not self._find_img('resumo_dados.png', conf=0.9):
            time.sleep(1)
        time.sleep(1)

        p.hotkey('alt', 'o')

        time.sleep(3)

        p.press('esc')

        time.sleep(5)

        self._escreve_relatorio_csv(
            f'{cod};OK', nome=andamentos)

        return 'ok'

    def _login(self, cod, retorna_erro_parametro=False, local=False, ignora_sem_parametro=False):
        def abre_tela_e_loga(cod):
            # espera abrir a janela de seleção de empresa
            while not self._find_img('trocar_empresa.png', pasta='imgs', conf=0.9):
                if self._find_img('trim_morreu.png', pasta='imgs', conf=0.9):
                    return 'conexao perdida'
                if self._find_img('conexao_perdida.png', pasta='imgs', conf=0.9):
                    return 'conexao perdida'
                if self._find_img('trk_internet_monitor.png', pasta='imgs', conf=0.9):
                    return 'conexao perdida'
                if self._find_img('reconecta_dominio.png', pasta='imgs', conf=0.9):
                    return 'dominio desconectou'

                p.press('f8')
                if self._find_img('trocar_empresa_2.png', pasta='imgs', conf=0.9):
                    break

            time.sleep(1)
            # clica para pesquisar empresa por código
            if self._find_img('codigo.png', pasta='imgs', conf=0.9):
                self._click_img('codigo.png', pasta='imgs', conf=0.9)

            p.write(cod)
            time.sleep(3)

            # confirmar empresa
            p.hotkey('alt', 'a')


        cod = str(cod)
        # espera a tela inicial do domínio
        while not self._find_img('inicial.png', pasta='imgs', conf=0.9):
            if self._find_img('trim_morreu.png', pasta='imgs', conf=0.9):
                return 'conexao perdida'
            if self._find_img('trk_internet_monitor.png', pasta='imgs', conf=0.9):
                return 'conexao perdida'
            if self._find_img('conexao_perdida.png', pasta='imgs', conf=0.9):
                return 'conexao perdida'
            if self._find_img('reconecta_dominio.png', pasta='imgs', conf=0.9):
                return 'dominio desconectou'
            if self._find_img('inicial_2.png', pasta='imgs', conf=0.9):
                break
            time.sleep(1)

        if self._find_img('onvio_desconectado.png', pasta='imgs', conf=0.99):
            self._click_img('onvio_desconectado.png', pasta='imgs', conf=0.99)
            time.sleep(2)

        p.click(833, 384)

        abre_tela_e_loga(cod)

        # enquanto a janela estiver aberta verifica exceções
        while self._find_img('trocar_empresa_2.png', pasta='imgs', conf=0.9):
            if self._find_img('fechar_tudo.png', pasta='imgs', conf=0.9):
                self._click_img('fechar_tudo.png', pasta='imgs', conf=0.9)
                for i in range(10):
                    p.press('esc')
                    time.sleep(1)
                abre_tela_e_loga(cod)

            time.sleep(1)
            if self._find_img('sem_parametro.png', pasta='imgs', conf=0.9):
                if retorna_erro_parametro:
                    return 'Sem parâmetros'
                if local:
                    self._escreve_relatorio_csv(f'{cod};Parametro não cadastrado para esta empresa',
                                           nome=andamentos, local=local)
                else:
                    self._escreve_relatorio_csv(f'{cod};Parametro não cadastrado para esta empresa',
                                           nome=andamentos)
                p.press('enter')
                time.sleep(1)
                while not self._find_img('parametros.png', pasta='imgs', conf=0.9):
                    time.sleep(1)
                p.press('esc')
                time.sleep(1)
                return False

            if not ignora_sem_parametro:
                if self._find_img('nao_existe_parametro.png', pasta='imgs', conf=0.9) or self._find_img(
                        'nao_existe_parametro_2.png', pasta='imgs', conf=0.9) or self._find_img(
                        'nao_existe_parametro_3.png', pasta='imgs', conf=0.9) or self._find_img(
                        'nao_existe_parametro_4.png', pasta='imgs', conf=0.9):
                    if retorna_erro_parametro:
                        return 'Sem parâmetros'
                    if local:
                        self._escreve_relatorio_csv(
                            f'{cod};Não existe parametro cadastrado para esta empresa',
                            nome=andamentos, local=local)
                    else:
                        self._escreve_relatorio_csv(
                            f'{cod};Não existe parametro cadastrado para esta empresa',
                            nome=andamentos)
                    p.press('enter')
                    time.sleep(1)
                    p.hotkey('alt', 'n')
                    time.sleep(1)
                    p.press('esc')
                    time.sleep(1)
                    p.hotkey('alt', 'n')
                    while self._find_img('trocar_empresa.png', pasta='imgs', conf=0.9):
                        time.sleep(1)
                    while self._find_img('trocar_empresa_2.png', pasta='imgs', conf=0.9):
                        time.sleep(1)
                    return False
            else:
                if self._find_img('nao_existe_parametro.png', pasta='imgs', conf=0.9) or self._find_img(
                        'nao_existe_parametro_2.png', pasta='imgs', conf=0.9) or self._find_img(
                        'nao_existe_parametro_3.png', pasta='imgs', conf=0.9) or self._find_img(
                        'nao_existe_parametro_4.png', pasta='imgs', conf=0.9):
                    while self._find_img('nao_existe_parametro.png', pasta='imgs', conf=0.9) or self._find_img(
                            'nao_existe_parametro_2.png', pasta='imgs', conf=0.9) or self._find_img(
                            'nao_existe_parametro_3.png', pasta='imgs', conf=0.9) or self._find_img(
                            'nao_existe_parametro_4.png', pasta='imgs', conf=0.9):
                        p.press('enter')
                        time.sleep(2)
                    while self._find_img('trocar_empresa_2.png', pasta='imgs', conf=0.9):
                        time.sleep(1)

                    time.sleep(2)
                    while self._find_img('tela_parametros.png', conf=0.9):
                        time.sleep(1)
                        p.press('esc')
                        time.sleep(1)

                    time.sleep(1)

            if (self._find_img('empresa_nao_usa_sistema.png', pasta='imgs', conf=0.9) or
                    self._find_img('empresa_nao_usa_sistema_2.png', pasta='imgs', conf=0.9) or
                    self._find_img('empresa_nao_usa_sistema_3.png', pasta='imgs', conf=0.9)):
                if local:
                    self._escreve_relatorio_csv(
                        f'{cod};Empresa não está marcada para usar este sistema', nome=andamentos,
                        local=local)
                else:
                    self._escreve_relatorio_csv(
                        f'{cod};Empresa não está marcada para usar este sistema', nome=andamentos)
                p.press('enter')
                time.sleep(1)
                p.press('esc', presses=5)
                while self._find_img('trocar_empresa.png', pasta='imgs', conf=0.9):
                    time.sleep(1)
                while self._find_img('trocar_empresa_2.png', pasta='imgs', conf=0.9):
                    time.sleep(1)
                return False

            if self._find_img('fase_dois_do_cadastro.png', pasta='imgs', conf=0.9) or self._find_img(
                    'fase_dois_do_cadastro_2.png', pasta='imgs', conf=0.9):
                p.hotkey('alt', 'n')
                time.sleep(1)
                p.hotkey('alt', 'n')

            if self._find_img('conforme_modulo.png', pasta='imgs', conf=0.9) or self._find_img('conforme_modulo_2.png',
                                                                                      pasta='imgs', conf=0.9):
                p.press('enter')
                time.sleep(1)

            if self._find_img('aviso_regime.png', pasta='imgs', conf=0.9):
                p.hotkey('alt', 'n')
                time.sleep(1)

            if self._find_img('aviso.png', pasta='imgs', conf=0.9):
                p.hotkey('alt', 'o')
                time.sleep(1)

            if self._find_img('erro_troca_empresa.png', pasta='imgs', conf=0.9):
                p.press('enter')
                time.sleep(1)
                p.press('esc', presses=5, interval=1)
                self._login(cod, retorna_erro_parametro, local, ignora_sem_parametro)

        if not self._verifica_empresa(cod):
            if local:
                self._escreve_relatorio_csv(f'{cod};Empresa não encontrada', nome=andamentos,
                                       local=local)
            else:
                self._escreve_relatorio_csv(f'{cod};Empresa não encontrada', nome=andamentos)
            p.press('esc')
            return False

        p.press('esc', presses=5)
        time.sleep(1)

        return True

    def click_img(self, img, pasta='imgs', conf=1.0, delay=1, timeout=20, button='left', clicks=1):
        img = os.path.join(pasta, img)
        try:
            aux = 0
            while True:
                box = p.locateCenterOnScreen(img, confidence=conf)
                if box:
                    p.click(p.locateCenterOnScreen(img, confidence=conf), button=button, clicks=clicks)
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

    def find_img(self, img, pasta='imgs', conf=1):
        try:
            path = os.path.join(pasta, img)
            if conf != 1:
                return p.locateCenterOnScreen(path, confidence=conf)
            else:
                return p.locateCenterOnScreen(path)
        except:
            return False

    _find_img = find_img

    def escreve_relatorio_csv(self, texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
        """Recebe um texto 'texto' junta com 'end' e escreve num arquivo 'nome'"""

        os.makedirs(local, exist_ok=True)

        try:
            f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
        except:
            f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

        f.write(texto + end)
        f.close()

    _escreve_relatorio_csv = escreve_relatorio_csv

    def verifica_empresa(self, cod):
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
            return False
        else:
            return True

    _verifica_empresa = verifica_empresa

    def stop(self):
        self.running = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setWindowTitle("Interface de Controle")
        self.setGeometry(0, 0, self.screen().size().width(), 40)

        central_widget = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        central_widget.setStyleSheet("""
            background-color: #3b4042;
            color: white;
        """)

        botao_largura = 150
        input_largura = 200
        exibicao_largura = 150

        # Seção Esquerda
        self.btn_selecionar = QPushButton("Selecionar Planilha")
        self.btn_selecionar.setFixedWidth(botao_largura)
        self.btn_selecionar.clicked.connect(self.selecionar_planilha)
        layout.addWidget(self.btn_selecionar)

        self.input_caminho = QLineEdit()
        self.input_caminho.setFixedWidth(input_largura)
        self.input_caminho.setPlaceholderText("Caminho da planilha selecionada")
        self.input_caminho.setDisabled(True)
        layout.addWidget(self.input_caminho)

        # Espaço entre a esquerda e o centro
        layout.addSpacerItem(QSpacerItem(20, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))

        # Seção Central (Inputs de Exibição) - Inicialmente Ocultos
        self.input_exibicao1 = QLineEdit()
        self.input_exibicao1.setFixedWidth(exibicao_largura)
        self.input_exibicao1.setDisabled(True)
        self.input_exibicao1.setVisible(False)  # Oculto inicialmente
        layout.addWidget(self.input_exibicao1)

        self.input_exibicao2 = QLineEdit()
        self.input_exibicao2.setFixedWidth(exibicao_largura)
        self.input_exibicao2.setVisible(False)  # Oculto inicialmente
        layout.addWidget(self.input_exibicao2)

        # Espaço entre o centro e a direita
        layout.addSpacerItem(QSpacerItem(20, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))

        # Seção Direita
        self.btn_iniciar = QPushButton("Iniciar")
        self.btn_iniciar.setFixedWidth(100)
        self.btn_iniciar.clicked.connect(self.iniciar)
        layout.addWidget(self.btn_iniciar)

        self.btn_parar = QPushButton("Parar")
        self.btn_parar.setFixedWidth(100)
        self.btn_parar.clicked.connect(self.parar)
        layout.addWidget(self.btn_parar)

        self.btn_sair = QPushButton("Sair")
        self.btn_sair.setFixedWidth(100)
        self.btn_sair.clicked.connect(self.sair)
        layout.addWidget(self.btn_sair)

        self.worker_thread = None
        self.file_path = None

    def selecionar_planilha(self):
        file_dialog = QFileDialog(self)
        caminho, _ = file_dialog.getOpenFileName(
            self, "Selecione uma planilha", "", "Arquivos Excel (*.xlsx)"
        )
        if caminho:
            self.file_path = caminho
            self.input_caminho.setText(caminho)

    def iniciar(self):
        if not self.file_path:
            QMessageBox.warning(self, "Aviso", "Selecione uma planilha antes de iniciar.")
            return
        # Exibir os inputs de exibição ao iniciar
        self.input_exibicao1.setVisible(True)
        self.input_exibicao2.setVisible(True)
        if not self.worker_thread or not self.worker_thread.isRunning():
            self.worker_thread = WorkerThread(self.file_path)
            # Conectar sinais para atualizar os campos de exibição
            self.worker_thread.update_signal_exibicao1.connect(self.atualizar_exibicao1)
            self.worker_thread.update_signal_exibicao2.connect(self.atualizar_exibicao2)
            self.worker_thread.start()

    def parar(self):
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.stop()

    def sair(self):
        self.parar()
        self.close()

    def atualizar_exibicao1(self, cod):
        self.input_exibicao1.setText('Empresa: ' + cod)

    def atualizar_exibicao2(self, processed_value):
        self.input_exibicao2.setText('Status: ' + processed_value)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
