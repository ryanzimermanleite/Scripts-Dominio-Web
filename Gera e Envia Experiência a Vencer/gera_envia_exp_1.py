# -*- coding: utf-8 -*-
import datetime, shutil, fitz, re, time, os, pyautogui as p, pyperclip as clip

from _comum.pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, _get_host_name, _escreve_header_csv, _kill_process_by_name
from _comum.dominio_comum import _login, _login_web, _abrir_modulo, _salvar_pdf, imagens


# guarda info da página anterior

def guarda_info(page, matchtexto_nome, matchtexto_cod):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    prevtexto_cod = matchtexto_cod
    return prevpagina, prevtexto_nome, prevtexto_cod


def cria_pdf(pasta_final_relatorios, nome, pdf, pagina1, pagina2):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        text = f'Relatório de Experiências Domínio Web - {comp} - {nome.replace("/", "")}.pdf'
        
        # Define o caminho para salvar o pdf
        os.makedirs(pasta_final_relatorios, exist_ok=True)
        arquivo = os.path.join(pasta_final_relatorios, text)
        
        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)
        
        new_pdf.save(arquivo)
        return arquivo


def separa(nome):

    pasta_final_relatorios = os.path.join(pasta_final, 'Relatórios de experiência separados Domínio Web')
    os.makedirs(pasta_final_relatorios, exist_ok=True)
    
    arquivo_base = 'V:\Setor Robô\Scripts Python\Domínio\Gera e Envia Experiência a Vencer\ignore\Relação de Empregados - Contratos_Vencimento_Modelo_Veiga.pdf'
    with fitz.open(arquivo_base) as pdf:
        # para cada página do pdf
        for page in pdf:
            # Pega o texto da pagina
            textinho = page.get_text('text', flags=1 + 2 + 8)
            # Procura o nome da empresa no texto do pdf
            if re.compile(nome).search(textinho):
                novo_relatorio = cria_pdf(pasta_final_relatorios, nome, pdf, page.number, page.number)
                break
    
    texto_pdf = ''
    with fitz.open(novo_relatorio) as pdf:
        # para cada página do pdf
        for page in pdf:
            # Pega o texto da pagina
            textinho = page.get_text('text', flags=1 + 2 + 8)
            texto_pdf += textinho + '\n'
        
    info_funcionarios = (re.compile(r'(.+)\n(\d+)\n(.+)\n(.+)\n1º Periodo\s+Nº Dias: (\d+) \n(\d\d/\d\d/\d\d\d\d)\nVencimento: (\d\d/\d\d/\d\d\d\d)\nProrrogação\s+Nº Dias: (\d+)\nVencimento: (\d\d/\d\d/\d\d\d\d)')
                         .findall(texto_pdf))

    return info_funcionarios
   
    
def escreve_dados(cod, nome):
    # cria a planilha de dados
    f = open(os.path.join('ignore', 'Dados.csv'), 'a', encoding='latin-1')
    f.write(f'{cod};CNPJ;{nome}\n')
    f.close()


def open_lista_dados(file, encode='latin-1'):
    # abre a planilha de dados
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        p.alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    print('>>> usando dados de ' + file.split('/')[-1])
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


def gera_arquivo(comp, cod='*', nome=''):
    # espera o botão de relatórios do domínio aparecer na tela
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    
    print('>>> Gerando o relatório')
    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('gerenciador_de_relatorios.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        
        # Relatórios
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        # gerador de relatórios
        p.press('i', presses=2)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(2)
    
    # escreve o nome da opção para selecionar
    time.sleep(0.5)
    
    print('>>> Buscando relatório')
    p.press('pgup', presses=20)
    contador = 0
    while not _find_img('relacao_empregados.png', conf=0.98):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        
        if _find_img('relacao_empregados_2.png', conf=0.98):
            break

        p.press('pgdn')
        time.sleep(0.5)
        contador += 1
        if contador > 50:
            print('Erro ao localizar o relatório, tentendo novamente')
            p.press('esc', presses=2)
            return False

    if _find_img('relacao_empregados.png', conf=0.97):
        _click_img('relacao_empregados.png', conf=0.97, clicks=2, timeout=1)
    else:
        _click_img('relacao_empregados_2.png', conf=0.97, clicks=2, timeout=1)

    timer = 0
    # espera aparecer o tipo do relatório que sera usado e depois clica nele
    while not _find_img('relatorio_modelo_veiga.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        
        print('>>> Buscando relatorio Veiga')
        time.sleep(1)
        timer += 1
        if timer > 5:
            p.press('pgdn')
            if _find_img('relacao_empregados.png', conf=0.97):
                _click_img('relacao_empregados.png', conf=0.97, clicks=2, timeout=1)
            if _find_img('relacao_empregados_2.png', conf=0.97):
                _click_img('relacao_empregados_2.png', conf=0.97, clicks=2, timeout=1)
        if timer > 30:
            p.press('pgup', presses=20)
            timer = 0

        if _find_img('relacao_empregados.png', conf=0.98):
            _click_img('relacao_empregados.png', conf=0.98, clicks=2, timeout=1)
        else:
            _click_img('relacao_empregados_2.png', conf=0.98, clicks=2, timeout=1)

    _click_img('relatorio_modelo_veiga.png', conf=0.9, timeout=1)

    while not _find_img('relatorio_modelo_veiga_tela.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        
        print('>>> Acessando relatorio Veiga')
        if _find_img('relacao_empregados.png', conf=0.98):
            _click_img('relacao_empregados.png', conf=0.98, clicks=2, timeout=1)
        else:
            _click_img('relacao_empregados_2.png', conf=0.98, clicks=2, timeout=1)
        _click_img('relatorio_modelo_veiga.png', conf=0.9, timeout=1)
        _click_img('relatorio_modelo_veiga_2.png', conf=0.9, timeout=1)
    
    # insere o código da empresa, '*' para selecionar todas
    time.sleep(1)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(cod)
    
    # insere '*' para selecionar todos os funcionários
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.press('*')

    hoje, data_subtraida = define_data()

    # insere a data de início e fim do contrato
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(data_subtraida)

    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(hoje)
    
    # executa
    time.sleep(0.5)
    p.hotkey('alt', 'e')

    if cod != '*':
        infos = separa(nome)
        print('>>> Aguardando gerar o relatório...')
        # enquanto o relatório não é gerado, verifica se aparece a mensagem dizendo que não possuí dados para emitir
        while not _find_img('contrato_experiencia.png', conf=0.9):
            if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
                return 'conexao perdida'
            if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
                return 'dominio desconectou'
            
            if _find_img('sem_dados.png', conf=0.9):
                for info in infos:
                    _escreve_relatorio_csv(';'.join([cod, nome, info[1], info[2], info[0], info[3], info[5], info[4], info[6], info[7], info[8], 'Não gerou relatório indivídual da empresa']), nome=andamentos, local=pasta_final)
                print('❌ Sem dados para emitir')
                p.press('enter')
                time.sleep(1)
                p.press('esc')
                time.sleep(1)
                return 'ok'

    # se gerar o relatório para todas as empresas só salva o pdf, se não tenta enviar o arquivo para o cliente
    if cod == '*':
        _salvar_pdf()
    else:
        if not _find_img('enviar_arquivo.png', conf=0.9):
            for info in infos:
                _escreve_relatorio_csv(';'.join([cod, nome, info[1], info[2], info[0], info[3], info[5], info[4], info[6], info[7], info[8], 'Não possuí opção de enviar o relatório para o cliente']), nome=andamentos, local=pasta_final)
            print('❌ Não possuí opção de enviar o relatório para o cliente')
            p.press('esc', presses=5)
            return 'ok'
            
        envia_experiencia(comp)
        for info in infos:
            _escreve_relatorio_csv(';'.join([cod, nome, info[1], info[2], info[0], info[3], info[5], info[4], info[6], info[7], info[8], 'Relatório enviado']), nome=andamentos, local=pasta_final)
        print('✔ Relatório enviado')

    # fechar qualquer possível tela aberta
    p.press('esc', presses=5)
    time.sleep(2)
    return 'ok'


def pega_empresas_com_exp():
    folder = 'C:\\'
    final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Gera e Envia Experiência a Vencer\\ignore'
    arquivo = 'Relação de Empregados - Contratos_Vencimento_Modelo_Veiga.pdf'
    try:
        os.remove(os.path.join(final_folder, arquivo))
    except:
        pass
    time.sleep(1)
    shutil.move(os.path.join(folder, arquivo), os.path.join(final_folder, arquivo))
    
    # Definir os padrões de regex
    padraozinho_nome1 = re.compile(r'Local\n(\d) - (.+)\n')
    padraozinho_nome2 = re.compile(r'Local\n(\d\d) - (.+)\n')
    padraozinho_nome3 = re.compile(r'Local\n(\d\d\d) - (.+)\n')
    padraozinho_nome4 = re.compile(r'Local\n(\d\d\d\d) - (.+)\n')
    prevtexto_nome = ''
    
    if os.path.exists(os.path.join('ignore', 'Dados.csv')):
        os.remove(os.path.join('ignore', 'Dados.csv'))
    
    # abre o pdf gerado no domínio com todas as empresas que possuem experiência a vencer
    with fitz.open(os.path.join(final_folder, arquivo)) as pdf:
        for page in pdf:
            andamento = f'Pagina = {str(page.number + 1)}'
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # Procura o nome da empresa no texto do pdf
                matchzinho_nome = padraozinho_nome1.search(textinho)
                if not matchzinho_nome:
                    matchzinho_nome = padraozinho_nome2.search(textinho)
                    if not matchzinho_nome:
                        matchzinho_nome = padraozinho_nome3.search(textinho)
                        if not matchzinho_nome:
                            matchzinho_nome = padraozinho_nome4.search(textinho)
                            if not matchzinho_nome:
                                continue
    
                # Guardar o nome da empresa
                matchtexto_nome = matchzinho_nome.group(2)
                # Guardar o código da empresa
                matchtexto_cod = matchzinho_nome.group(1)
                
                # se a empresa for igual a anterior, pula para à próxima, pois ela já foi inserida
                if matchtexto_nome == prevtexto_nome:
                    continue
                
                # coloca a empresa na planilha de dados e guarda a informação na variável para comparar com a próxima
                escreve_dados(matchtexto_cod, matchtexto_nome)
                prevtexto_nome = matchtexto_nome
            except:
                _escreve_relatorio_csv(andamento, nome='Erros', local=pasta_final)
                continue
    
    # seleciona a planilha gerada para usar no script
    empresas = open_lista_dados(os.path.join('ignore', 'Dados.csv'))
    
    return empresas


def envia_experiencia(comp):
    print('>>> Enviando relatório')
    # enquanto a janela de enviar não aparece clica no botão de enviar
    while not _find_img('publicar_doc.png', conf=0.9):
        _click_img('enviar_arquivo.png', conf=0.9)
        time.sleep(2)
    
    # verifica a pasta que está selecionada para inserir o arquivo no sistema
    # if not _find_img('pasta_pessoal_outros.png', conf=0.9):
    _click_img('drop.png', conf=0.9)
    time.sleep(0.5)
    _click_img('selecao_pasta_pessoal_outros.png', conf=0.9, timeout=2)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    
    # vai para o campo do nome do arquivo
    p.press('tab')
    time.sleep(1)
    
    # pega o nome do arquivo com o clipboard para garantir que não irá faltar os caracteres com acento
    while True:
        try:
            clip.copy(f'Relação de Empregados - Contratos_Vencimento_Modelo_Veiga - {comp}')
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
    
    # cola o nome do arquivo
    time.sleep(1)
    
    # enquanto a janela de enviar não aparece clica no botão de enviar
    while _find_img('publicar_doc.png', conf=0.9):
        # confirma o envio do arquivo
        p.hotkey('Alt', 'g')
        time.sleep(5)
        if _find_img('erro_nome.png', conf=0.9):
            _click_img('erro_nome.png', conf=0.9)
            p.press('enter')
            time.sleep(3)
            _click_position_img('campo_nome.png', '+', pixels_x=110, conf=0.9)
            time.sleep(0.2)
            # pega o nome do arquivo com o clipboard para garantir que não irá faltar os caracteres com acento
            while True:
                try:
                    clip.copy(f'Relação de Empregados - Contratos_Vencimento_Modelo_Veiga - {comp}')
                    break
                except:
                    pass

            # cola o nome do arquivo
            time.sleep(0.5)
            p.hotkey('ctrl', 'v')
            time.sleep(1)


def define_data():
    # Obter a data atual
    hoje = datetime.date.today()

    # Subtrair 90 dias da data atual
    data_subtraida = hoje - datetime.timedelta(days=90)

    hoje = hoje.strftime('%d/%m/%Y')
    data_subtraida = data_subtraida.strftime('%d/%m/%Y')

    return hoje.replace('//', '/'), data_subtraida.replace('//', '/')


@_time_execution_monitor_db
def run(controle):
    if novo == 'Não':
        # abre o Domínio Web e o módulo, no caso será o módulo Folha
        _login_web()
        _abrir_modulo('folha')
        
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index=index, tempos=tempos, tempo_execucao=tempo_execucao, controle=controle, usando_bd=True,
                                         nome_rotina=andamentos +  f' - {_get_host_name()}', planilha=os.path.join(pasta_final, andamentos + '.csv'))
        
        while True:
            # abre a empresa no domínio
            if not _login(empresa, andamentos, local=pasta_final):
                break
            else:
                # gera o arquivo específico da empresa
                resultado = gera_arquivo(comp, cod=empresa[0], nome=empresa[2])
                if resultado == 'conexao perdida':
                    _kill_process_by_name('TRInternetMonitor')
                    _kill_process_by_name('AppController')
                    time.sleep(2)
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
                    continue
                if resultado == 'dominio desconectou':
                    _click_img('reconecta_dominio.png', pasta=imagens, conf=0.9)
                    p.hotkey('alt', 'n')
                    
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
                    continue
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
                    
                if resultado == 'ok':
                    break
    
    _escreve_header_csv(';'.join(['CÓD. EMPRESA', 'RAZÃO SOCIAL', 'SITUAÇÃO RELATÓRIO', 'CÓD. FUNCIONÁRIO', 'NOME', 'FUNÇÃO', 'LOCAL', 'ADMISSÃO', 'DIAS 1º PERÍODO', 'VENCIMENTO', 'PRORROGAÇÃO', 'VENCIMENTO']), nome=andamentos, local=pasta_final)
    

if __name__ == '__main__':
    # define o nome da planilha de andamentos
    
    dia = datetime.datetime.now().strftime('%d')
    mes = datetime.datetime.now().strftime('%m')
    ano = datetime.datetime.now().strftime('%Y')
    
    comp = f'{str(dia)}-{str(mes)}-{str(ano)}'.replace('/', '-')
    empresas = ''
    index = ''
    
    pasta_final = '\\'.join([r'\\vpsrv03\Arq_Robo\Domínio\Experiência a vencer', str(ano), f'{str(mes)}-{str(ano)}', f'{str(dia)}-{str(mes)}-{str(ano)}'])
    andamentos = 'Gera e Envia Experiência a Vencer'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'
    
    # pergunta se deve gerar uma nova planilha de dados
    novo = p.confirm(title='Script incrível', text='Gerar nova planilha de dados?', buttons=('Sim', 'Não'))
    
    # se não for gerar uma nova planilha, seleciona a que já existe e pergunta se vai continuar de onde parou
    if novo == 'Não':
        empresas = _open_lista_dados()
        index = _where_to_start(tuple(i[0] for i in empresas))
        if index is not None:
            run(controle)
            
    # gera uma nova planilha e a seleciona
    if novo == 'Sim':
        # abre o Domínio Web e o módulo, no caso será o módulo Folha

        index = 0

        _login_web()
        _abrir_modulo('folha')
        # a função de gerar o relatório, pode ser usada para gerar individualmente para cada empresa ou geral, por padrão ela gera o relatório geral
        gera_arquivo(comp)

        empresas = pega_empresas_com_exp()
        run(controle)
