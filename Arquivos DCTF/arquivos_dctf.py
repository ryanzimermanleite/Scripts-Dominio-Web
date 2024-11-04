# -*- coding: utf-8 -*-
import datetime, os, time, shutil, pyautogui as p

from _comum.pyautogui_comum import _find_img, _click_img
from _comum.comum_comum import _indice, _time_execution_monitor_db, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _concatena, _kill_process_by_name, _get_host_name
from _comum.dominio_comum import _login_web, _abrir_modulo, _login, imagens


def arquivos_dctf(empresa, periodo):
    try:
        cod, cnpj, nome, regime, movimento = empresa
    except:
        cod, cnpj, nome, regime = empresa
        movimento = False


    nome_arquivo = 'M:\DCTF_{}.RFB'.format(cod)
    
    # aguarda a tela abrir
    while not _find_img('dctf_mensal.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida', ''
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou', ''

        if _find_img('dctf_mensal_2.png', conf=0.9):
            break
        if _find_img('dctf_mensal_3.png', conf=0.9):
            break
        # relatórios
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        # informativos
        p.press('n')
        time.sleep(0.5)
        # federais
        p.press('f')
        time.sleep(0.5)
        # dctf
        p.press('d')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        # mensal
        p.press('m')
        time.sleep(2)

    # digita o período
    p.write(periodo)
    time.sleep(1)
    
    # desce para inserir o nome do arquivo
    p.press('tab')
    time.sleep(1)
    
    # apaga o nome que estiver digitado
    p.press('delete', presses=50)
    time.sleep(1)
    
    # digita o nome do arquivo
    p.write(nome_arquivo)
    time.sleep(1)
    
    # abre outros dados
    p.hotkey('alt', 'o')
    
    # aguarda a tela abrir
    while not _find_img('outros_dados.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida', ''
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou', ''
        if _find_img('outros_dados_2.png', conf=0.9):
            break
        if _find_img('outros_dados_3.png', conf=0.9):
            break

        # verifica a competência
        if _find_img('data_inicio.png', conf=0.9) or _find_img('data_inicio_2.png', conf=0.9):
            _click_img('data_inicio.png', conf=0.9, timeout=1)
            _click_img('data_inicio_2.png', conf=0.9, timeout=1)
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, regime, 'Competência informada menor que a data de início efetivo das atividades.']), nome=andamentos)
            print('❗ Competência informada menor que a data de início efetivo das atividades.')
            time.sleep(1)
            p.press('esc', presses=5)
            return 'ok', nome_arquivo

        time.sleep(2)

    if _find_img('declarante_nao_contribui_irpj.png'):
        regime = 'Declarante não é contribuinte do IRPJ'
    else:
        # configura o regime, forma de tributação
        if regime == 'Lucro Real':
            if not _find_img('real_estimativa.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_real_estimativa.png', conf=0.99)
                p.press('enter')

        if regime == 'Lucro Presumido':
            if not _find_img('presumido.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_presumido.png', conf=0.99)
                p.press('enter')
                
        if regime == 'Imunes':
            if not _find_img('imune.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_imune.png', conf=0.99)
                p.press('enter')
            
        if regime == 'Isentas':
            if not _find_img('isenta.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_isenta.png', conf=0.99)
                p.press('enter')

        if regime == 'Simples Nacional':
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, regime, 'Empresa Simples Nacional não gera arquivo.']), nome=andamentos)
            print('❗ Empresa Simples Nacional não gera arquivo.')
            return 'ok', nome_arquivo

        # configura o regime, qualificação
        time.sleep(1)
        p.click(1215, 217)
        time.sleep(0.5)
        # pj em geral
        p.write('PJ em Geral')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(1)
    
    # levantou balancete/balanço de suspensão e redução
    if regime == 'Lucro Real':
        if not _find_img('levantou_marcado.png'):
            p.click(414, 421)
            time.sleep(0.5)
    else:
        if _find_img('levantou_marcado.png'):
            p.click(414, 421)
            time.sleep(0.5)
        
    # situação da pj no mês da declaração
    p.click(1214, 488)
    time.sleep(0.5)
    # pj não se enquadra em nenhuma das situações anteriores no mês da declaração
    p.write('PJ nao')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    
    # criterio de reconhecimento das variações monetárias dos direitos de crédito e das obrigações do contribuinte, em função da taxa de câmbio
    if not _find_img('criterio_marcado.png'):
        p.click(413, 541)
        time.sleep(1)
    
    if movimento == 'Sem movimento':
        if not _find_img('pj_inativa_marcado.png'):
            p.click(412, 601)
            time.sleep(1)
    
    # sem alteração do regime ou não se aplica se for janeiro
    p.click(1215, 547)
    time.sleep(0.5)
    
    if _find_img('nao_se_aplica.png', conf=0.99):
        _click_img('nao_se_aplica.png', conf=0.99)
        time.sleep(1)
        
    """if periodo[:2] == '01':
        if _find_img('nao_se_aplica.png', conf=0.99):
            _click_img('nao_se_aplica.png', conf=0.99)
            time.sleep(1)
    else:
        if _find_img('sem_alteracao.png', conf=0.99):
            _click_img('sem_alteracao.png', conf=0.99)
            time.sleep(1)"""
    
    # ok
    p.hotkey('alt', 'o')
    time.sleep(1)
    
    # aguarda a tela fechar
    while _find_img('outros_dados.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida', ''
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou', ''
        time.sleep(1)
    
    # exportar arquivo
    p.hotkey('alt', 'x')
    
    print('>>> Gerando arquivo')
    # aguarda o arquivo gerar
    while _find_img('dctf_mensal.png', conf=0.9) or _find_img('dctf_mensal_2.png', conf=0.9) or _find_img('dctf_mensal_3.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida', ''
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou', ''
        time.sleep(1)
        ocorrencias = [('nao_gerou_arquivo.png', 'Não gerou arquivo', 'ok'),
                       ('saldo_nao_calculado.png', 'Saldo dos impostos não foi calculado no período - ' + periodo, 'ok'),
                       ('saldo_nao_calculado_2.png', 'Saldo dos impostos não foi calculado no período - ' + periodo, 'ok'),
                       ('saldo_nao_calculado_3.png', 'Saldo dos impostos não foi calculado no período - ' + periodo, 'ok'),
                       ('nao_tem_parametro.png', 'Não existe parametro para a vigência - ' + periodo, 'ok'),
                       ('exportacao_cancelada.png', 'Exportação cancelada', 'ok'),
                       ('final_da_exportacao.png', 'Arquivo gerado', 'arquivo gerado'),
                       ('final_da_exportacao_2.png', 'Arquivo gerado', 'arquivo gerado'),
                       ('final_da_exportacao_3.png', 'Arquivo gerado', 'arquivo gerado')]

        for ocorrencia in ocorrencias:
            if _find_img(ocorrencia[0], conf=0.9):
                p.press('enter')
                _escreve_relatorio_csv(';'.join([cod, cnpj, nome, regime, ocorrencia[1]]), nome=andamentos)
                print(f'❗ {ocorrencia[1]}')
                time.sleep(1)
                p.press('esc', presses=5)
                return ocorrencia[2], nome_arquivo
        
        if _find_img('imune_irpj.png', conf=0.9) or _find_img('imune_irpj_2.png', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('enter')
            resultado = mudar_regime('Imunes')
            if resultado == 'conexao perdida':
                return 'conexao perdida', ''
            if resultado == 'dominio desconectou':
                return 'dominio desconectou', ''
            
            regime = f'AE: {regime}, Domínio: Imunes'

        if _find_img('isenta_alerta.png', conf=0.9) or _find_img('isenta_alerta_2.png', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('enter')
            resultado = mudar_regime('Isentas')
            if resultado == 'conexao perdida':
                return 'conexao perdida', ''
            if resultado == 'dominio desconectou':
                return 'dominio desconectou', ''
            regime = f'AE: {regime}, Domínio: Isentas'
        
    print('❗ Erro inesperado')
    p.press('esc', presses=5)
    time.sleep(3)
    return 'ok', nome_arquivo


def mudar_regime(regime_certo):
    # abre outros dados
    p.hotkey('alt', 'o')
    
    # aguarda a tela abrir
    while not _find_img('outros_dados.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        if _find_img('outros_dados_2.png', conf=0.9):
            break
        if _find_img('outros_dados_3.png', conf=0.9):
            break
    
    if regime_certo == 'Imunes':
        if not _find_img('imune.png'):
            p.click(1215, 192)
            time.sleep(0.5)
            _click_img('seleciona_imune.png', conf=0.99)
            p.press('enter')
    
    if regime_certo == 'Isentas':
        if not _find_img('isenta.png'):
            p.click(1215, 192)
            time.sleep(0.5)
            _click_img('seleciona_isenta.png', conf=0.99)
            p.press('enter')
        
    # ok
    p.hotkey('alt', 'o')
    time.sleep(1)
    
    # aguarda a tela fechar
    while _find_img('outros_dados.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)

    # aguarda a tela fechar
    while _find_img('outros_dados_2.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)

    # aguarda a tela fechar
    while _find_img('outros_dados_3.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)

    # exportar arquivo
    p.hotkey('alt', 'x')
    
    
def mover_arquivo(nome_arquivo):
    nome_arquivo = nome_arquivo.replace('M:\DCTF', 'DCTF')
    os.makedirs('execução/Arquivos', exist_ok=True)
    final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Arquivos DCTF\\execução\\Arquivos'
    final_folder_editado = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Arquivos DCTF\\execução\\Arquivos editados'
    folder = 'C:\\'

    edita_arquivo(nome_arquivo, os.path.join(folder, nome_arquivo), os.path.join(final_folder, nome_arquivo), final_folder_editado)
    

def edita_arquivo(nome_arquivo, arquivo_original, arquivo_final, arquivo_final_editados):
    contador = 0
    conteudo_modificado = ''
    with open(arquivo_original, 'r') as arquivo:
        # Leia o conteúdo do arquivo
        conteudo = arquivo.readlines()
        
        # verifica os números de chave que podem já existir no arquivo original
        for count, linha in enumerate(conteudo[:-1], start=1):
            codigo_capturado = linha[34:38]
            if str(codigo_capturado) == '5952' or str(codigo_capturado) == '1708' or str(codigo_capturado) == '3208' or str(codigo_capturado) == '5706' or str(codigo_capturado) == '8045' or str(codigo_capturado) == '2991':
                contador += 1
                continue
            else:
                conteudo_modificado += linha
            
            linhas = _concatena(str(count + 1 - contador), 4, 'antes', '0')
        
    if contador > 0:
        os.makedirs(arquivo_final_editados, exist_ok=True)
        
        ultima_linha = conteudo[-1].replace('\n', '').strip()
        ultima_linha_limpa = ultima_linha[:-4]
        ultima_linha_editada = ultima_linha_limpa + linhas + '\n'
        
        """while True:
            try:"""
        # Abra o arquivo para escrita
        with open(os.path.join(arquivo_final_editados, nome_arquivo), 'w') as arquivo:
            # Escreva o conteúdo modificado de volta no arquivo
            arquivo.writelines(conteudo_modificado + ultima_linha_editada)
            """break
        except PermissionError:
            print('Tentando salvar o novo arquivo...')"""
                
        print('❗ Arquivo editado')
        shutil.move(arquivo_original, arquivo_final)
        _escreve_relatorio_csv(f'{nome_arquivo};Arquivo editado', nome='Relatório de arquivos editados')
    else:
        shutil.move(arquivo_original, arquivo_final)
        _escreve_relatorio_csv(f'{nome_arquivo};Arquivo original', nome='Relatório de arquivos editados')


#TRInternetMonitor.exe
#AppController.exe
@_time_execution_monitor_db
def run(controle):
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _login_web()
    _abrir_modulo('escrita_fiscal')

    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index=index, tempos=tempos,
                                         tempo_execucao=tempo_execucao, controle=controle, usando_bd=True, nome_rotina=andamentos +  f' - {_get_host_name()}',
                                         planilha=os.path.join('execução', andamentos + '.csv'))

        while True:
            if not _login(empresa, andamentos):
                break
            else:
                resultado, nome_arquivo = arquivos_dctf(empresa, periodo)
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

                if resultado == 'arquivo gerado':
                    mover_arquivo(nome_arquivo)
                    break
                
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
    
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'ok':
                    break
    

if __name__ == '__main__':
    periodo = p.prompt(text='Qual o período do arquivo', title='Script incrível', default='00/0000')
    empresas = _open_lista_dados()

    andamentos = 'Gerador de arquivos DCTF Domínio'
    controle = f'V:\\Setor Robô\\Scripts Python\\_comum\\_Monitoramento\\{andamentos} - {_get_host_name()}.txt'

    run(controle)
