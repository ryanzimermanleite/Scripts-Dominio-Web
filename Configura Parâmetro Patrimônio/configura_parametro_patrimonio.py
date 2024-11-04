# -*- coding: utf-8 -*-
import datetime, os, time, shutil, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status, _kill_process_by_name
from dominio_comum import _login_web, _abrir_modulo, _login, _verifica_empresa

imagens = "V:\\_imagens_comum_python\\imgs_comum_dominio"


def verifica_parametro(regime, cnpj, periodo):
    regime = regime.lower()
    print('>>> Aguardando tela de parâmetros')
    print(regime)
    while not _find_img('parametros_2', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        
        if _find_img('parametros_2.png', conf=0.9):
            break
        p.hotkey('alt', 'c')
        time.sleep(1)
        p.press('p')
        time.sleep(1)

    print('>>> Configurando parâmetros')
    # seleciona o periodo inicial
    p.press('tab')
    time.sleep(1)
    p.write(periodo)
    time.sleep(1)
    
    p.press('tab')
    p.press('tab')
    
    # marca a opção calcular a depreciação considerando todos os dias do mês na competência que os bens foram adquiridos ou baixados
    while not _find_img('configuracoes_aba_geral.png', conf=0.99):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        print('>>> Marcando calcular a depreciação')
        _click_img('calcular_depreciacao.png', conf=0.99)
        time.sleep(1)

    # vai até a aba impostos
    _click_img('aba_impostos.png', conf=0.99)
    while not _find_img('descricao_impostos.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    print('>>> Configurando aba de impostos')
    _click_position_img('descricao_impostos.png', '+', pixels_x=100, conf=0.99, clicks=3)
    time.sleep(1)
    p.write('VIGENCIA')

    
    # vai até a aba ganhos de capital
    print('>>> Indo para a aba de ganhos de capital')
    _click_img('aba_ganhos_capital.png', conf=0.99)
    while not _find_img('tela_ganhos_capital.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    print('>>> Configurando aba de ganhos de capital')
    if _find_img('opcao_integrar_ganhos_capital_marcado.png', conf=0.99):
        _click_img('opcao_integrar_ganhos_capital_marcado.png', conf=0.99)
    if _find_img('opcao_integrar_2_ganhos_capital_marcado.png', conf=0.99):
        _click_img('opcao_integrar_2_ganhos_capital_marcado.png', conf=0.99)
        
        
    # vai até a aba pis cofins
    print('>>> Indo para a aba de pis cofins')
    _click_img('aba_pis_cofins.png', conf=0.99)
    while not _find_img('tela_pis_cofins.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    print('>>> Configurando aba de pis cofins')
    if regime == 'lucro real':
        if _find_img('opcao_calcular_depreciacao.png', conf=0.99):
            _click_img('opcao_calcular_depreciacao.png', conf=0.99)
    else:
        if _find_img('opcao_calcular_depreciacao_marcado.png', conf=0.99):
            _click_img('opcao_calcular_depreciacao_marcado.png', conf=0.99)
    

    # vai até a aba icms
    print('>>> Indo para a aba de ICMS')
    _click_img('aba_icms.png', conf=0.99)
    while not _find_img('tela_icms.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    
    print('>>> Configurando a aba de ICMS')
    if regime == 'simples nacional':
        if _find_img('opcao_calcular_icms_marcado.png', conf=0.99):
            _click_img('opcao_calcular_icms_marcado.png', conf=0.99)
    else:
        for imagem in ['opcao_calcular_icms.png', 'opcao_gerar_info_credito.png', 'opcao_calcular_conforme_pro_rata_die_marcado.png', 'opcao_nao_considerar_calculado_credito_icms.png',
                       'opcao_postergar_a_parcela_marcado.png', 'opcao_realizar_lancamento.png', 'opcao_utilizar_indexador_marcado.png']:
            if _find_img(imagem, conf=0.99):
                _click_img(imagem, conf=0.99)
            
    if _find_img('opcao_calcular_debito_icms_marcado.png', conf=0.99):
        _click_img('opcao_calcular_debito_icms_marcado.png', conf=0.99)
            
    # vai até > Contabilidade > Geral
    print('>>> Indo para a aba geral da contabilidade')
    _click_img('aba_contabilidade.png', conf=0.99)
    while not _find_img('aba_contabilidade_aberta.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    
    print('>>> Configurando a aba geral da contabilidade')
    # marca as opções da contabilidade
    while not _find_img('configuracoes_aba_contabilidade_marcada.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        print('>>> Marcando lançamentos contábeis')
        for checkbox in ['gerar_lancamento_contabeis.png', 'gerar_lancamento_contabeis_2.png', 'gerar_lancamento_contabeis_3.png', 'gerar_lancamento_contabeis_4.png']:
            _click_img(checkbox, conf=0.99)
        time.sleep(1)
    
    if _find_img('empresa_nao_selecionada.png', conf=0.9):
        _click_img('empresa_nao_selecionada.png', conf=0.9)
        time.sleep(1)
        numero_empresa = cnpj[11:12]
        p.press('down', presses=int(numero_empresa), interval=0.5)
        p.press('enter')
    time.sleep(1)

    # vai até > Contabilidade > Tipos de Lançamentos
    print('>>> Indo para a aba tipo de lançamentos da contabilidade')
    _click_img('aba_contabilidade.png', conf=0.99)
    while not _find_img('aba_tipos_de_lancamentos.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    _click_img('aba_tipos_de_lancamentos.png', conf=0.99)
    while not _find_img('aba_tipos_de_lancamentos_aberta.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)

    # marca as opções da aba tipos de lançamentos
    print('>>> Marcando tipos de lançamentos')

    opcoes = [['baixas_correta.png', 'baixas.png', 'opcao_baixas.png', 222], ['credito_correta.png', 'credito.png', 'opcao_credito.png', 160],
              ['aquisicao_correta.png', 'aquisicao.png', 'opcao_aquisicao.png', 180], ['transferencia_correta.png', 'transferencia.png', 'opcao_transferencia.png', 150]]

    for opcao in opcoes:
        print('Entra no loop')
        if not _find_img(opcao[0], conf=0.99):
            print('Não tem a opção certa', opcao[0])
            _click_position_img(opcao[1], '+', pixels_x=opcao[3], conf=0.99)
            print('Selecionando a opção', opcao[0])
            while not _find_img(opcao[2], conf=0.99):
                if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
                    return 'conexao perdida'
                if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
                    return 'dominio desconectou'
                print('aguardando a opção', opcao[0])
                time.sleep(1)
            print('clicando na opção', opcao[0])
            _click_img(opcao[2], conf=0.99, timeout=-1)
            print('opção clicada', opcao[0])
        
        print('Clicando na aba')
        _click_img('aba_tipos_de_lancamentos_2.png', conf=0.99, timeout=-1)
        print('Cliquei na aba')
        
    # vai até > Contabilidade > Impostos
    _click_img('aba_contabilidade.png', conf=0.95)
    while not _find_img('aba_cont_impostos.png', conf=0.99):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    _click_img('aba_cont_impostos.png', conf=0.99)
    while not _find_img('aba_cont_impostos_aberta.png', conf=0.9):
        if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
            return 'conexao perdida'
        if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
            return 'dominio desconectou'
        time.sleep(1)
    
    p.press('tab')
    print('>>> Marcando impostos da contabilidade')
    for codigo in [['597', 'PIS', 'Conta a recuperar'], ['182', 'PIS', 'Histórico'], ['609', 'COFINS', 'Conta a recuperar'], ['182', 'COFINS', 'Histórico']]:
        p.write(codigo[0])
        p.press('tab')
        time.sleep(1)
        for descricao in ['conta_invalida.png', 'historico_invalido.png']:
            if _find_img(descricao, conf=0.9):
                _click_img(descricao, conf=0.9, timeout=-1)
                p.press('enter')
                time.sleep(1)
                p.hotkey('alt', 'f')
                while not _find_img('deseja_gravar.png', conf=0.9):
                    if _find_img('conexao_perdida.png', pasta=imagens, conf=0.9):
                        return 'conexao perdida'
                    if _find_img('reconecta_dominio.png', pasta=imagens, conf=0.9):
                        return 'dominio desconectou'
                    time.sleep(1)
    
                _click_img('deseja_gravar.png', conf=0.9)
                p.hotkey('alt', 'n')
                return f'{codigo[1]} - {codigo[2]} {descricao.split("_")[1].replace(".png", "").replace("invalid", "inválid")}: {codigo[0]}'
    
    p.hotkey('alt', 'o')
    time.sleep(1)
    p.hotkey('alt', 'f')
    """while not _find_img('deseja_gravar.png', conf=0.9):
        time.sleep(1)
        
    _click_img('deseja_gravar.png', conf=0.9)
    p.hotkey('alt', 'S')"""
        
    return 'Parâmetros configurados'


@_time_execution
@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('patrimonio', usuario='joao', senha='1234')
    
    if _find_img('nao_existe_parametro_2.png', conf=0.9):
        _click_img('nao_existe_parametro_2.png', conf=0.9)
        p.press('enter')
    
    while not _find_img('parametros_2.png', conf=0.9):
        if _find_img('inicial.png', conf=0.9) or _find_img('inicial_2.png', conf=0.9):
            break
        time.sleep(1)
    
    time.sleep(5)
    for imagem in ['parametros_modulos.png', 'parametros_modulos_2.png', 'parametros_modulos_3.png']:
        if _find_img(imagem, conf=0.95):
            _click_img(imagem, conf=0.95)
            p.press('enter')
    
    p.hotkey('alt', 'f')
    while _find_img('parametros_2.png', conf=0.9):
        if _find_img('deseja_gravar.png', conf=0.9):
            _click_img('deseja_gravar.png', conf=0.9)
            p.hotkey('alt', 'n')
    
    _login(['1', '24982859000101', 'FACILITA SISTEMAS DE TECNOLOGIA LTDA'], andamentos)
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)
        cod, cnpj, nome, regime = empresa
        

        alerta = ''
        while True:
            resultado = _login(empresa, andamentos, retorna_erro_parametro=True)
            if not resultado:
                break
            
            if resultado == 'Sem parâmetros':
                _click_img('nao_existe_parametro_2.png', conf=0.9)
                p.press('enter')
                while not _find_img('parametros_2.png', conf=0.9):
                    time.sleep(1)
                
                time.sleep(5)
                for imagem in [('parametros_modulos.png', f'Conforme o módulo contabilidade os lançamentos contábeis serão gravados na empresa {cod}, porém os modulos Folha, Escrita Fiscal, estão definidos para empresas diferentes. Favor reconfigurar os parâmetros desses módulos.'),
                               ('parametros_modulos_2.png', f'Conforme o módulo contabilidade os lançamentos contábeis serão gravados na empresa {cod}, porém o modulo Escrita Fiscal, está definido para uma empresa diferente. Favor reconfigurar o parâmetro desse módulo.'),
                               ('parametros_modulos_3.png', f'Conforme o módulo contabilidade os lançamentos contábeis serão gravados na empresa {cod}, porém o módulo Folha, está definido para uma empresa diferente. Favor reconfigurar o parãmetro desse módulo.')]:
                    if _find_img(imagem[0], conf=0.95):
                        _click_img(imagem[0], conf=0.95)
                        p.press('enter')
                        alerta = imagem[1]
                        
                p.hotkey('alt', 'f')
                while _find_img('parametros_2.png', conf=0.9):
                    if _find_img('deseja_gravar.png', conf=0.9):
                        _click_img('deseja_gravar.png', conf=0.9)
                        p.hotkey('alt', 'n')
                
                if not _verifica_empresa(cod):
                    resultado = 'Empresa não encontrada'
                    print('❌ Empresa não encontrada')
                    break
                    
                print('>>> Sem parâmetros')
                resultado = verifica_parametro(regime, cnpj, periodo)
                if resultado == 'conexao perdida':
                    _kill_process_by_name('TRInternetMonitor')
                    _kill_process_by_name('AppController')
                    time.sleep(2)
                    _login_web()
                    _abrir_modulo('patrimonio', usuario='joao', senha='1234')
                    continue
                
                if resultado == 'dominio desconectou':
                    _click_img('reconecta_dominio.png', pasta=imagens, conf=0.9)
                    p.hotkey('alt', 'n')

                    _login_web()
                    _abrir_modulo('patrimonio', usuario='joao', senha='1234')
                    continue
                    
                break
            else:
                print('>>> Parâmetros ok')
                resultado = 'Parâmetros ok'
                break
        
        if resultado:
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{regime};{resultado};{alerta}', nome=andamentos)


if __name__ == '__main__':
    periodo = '01/2022'
    empresas = _open_lista_dados()
    andamentos = 'Verifica Parâmetro Patrimônio'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
