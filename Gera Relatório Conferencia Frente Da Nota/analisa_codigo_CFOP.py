from sys import path

path.append(r'..\..\_comum')
from comum_comum import _escreve_relatorio_csv

import codecs, re, os
from bs4 import BeautifulSoup


caminho = f'V:\Setor Robô\Scripts Python\Domínio\Gera Relatório Conferencia Frente Da Nota\\execução\\notas_html\\'

def run():
    for arquivo in os.listdir(caminho):
        f = codecs.open(caminho + arquivo, 'r', 'latin-1')
        document = BeautifulSoup(f, features="lxml").get_text()
        print('Analisando arquivo: ' + arquivo)

        x = re.compile(r'(outras\s+)(1949)').search(document)
        y = re.compile(r'(outras\s+)(5949)').search(document)

        #z = re.compile(r'(outras\s+)(3551)').search(document)##
        
        if x or y:
            _escreve_relatorio_csv(';'.join([arquivo, 'CFOP Encontrado']), nome='Relatorio CFOP')
       
if __name__ == '__main__':
    run()