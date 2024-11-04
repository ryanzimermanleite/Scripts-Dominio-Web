from sys import path

path.append(r'..\..\_comum')
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

import codecs, re, os
from bs4 import BeautifulSoup


caminho = f'V:\Setor Robô\Scripts Python\Domínio\Gera Relatório Conferencia Frente Da Nota\\execução\\notas_html\\'

def run():
    for arquivo in os.listdir(caminho):
        f = codecs.open(caminho + arquivo, 'r', 'latin-1')
        document = BeautifulSoup(f, features="lxml").get_text()
        print('Analisando arquivo: ' + arquivo)
        textinho = re.compile(r'((.+)\n.+\n.+\n.+\n.+\nCFOP)').findall(document)

        for i, elemento in enumerate(textinho):
                if elemento in textinho[i+1:]:
                    print(f'====== {arquivo} ======')
                    print(f'Elemento repetido: {elemento[0]}')
                    print('Nota Repetida:', elemento[0][0] + '\n')

                    _escreve_relatorio_csv(';'.join([arquivo, elemento[0][0], 'Nota Duplicada']), nome='Relatorio Notas Duplicadas')

if __name__ == '__main__':
    run()