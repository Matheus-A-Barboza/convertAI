import xmltodict as xml
import os
import pandas as pd

def pegar_infos(nome_arquivo, valores):
    print(f'Pegou as Infos de: {nome_arquivo}')
    with open(f'nfs/{nome_arquivo}', 'rb') as arquivo_xml:
        dic_arquivo = xml.parse(arquivo_xml)
        if 'NFe' in dic_arquivo:
            infos_nf = dic_arquivo['NFe']['infNFe']
        else:
            infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']
        numero_nota = infos_nf['@Id']
        empresa_emissora = infos_nf['emit']['xNome']
        nome_cliente = infos_nf['dest']['xNome']
        logradouro = infos_nf['dest']['enderDest']['xLgr']
        bairro = infos_nf['dest']['enderDest']['xBairro']
        municipio = infos_nf['dest']['enderDest']['xMun']
        uf = infos_nf['dest']['enderDest']['UF']
        if 'vol' in infos_nf['transp']:
            peso = infos_nf['transp']['vol']['pesoB']
        else:
            peso = 'Não Informado'
        valores.append([numero_nota, empresa_emissora, nome_cliente, logradouro, bairro, municipio, uf,
                        peso])


lista_arquivos = os.listdir('nfs')

colunas = ['numero_nota', 'empresa_emissora', 'nome_cliente', 'logradouro', 'bairro', 'municipio', 'uf', 'peso']
valores = []
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('NotasFicais.xlsx', index=False)
print('Arquivo Xlsx importado.')