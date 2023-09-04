import json
import xmltodict as xml
import os
import pandas as pd


def pegar_infos(nome_arquivo, valores):
    print(f'Pegou as Infos de: {nome_arquivo}')
    with open(f'nfs/{nome_arquivo}', 'rb') as arquivo_xml:
        dic_arquivo = xml.parse(arquivo_xml)
        print(json.dumps(dic_arquivo, indent=4))
        if 'NFe' in dic_arquivo:
            infos_nf = dic_arquivo['NFe']['infNFe']
        else:
            infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']
        numeroNota = infos_nf['@Id']
        empresaEmissora = infos_nf['emit']['xNome']
        nomeCliente = infos_nf['dest']['xNome']
        logradouro = infos_nf['dest']['enderDest']['xLgr']
        bairro = infos_nf['dest']['enderDest']['xBairro']
        municipio = infos_nf['dest']['enderDest']['xMun']
        uf = infos_nf['dest']['enderDest']['UF']
        valorNota = infos_nf['total']['ICMSTot']['vBC']
        valorICMS = infos_nf['total']['ICMSTot']['vICMS']
        # valorDesconto = infos_nf['cobr']['fat']['vDesc']
        valorImposto = infos_nf['total']['ICMSTot']['vTotTrib']
        DtVencimento = infos_nf['cobr']['dup']['dVenc']

        if 'vDesc' in infos_nf['cobr']:
            valorDesconto = infos_nf['cobr']['fat']['vDesc']
        else:
            valorDesconto = '0.0'

        if 'COFINS' in infos_nf['det']:
            COFINS = infos_nf['det']['imposto']['COFINS']['COFINSAliq']['vBC']
        else:
            COFINS = 'Nao Informado'

        if 'prod' in infos_nf['det']:
            qntProd = infos_nf['det']['prod']['qCom']
        else:
            qntProd = 'Nao Informado'


        if 'vol' in infos_nf['transp']:
            peso = infos_nf['transp']['vol']['pesoB']
        else:
            peso = 'Não Informado'
    # valores.append([numeroNota, empresaEmissora, nomeCliente, endereco, produtos, valorNota, peso])
    valores.append([numeroNota, empresaEmissora, nomeCliente, logradouro, bairro, municipio, uf,
                    peso, valorNota, DtVencimento, valorICMS, valorImposto, valorDesconto, COFINS, qntProd])


lista_arquivos = os.listdir('nfs')

colunas = ['numeroNota', 'empresaEmissora', 'nomeCliente', 'logradouro', 'bairro', 'municipio', 'uf',
           'peso', 'valorNota', 'DtVencimento', 'valorICMS', 'valorImposto', 'valorDesconto', 'COFINS', 'qntProd']
valores = []
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('NotasFicais.xlsx', index=False)
print('Arquivo Xlsx importado.')