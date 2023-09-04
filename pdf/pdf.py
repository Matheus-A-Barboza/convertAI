from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from xml.etree import ElementTree as ET
import os
import pandas as pd

# Carregar o conteúdo do arquivo XML (substitua 'seu_arquivo.xml' pelo nome do arquivo XML)

def pegar_infos(nome_arquivo, valores):
    print(f'Pegou as Infos de: {nome_arquivo}')
    with open(f'nfs/{nome_arquivo}', 'rb') as arquivo_xml:
        arquivo_xml = 'C:/Users/MatheusA/PycharmProjects/ConversorNF/pdf/nfs'
        tree = ET.parse(arquivo_xml)
        root = tree.getroot()

# Inicializar o documento PDF
        pdf_file = 'saida.pdf'
        doc = SimpleDocTemplate(pdf_file, pagesize=letter)
        story = []

        # Estilos para os parágrafos
        styles = getSampleStyleSheet()
        normal_style = styles['Normal']
        title_style = styles['Title']

# Adicionar título ao PDF (opcional)
        title = Paragraph("Documento PDF Gerado a partir de XML", title_style)
        story.append(title)
        story.append(Spacer(1, 20))  # Espaço entre o título e o conteúdo

        # Iterar através das tags XML e adicionar ao PDF como parágrafos
        for element in root.iter():
            text = ET.tostring(element, encoding='unicode')
            paragraph = Paragraph(text, normal_style)
            story.append(paragraph)

        # Construir o PDF
        doc.build(story)
        print(f'PDF gerado com sucesso: {pdf_file}')

        lista_arquivos = os.listdir('nfs')

        colunas = ['numero_nota', 'empresa_emissora', 'nome_cliente', 'logradouro', 'bairro', 'municipio', 'uf', 'peso']
        valores = []
        for arquivo in lista_arquivos:
            pegar_infos(arquivo, valores)

        tabela = pd.DataFrame(columns=colunas, data=valores)
        tabela.to_pdf('NotasFicais.xlsx', index=False)
        print('Arquivo Xlsx importado.')