from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QPushButton, QListWidget, QVBoxLayout, QLabel, QWidget, QFileDialog, \
    QHBoxLayout, QMessageBox
import sys
import xmltodict as xml
import pandas as pd


class AppXMLtoExcel(QWidget):
    def __init__(self):
        super().__init__()

        self.xml_files = []

        self.setWindowTitle('XML to Excel')
        self.setGeometry(100, 100, 400, 600)

        self.txt_logo = QLabel('ConversorNFs')
        self.txt_logo.setAlignment(Qt.AlignCenter)  # Centralizar o texto

        self.btn_selecionar = QPushButton('Selecionar Arquivos XML')
        self.btn_selecionar.clicked.connect(self.adicionar_xml)

        self.btn_converter = QPushButton('Converter Notas Fiscais')
        self.btn_converter.clicked.connect(self.converter_notas)

        self.lst_xml = QListWidget()

        layout = QVBoxLayout()

        # Criar um layout horizontal para centralizar elementos
        horizontal_layout = QHBoxLayout()
        horizontal_layout.addWidget(self.txt_logo)
        layout.addLayout(horizontal_layout)

        layout.addWidget(self.btn_selecionar)
        layout.addWidget(self.btn_converter)
        layout.addWidget(self.lst_xml)

        self.setLayout(layout)

    def adicionar_xml(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        files, _ = QFileDialog.getOpenFileNames(self, 'Selecionar Arquivos XML', '',
                                                'Arquivos XML (*.xml);;Todos os Arquivos (*)', options=options)

        if files:
            self.xml_files.extend(files)
            self.lst_xml.clear()
            self.lst_xml.addItems(files)

    def converter_notas(self):
        if not self.xml_files:
            QMessageBox.warning(self, "Nenhum Arquivo Selecionado", "Nenhum arquivo foi selecionado.")
            return

        valores = []
        for arquivo in self.xml_files:
            self.pegar_infos(arquivo, valores)

        colunas = ['numeroNota', 'empresaEmissora', 'nomeCliente', 'logradouro', 'bairro', 'municipio', 'uf',
                   'peso', 'valorNota', 'DtVencimento', 'valorICMS', 'valorImposto', 'valorDesconto', 'COFINS',
                   'qntProd']

        tabela = pd.DataFrame(columns=colunas, data=valores)
        tabela.to_excel('NotasFiscais.xlsx', index=False)

        QMessageBox.information(self, "Conversão Concluída",
                                "Notas fiscais convertidas com sucesso para 'NotasFiscais.xlsx'.")

    def pegar_infos(self, nome_arquivo, valores):
        print(f'Pegou as Infos de: {nome_arquivo}')
        with open(nome_arquivo, 'rb') as arquivo_xml:
            dic_arquivo = xml.parse(arquivo_xml)
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
            valorDesconto = infos_nf['total']['ICMSTot'].get('vDesc', '0.0')
            valorImposto = infos_nf['total']['ICMSTot']['vTotTrib']
            DtVencimento = infos_nf['cobr']['dup']['dVenc']

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

        valores.append([numeroNota, empresaEmissora, nomeCliente, logradouro, bairro, municipio, uf,
                        peso, valorNota, DtVencimento, valorICMS, valorImposto, valorDesconto, COFINS, qntProd])


if __name__ == '__main__':
    app = QApplication()
    app_lista_tarefas = AppXMLtoExcel()
    app_lista_tarefas.show()
    sys.exit(app.exec())