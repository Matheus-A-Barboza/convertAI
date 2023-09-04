from PySide6.QtWidgets import QApplication, QPushButton, QListWidget, QVBoxLayout, QLabel, QWidget, QFileDialog
import sys


class AppXMLtoExcel(QWidget):
    def __init__(self):
        super().__init__()

        self.xml = []

        self.setWindowTitle('XML to Excel')
        self.setGeometry(100, 100, 400, 600)

        self.txt_logo = QLabel('XML to Excel')

        self.btn_selecionar = QPushButton('Selecionar Arquivos')
        self.btn_selecionar.clicked.connect(self.adiconar_xml)

        self.lst_xml = QListWidget()

        layout = QVBoxLayout()
        layout.addWidget(self.btn_selecionar)


        self.setLayout(layout)

    def adiconar_xml(self):

        opcoes = QFileDialog.Options()
        opcoes |= QFileDialog.ReadOnly

        files, _ = QFileDialog.getOpenFileNames(self, 'Selecionar Arquivos XML', '', 'Arquivos .xml (*,.xml);;Todos os Arquivos (*)', options=opcoes)

        if files:
            self.xml.extend(files)

if __name__ == '__main__':
    app = QApplication()
    app_lista_tarefas = AppXMLtoExcel()
    app_lista_tarefas.show()
    sys.exit(app.exec())