from buscadorArquivo import buscadorDeArquivo
import pandas as pd
import os

class DadosFuncionario:
    def __init__(self):
        self.nome = []
        self.setor = []
        self.funcao = []
        self.celular = []
        self.telefone = []
        self.ramal = []
        self.email = []

    def executarleitura(self):
        # Abrir o arquivo
        file_path = buscadorDeArquivo()
        data = pd.read_excel(file_path, dtype=str)  # tudo como texto

        # Dados do usuário
        self.nome = data['Nome_Funcionario'].tolist()
        self.setor = data['Setor'].tolist()
        self.funcao = data['Funcao'].tolist()
        self.celular = data['Celular Institucional'].tolist()
        self.telefone = data['Contato_Empresarial'].tolist()
        self.ramal = data['Ramal'].tolist()
        self.email = data['Email'].tolist()

    def exibir_dados(self):
        # Imprimir dados do usuário
        print("Nome:", ', '.join(self.nome))
        print("Setor:", ', '.join(self.setor))
        print("Função:", ', '.join(self.funcao))
        print("Celular:", ', '.join(self.celular))
        print("Telefone:", ', '.join(self.telefone))
        print("Ramal:", ', '.join(self.ramal))
        print("E-mail:", ', '.join(self.email))

if __name__ == "__main__":
    df = DadosFuncionario()
    df.executarleitura()
    df.exibir_dados()