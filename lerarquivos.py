from buscadorArquivo import buscadorDeArquivo  # Importa a função de buscadorDeArquivo do arquivo buscadorArquivo.py
import pandas as pd  # Importa a biblioteca pandas
import os  # Importa a biblioteca os

def executarleitura():
    # Abrir o arquivo
    # file_path = buscadorDeArquivo()
    file_path = os.path.join(os.getcwd(), 'Dados p. preencher.xlsx')
    data = pd.read_excel(file_path, dtype=str)  # leitura do arquivo excel em formato texto

    # Dados do usuário para lsita
    nome = data['Nome_Funcionario'].tolist()
    setor = data['Setor'].tolist()
    funcao = data['Funcao'].tolist()
    celular = data['Celular Institucional'].tolist()
    telefone = data['Contato_Empresarial'].tolist()
    ramal = data['Ramal'].tolist()
    email = data['Email'].tolist()

    return nome, setor, funcao, celular, telefone, ramal, email

def exibir_dados(nome, setor, funcao, celular, telefone, ramal, email):
    # Imprimir dados do usuário
    print("{}, {}".format(nome[0], nome[1]))
    print("{}, {}".format(setor[0], setor[1]))
    print("{}, {}".format(funcao[0], funcao[1]))
    print("{}, {}".format(celular[0], celular[1]))
    print("{}, {}".format(telefone[0], telefone[1]))
    print("{}, {}".format(ramal[0], ramal[1]))
    print("{}, {}".format(email[0], email[1]))

if __name__ == "__main__":
    nome, setor, funcao, celular, telefone, ramal, email = executarleitura()
    exibir_dados(nome, setor, funcao, celular, telefone, ramal, email)