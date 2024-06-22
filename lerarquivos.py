import pandas as pd
import os

def executarleitura():
    # Abrir o arquivo
    file_path = os.path.join(os.getcwd(), 'Dados p. preencher.xlsx')
    data = pd.read_excel(file_path, dtype=str)  # tudo como texto

    # Dados do usuário
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
    print("Nome:", ', '.join(nome))
    print("Setor:", ', '.join(setor))
    print("Função:", ', '.join(funcao))
    print("Celular:", ', '.join(celular))
    print("Telefone:", ', '.join(telefone))
    print("Ramal:", ', '.join(ramal))
    print("E-mail:", ', '.join(email))

if __name__ == "__main__":
    nome, setor, funcao, celular, telefone, ramal, email = executarleitura()
    exibir_dados(nome, setor, funcao, celular, telefone, ramal, email)
