import win32print
import win32api
import os
import shutil
from time import sleep
from datetime import date


KILOBYTE = 1024

# Lista de impressora
def escolher_impressora(lst_impressora):
    opcao = 0
    while True:
        for key, impressora in enumerate(lst_impressora): 
            print(f'[{key + 1 } - {impressora}')
        print('=' * 10)

        try:
            opcao = int(input('Opção >>> '))
        
            # verificando se a escolha esta no intervalo
            if opcao > 0  and opcao <= len(lst_impressora):
                break
            else:
                print('Por favor informe a opção corretamente!')
        except ValueError:
            print('<<< ERRO! Digite um valor inteiro >>>')
        
    return opcao - 1


# Função que fica responsável por listar os arquivos PDF
def meusPdf(caminho, flag_doc = False):
    listaarquivos = os.listdir(caminho)

    file_pdf = []

    for file in  listaarquivos:
        extensao = file[-3:]
        kilobytes = os.path.getsize(rf'{caminho}/{file}') / 1024

        # Verifica se possui extensao do arquivo
        if extensao.lower() == 'pdf':
            if (isCCIR(caminho, file) or isITR(file)) and flag_doc:
                file_pdf.append(file)
    return file_pdf

# Funcao que verifica se é um CCIR
def isCCIR(caminho, file):
    return file.upper() in 'CCIR' and os.path.getsize(rf'{caminho}/{file}') / KILOBYTE

# Funcao que verifica se é um ITR
def isITR(file):
    return file.upper() in 'ITR'


lista_impressoras = win32print.EnumPrinters(2)

print('~~~~~~ LISTA DE IMPRESSORA ~~~~~~~~')

escolha_impressora = escolher_impressora(lista_impressoras)

# caminho 
while True:
    caminho = str(input('Copie aqui a pasta dos PDF: ')).strip()

    if os.path.isdir(caminho):
        if caminho[-1] != '\\' or caminho[-1] != '/':
            caminho = f'{caminho}\\'
        break
    else:
        print(f'O caminho {caminho} nao foi entrado!')

# escolher a impressora que vai imprimir os PDF
win32print.SetDefaultPrinter(lista_impressoras[escolha_impressora][2])

# Listagem de arquivos
files_PDF = meusPdf(caminho, True)

# realizando a impressao
tot_impresso = tot_impresso_erro = 0
print(f'IMPRESSORA SELECIONA: {win32print.GetDefaultPrinter()}')
print(f'Pasta selecionada: {caminho}')
for pdf in files_PDF:
    print(f'Caminho do arquivo -> {caminho}{pdf}')
    print(f'Imprindo o arquivo {pdf} na impressora {win32print.GetDefaultPrinter()}')
    print('*' * 30)

    # Executando a impressão do arquivo
    try:
        win32api.ShellExecute(0, "print", pdf, None, caminho, 0)
        sleep(0.7)

        # criando uma pasta
        """dir_impresso = f'{caminho}impressos{date.today()}'

        # Pegando os arquivo que foram impressos criando uma nova pasta para os arquivo que foram impresso
        if not os.path.isdir(dir_impresso):
            os.makedirs(dir_impresso)

        # Movendo o arquivo
        try:  
            shutil.move(caminho + pdf, dir_impresso)
        except:
            print(f'Erro ao mover o arquivo [{pdf}]')"""
        
        tot_impresso += 1
    except:
        print(f'Houve um erro para imprimir o arquivo {pdf}')
        tot_impresso_erro += 1

print(f"Todos os {len(files_PDF)} arquivos foram impressos")
print(f'{tot_impresso} com sucesso')
print(f'{tot_impresso_erro} com erro')