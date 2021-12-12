import win32print
import win32api
import os

def meusPdf(caminho):
    listaarquivos = os.listdir(caminho)

    file_pdf = []

    for file in  listaarquivos:
        extensao = file[-3:]

        # Verifica se possui extensao do arquivo
        if extensao.lower() == 'pdf':
            file_pdf.append(file)
    return file_pdf


print(meusPdf(r'C:\Users\Danilo Santana\Downloads\imprimir\Teste'))