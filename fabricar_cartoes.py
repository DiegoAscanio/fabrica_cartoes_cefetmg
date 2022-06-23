import pandas as pd 
import imgkit
from datetime import datetime, timedelta
import code128
import random
from unidecode import unidecode
import math
import numpy as np
from io import BytesIO
import base64
import pdb
import zipfile
import sys
from PIL import Image
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
import os

# carregar template da frente do cartao
with open('frente.tpl', 'r') as file:
    template_frente = file.read()

# carregar template do verso do cartao
with open('verso.tpl', 'r') as file:
    template_verso = file.read()

# definir opcoes do imgkit
img_options = {
    'format': 'png',
    'crop-x': 10,
    'crop-y': 10,
    'crop-h': 662,
    'crop-w': 1032,
}

# criar dicionario de arquivos do lote
arquivos_lote = {}

# define uma janela base para a interface gráfica
janela = tk.Tk()
janela.withdraw()

def carregar_arquivos_lote(arquivo_zip):
    global arquivos_lote
    arquivos_lote = {}
    with zipfile.ZipFile(arquivo_zip) as zip_ref:
        filelist = [
            info.filename for info in zip_ref.filelist
        ]
        for filename in filelist:
            arquivos_lote[filename] = zip_ref.read(filename)

def carregar_dataframe():
    lista_arquivos = arquivos_lote.keys()
    for nome_arquivo in lista_arquivos:
        if '.xls' in nome_arquivo: # encontramos o arquivo excel
            arquivo_excel = nome_arquivo
            break
    return pd.read_excel(arquivos_lote[arquivo_excel], dtype='str')

def tamanho_fonte(nome):
    if len(nome) < 42:
        tamanho = 36
    elif 42 <= len(nome) < 125:
        tamanho = 24
    else:
        tamanho = 18
    return str(tamanho) + 'px'

def normalizar(nome, underscore=False):
    if underscore:
        return unidecode(nome).upper().replace(' ', '_')
    else:
        return unidecode(nome).upper()

def completar(codigo_barras):
    return str(codigo_barras).zfill(16)

def carregar_foto(foto):
    im = Image.open(foto)
    buffered = BytesIO()
    im.save(buffered, format = 'JPEG')
    b64 = 'data:image/png;base64,'
    b64 += base64.b64encode(buffered.getvalue()).decode()
    return b64

def gerar_codigo_barras(codigo_barras, branco = True):
    if not branco:
        bar = code128.image(
            completar(codigo_barras)
        ).convert('RGBA')
        pixdata = bar.load()
        width, height = bar.size
        for y in range(height):
            for x in range(width):
                if pixdata[x, y] == (255, 255, 255, 255):
                    pixdata[x, y] = (255, 255, 255, 0)
    else:
        bar = code128.image(
            completar(codigo_barras)
        )
    buffered = BytesIO()
    bar.save(buffered, format = 'PNG')
    b64 = 'data:image/png;base64,'
    b64 += base64.b64encode(buffered.getvalue()).decode()
    return b64

def gerar_frente(nome, foto, sigla_curso, data_ingresso, data_termino, unidade):
    html = template_frente
    html = html.replace('@tamanho-fonte', tamanho_fonte(nome))
    html = html.replace('@foto', carregar_foto(BytesIO(arquivos_lote[foto])))
    html = html.replace('@nome', normalizar(nome))
    html = html.replace('@sigla-curso', sigla_curso)
    html = html.replace('@data-ingresso', data_ingresso)
    html = html.replace('@data-termino', data_termino)
    html = html.replace('@unidade', unidade)
    return html

def gerar_verso(matricula, curso, codigo_barras):
    html = template_verso
    html = html.replace('@matricula', matricula)
    html = html.replace('@curso', normalizar(curso))
    html = html.replace('@codigo-barras', gerar_codigo_barras(codigo_barras))
    html = html.replace('@numero-codigo-barras', codigo_barras)
    return html

def gerar_img_cartao(nome, foto, sigla_curso, data_ingresso, data_termino, unidade, matricula, curso, codigo_barras):
    html_frente = gerar_frente(nome, foto, sigla_curso, data_ingresso, data_termino, unidade)
    frente = Image.open(
        BytesIO(
            imgkit.from_string(
                html_frente,
                False,
                options = img_options
            )
        )
    ).convert('RGB')

    html_verso = gerar_verso(matricula, curso, codigo_barras)
    verso = Image.open(
        BytesIO(
            imgkit.from_string(
                html_verso,
                False,
                options = img_options
            )
        )
    ).convert('RGB')
    return frente, verso

def salvar_pdf(batch_cartoes): 
    arquivo_pdf = filedialog.asksaveasfilename(
        initialdir = '~/Desktop/',
        initialfile = 'cartoes.pdf',
        title = 'Salvar arquivo pdf dos cartões',
        filetypes = [('Arquivo pdf', '.pdf')]
    )
    head = batch_cartoes[0]
    left = batch_cartoes[1:]
    head.save(
        arquivo_pdf,
        save_all = True,
        append_images = left
    )
    return arquivo_pdf

def gerar_cartoes(df):
    df['INGRESSO'] = pd.to_datetime(df['INGRESSO'])
    df['TÉRMINO PREVISTO'] = pd.to_datetime(df['TÉRMINO PREVISTO'])
    batch_cartoes = []
    for i in range(len(df)):
        frente, verso = gerar_img_cartao(
            df['NOME'].at[i],
            df['FOTO DIGITAL'].at[i],
            df['SIGLA '].at[i],
            df['INGRESSO'].at[i].strftime('%d/%m/%Y'),
            df['TÉRMINO PREVISTO'].at[i].strftime('%d/%m/%Y'),
            df['UNIDADE'].at[i],
            df['MATRICULA'].at[i],
            df['CURSO'].at[i],
            df['CÓDIGO DE BARRAS'].at[i]
        )
        batch_cartoes.append(frente)
        batch_cartoes.append(verso)
    return batch_cartoes


def main():
    # 1. solicita ao usuário para carregar o arquivo zip
    arquivo_zip = filedialog.askopenfile(
        initialdir = '~',
        title = 'Selecione o arquivo zip do Lote de Cartões',
        filetypes = [('Arquivos zip', '.zip')]
    ).name
    # 2. carrega os arquivos do lote
    carregar_arquivos_lote(arquivo_zip)
    # 3. carrega o dataframe de cartões
    df = carregar_dataframe()
    # 4. gera os cartões
    batch_cartoes = gerar_cartoes(df)
    # 5. informa o usuário do sucesso da operacao
    messagebox.showinfo(title = None, message = 'Cartões gerados com sucesso!')
    # 6. solicita ao usuario o nome do arquivo pdf para salvar
    arquivo_pdf = salvar_pdf(batch_cartoes)
    # 6. informa o usuário onde os cartoes estao salvos 
    messagebox.showinfo(title = None, message = 'Cartões salvos em {}'.format(arquivo_pdf))
    janela.destroy()
    return

if __name__ == "__main__":
    main()
