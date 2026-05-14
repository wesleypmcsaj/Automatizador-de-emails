import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials
import gspread
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from send2trash import send2trash

import imaplib
import email
from email.header import decode_header
from fpdf import FPDF
import os
from datetime import datetime

import win32print
import win32api

import shutil
import re
import tkinter as tk
from tkinter import messagebox, ttk
import pdfkit

# Atualizado e funcional 19/12/2025

# Código 0
# Código 1
# Código 2
# Código 2.1
# Código 3
# Código 4

print("Copiando informações da tabela do Google Drive, por favor aguarde...")

time.sleep(2)

print("Dados Copiados com sucesso! Baixando os arquivos nescessários no Google Drive para enviar como anexo.")


config_pdf = pdfkit.configuration(
    wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
)

# Código 0 - criação e envio dos e-mails

# Configuração das credenciais do Google
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/gmail.send'
]
SERVICE_ACCOUNT_FILE = 'credenciais.json'  # Substitua pelo caminho para suas credenciais

# Autenticação com as APIs do Google
credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
sheets_service = build('sheets', 'v4', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)
gspread_client = gspread.authorize(credentials)

# ID da planilha e nome da aba
SPREADSHEET_ID = '16_zlC5bRdyGTqFcVFvRIBCYzP-fjoPN9i64tD5DGe5c'
SHEET_NAME = 'emails'

# Pasta local para armazenar anexos e comprovantes
ANEXOS_DIR = 'anexos e-mails'
COMPROVANTES_DIR = 'comprovantes de envio'

# Criar pastas se não existirem
os.makedirs(ANEXOS_DIR, exist_ok=True)
os.makedirs(COMPROVANTES_DIR, exist_ok=True)



# Lendo os dados da planilha do Google Sheets
sheet = gspread_client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
dados = sheet.get_all_records()


#Chat falou que é aqui que eu adiciono o meu modelo de HTML
# Carregar modelo HTML do email
with open("modelodeemailfinal.html", "r", encoding="utf-8") as f:
    template_html = f.read()


# Função para baixar arquivos do Google Drive
def baixar_anexo(nome_arquivo):
    query = f"name = '{nome_arquivo}'"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    if not files:
        print(f"Arquivo {nome_arquivo} não encontrado no Google Drive.")
        return None

    file_id = files[0]['id']
    request = drive_service.files().get_media(fileId=file_id)
    file_path = os.path.join(ANEXOS_DIR, nome_arquivo)

    with open(file_path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            print(f"Download de {nome_arquivo}: {int(status.progress() * 100)}% concluído.")

    return file_path


# Função para enviar e-mails
def enviar_email(destinatario, assunto, corpo, anexo):
    # Configuração do servidor SMTP do Gmail
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    # Essa conta abaixo eu criei para enviar os emails e logo abaixo tenho uma conta só para receber o backup desses emails enviados.

    email_usuario = 'protocolosexpedientesajpmc@gmail.com'  # Substitua pelo seu e-mail
    email_senha = 'inserir senha de app'  # Substitua pela sua senha ou app password

    # Criar o e-mail
    msg = MIMEMultipart()
    msg['From'] = email_usuario
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Adicionar CCO

    # msg['Bcc'] = 'protocolosexpedientepmcsaj@gmail.com' Esse é o e-mail anterior.

    msg['Bcc'] = 'copiadeemailsexpedientesaj@gmail.com'

    # Corpo do e-mail
    msg.attach(MIMEText(corpo, 'html'))

    # Anexar arquivo
    if anexo:
        with open(anexo, 'rb') as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(anexo))
            msg.attach(part)

    with open("pmcLOGO.png", "rb") as img:
        mime_img = MIMEImage(img.read())
        mime_img.add_header('Content-ID', '<logo>')
        mime_img.add_header('Content-Disposition', 'inline', filename="pmcLOGO.png")
        msg.attach(mime_img)

    # 🔼 FIM DA PARTE DO LOGO

    # Enviar o e-mail
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_usuario, email_senha)
            server.send_message(msg)
        print(f"E-mail enviado para {destinatario}.")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")


# Função para salvar comprovante com informações adicionais
def salvar_comprovante(destinatario, assunto, corpo_html, anexo=None):
    options = {
        'encoding': 'UTF-8',
        'enable-local-file-access': None
    }

    data_envio = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    info_anexo = ""

    if anexo:
        nome_arquivo = os.path.basename(anexo)
        tamanho_bytes = os.path.getsize(anexo)
        tamanho_kb = tamanho_bytes / 1024

        info_anexo = f"""
            <br>
            <b>Arquivo anexado:</b> {nome_arquivo}<br>
            <b>Tamanho do arquivo:</b> {tamanho_kb:.2f} KB
            """

    html_pdf = corpo_html.replace(
        "cid:logo",
        "file:///C:/Users/wesley/PycharmProjects/EMAIL/pmcLOGO.png"
    )

    html_comprovante = f"""
        {html_pdf}
        
    
    <br><br>
    <hr>

    <p style="font-size:12px">
    <b>Comprovante de envio</b><br>
    Destinatário: {destinatario}<br>
    Assunto: {assunto}<br>
    Data/Hora do envio: {data_envio}s
     
    {info_anexo}
    </p>
    """
    #Esse trecho de info anexo (nome do arquivo e tamanho acima), eu adicionei em 30/04/2026
    nome_arquivo_pdf = f"comprovante_{destinatario}_{int(time.time())}.pdf"

    caminho_comprovante = os.path.join(
        COMPROVANTES_DIR,
        nome_arquivo_pdf
    )

    pdfkit.from_string(html_comprovante, caminho_comprovante, configuration=config_pdf, options=options)

    print(f"Comprovante salvo em {caminho_comprovante}.")


# Processar cada linha da planilha
for linha in dados:
    destinatario = linha['Destinatario']
    assunto = linha['Assunto']
    secretaria = linha['Secretaria_responsavel_pela_resposta']
    nome_doc_anexado = linha['Nome_do_documento_a_ser_anexado']
    nome_doc_sec = linha['Nome_do_doc_enviado_pela_sec']
    numero_pa_pmc = linha.get('numero_pa_pmc',
                              'Sem número de PA')  # Pegando o valor da nova coluna (se não existir, usa 'Sem número de PA')

    corpo_html = template_html.format(
        numero_pa_pmc=numero_pa_pmc,
        secretaria=secretaria,
        nome_doc_anexado=nome_doc_anexado,
        nome_doc_sec=nome_doc_sec
    )

    # Baixar o anexo
    caminho_anexo = baixar_anexo(nome_doc_anexado)

    # Enviar o e-mail
    enviar_email(destinatario, assunto, corpo_html, caminho_anexo)

    # Salvar comprovante
    salvar_comprovante(destinatario, assunto, corpo_html, caminho_anexo)
    # Na linha acima eu adicionei o trecho "caminho_anexo" dentro em 30/04/2026
time.sleep(2)


#Código 3 - impressão dos arquivos


#print("Abaixo segue a listagem das impressoras disponíveis:")

# Lista todas as impressoras disponíveis
#lista_impressoras = win32print.EnumPrinters(2)
#print("---------------------")
#for i, impressora in enumerate(lista_impressoras):
#    print(f"{i}: {impressora[2]}")  # O nome da impressora está na posição [2]
#print("---------------------")


ARQUIVO_IMPRESSORA = "impressora_utilizada.txt"
CAMINHO_ARQUIVOS = r"C:\Users\wesley\PycharmProjects\EMAIL\comprovantes de envio"

# Teste rápido de impressão
def teste_impressao(printer_name):
    try:
        win32print.SetDefaultPrinter(printer_name)
        temp_file = os.path.join(os.environ["TEMP"], "teste_impressao.txt")
        with open(temp_file, "w") as f:
            f.write("Teste de impressão.")
        win32api.ShellExecute(0, "print", temp_file, None, os.environ["TEMP"], 0)
        return True
    except Exception as e:
        print(f"Erro no teste de impressão: {e}")
        return False

# Lê a impressora salva anteriormente
def ler_impressora_salva():
    if os.path.exists(ARQUIVO_IMPRESSORA):
        with open(ARQUIVO_IMPRESSORA, "r", encoding="utf-8") as f:
            linhas = f.readlines()
            if linhas:
                return linhas[-1].strip()
    return None

# Salva a impressora escolhida
def salvar_impressora(printer_name):
    with open(ARQUIVO_IMPRESSORA, "a", encoding="utf-8") as f:
        f.write(f"{printer_name}\n")

# Verifica fila de impressão
def obter_numero_jobs():
    try:
        printer_info = win32print.OpenPrinter(win32print.GetDefaultPrinter())
        jobs = win32print.EnumJobs(printer_info, 0, -1, 1)
        win32print.ClosePrinter(printer_info)
        return len(jobs)
    except Exception as e:
        print(f"Erro ao acessar a fila de impressão: {e}")
        return -1

# Verifica se o arquivo está desbloqueado
def arquivo_desbloqueado(caminho_arquivo, timeout=30):
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < timeout:
        try:
            with open(caminho_arquivo, "r"):
                return True
        except PermissionError:
            time.sleep(5)
    return False

# Impressão automática (sem interface)
def imprimir_com_esta(printer_name):
    try:
        win32print.SetDefaultPrinter(printer_name)
    except Exception as e:
        print(f"Falha ao definir impressora {printer_name}: {e}")
        return False

    lista_arquivos = sorted(os.listdir(CAMINHO_ARQUIVOS))
    arquivos_nao_impressos = []

    if not lista_arquivos:
        print("Nenhum arquivo encontrado para impressão.")
        return True

    for arquivo in lista_arquivos:
        caminho_arquivo = os.path.join(CAMINHO_ARQUIVOS, arquivo)
        try:
            win32api.ShellExecute(0, "print", caminho_arquivo, None, CAMINHO_ARQUIVOS, 0)
        except Exception as e:
            print(f"Erro ao imprimir {arquivo}: {e}")
            arquivos_nao_impressos.append(arquivo)
            continue

        tempo_inicio = time.time()
        while obter_numero_jobs() == 0:
            if time.time() - tempo_inicio > 30:
                arquivos_nao_impressos.append(arquivo)
                break
            time.sleep(1)

        tempo_inicio = time.time()
        while obter_numero_jobs() > 0:
            if time.time() - tempo_inicio > 120:
                arquivos_nao_impressos.append(arquivo)
                break
            time.sleep(5)

        if os.path.exists(caminho_arquivo):
            if not arquivo_desbloqueado(caminho_arquivo):
                arquivos_nao_impressos.append(arquivo)
                continue
            send2trash(caminho_arquivo)
        else:
            print(f"Aviso: O arquivo '{arquivo}' já não existe mais.")



    if arquivos_nao_impressos:
        print("Arquivos não impressos:")
        for arq in arquivos_nao_impressos:
            print(f"- {arq}")
        return False

    print("Todos os arquivos foram impressos com sucesso.")
    return True

# Interface para escolher nova impressora
def abrir_interface():
    def imprimir():
        impressora = impressora_var.get()
        if not impressora:
            messagebox.showerror("Erro", "Nenhuma impressora selecionada.")
            return

        if not teste_impressao(impressora):
            messagebox.showerror("Erro", f"A impressora '{impressora}' não pôde ser usada.")
            return

        salvar_impressora(impressora)
        janela.destroy()
        imprimir_com_esta(impressora)

    janela = tk.Tk()
    janela.title("Selecionar Impressora")
    janela.geometry("500x200")

    tk.Label(janela, text="Impressora salva não funcionou.\nSelecione outra impressora:").pack(pady=10)

    lista_impressoras = [info[2] for info in win32print.EnumPrinters(2)]
    impressora_var = tk.StringVar()
    impressora_var.set(win32print.GetDefaultPrinter())

    menu = ttk.Combobox(janela, textvariable=impressora_var, values=lista_impressoras, state="readonly", width=50)
    menu.pack()

    tk.Button(janela, text="Imprimir com esta", command=imprimir, bg="#4CAF50", fg="white", font=("Arial", 12)).pack(pady=20)
    janela.mainloop()


# ============================
# PONTO DE ENTRADA DO SCRIPT
# ============================

impressora_salva = ler_impressora_salva()

if impressora_salva and impressora_salva in [info[2] for info in win32print.EnumPrinters(2)]:
    print(f"Tentando imprimir automaticamente com: {impressora_salva}")
    if not teste_impressao(impressora_salva):
        print("Falha ao imprimir com a impressora salva. Abrindo interface...")
        abrir_interface()
    else:
        imprimir_com_esta(impressora_salva)
else:
    print("Impressora salva não encontrada. Abrindo interface...")
    abrir_interface()


#Código 4 - mover pela última vez os comprovantes de email


# Caminhos das pastas
pasta_comprovantes = r'C:\Users\wesley\PycharmProjects\EMAIL\comprovantes de envio'
pasta_anexos = r'C:\Users\wesley\PycharmProjects\EMAIL\anexos e-mails'

# Função para enviar arquivos de uma pasta para a lixeira
def limpar_pasta(caminho_pasta):
    if os.path.exists(caminho_pasta):
        # Itera sobre todos os arquivos da pasta
        for arquivo in os.listdir(caminho_pasta):
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            # Verifica se é um arquivo e não uma pasta
            if os.path.isfile(caminho_arquivo):
                try:
                    # Envia o arquivo para a lixeira
                    send2trash(caminho_arquivo)
                    print(f"Arquivo {arquivo} enviado para a lixeira.")
                except Exception as e:
                    print(f"Erro ao mover o arquivo {arquivo}: {e}")
    else:
        print(f"A pasta {caminho_pasta} não foi encontrada.")

# Limpa os arquivos das pastas
limpar_pasta(pasta_comprovantes)
limpar_pasta(pasta_anexos)

print("E-mails enviados, comprovantes baixados e imprimidos com sucesso!")
