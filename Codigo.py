# importação das bibliotecas que serão utilizadas.
import PyPDF2 as pdf
from pathlib import Path
import openpyxl
import smtplib
from email.message import EmailMessage

nome_arquivo = "D:\ProjetosPython\EnvioAutomatico_EspelhoPonto\espelho_m3.pdf" # salvar o arquivo
arquivo_pdf = pdf.PdfReader(nome_arquivo)        # função que lê arquivos

# gerando uma folha de ponto para cada funcionário em arquivo separado.
i = 1
for pagina in arquivo_pdf.pages:    # para cada página dentro do arquivo completo:
    arquivo_novo = pdf.PdfWriter()  # cria novos arquivos
    arquivo_novo.add_page(pagina)   # adciona cada página

    texto_pagina = pagina.extract_text()  # extração do texto
    pos_nome = texto_pagina.find('Nome')  # encontra o nome
    pos_cracha = texto_pagina.find('Crachá') # encontra o crachá
    nome = texto_pagina[pos_nome+5:pos_cracha].strip() # cria o arquivo com o nome da pessoa

    # salvando o arquivo
    with Path(f'D:\ProjetosPython\EnvioAutomatico_EspelhoPonto\Folhas\{nome}.pdf').open(mode='wb') as arquivo_final:
        arquivo_novo.write(arquivo_final)
        i += 1

# Caminho para o arquivo PDF
caminho_arquivo = "D:\ProjetosPython\EnvioAutomatico_EspelhoPonto\espelho_m3.pdf"

# Caminho para o arquivo Excel contendo endereços de e-mail
lista_emails = "D:\ProjetosPython\EnvioAutomatico_EspelhoPonto\Lista_emails.xlsx"

# Leia os endereços de e-mail do arquivo Excel
lendo_arquivo = openpyxl.load_workbook(lista_emails)
linhas = lendo_arquivo.active

# Iterar sobre cada linha na planilha do Excel
for linha in linhas.iter_rows(values_only=True):
    funcionario, email = linha

    # Gere o nome do arquivo PDF com base no nome da planilha do Excel
    espelho_ponto = f"D:\ProjetosPython\EnvioAutomatico_EspelhoPonto\Folhas\{funcionario}.pdf"
    # print(pdf_file_name)

    # Crie uma nova mensagem de e-mail
    message = EmailMessage()
    message["Subject"] = "Folhas de ponto"
    message["From"] = "e-maildoresponsavel"
    message["To"] = email

    # Anexe o arquivo PDF ao e-mail
    with open(espelho_ponto, "rb") as pdf_file:
        message.add_attachment(pdf_file.read(), maintype="application", subtype="pdf", filename=espelho_ponto)

    # Envie o e-mail
    with smtplib.SMTP("provedor.e-mail", 587) as smtp_server:
        smtp_server.starttls()
        smtp_server.login("emaildoresponsavel", "senha_email")
        smtp_server.send_message(message)




