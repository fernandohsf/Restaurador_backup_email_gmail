import os
import time
import shutil
import pdfkit
import mailbox
import requests
from PyPDF2 import PdfMerger
from PIL import Image
from win32com import client
from email.utils import parseaddr, getaddresses
from email import message_from_bytes
from email.policy import default
from tipos_conteudos_email.calendario import tipo_calendario
from bs4 import BeautifulSoup

def decodificar(texto):
    try:
        return texto.decode("UTF-8", errors="replace")
    except UnicodeDecodeError:
        return texto.decode('latin-1')
    except Exception as e:
        print(f"Erro ao corrigir codificação: {e}")
        return ""

def agrupar_pdf_anexos(pasta_email, numero_email, caminho_pdf, anexos_gerados):
    pdf_temp = os.path.join(pasta_email, f"Email_{numero_email}_temp.pdf")
    os.rename(caminho_pdf, pdf_temp)
    anexos_gerados.insert(0, pdf_temp)

    with PdfMerger() as pdf_agrupado:
        for anexo in anexos_gerados:
            pdf_agrupado.append(anexo)
        pdf_agrupado.write(caminho_pdf)
        
    for anexo in anexos_gerados:
        os.remove(anexo)

def converter_para_pdf(extensao, caminho_anexo, anexo_convertido_pdf):
    extensao = extensao.replace('.', '')
    if extensao in ["png", "jpg", "jpeg", "bmp", "webp"]:
        with Image.open(caminho_anexo) as img:
            img.convert("RGB").save(anexo_convertido_pdf)
            os.remove(caminho_anexo)

    elif extensao == "txt":
        with open(caminho_anexo, "r", encoding="utf-8") as f:
            texto = f.read()
        html = f"<pre>{texto}</pre>"
        pdfkit.from_string(html, anexo_convertido_pdf)
        os.remove(caminho_anexo)

    elif extensao == "docx":
        word = client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(caminho_anexo)
        doc.SaveAs(anexo_convertido_pdf, FileFormat=17)  # 17 É O CÓDIGO PARA PDF
        doc.Close()
        word.Quit()
        os.remove(caminho_anexo)

    elif extensao == "pptx":
        powerpoint = client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 0
        presentation = powerpoint.Presentations.Open(caminho_anexo, WithWindow=False)
        presentation.SaveAs(anexo_convertido_pdf, 32)  # 32 É O CÓDIGO PARA PDF
        presentation.Close()
        powerpoint.Quit()
        os.remove(caminho_anexo)

    elif extensao == "xlsx":
        excel = client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(caminho_anexo)
        workbook.SaveAs(anexo_convertido_pdf, FileFormat=57)  # 57 É O CÓDIGO PARA PDF
        workbook.Close()
        excel.Quit()
        os.remove(caminho_anexo)

    return anexo_convertido_pdf

def mapear_extensao(extensao):
    if extensao == "vnd.openxmlformats-officedocument.wordprocessingml.document":
        return ".docx"
    if extensao == "vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        return ".xlsx"
    if extensao == "vnd.openxmlformats-officedocument.presentationml.presentation":
        return ".pptx"
    if extensao == "plain":
        return ".txt"
    return f".{extensao}"

def corpo_email(mensagem, pasta_email, numero_email):
    html_content = ""
    soup = None
    anexos = []
    imagens_corpo = []
    corpo_html = None

    tem_html = any(parte.get_content_type() == "text/html" for parte in mensagem.walk())

    def salvar_arquivo(caminho, conteudo):
        try:
            with open(caminho, "wb") as file:
                file.write(conteudo)
        except Exception as e:
            print(f"Erro ao salvar arquivo {caminho}: {e}")

    if mensagem.is_multipart():
        for parte in mensagem.walk():
            content_type = parte.get_content_type() or ""
            content_disposition = parte.get("Content-Disposition", "")

            try:
                if content_type == "text/html" and tem_html:
                    corpo_html = parte.get_payload(decode=True)
                    html_decodificado = decodificar(corpo_html)

                    # SANITIZAR O HTML USANDO BEAUTIFULSOUP
                    soup = BeautifulSoup(html_decodificado, "html.parser")
                    for tag in soup(["script", "style"]):
                        tag.decompose()

                    # PROCESSAMENTO DE IMAGENS EXTERNAS QUE ESTÃO NO HTML
                    for i, img_tag in enumerate(soup.find_all("img")):
                        src = img_tag.get("src")
                        if src and src.startswith("http"):
                            try:
                                resposta = requests.get(src, stream=True)
                                if resposta.status_code == 200:
                                    extensao = src.split(".")[-1].split("?")[0]
                                    if extensao not in ["png", "jpg", "jpeg", "gif", "bmp", "webp"]:
                                        extensao = "png"

                                    nome_imagem = f"imagem_externa_{numero_email}_{i}.{extensao}"
                                    caminho_imagem = os.path.join(pasta_email, nome_imagem)

                                    with open(caminho_imagem, "wb") as img:
                                        for chunk in resposta.iter_content(1024):
                                            img.write(chunk)
                                    img_tag["src"] = caminho_imagem.replace("\\", "/")
                                    imagens_corpo.append(caminho_imagem)

                            except Exception as e:
                                print(f"Erro ao baixar imagem externa: {e}")
                    html_content += soup.prettify()
    
                elif content_type == "text/plain" and not tem_html:
                    if not corpo_html:
                        texto_plain = parte.get_payload(decode=True)
                        html_content += f"<p>{decodificar(texto_plain)}</p>"
                
                if content_type == "text/calendar":
                    calendario = parte.get_payload(decode=True)
                    html_content += tipo_calendario(calendario)

                nome_anexo = parte.get_filename()
                extensao = content_type.split("/")[-1]

                # TRATAMENTO DE IMAGENS IMBUTIDAS NO E-MAIL
                if "inline" in content_disposition and "image" in content_type:
                    if extensao not in ["png", "jpg", "jpeg", "gif", "bmp", "webp"]:
                        extensao = "png"
                    
                    caminho_imagem = os.path.join(pasta_email, f"{nome_anexo}.{extensao}")
                    salvar_arquivo(caminho_imagem, parte.get_payload(decode=True))
                    img_tags = soup.find_all("img", alt=True)
                    for img_tag in img_tags:
                        if nome_anexo in img_tag["alt"]:
                            img_tag["src"] = caminho_imagem.replace("\\", "/")
                            html_content = soup.prettify()
                        else:
                            html_content = html_content.replace(f"cid:{nome_anexo}", caminho_imagem)
                    imagens_corpo.append(caminho_imagem)

                # TRATAMENTO PARA ANEXOS QUE NÃO SEJAM CALENDÁRIOS
                elif "attachment" in content_disposition and "ics" not in content_type:
                    extensao = mapear_extensao(extensao)
                    if not nome_anexo.endswith(f".{extensao}"):
                        nome_anexo = os.path.splitext(nome_anexo)[0] + f".{extensao}"
                        
                    caminho_anexo = os.path.join(pasta_email, f"Email_{numero_email}_Anexo_{nome_anexo}")
                    salvar_arquivo(caminho_anexo, parte.get_payload(decode=True))
                    anexos.append(converter_para_pdf(extensao, caminho_anexo, caminho_anexo.replace(extensao, ".pdf")))

                # TRATAMENTO DE IMAGENS QUE NÃO ESTÃO ESPECIFICADAS COMO ANEXO
                elif "image" in content_type and not content_disposition:
                    if extensao not in ["png", "jpg", "jpeg", "gif", "bmp", "webp"]:
                        extensao = "png"

                    nome_imagem_assinatura = f"assinatura_{numero_email}_{nome_anexo or 'imagem'}.{extensao}"
                    caminho_imagem_assinatura = os.path.join(pasta_email, nome_imagem_assinatura)

                    salvar_arquivo(caminho_imagem, parte.get_payload(decode=True))
                    anexos.append(caminho_imagem_assinatura)
                    html_content += f'<img src="{caminho_imagem_assinatura}" alt="Imagem de assinatura">'
                    
            except Exception as e:
                print(f"Erro ao processar parte do e-mail: {e}")

    else:
        try:
            corpo = mensagem.get_payload(decode=True)
            html_content += f"<p>{decodificar(corpo)}</p>"
        except Exception as e:
            print(f"Erro ao processar corpo de e-mail não multipart: {e}")

    return html_content, anexos, imagens_corpo

def salvar_email_como_pdf(mensagem, pasta_email, numero_email, tela):
    # CABEÇALHO E-MAIL
    assunto = mensagem.get("subject", "Sem Assunto")

    nome_remetente, email_remetente = parseaddr(mensagem.get("from", "Desconhecido"))
    if nome_remetente and email_remetente:
        remetente = f"{nome_remetente} ({email_remetente})"
    elif email_remetente:
        remetente = email_remetente
    elif nome_remetente:
        remetente = nome_remetente
    else:
        remetente = "Desconhecido"

    lista_destinatarios = getaddresses(mensagem.get_all("to", []))
    destinatarios = []
    for nome, email in lista_destinatarios:
        if nome and email:
            destinatarios.append(f"{nome} ({email})")
        elif email:
            destinatarios.append(email)
        elif nome:
            destinatarios.append(nome)

    destinatarios_formatados = ", ".join(destinatarios) if destinatarios else "Desconhecido"

    data = mensagem.get("date", "Desconhecida")

    # INÍCIO DO HMTL
    html_content = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }}
            .header {{ margin-bottom: 20px; }}
            .header div {{ margin: 5px 0; }}
            .email-body {{ border: 1px solid #ddd; padding: 15px; background-color: #f9f9f9; margin-top: 20px; }}
            img {{ max-width: 100%; height: auto; }}
        </style>
    </head>
    <body>
        <div class="header">
            <div><strong>Assunto:</strong> {assunto}</div>
            <div><strong>De:</strong> {remetente}</div>
            <div><strong>Para:</strong> {destinatarios_formatados}</div>
            <div><strong>Data:</strong> {data}</div>
        </div>
        <div class="email-body">
    """
    email, anexos_gerados, imagens_corpo = corpo_email(mensagem, pasta_email, numero_email)

    html_content += email

    # FECHAMENTO HMTL
    html_content += """
        </div>
    </body>
    </html>
    """

    # CAMINHO DO EXECUTÁVEL WKHTMLTOPDF
    caminho_wkhtmltopdf = f"{tela.caminhoExe}\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
    config = pdfkit.configuration(wkhtmltopdf=caminho_wkhtmltopdf)

    caminho_pdf = os.path.join(pasta_email, f"Email_{numero_email}.pdf")

    options = {
        'enable-local-file-access': None
    }

    try:
        pdfkit.from_string(html_content, caminho_pdf, configuration=config, options=options)
    except Exception as e:
        tela.adicionar_mensagem(f"Falha na criação do pdf. Tentando novamente.")
        time.sleep(1)
        tela.adicionar_mensagem("Criado com sucesso.")
    finally:
        if len(anexos_gerados)>0:
            agrupar_pdf_anexos(pasta_email, numero_email, caminho_pdf, anexos_gerados)
        
        for imagem in imagens_corpo:
            os.remove(imagem)

        tela.adicionar_mensagem(f"E-mail {numero_email} salvo em PDF.")

def processar_mbox_html(tela):
    try:
        #pasta_destino = "G:\\Drives compartilhados\\SUPER. EXEC - UTI\\Google Workspace\\E-mail Restaurado (Backup)" 
        pasta_destino = "C:\\Downloads"
        pasta_saida = ""
        tela.adicionar_mensagem("Preparando o arquivo para leitura.")
        time.sleep(1)
        tela.adicionar_mensagem("O tempo pode variar dependendo do tamanho do arquivo, por favor aguarde.")
        mbox = mailbox.mbox(tela.arquivo_mbox)
        total_emails = len(mbox)
        for i, mensagem in enumerate(mbox, start=1):
            try:
                if isinstance(mensagem, mailbox.mboxMessage):
                    mensagem = message_from_bytes(mensagem.as_bytes(), policy=default)
                else:
                    mensagem = message_from_bytes(mensagem.as_bytes())

                # CRIAÇÃO DA PASTA COM O NOME DO E-MAIL
                if i == 1:
                    tipo_email = mensagem.get("X-Gmail-Labels")
                    if "Enviado" in tipo_email:
                        pasta_saida = mensagem.get("From", "Desconhecido").split("<")[1].strip()
                        pasta_saida = pasta_saida.replace(">", "")
                    else:
                        pasta_saida = mensagem.get("Delivered-To", "Desconhecido")
                        if pasta_saida == 'Desconhecido':
                            pasta_saida = mensagem.get("To", "Desconhecido")
                    tela.atualizar_titulo(f"E-mail: {pasta_saida} \nConteúdo: {total_emails} e-mails.")
                    pasta_email = os.path.join(pasta_destino, pasta_saida)
                    if not os.path.exists(pasta_email):
                        os.makedirs(pasta_email)
                    tela.adicionar_mensagem(f"Iniciando recuperação.")
                time.sleep(1)

                salvar_email_como_pdf(mensagem, pasta_email, i, tela)

            except Exception as e:
                tela.adicionar_mensagem(f"Erro ao processar mensagem email {i}: {e}")

        tela.adicionar_mensagem("Processamento concluído.")
    except Exception as e:
        tela.adicionar_mensagem(f"Erro ao abrir o arquivo MBOX: {e}")
    finally:
        tela.botao_selecionar.config(state="normal")
    
        mbox.close()

        if os.path.exists(tela.pasta_temp):
            shutil.rmtree(tela.pasta_temp)