import os
import time
import pdfkit
import mailbox
from email import message_from_bytes
from email.policy import default
from tipos_conteudos_email.calendario import tipo_calendario
from bs4 import BeautifulSoup

def decodificar(texto):
    try:
        return texto.decode("UTF-8", errors="replace")
    except Exception as e:
        print(f"Erro ao corrigir codificação: {e}")
        return ""
    
def corpo_email(mensagem, pasta_email, numero_email):
    html_content = ""
    if mensagem.is_multipart():
        for parte in mensagem.walk():
            content_type = parte.get_content_type() or ""
            content_disposition = parte.get("Content-Disposition", "")

            if content_type == "text/html":
                corpo_html = parte.get_payload(decode=True)
                html_decodificado = decodificar(corpo_html)
                # Sanitizar o HTML usando BeautifulSoup
                soup = BeautifulSoup(html_decodificado, "html.parser")

                # Remover tags indesejadas, como scripts ou estilos maliciosos
                for tag in soup(["script", "style"]):
                    tag.decompose()  # Remove o conteúdo dessas tags

                # Adicionar o HTML sanitizado ao conteúdo final
                html_content += str(soup)
            
            elif content_type == "text/plain":
                texto_plain = parte.get_payload(decode=True)
                html_content += f"<p>{decodificar(texto_plain)}</p>"
            
            elif content_type == "text/calendar":
                calendario = parte.get_payload(decode=True)
                html_content += tipo_calendario(calendario)

            nome_anexo = parte.get_filename()
            extensao = content_type.split("/")[-1]
            if "inline" in content_disposition and "image" in content_type:
                if extensao not in ["png", "jpg", "jpeg", "gif", "bmp", "webp"]:
                    extensao = "png"
                
                caminho_imagem = os.path.join(pasta_email, f"{nome_anexo}.{extensao}")
                html_content = html_content.replace(f"cid:{nome_anexo}", caminho_imagem)

            elif "attachment" in content_disposition and "ics" not in content_type:
                if "vnd.openxmlformats-officedocument.wordprocessingml.document" == extensao:
                    extensao = ".docx"

                elif "vnd.openxmlformats-officedocument.presentationml.presentation" == extensao:
                    extensao = ".pptx"

                elif "plain" == extensao:
                    extensao = ".txt"
                    
                extensao = f".{extensao}"

                if not nome_anexo.endswith(f".{extensao}"):
                    nome_anexo = os.path.splitext(nome_anexo)[0] + f".{extensao}"
                caminho_imagem = os.path.join(pasta_email, f"Email_{numero_email}_Anexo_{nome_anexo}")
                with open(caminho_imagem, "wb") as img:
                    img.write(parte.get_payload(decode=True))

    else:
        corpo = mensagem.get_payload(decode=True)
        html_content += f"<p>{decodificar(corpo)}</p>"
    
    return html_content

def salvar_email_como_pdf(mensagem, caminho_pdf, pasta_email, numero_email, tela):
    # Cabeçalho do e-mail
    assunto = mensagem.get("subject", "Sem Assunto")
    remetente = mensagem.get("from", "Desconhecido")
    destinatario = mensagem.get("to", "Desconhecido")
    data = mensagem.get("date", "Desconhecida")

    # Início do HTML
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
            <div><strong>Para:</strong> {destinatario}</div>
            <div><strong>Data:</strong> {data}</div>
        </div>
        <div class="email-body">
    """
    html_content += corpo_email(mensagem, pasta_email, numero_email)

    # Fechar HTML
    html_content += """
        </div>
    </body>
    </html>
    """

    # Caminho do executável wkhtmltopdf
    caminho_wkhtmltopdf = f"{tela.caminhoExe}\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"

    # Configurar o pdfkit com o caminho do wkhtmltopdf
    config = pdfkit.configuration(wkhtmltopdf=caminho_wkhtmltopdf)

    # Salvar como PDF usando pdfkit com a configuração
    options = {
    "no-images": "",  # Não carregar imagens externas
    "disable-local-file-access": "",  # Bloqueia acesso a arquivos locais
    }
    pdfkit.from_string(html_content, caminho_pdf, configuration=config, options=options)
    tela.adicionar_mensagem(f"E-mail {numero_email} salvo em PDF.")

def processar_mbox_html(tela):
    try:
        #pasta_destino = "G:\\Drives compartilhados\\SUPER. EXEC - UTI\\Google Workspace\\E-mail Restaurado (Backup)" 
        pasta_destino = "C:\\Downloads"
        pasta_saida = ""
        mbox = mailbox.mbox(tela.arquivo_mbox)
        total_emails = len(mbox)
        for i, mensagem in enumerate(mbox, start=1):
            try:
                if isinstance(mensagem, mailbox.mboxMessage):
                    mensagem = message_from_bytes(mensagem.as_bytes(), policy=default)
                else:
                    mensagem = message_from_bytes(mensagem.as_bytes())

                # Criar pasta com o endereço de e-mail.
                if i == 1:
                    pasta_saida = mensagem.get("to", "Desconhecido")
                    tela.atualizar_titulo(f"E-mail: {pasta_saida} \nConteúdo: {total_emails} e-mails.")
                    pasta_email = os.path.join(pasta_destino, pasta_saida)
                    if not os.path.exists(pasta_email):
                        os.makedirs(pasta_email)
                    tela.adicionar_mensagem(f"Iniciando recuperação.")
                time.sleep(1)

                # Caminho do PDF
                caminho_pdf = os.path.join(pasta_email, f"email_{i}.pdf")
                salvar_email_como_pdf(mensagem, caminho_pdf, pasta_email, i, tela)

            except Exception as e:
                if "Exit with code 1 due to network error" in str(e):
                    tela.adicionar_mensagem("Erro na tentativa de obtenção dos dados. Tentando novamente.")
                    time.sleep(1)
                    tela.adicionar_mensagem("Restaurado com sucesso.")
                else:
                    tela.adicionar_mensagem(f"Erro ao processar mensagem email {i}: {e}")

        tela.adicionar_mensagem("Processamento concluído.")
    except Exception as e:
        tela.adicionar_mensagem(f"Erro ao abrir o arquivo MBOX: {e}")
    finally:
        tela.botao_selecionar.config(state="normal")