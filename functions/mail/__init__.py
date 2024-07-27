# SDG
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from email.mime.image import MIMEImage
from functions.secrets import tenant_id, client_id, client_secret
from tkinter import messagebox
from email.utils import make_msgid

import os
import mimetypes
import smtplib
import base64
import ssl
import msal


EMAIL_ADDRESS = 'vendas@comagro.com.br'  # Endereço de email do remetente


def add_attachment(msg, filepath):
    """    Lê e adiciona anexo ao email    """
    if not os.path.isfile(filepath):
        # Caso não haja arquivo no caminho especificado
        messagebox.showinfo("Erro!", f"Arquivo não encontrado. Entrar em contato com o TI.")
        return False

    filenumber = filepath.split('\\')[-1].split('.')[0]
    filename = f'cotacao_comagro-{filenumber}'

    # Determine the content type of the file
    ctype, encoding = mimetypes.guess_type(filepath)

    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'

    # Salva o tipo e subtipo do arquivo
    maintype, subtype = ctype.split('/', 1)

    if subtype == 'vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        subtype = 'xlsx'

    # Read the file and set the correct MIME type
    with open(filepath, 'rb') as file:
        file_data = file.read()
        if maintype == 'image':
            # For images, use MIMEImage
            attachment = MIMEImage(file_data, _subtype=subtype)
            attachment.add_header('Content-ID', '<{}>'.format(os.path.basename(filepath)))
            attachment.add_header('Content-Disposition', 'inline', filepath=os.path.basename(filepath))
        else:
            if subtype == 'pdf' or subtype == 'xlsx':
                pass
            else:
                # For all other file types, use MIMEApplication
                maintype = 'text'
                subtype = 'csv'
            attachment = MIMEApplication(file_data, _subtype=subtype)
            attachment.add_header(
                'Content-Disposition', 'attachment', filename=f'{filename}.{subtype}',
                filepath=os.path.basename(filepath))
        msg.attach(attachment)


def mail_bohe(msg, user, image_c_id):
    """ Creates email's body

    Args:
        msg (EmailMessage): EmailMessage object to add the body
        user (String): User that asked for the quotation
        image_c_id (String): Content-ID of the image to be attached

    Returns:
        EmailMessage: Returns the EmailMessage object with the body
    """
    signature = ''

    try:
        signature = f'<img src="cid:{image_c_id[1:-1]}", style="max-width: 500px; text-align: left;">'
    except:
        messagebox.showinfo(
            "Erro!",
            "Assinatura não encontrada. Entrar em contato com o TI.\nO programa continuará normalmente.")

    # Create the email footer
    footer = ("<div style='border:none;border-bottom:solid windowtext 1.0pt;padding:0cm 0cm 1.0pt 0cm'><p class='MsoNormal' style='border:none;padding:0cm'><span><u></u>&nbsp;<u></u></span></p></div>" +
              f"<p></p><p style='max-width: 70%;font-size: 12pt;'>Atenciosamente,<br/></p>") + signature

    # Create the email body and add the footer
    body = f"<p style='max-width: 70%;font-size: 13pt;'>Olá!" + \
        "</br>Envio a seguir uma cotação na planilha em anexo." + \
        "</br>Solicito sua resposta assim que possível.</p>" + \
        footer

    # body += f'<br><img src="cid:{image_c_id[1:-1]}">'

    # Add the body, with the footer, to the email
    msg.add_alternative(body, subtype='html')

    # Open the user image and attach it to the email
    with open(f'images\\{user}.png', 'rb') as img:
        # Guess the content type of the image and split into main and subtype
        maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
        # Attach the image to the email with the specified Content-ID
        msg.get_payload()[1].add_related(img.read(), maintype=maintype, subtype=subtype, cid=image_c_id)

    return msg


def get_oauth2_token():
    # URL de token do Azure Entra
    url = f"https://login.microsoftonline.com/{tenant_id}"

    # Escopo para enviar email
    scope = ['https://graph.microsoft.com/.default']

    # Autenticação usando MSAL
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=url,
        client_credential=client_secret
    )

    # Obtém o token de acesso
    token_response = app.acquire_token_for_client(scopes=scope)

    if 'access_token' in token_response:
        return token_response['access_token']
    else:
        return None


def encode_oauth2_string(username, access_token):
    auth_string = f'user={username}\1auth=Bearer {access_token}\1\1'
    return base64.b64encode(auth_string.encode()).decode()


def send_email(xslx_path, pdf_path, emails, user):
    """Envia o email ao cliente

    Args:
        xslx_path (string): path of the actual quotation
        pdf_path (string): path of the pdf quotation
        emails (list): list of emails to send the quotation
        user (string): user that asked for the quotation
    """
    # Gets quotation number
    number = xslx_path.split('\\')[-1].split('.')[0]

    # Cria um corpo de email e define o assunto
    msg = EmailMessage()
    msg['Subject'] = f'Cotação Comagro nº {number}'

    # Cria um texto plain e evita erro no get_playload
    msg.set_content('If you can see this, please consider updating your email client to support HTML emails.')

    # Cria um Content-ID para a imagem
    image_cid = make_msgid(domain='comagro.com.br')

    # Insert the email header and body
    msg = mail_bohe(msg, user, image_cid)

    # Attach the xslx to the email
    a = add_attachment(msg, f'{xslx_path}')

    # If the file was not found, return False
    if a is False:
        return False

    # Attach the pdf to the email
    a = add_attachment(msg, f'{pdf_path}')

    # If the file was not found, return False
    if a is False:
        return False

    access_token = get_oauth2_token()

    if access_token is None:
        messagebox.showinfo("Erro!", "Não foi possível obter o token de acesso.")
        return False

    # Send the email to each recipient
    with smtplib.SMTP('comagro-com-br.mail.protection.outlook.com', 25) as smtp:
        smtp.ehlo()
        smtp.starttls(context=ssl.create_default_context())  # Inicia a conexão TLS
        smtp.ehlo()

        auth_string = encode_oauth2_string(EMAIL_ADDRESS, access_token)

        try:
            smtp.docmd('AUTH XOAUTH2', auth_string)
        except smtplib.SMTPException as e:
            messagebox.showinfo("Erro de Autenticação!", f"[Erro]: {
                str(e)}\nNão foi possível autenticar usando OAuth2.")

        for email in emails:
            sender = EMAIL_ADDRESS
            to = email

            # Converte o e-mail em um formato legível para o smtplib
            msg['From'] = sender  # Adiciona o remetente ao cabeçalho
            msg['To'] = to  # Adiciona o destinatário ao cabeçalho
            raw = msg.as_string()

            # Envia o e-mail
            try:
                smtp.sendmail(sender, to, raw)
            except smtplib.SMTPException as e:
                messagebox.showinfo("Erro de Envio!", f"[Erro]: {str(e)}\nEmail: {email} Lista: {
                    emails}\n\nO programa continuará a enviar os emails restantes.")
