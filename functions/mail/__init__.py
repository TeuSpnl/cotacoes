# SDG
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from email.mime.image import MIMEImage
from functions.secrets import senha
from tkinter import messagebox
# from functions.calen import log_dev
from email.utils import make_msgid

import os
import mimetypes
import smtplib

EMAIL_ADDRESS = 'noreply-faturas@comagro.com.br'  # Endereço de email do remetente
EMAIL_PASS = senha  # Senha do email do remetente


def add_attachment(msg, filename):
    """    Lê e adiciona anexo ao email    """

    if not os.path.isfile(filename):
        # Caso não haja arquivo no caminho especificado
        messagebox.showinfo(
            "Erro!", f"Planilha não encontrada. Entrar em contato com o magnífico TI.")
        return

    # Determine the content type of the file
    ctype, encoding = mimetypes.guess_type(filename)

    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'

    # Salva o tipo e subtipo do arquivo
    maintype, subtype = ctype.split('/', 1)

    # Read the file and set the correct MIME type
    with open(filename, 'rb') as file:
        file_data = file.read()
        if maintype == 'image':
            # For images, use MIMEImage
            attachment = MIMEImage(file_data, _subtype=subtype)
            attachment.add_header(
                'Content-ID', '<{}>'.format(os.path.basename(filename)))
            attachment.add_header('Content-Disposition',
                                  'inline', filename=os.path.basename(filename))
        else:
            # For all other file types, use MIMEApplication
            attachment = MIMEApplication(file_data, _subtype=subtype)
            attachment.add_header(
                'Content-Disposition', 'attachment', filename=os.path.basename(filename))
        msg.attach(attachment)


def mail_bohe(msg, image_c_id):
    """    Insere o header e o corpo no email    """

    # Create the email body and add the footer
    body = f"<p style='max-width: 70%;font-size: 13pt;'>Olá!" + \
        "</br>Envio a seguir uma cotação na planilha em anexo." + \
        "</br>Solicito sua resposta assim que possível.</p>" + \
        footer

    # Create the email footer
    footer = ("<div style='border:none;border-bottom:solid windowtext 1.0pt;padding:0cm 0cm 1.0pt 0cm'><p class='MsoNormal' style='border:none;padding:0cm'><span><u></u>&nbsp;<u></u></span></p></div>" +
              "<p></p><p style='max-width: 70%;font-size: 12pt;'>Atenciosamente,<br/>")

    body += f'<br><img src="cid:{image_c_id[1:-1]}">'

    # Add the body, with the footer, to the email
    msg.add_alternative(body, subtype='html')

    return msg


def send_email(filepath, emails):
    """Envia o email ao cliente

    Args:
        filepath (string): path of the actual quotation
        emails (list): list of emails to send the quotation
    """
    
    ################################################### DELETE THIS SECTION ###################################################
    messagebox.showinfo(f"Os emails são {emails}")

    emails = ['ti@comagro.com.br', 'mateusspinola@comagro.com.br']
    ################################################### DELETE THIS SECTION ###################################################

    filepath = f"./../../{filepath}"
    
    # Cria um corpo de email e define o assunto
    msg = EmailMessage()
    msg['Subject'] = f'Cotação Comagro'
    msg['From'] = EMAIL_ADDRESS

    # Cria um texto plain e evita erro no get_playload
    msg.set_content(
        'If you can see this, please consider updating your email client to support HTML emails.')

    # Cria um Content-ID para a imagem
    image_cid = make_msgid(domain='comagro.com.br')

    # Insert the header and body in the email
    msg = mail_bohe(msg, image_cid)

    # Attach the image and the csv to the email
    add_attachment(msg, f'./../../images/signature.png')
    add_attachment(msg, f'./../../cotacoes/{filepath}.csv')

    cloud = 'nuvem@comagro.com.br'

    # Send the email to each recipient
    with smtplib.SMTP_SSL('smtp.hostinger.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASS)
        for email in emails:
            msg['To'] = email
            smtp.send_message(msg)
        msg['To'] = cloud
        smtp.send_message(msg)

    smtp.quit()
