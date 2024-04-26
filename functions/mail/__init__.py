# SDG
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from email.mime.image import MIMEImage
from functions.secrets import senha
from tkinter import messagebox
from email.utils import make_msgid

import os
import mimetypes
import smtplib

EMAIL_ADDRESS = 'noreply-faturas@comagro.com.br'  # Endereço de email do remetente
EMAIL_PASS = senha  # Senha do email do remetente


def add_attachment(msg, filepath):
    """    Lê e adiciona anexo ao email    """
    if not os.path.isfile(filepath):
        # Caso não haja arquivo no caminho especificado
        messagebox.showinfo("Erro!", f"Planilha não encontrada. Entrar em contato com o magnífico TI.")
        return False

    # Determine the content type of the file
    ctype, encoding = mimetypes.guess_type(filepath)

    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'

    # Salva o tipo e subtipo do arquivo
    maintype, subtype = ctype.split('/', 1)

    # Read the file and set the correct MIME type
    with open(filepath, 'rb') as file:
        file_data = file.read()
        if maintype == 'image':
            # For images, use MIMEImage
            attachment = MIMEImage(file_data, _subtype=subtype)
            attachment.add_header('Content-ID', '<{}>'.format(os.path.basename(filepath)))
            attachment.add_header('Content-Disposition', 'inline', filepath=os.path.basename(filepath))
        else:
            # For all other file types, use MIMEApplication
            maintype = 'text'
            subtype = 'csv'
            attachment = MIMEApplication(file_data, _subtype=subtype)
            attachment.add_header(
                'Content-Disposition', 'attachment', filename='cotacao_comagro.csv',
                filepath=os.path.basename(filepath))
        msg.attach(attachment)


def mail_bohe(msg, image_c_id):
    """    Insere o header e o corpo no email    """
    signature = ''

    try:
        signature = f'<img src="cid:{image_c_id[1:-1]}", style="max-width: 500px; text-align: left;">'
    except:
        messagebox.showinfo(
            "Erro!",
            "Assinatura não encontrada. Entrar em contato com o magnífico TI.\nO programa continuará normalmente.")

    # Create the email footer
    footer = ("<div style='border:none;border-bottom:solid windowtext 1.0pt;padding:0cm 0cm 1.0pt 0cm'><p class='MsoNormal' style='border:none;padding:0cm'><span><u></u>&nbsp;<u></u></span></p></div>" +
              "<p></p><p style='max-width: 70%;font-size: 12pt;'>Atenciosamente,<br/>") + signature

    # Create the email body and add the footer
    body = f"<p style='max-width: 70%;font-size: 13pt;'>Olá!" + \
        "</br>Envio a seguir uma cotação na planilha em anexo." + \
        "</br>Solicito sua resposta assim que possível.</p>" + \
        footer

    # body += f'<br><img src="cid:{image_c_id[1:-1]}">'

    # Add the body, with the footer, to the email
    msg.add_alternative(body, subtype='html')

    # Open the signature image and attach it to the email
    with open('images\\signature.png', 'rb') as img:
        # Guess the content type of the image and split into main and subtype
        maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')
        # Attach the image to the email with the specified Content-ID
        msg.get_payload()[1].add_related(img.read(), maintype=maintype, subtype=subtype, cid=image_c_id)

    return msg


def send_email(filepath, emails):
    """Envia o email ao cliente

    Args:
        filepath (string): path of the actual quotation
        emails (list): list of emails to send the quotation
    """
    # Cria um corpo de email e define o assunto
    msg = EmailMessage()
    msg['Subject'] = f'Cotação Comagro'

    # Cria um texto plain e evita erro no get_playload
    msg.set_content('If you can see this, please consider updating your email client to support HTML emails.')

    # Cria um Content-ID para a imagem
    image_cid = make_msgid(domain='comagro.com.br')

    # Insert the email header and body
    msg = mail_bohe(msg, image_cid)

    # Attach the csv to the email
    a = add_attachment(msg, f'{filepath}')

    # If the file was not found, return False
    if a is False:
        return False

    # Send the email to each recipient
    with smtplib.SMTP_SSL('smtp.hostinger.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASS)

        for email in emails:
            sender = EMAIL_ADDRESS
            to = email

            # Converts the email to a legible archive for the smtplib
            raw = f"From: {sender}\r\nTo: {to}\r\n{msg.as_string()}"

            # Send the mail
            try:
                smtp.sendmail(sender, to, raw)
            except Exception as e:
                messagebox.showinfo("Erro!", f"[Erro]: {e.__class__}\nEmail: {email} Lista: {
                                    emails}\n\nO programa continuará a enviar os emails restantes.")
