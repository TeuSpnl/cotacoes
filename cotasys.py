from tkinter import *
from tkinter import ttk
from tkinter import messagebox, Toplevel
from tkinter.filedialog import asksaveasfilename
from functions.mail import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer, Paragraph
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
import os
import csv
import re

# Headers for the table
xAxis = ["Código", "Nome", "Quantidade"]

# Cells will hold the string vars and the entries
cells = {}

# Store Entry widgets to manage them later
entries = []


def setup_top_fields():
    # Create a frame to hold the new fields at the top of the window
    top_frame = Frame(root, bg='white')
    top_frame.grid(row=0, column=0, columnspan=4, sticky='ew', padx=10, pady=(15, 5))

    # Máquinas entry field
    Label(top_frame, text="Máquinas:", bg='white').pack(side='left', padx=10)
    maquinas_entry = Entry(top_frame, width=20, bg='#FFFFF9')
    maquinas_entry.pack(side='left')
    maquinas_entry.bind('<Return>', focus_next_entry)  # Bind the Return key to focus the next entry

    # Listbox for selecting names
    names_combobox = ttk.Combobox(
        top_frame, values=["Josuilton", "Jucilande", "Palmiro"],
        state="readonly", width=15, background='#FFFFF9')
    names_combobox.pack(side='right', padx=10)
    names_combobox.bind('<Return>', focus_one)  # Bind the Return key to focus the next entry
    Label(top_frame, text="Usuário:", bg='white').pack(side='right', padx=10)

    maquinas_entry.focus()
    return maquinas_entry, names_combobox


def validate_inputs():
    if not maquinas_entry.get().strip():
        messagebox.showerror("Erro!", "Por favor, insira o nome das máquinas.")
        return False
    if not user.get():
        messagebox.showerror("Erro!", "Por favor, selecione o usuário.")
        return False
    return True


def validate_positive_integer(P):
    # Check if the string is a digit and convert to integer to check if it's greater than 0
    if P.isdigit() and int(P) > 0:
        return int(P)
    elif P == "":  # Allow empty field to be able to clear the entry or during typing
        return True
    return False


def refresh_grid_display():
    """ Function to refresh the grid display after adding or removing rows
    """
    for index, row_entries in enumerate(entries):
        for col_index, entry in enumerate(row_entries[1:], start=1):
            entry.grid(row=index + 2, column=col_index)  # Re-grid the entry at the correct position


def update_row_labels():
    """ Function to update the row number labels after adding or removing rows
    """

    # First, clear all existing row number labels
    for row in entries:
        row[0].grid_forget()  # This forgets the grid configuration for the label
        row[0].destroy()      # This completely destroys the label widget

    # Now recreate the row number labels with correct numbering
    for index, row in enumerate(entries):
        # Create a new label for the row number with updated index
        label = Label(root, text=str(index + 1), width=5, bg='white')
        label.grid(row=index + 2, column=0)  # Row index in grid starts from 1
        row[0] = label  # Update the first element in the row with the new label
        refresh_grid_display()


def add_row(event=None):
    """ Function to add rows
    """

    row_number = len(entries) + 1
    row_entries = []

    if row_number <= 50:
        # Create a label for the row number
        label = Label(root, text=str(row_number), width=5, bg='white')
        label.grid(row=row_number + 1, column=0)
        row_entries.append(label)

        # Register the validation funcion to check if the input is a positive integer
        validate_int = root.register(validate_positive_integer)

        # Create an entry for each column
        for xcoor, x in enumerate(xAxis):
            var = StringVar(root, '')
            e = Entry(root, textvariable=var, width=30, bg='#FFFFF9')

            if x == 'Quantidade':
                # Apply validation only to 'Quantidade' column
                # The validate="key" option causes the validatecommand to be triggered on every keystroke, preventing the user from even entering invalid characters, such as non-digits or zero
                e.config(validate='key', validatecommand=(validate_int, '%P'))

            e.grid(row=row_number + 1, column=xcoor + 1)
            e.bind('<Return>', focus_next_entry)  # Bind the Return key to focus the next entry
            row_entries.append(e)

        row_entries[-1].bind('<Return>', add_row)  # Bind the Return key to add a new row

        entries.append(row_entries)
        row_entries[1].focus()
    else:
        messagebox.showinfo("Não permitido", "Limite de 50 itens atingido")


def remove_row():
    """ Function to remove rows
    """
    if len(entries) > 1:
        last_row = entries.pop()
        for entry in last_row:
            entry.destroy()
        update_row_labels()
    else:
        messagebox.showinfo("Erro!", "Não é possível remover mais linhas.")


def clear_all():
    """ Function to clear all entries in the grid
    """
    user.set('')
    maquinas_entry.delete(0, END)

    for row in entries[1:]:
        for entry in row:
            entry.grid_forget()
            entry.destroy()

    entries.clear()
    update_row_labels()
    add_row()

    maquinas_entry.focus()


def focus_next_entry(event):
    """Move focus to the next widget or add a new row if it's the last widget in the row."""
    widget = event.widget
    if widget.tk_focusNext() != widget:
        widget.tk_focusNext().focus()
    else:
        add_row()
    return "break"  # Prevent the default 'Return' behavior


def focus_one(event):
    """Move focus to the first widget."""
    entries[0][1].focus()
    return "break"  # Prevent the default 'Return' behavior


def finalize():
    """ Function to finalize the quotation and display the data in a new window
    """
    if not validate_inputs():
        return

    new_window = Toplevel(root)
    new_window.title("Revisão")
    new_window.configure(bg='white', padx=10, pady=10)
    new_window.grab_set()

    # Filter out completely empty rows and update entries list
    temp_entries = []  # Temporary list to hold non-empty entries
    for row in entries:
        if any(e.get().strip() for e in row[1:]):  # Check if there's any non-empty cell
            temp_entries.append(row)
        else:
            for e in row:  # Remove widgets of empty rows from grid
                e.grid_forget()
                e.destroy()

    entries[:] = temp_entries  # Update the main entries list
    update_row_labels()

    if not entries:
        new_window.destroy()
        messagebox.showinfo("Erro!", "Planilha vazia. Por favor, adicione dados.")

        if len(entries) == 0:
            add_row()

        return

    # Display the machine and user fields
    Label(new_window, text=f"Máquinas: {maquinas_entry.get().replace(';', ', ')}",
          background='#FFFFF9').grid(row=0, column=1, columnspan=2, pady=5)
    Label(new_window, text=f"Usuário: {user.get()}", background='#FFFFF9').grid(row=0, column=2, columnspan=2, pady=5)

    # Display the headers
    for i, header in enumerate(xAxis):
        Label(new_window, text=header, bg='white', width=15).grid(row=1, column=i+1, padx=2, pady=2)

    # Display data as sheet
    for row_index, row_entries in enumerate(entries, start=1):
        # Row number label
        Label(new_window, text=str(row_index), bg='white', width=5).grid(row=row_index + 1, column=0, padx=2, pady=2)

        for col_index, entry in enumerate(row_entries[1:]):
            if not entry.get().strip():
                new_window.destroy()
                messagebox.showinfo("Erro!", "Por favor, preencha todos os campos.")
                return

            Label(new_window, text=entry.get(), bg="#f0f0f0", width=30).grid(
                row=row_index + 1, column=col_index + 1, padx=2, pady=2)

    # Buttons for actions
    btn_frame = Frame(new_window, bg='white')
    btn_frame.grid(row=len(entries) + 2, column=2, pady=(15, 0))

    Button(btn_frame, text="Voltar", command=new_window.destroy).pack(side=LEFT, padx=10)
    # Button(btn_frame, text="Salvar como PDF", command=lambda: guide_save_pdf(new_window)).pack(side=LEFT, padx=10)
    Button(btn_frame, text="Concluir", command=lambda: guide_finalize(new_window)).pack(side=RIGHT, padx=10)


def guide_save_pdf(new_window):
    """Function to guide the user to save the file as PDF
    """
    new_window.configure()
    filepath = export_to_csv(new_window)
    get_pdf_path(new_window, filepath)


def guide_finalize(new_window):
    filepath = export_to_csv(new_window, TRUE)
    gather_emails_and_send(filepath, save_as_pdf(filepath))
    clear_all()


def get_next_filename(type='csv'):
    """Function to get the next filename for the CSV or PDF file

    Returns:
        String: Path of the next file
    """

    directory = "\\\\Servidor\\Users\\Pichau\\Documents\\Drive Comagro\\Cotacoes"
    if not os.path.exists(directory):
        os.makedirs(directory)

    if user.get() == 'Josuilton':
        nuser = 1
    elif user.get() == 'Jucilande':
        nuser = 2
    elif user.get() == 'Palmiro':
        nuser = 3

    files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f)) and f.endswith('.csv')]

    highest_number = 0
    for file in files:
        try:
            n = file.split('.')[0]
            number = int(n[-5:])
            if number >= highest_number:
                highest_number = number
        except ValueError:
            continue

    if type == 'pdf':
        return os.path.join(directory, f"{nuser:02d}{highest_number:05d}.pdf")
    else:
        return os.path.join(directory, f"{nuser:02d}{highest_number + 1:05d}.csv")


def export_to_csv(new_window, i=FALSE):
    """ Function to export the data to a CSV file

    Args:
        new_window (Tk Window): The window to be destroyed
    """
    if i:
        new_window.destroy()

    filepath = get_next_filename()

    # Data list to hold the data for the CSV file
    data = []

    # Gather data from entries and include in the data list
    for row_entries in entries:
        row_data = [entry.get() for entry in row_entries[1:]]
        data.append(row_data)
    
    # Include an empty row for spacing
    data.append('')
    
    # Include quotation number
    nquot = filepath.split('\\')[-1].split('.')[0]
    data.append(['Cotação: ', f'#{nquot}'])
    
    # Include an empty row for spacing
    data.append('')

    # Include Machines field
    data.append(['Máquinas'])
    data.append(maquinas_entry.get().split(';'))

    # Include an empty row for spacing
    data.append('')

    # Include User field
    data.append(['Solicitante'])
    data.append([user.get()])

    # Find the maximum length of any row in the data
    max_columns = max(len(row) for row in data)

    # Create a model row with the maximum number of columns
    xAxis.extend(['' for _ in range(max_columns - 3)])

    # Create a DataFrame and write to CSV
    df = pd.DataFrame(data, columns=xAxis)
    df.to_csv(filepath, index=False, encoding='utf-8-sig', sep=';')  # Save to CSV without the index

    return filepath


def open_pdf_window(filepath):
    """Function to open a new window to ask if the user wants to save the file as PDF

    Args:
        filepath (String): Path of the CSV file to be saved as PDF
    """

    new_window = Toplevel(root)
    new_window.title("Salvar como PDF?")
    new_window.geometry("200x125")
    new_window.configure(bg='white')
    new_window.grab_set()
    Label(new_window, text="Salvar como PDF?").pack(pady=10)

    Button(
        new_window, text="Sim", command=lambda: get_pdf_path(new_window, filepath, TRUE),
        width=10, bg='#FFFFF9').pack(
        side="left", padx=5, pady=5)
    Button(new_window, text="Não", command=new_window.destroy,
           width=10, bg='#FFFFF9').pack(side="right", padx=5, pady=5)


def get_pdf_path(new_window, csv_path, i=FALSE):
    """Function to get the path where the PDF file should be saved

    Args:
        new_window (TK Window): The window to be destroyed
        csv_path (String): Path of the CSV file to be saved as PDF
        i (boolean): [Optional] True if the window should be destroyed

    Returns:
        String: Path of the PDF file
    """
    if i:
        new_window.destroy()

    pdf_path = asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

    save_as_pdf(csv_path, pdf_path)


def save_as_pdf(csv_path, pdf_path=''):
    """ Function to save the CSV file as PDF

    Args:
        csv_path (String): Path of the CSV file to be saved as PDF
        pdf_path (String): [Optional] Path of the PDF file
    """

    nquot = ''

    if pdf_path.strip() == '':
        pdf_path = get_next_filename('pdf')
        nquot = pdf_path.split('\\')[-1].split('.')[0]
        

    styles = getSampleStyleSheet()

    # Create a PDF document with a specific filename and page size
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)

    # Convert hexadecimal colors to Color objects
    header_color = colors.HexColor('#585859')
    row_color = colors.HexColor('#F2AC29')

    # Create a list to hold the PDF elements
    elements = []

    # Add the logo at the top
    try:
        logo = Image('images\\logo.png')
        logo.drawHeight = 1.25 * inch * logo.drawHeight / logo.drawWidth
        logo.drawWidth = 1.25 * inch
        logo.hAlign = 'LEFT'
        elements.append(logo)
    except:
        pass

    # Add some space
    elements.append(Spacer(1, 0.25 * inch))
    
    if nquot != '':
        # Add quotation number
        elements.append(Paragraph(f"Cotação: <b>#{nquot}</b>", styles['Normal']))
    
    # Add some space
    elements.append(Spacer(1, 0.25 * inch))

    # Add a context to the PDF
    elements.append(Paragraph(f"Favor, cotar os itens abaixo referentes à(s) máquina(s) {
                    maquinas_entry.get()}.", styles['Normal']))
    elements.append(Paragraph("Frete e outras observações a combinar.", styles['Normal']))

    # Add some space
    elements.append(Spacer(2, 0.25 * inch))

    # List to hold the data for the table
    data = []

    # Read the CSV file and append each row to the data list
    with open(csv_path, newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.reader(csvfile)
        keep = True
        while keep:
            for row in reader:
                if row[0].replace(';', '').strip() == '':
                    keep = False
                    break
                if row:  # Ensure the row is not empty
                    # Split the single string in the row by ';' to form a list of fields
                    fields = row[0].split(';') if len(row) == 1 else row
                    fields = [f.upper() for f in fields if f != '']  # Remove empty strings from the row

                    data.append(fields)

    # Create a table with the data
    table = Table(data)

    # Define a style for the table with specific properties for each cell
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), header_color),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Font for the header
        ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
        ('BACKGROUND', (0, 1), (-1, -1), row_color),  # Background for other rows
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Grid lines for all cells
        ('BOX', (0, 0), (-1, -1), 1, colors.black),  # Border around the table
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Font for the rest of the table
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 5)  # Padding in each cell
    ])
    table.setStyle(style)

    # Append the table to the elements
    elements.append(table)

    # Add some space
    elements.append(Spacer(4, 0.25 * inch))

    # Add a signature and closing message
    elements.append(Paragraph("Atenciosamente,", styles['Normal']))

    try:
        signature = Image('images\\Jhonatan.png')
        signature.drawHeight = 4 * inch * signature.drawHeight / signature.drawWidth
        signature.drawWidth = 4 * inch
        signature.hAlign = 'LEFT'
        elements.append(signature)
    except:
        pass

    # Build the PDF
    doc.build(elements)

    if pdf_path:
        # Inform the user that the PDF has been saved
        messagebox.showinfo("PDF salvo", f"PDF salvo em {pdf_path}")

    return pdf_path


def is_valid_email(email):
    """Function to validate email addresses

    Args:
        email (list): list of email addresses

    Returns:
        Boolean: True if the email is valid, False otherwise
    """

    if not email.strip():  # Check if the email is not just empty or spaces
        return False
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None


def gather_emails_and_send(filepath, pdf_path):
    """Function to gather all emails and call send_email()

    Args:
        filepath (string): path of the actual quotation
        pdf_path (string): path of the PDF file
    """

    email = ['compras@comagro.com.br']
    a = send_email(filepath, pdf_path, email, user.get())
    if a is False:
        messagebox.showinfo("Erro!", f"Arquivo não encontrado. Entrar em contato com o magnífico TI.")
    else:
        messagebox.showinfo("Sucesso!", f"Arquivo exportado para {filepath} e enviado para email")

    open_pdf_window(filepath)
    os.remove(pdf_path)


# Make a new Top Level Element (Window)
root = Tk()
title = "Cotasys - Sistema de Cotação de Preços"
root.title(title)
root.configure(bg='white')

# Display the X-axis labels with enumerate
for i, x in enumerate(xAxis):
    label = Label(root, text=x, width=15, background='white')
    label.grid(row=1, column=i + 1, sticky='sn', pady=10)


# Initial rows count
initial_rows = 1

# Initialize rows
for _ in range(initial_rows):
    add_row()


# Initialize user and macchine fields and store references to the widgets
maquinas_entry, user = setup_top_fields()


# Buttons
Button(root, text="Limpar Tudo", command=clear_all).grid(column=1, row=53, columnspan=2, pady=(15, 5))
Button(root, text="Finalizar", command=finalize).grid(column=2, row=53, columnspan=2, pady=(15, 5))
Button(root, text="+", command=add_row, width=5).grid(column=53, row=1, padx=10)

# Run the Mainloop
root.mainloop()
