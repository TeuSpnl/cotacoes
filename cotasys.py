from tkinter import *
from tkinter import messagebox, Toplevel
from tkinter.filedialog import asksaveasfilename
from functions.mail import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer
from reportlab.lib.units import inch
import pandas as pd
import os
import csv
import re

# Headers for the table
xAxis = ["Código", "Nome", "Quantidade"]

# Initial rows count
initial_rows = 1

# Cells will hold the string vars and the entries
cells = {}

# Store Entry widgets to manage them later
entries = []

# Make a new Top Level Element (Window)
root = Tk()
title = "Sistema de Cotação"
root.title(title)
root.configure(bg='white')

# Display the X-axis labels with enumerate
for i, x in enumerate(xAxis):
    label = Label(root, text=x, width=15, background='white')
    label.grid(row=0, column=i + 1, sticky='sn', pady=10)


def refresh_grid_display():
    for index, row_entries in enumerate(entries):
        for col_index, entry in enumerate(row_entries[1:], start=1):
            entry.grid(row=index + 1, column=col_index)  # Re-grid the entry at the correct position


def update_row_labels():
    # First, clear all existing row number labels
    for row in entries:
        row[0].grid_forget()  # This forgets the grid configuration for the label
        row[0].destroy()      # This completely destroys the label widget

    # Now recreate the row number labels with correct numbering
    for index, row in enumerate(entries):
        # Create a new label for the row number with updated index
        label = Label(root, text=str(index + 1), width=5, bg='white')
        label.grid(row=index + 1, column=0)  # Row index in grid starts from 1
        row[0] = label  # Update the first element in the row with the new label
        refresh_grid_display()


def add_row():
    """ Function to add rows
    """

    row_number = len(entries) + 1
    row_entries = []

    if row_number <= 50:
        # Create a label for the row number
        label = Label(root, text=str(row_number), width=5, bg='white')
        label.grid(row=row_number, column=0)
        row_entries.append(label)

        # Create an entry for each column
        for xcoor, x in enumerate(xAxis):
            var = StringVar(root, '')
            e = Entry(root, textvariable=var, width=30, bg='#FFFFF9')
            e.grid(row=row_number, column=xcoor + 1)
            row_entries.append(e)
        entries.append(row_entries)
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


# Initialize rows
for _ in range(initial_rows):
    add_row()


def clear_all():
    for row in entries:
        for entry in row:
            entry.delete(0, END)
            entry.insert(0, "")


def open_pdf_window(filepath):
    new_window = Toplevel(root)
    new_window.title("Salvar como PDF?")
    new_window.geometry("200x125")
    new_window.configure(bg='white')
    Label(new_window, text="Salvar como PDF?").pack(pady=10)

    Button(
        new_window, text="Sim", command=lambda: save_as_pdf(new_window, filepath),
        width=10, bg='#FFFFF9').pack(
        side="left", padx=5, pady=5)
    Button(new_window, text="Não", command=new_window.destroy,
           width=10, bg='#FFFFF9').pack(side="right", padx=5, pady=5)


def finalize():
    new_window = Toplevel(root)
    new_window.title("Revisão")
    new_window.configure(bg='white', padx=10, pady=10)

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

    # Display the headers
    for i, header in enumerate(xAxis):
        Label(new_window, text=header, bg='white', width=15).grid(row=0, column=i+1, padx=2, pady=2)

    # Display data as sheet
    for row_index, row_entries in enumerate(entries, start=1):
        # Row number label
        Label(new_window, text=str(row_index), bg='white', width=5).grid(row=row_index, column=0, padx=2, pady=2)

        for col_index, entry in enumerate(row_entries[1:]):
            if not entry.get().strip():
                entry.delete(0, END)
                entry.insert(0, "N/A")

            Label(new_window, text=entry.get(), bg="#f0f0f0", width=30).grid(
                row=row_index, column=col_index + 1, padx=2, pady=2)

    # Buttons for actions
    btn_frame = Frame(new_window, bg='white')
    btn_frame.grid(row=len(entries) + 1, columnspan=len(xAxis) + 1, pady=(15, 0))

    Button(btn_frame, text="Concluir", command=lambda: export_to_csv(new_window)).pack(side=RIGHT, padx=10)
    Button(btn_frame, text="Voltar", command=new_window.destroy, bg='white').pack(side=LEFT, padx=10)


def get_next_filename():
    """Function to get the next filename for the CSV file

    Returns:
        String: Path of the next CSV file
    """

    directory = "cotacoes"
    if not os.path.exists(directory):
        os.makedirs(directory)
    files = [f for f in os.listdir(directory) if os.path.isfile(
        os.path.join(directory, f)) and f.endswith('.csv')]
    highest_number = 0
    for file in files:
        try:
            number = int(file.split('.')[0])
            if number > highest_number:
                highest_number = number
        except ValueError:
            continue
    return os.path.join(directory, f"{highest_number + 1}.csv")


def save_as_pdf(new_window, csv_path):
    new_window.destroy()

    pdf_path = asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

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
        logo.hAlign = 'CENTER'
        elements.append(logo)
    except:
        pass

    # Add some space after the logo - adjust as necessary
    elements.append(Spacer(1, 0.25 * inch))

    # List to hold the data for the table
    data = []

    # Read the CSV file and append each row to the data list
    with open(csv_path, newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if row:  # Ensure the row is not empty
                # Split the single string in the row by ';' to form a list of fields
                fields = row[0].split(';') if len(row) == 1 else row
                data.append(fields)

    # Create a table with the data
    table = Table(data)

    # Define a style for the table with specific properties for each cell
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), header_color),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Font for the header
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
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

    # Build the PDF
    doc.build(elements)

    if 1 == 1:
        # Save data as PDF (simulation)
        messagebox.showinfo("Salvar como PDF", f"PDF salvo em {pdf_path}")


def export_to_csv(new_window):
    filepath = get_next_filename()

    # Gather data from entries
    data = []
    for row_entries in entries:
        row_data = [entry.get() for entry in row_entries[1:]]
        data.append(row_data)

    # Create a DataFrame and write to CSV
    df = pd.DataFrame(data, columns=xAxis)
    df.to_csv(filepath, index=False, encoding='utf-8-sig', sep=';')  # Save to CSV without the index

    gather_emails_and_send(filepath)
    new_window.destroy()


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


def gather_emails_and_send(filepath):
    """Function to gather all emails and call send_email()

    Args:
        filepath (string): path of the actual quotation
    """

    email = ['compras@comagro.com.br']
    a = send_email(filepath, email)
    if a is False:
        messagebox.showinfo("Erro!", f"Arquivo não encontrado. Entrar em contato com o magnífico TI.")
    else:
        messagebox.showinfo("Sucesso!", f"Arquivo exportado para {filepath} e enviado para email")
    open_pdf_window(filepath)


# Buttons
Button(root, text="Limpar Tudo", command=clear_all).grid(column=1, row=53, columnspan=2, pady=(10, 5))
Button(root, text="Finalizar", command=finalize).grid(column=2, row=53, columnspan=2, pady=(10, 5))
Button(root, text="+", command=add_row, width=5).grid(column=53, row=1, padx=10)

# Run the Mainloop
root.mainloop()
