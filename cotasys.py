from tkinter import *
from tkinter import messagebox, Toplevel, filedialog
from tkinter.filedialog import asksaveasfilename
from functions.mail import *
import pandas as pd
import os
import re

# Define X and Y Axis
xAxis = ["Código", "Nome", "Quantidade"]

# Initial rows count
initial_rows = 3

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

# Function to add rows


def add_row():
    row_number = len(entries) + 1
    row_entries = []

    if row_number < 51:
        # Y-axis label
        label = Label(root, text=str(row_number),
                      width=5, background='white')
        label.grid(row=row_number, column=0)

        for xcoor, x in enumerate(xAxis):
            var = StringVar(root, '')
            e = Entry(root, textvariable=var, width=30, bg='#FFFFF9')
            e.grid(row=row_number, column=xcoor + 1)
            row_entries.append(e)
        entries.append(row_entries)
    else:
        # Cancel adding button
        messagebox.showinfo("Não permitido", f"Limite de 50 itens atingido")


def remove_row():
    if len(entries) > 1:
        last_row = entries.pop()
        for entry in last_row:
            entry.destroy()
    else:
        messagebox.showinfo("Erro!", "Não é possível remover mais linhas.")


# Initialize rows
for _ in range(initial_rows):
    add_row()


def clear_all():
    for var in cells.values():
        var.set('')


def open_pdf_window():
    new_window = Toplevel(root)
    new_window.title("Salvar como PDF?")
    new_window.geometry("200x125")
    new_window.configure(bg='white')
    Label(new_window, text="Salvar como PDF?").pack(pady=10)

    Button(
        new_window, text="Sim", command=lambda: save_as_pdf(new_window),
        width=10, bg='#FFFFF9').pack(
        side="left", padx=5, pady=5)
    Button(new_window, text="Não", command=new_window.destroy,
           width=10, bg='#FFFFF9').pack(side="right", padx=5, pady=5)


def save_as_pdf(new_window):
    new_window.destroy()
    file_path = asksaveasfilename(defaultextension=".pdf", filetypes=[
                                  ("PDF files", "*.pdf")])
    if file_path:
        # Save data as PDF (simulation)
        messagebox.showinfo("Salvar como PDF", f"PDF salvo em {file_path}")


def finalize():
    new_window = Toplevel(root)
    new_window.title("Revisão")
    new_window.configure(bg='white', padx=10, pady=10)

    # Display the headers
    for i, header in enumerate(xAxis):
        Label(new_window, text=header, bg='white',
              width=15).grid(row=0, column=i+1, padx=2, pady=2)

    # Display data as sheet
    for row_index, row_entries in enumerate(entries, start=1):
        # Row number label
        Label(new_window, text=str(row_index), bg='white', width=5).grid(
            row=row_index, column=0, padx=2, pady=2)
        for col_index, entry in enumerate(row_entries):
            value = entry.get()
            Label(new_window, text=value, bg="#f0f0f0", width=30).grid(
                row=row_index, column=col_index + 1, padx=2, pady=2)

    # Buttons for actions
    btn_frame = Frame(new_window, bg='white')
    btn_frame.grid(row=len(entries) + 1,
                   columnspan=len(xAxis) + 1, pady=(15, 0))

    Button(btn_frame, text="Concluir", command=lambda: export_to_csv()).pack(
        side=RIGHT, padx=10)
    Button(btn_frame, text="Voltar", command=new_window.destroy,
           bg='white').pack(side=LEFT, padx=10)


def get_next_filename():
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


def export_to_csv():
    filepath = get_next_filename()

    # Gather data from entries
    data = []
    for row_entries in entries:
        row_data = [entry.get() for entry in row_entries]
        data.append(row_data)

    # Create a DataFrame and write to CSV
    df = pd.DataFrame(data, columns=xAxis)
    df.to_csv(filepath, index=False, encoding='utf-8-sig', sep=';')  # Save to CSV without the index

    email_entry(filepath)

# Function to validate email addresses


def is_valid_email(email):
    if not email.strip():  # Check if the email is not just empty or spaces
        return False
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None


def email_entry(filepath):
    email_window = Toplevel(root)
    email_window.title("Adicionar emails")
    email_window.configure(bg='white')
    label = Label(email_window, text="Adicionar emails", font=("Arial", 14))
    label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

    # List to store all email entry widgets
    email_entries = []

    # Function to add a new email entry field

    def add_email_entry():
        entry = Entry(email_window, width=25)
        entry.grid(row=len(email_entries) + 1, column=0, padx=5, pady=5, sticky="we")
        email_entries.append(entry)
        entry.bind('<KeyRelease>', validate_emails)  # Bind key release event to validate emails
        update_remove_button_state()
        update_layout()

    # Function to remove the last email entry field
    def remove_email_entry():
        if email_entries:
            entry_to_remove = email_entries.pop()
            entry_to_remove.destroy()
            update_remove_button_state()
            update_layout()
            validate_emails(None)  # Validate emails after removal

    def validate_emails(event):
        all_valid = all(is_valid_email(entry.get()) for entry in email_entries if entry.get())
        send_button.config(state='normal' if all_valid else 'disabled')

    # Function to gather all emails and call send_email (placeholder for your actual send_email function)
    def gather_emails_and_send():
        emails = [entry.get() for entry in email_entries if entry.get()]
        a = send_email(filepath, emails)
        if a is False:
            messagebox.showinfo("Erro!", f"Arquivo não encontrado. Entrar em contato com o magnífico TI.")
        else:
            messagebox.showinfo("Sucesso!", f"Arquivo exportado para {filepath} e enviado para email")
        email_window.destroy()
        open_pdf_window()

    # "+" button to add new email entries
    add_button = Button(email_window, text="+", command=add_email_entry)
    add_button.grid(row=1, column=1, padx=5, pady=5)

    # "-" button to remove the last email entry
    remove_button = Button(email_window, text="-",
                           command=remove_email_entry, state='disabled')
    remove_button.grid(row=2, column=1, padx=5, pady=5)

    # "Send" button to gather emails and send
    send_button = Button(email_window, text="Enviar",
                         command=gather_emails_and_send, state='disabled')
    send_button.grid(row=50, column=0, columnspan=2, padx=5, pady=5)

    # Function to update the state of the "-" button
    def update_remove_button_state():
        remove_button.config(state='normal' if len(email_entries) > 1 else 'disabled')

    def update_layout():
        for index, entry in enumerate(email_entries):
            entry.grid(row=index+1, column=0)

    # Initial email entry
    add_email_entry()


# Buttons
Button(root, text="Limpar Tudo", command=clear_all).grid(
    column=1, row=52, columnspan=2, pady=(10, 5))
Button(root, text="Finalizar", command=finalize).grid(
    column=2, row=52, columnspan=2, pady=(10, 5))
Button(root, text="+", command=add_row,
       width=5).grid(column=4, row=0, padx=10)
Button(root, text="-", command=remove_row,
       width=5).grid(column=4, row=1, padx=10)

# Run the Mainloop
root.mainloop()
