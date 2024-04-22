from tkinter import *
from tkinter import messagebox, Toplevel, filedialog
from tkinter.filedialog import asksaveasfilename
import csv
import os

# Define X and Y Axis
xAxis = ["Nome", "Código", "Quantidade"]

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
    if entries:
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

    Button(new_window, text="Sim", command=lambda: save_as_pdf(
        new_window), width=10, bg='#FFFFF9').pack(side="left", padx=5, pady=5)
    Button(new_window, text="Não", command=new_window.destroy,
           width=10, bg='#FFFFF9').pack(side="right", padx=5, pady=5)


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


def save_as_pdf(new_window):
    new_window.destroy()
    file_path = asksaveasfilename(defaultextension=".pdf", filetypes=[
                                  ("PDF files", "*.pdf")])
    if file_path:
        # Save data as PDF (simulation)
        messagebox.showinfo("Salvar como PDF", f"PDF salvo em {file_path}")


def get_next_filename():
    directory = "cotacoes"
    if not os.path.exists(directory):
        os.makedirs(directory)
    files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f)) and f.endswith('.csv')]
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

    # Write data to CSV
    with open(filepath, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(xAxis)  # Write header
        writer.writerows(data)  # Write data rows
    messagebox.showinfo("Sucesso!",
                        f"Arquivo exportado para {filepath} e enviado para email")

    open_pdf_window()


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
