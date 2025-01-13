import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

DEFAULT_FILE = "date.ods"

# Global variables
file_url = DEFAULT_FILE
df = None

def load_file(file):
    """Load the file into the DataFrame."""
    global df
    try:
        if file.endswith(".ods"):
            df = pd.read_excel(file, engine="odf", sheet_name="date")
        elif file.endswith(".xlsx"):
            df = pd.read_excel(file, engine="openpyxl", sheet_name="date")
        else:
            messagebox.showerror("Errore", "Formato di file non supportato")
            return False

        materials = df['Material'].dropna().unique().tolist()
        combo_material['values'] = materials
        combo_material.set("")
        return True
    except Exception as e:
        messagebox.showerror("Errore", f"Impossibile caricare il file: {e}")
        return False

def upload_file():
    """Allow the user to select and load a file."""
    global file_url
    file = filedialog.askopenfilename(filetypes=[("File ODS", "*.ods"), ("File Excel", "*.xlsx")])
    if file:
        file_url = file
        entry_file.set(file)
        if not load_file(file):
            messagebox.showerror("Errore", "Impossibile caricare il file selezionato")

def update_thickness(event):
    """Update thicknesses based on selected material."""
    if df is None:
        return
    
    selected_material = combo_material.get()
    if selected_material:
        filtered_thicknesses = df[df['Material'] == selected_material]['Espesor'].dropna().unique().tolist()
        combo_thickness['values'] = filtered_thicknesses
        combo_thickness.set('')

def date():
    try:
        global df
        material = combo_material.get()
        thickness = combo_thickness.get()
        
        if thickness:
            thickness = float(thickness)
        else:
            thickness = None

        if material and thickness is not None:
            filtered_df = df[(df['Material'] == material) & (df['Espesor'] == thickness)]
        else:
            if material:
                filtered_df = df[df['Material'] == material]
            else:
                filtered_df = df

        if filtered_df.empty:
            messagebox.showwarning("Nessun risultato", "Non sono stati trovati dati con i criteri selezionati.")
            return None  
        else:
            # Convert the filtered DataFrame to a dictionary
            filtered_dict = filtered_df.to_dict(orient='records')

            # Optionally, you can process `filtered_dict` here if needed
            print(filtered_dict)  # For debugging or checking the result

            return filtered_dict

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")
        return None


# Main window
window = tk.Tk()
window.title("App di Calcoli")
window.geometry("400x300")
window.config(bg="#f4f4f4")
font = ('Helvetica', 12)

# Primary frame
frame_primary = tk.Frame(window, bg="#f4f4f4", padx=10, pady=10)
frame_primary.pack(fill='x', padx=10, pady=5)

entry_file = tk.StringVar(value=file_url)  
tk.Label(frame_primary, text="File Excel:", font=font, bg="#f4f4f4").grid(row=0, column=0, sticky="w", padx=5, pady=2)
tk.Entry(frame_primary, textvariable=entry_file, width=30, state='readonly', font=font).grid(row=0, column=1, padx=5, pady=2)
tk.Button(frame_primary, text="Carica un altro file", command=upload_file, font=font, relief="raised", padx=5, pady=5).grid(row=1, column=0, columnspan=2, pady=5)

# Secondary frames
frame_secondary = tk.Frame(window, bg="#f4f4f4", padx=10, pady=10)
frame_secondary.pack(fill='x', padx=10, pady=5)

# Material section
material_frame = tk.Frame(frame_secondary, bg="#f4f4f4", padx=5, pady=5)
material_frame.grid(row=0, column=0, padx=10)

tk.Label(material_frame, text="Seleziona Materiale:", font=font, bg="#f4f4f4").pack(anchor="w", pady=2)
combo_material = ttk.Combobox(material_frame, width=20, font=font)
combo_material.pack(pady=2)
combo_material.bind("<<ComboboxSelected>>", update_thickness)

# Thickness section
thickness_frame = tk.Frame(frame_secondary, bg="#f4f4f4", padx=5, pady=5)
thickness_frame.grid(row=0, column=1, padx=10)

tk.Label(thickness_frame, text="Seleziona Spessore:", font=font, bg="#f4f4f4").pack(anchor="w", pady=2)
combo_thickness = ttk.Combobox(thickness_frame, width=20, font=font)
combo_thickness.pack(pady=2)

# Button to trigger the calculation
tk.Button(window, text="Calcolare", command=date, font=font, relief="raised", padx=10, pady=5).pack(pady=10)

# Load default file if it exists
if os.path.exists(DEFAULT_FILE):
    if not load_file(DEFAULT_FILE):
        messagebox.showwarning("Attenzione", "Il file predefinito non può essere caricato. Seleziona un altro file.")
else:
    messagebox.showinfo("Informazione", "Il file predefinito non esiste. Seleziona un file.")

# Show window
window.mainloop()
