import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

# Ruta del archivo predeterminado
ARCHIVO = "date.ods"  # O "date.xlsx" si usas archivo Excel

# Variable global para almacenar la ruta del archivo
archivo_cargado = ARCHIVO

# Variable global para almacenar los datos filtrados como diccionario
datos_filtrados_dict = {}

def cargar_archivo():
    global archivo_cargado  # Usamos la variable global para almacenar la ruta
    archivo = filedialog.askopenfilename(filetypes=[("Archivos ODS", "*.ods"), ("Archivos Excel", "*.xlsx")])
    if archivo:
        archivo_cargado = archivo  # Si se selecciona un archivo, actualizamos la variable global
        entrada_archivo.set(archivo_cargado)  # Actualizamos la ruta mostrada en la interfaz

def actualizar_espesores(event):
    material_usuario = combo_material.get()
    
    if material_usuario:
        # Filtrar los espesores según el material
        espesores_filtrados = df[df['Material'] == material_usuario]['Espesor'].dropna().unique().tolist()
        # Actualizar los valores del combobox de espesor
        combo_espesor['values'] = espesores_filtrados
        combo_espesor.set('')  # Limpiar la selección de espesor

def recolectar_datos():
    try:
        global datos_filtrados_dict
        archivo = archivo_cargado  
        material_usuario = combo_material.get() 
        espesor_usuario = combo_espesor.get()  
        
        # Verificar si el espesor está vacío
        if espesor_usuario:
            espesor_usuario = float(espesor_usuario)  
        else:
            espesor_usuario = None 

        # Leer el archivo ODS (si es un archivo LibreOffice) o Excel
        if archivo.endswith(".ods"):
            df = pd.read_excel(archivo, engine="odf", sheet_name="date")  
        else:
            df = pd.read_excel(archivo, sheet_name="date")  

        # Filtrar los datos por material y espesor si es necesario
        if material_usuario and espesor_usuario is not None:
            df_filtrado = df[(df['Material'] == material_usuario) & (df['Espesor'] == espesor_usuario)]
        else:
            df_filtrado = df  

        # Verificar si se encontraron resultados
        if df_filtrado.empty:
            messagebox.showwarning("Nessun risultato", "Non sono stati trovati dati con i criteri selezionati.") 
            return None  
        else:
            # Convertir el dataframe filtrado a un diccionario de Python sin columnas
            datos_filtrados_dict = df_filtrado.to_dict(orient='records')
            return datos_filtrados_dict

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")
        return None

def calcular_tiempo_corte():
    datos_filtrados_dict = recolectar_datos()

    if not datos_filtrados_dict:
        return None  

    try:
        # Obtener el perímetro ingresado y convertirlo a float
        perimetro_usuario = entrada_perimetro.get()
        cantidad_aujeros = entrada_aujeros.get()
        
        # Verificar que el usuario ingresó un perímetro
        if not perimetro_usuario:
            messagebox.showerror("Errore", "Per favore, inserisci un perimetro.")
            return None
        
        # Verificar que el usuario ingresó la cantidad de agujeros
        if not cantidad_aujeros:
            messagebox.showerror("Errore", "Per favore, inserisci la quantità di fori.")
            return None
        
        # Convertir valores a float
        perimetro_usuario = float(perimetro_usuario)  # mm
        cantidad_aujeros = float(cantidad_aujeros)    


        
        # Recuperar el valor de 'CW ' y los tiempos desde el diccionario
        cw = datos_filtrados_dict[0]['CW ']  # mm * h
        tiempo_1 = datos_filtrados_dict[0][1]  # h 
        tiempo_2 = datos_filtrados_dict[0][2]  # h

        # Verificar que 'cw' no sea cero para evitar la división por cero
        if cw == 0:
            messagebox.showerror("Errore", "Il valore di CW non può essere zero.")
            return None

        # Verificar que los tiempos no sean nulos
        if tiempo_1 is None or tiempo_2 is None:
            messagebox.showerror("Errore", "I tempi non sono validi.")
            return None
        
        tiempo_corte = perimetro_usuario / cw  # Tiempo de corte en horas
        tiempo_aujeros = cantidad_aujeros * (tiempo_1 + tiempo_2)  
        tiempo_total = tiempo_corte + tiempo_aujeros 

        # Convertir el tiempo total a minutos
        tiempo_total_min = tiempo_total * 60

        # Crear el diccionario con los resultados
        tiempo_corte_dict = {
            "tiempo_corte_horas": tiempo_corte,
            "tiempo_aujeros_horas": tiempo_aujeros,
            "tiempo_total_horas": tiempo_total,
            "tiempo_total_minutos": tiempo_total_min
        }

        # Retornar el diccionario de resultados
        return tiempo_corte_dict

    except ValueError:
        # Si no se pudo convertir el perímetro o la cantidad de agujeros a números, mostrar un error
        messagebox.showerror("Errore", "Il valore del perimetro o della quantità di fori deve essere un numero valido.")
        return None
    except Exception as e:
        # Capturar cualquier otro error y mostrarlo
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")
        return None

def calcular_consumo_gas():
    try:
        # Primero recolectamos los datos
        datos_filtrados_dict = recolectar_datos()
        if not datos_filtrados_dict:
            return None 

        # Calculamos tiempos
        tiempo_corte_dict = calcular_tiempo_corte()
        if not tiempo_corte_dict:
            return None  

        # Obtener el contenido neto del pack ingresado
        neto_pack_usuario = neto_pack.get()
        
        # Verificar que el usuario ingresó el contenido neto del pack
        if not neto_pack_usuario:
            messagebox.showerror("Errore", "Per favore, inserisci il contenuto netto del pack.")  
            return None
        
        # Convertir el contenido neto a float
        try:
            neto_pack_usuario = float(neto_pack_usuario)
        except ValueError:
            messagebox.showerror("Errore", "Il contenuto netto deve essere un numero valido.")
            return None

        # Obtener el tiempo total en horas y la duración del pack
        tiempo_corte_hora = tiempo_corte_dict['tiempo_total_horas']
        duracion_pack = datos_filtrados_dict[0]['Duracion'] 
        
        # Verificar si la duración está vacía o no existe
        if duracion_pack is None or duracion_pack == "":
            messagebox.showerror("Errore", "Il campo 'Durazione' è vuoto o mancante. Per favore, inserisci la durata del pack.")
            return None
        
        # Verificar que la duración no sea cero
        if duracion_pack == 0:
            messagebox.showerror("Errore", "La durata del pack non può essere zero.")
            return None
        
        # Calcular el consumo
        consumo = (tiempo_corte_hora * neto_pack_usuario) / duracion_pack
        
        return consumo  # m3

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")
        return None

def calcular_costos_general():
    # Primero recolectamos los datos
    datos_filtrados_dict = recolectar_datos()
    
    if not datos_filtrados_dict:
        return None 
    
    # Calculamos los tiempos de corte
    tiempo_corte_dict = calcular_tiempo_corte()
    if not tiempo_corte_dict:
        return None  

    # Calculamos el consumo de gas
    cantidad_gas_dict = calcular_consumo_gas()
    if cantidad_gas_dict is None:
        return None 

    # Obtenemos los valores de entrada del usuario
    costo_pack_usuario = entrada_costo_pack.get()
    costo_maquina_usuario = entrada_maquina.get()
    ancho_usuario = entrada_ancho.get()
    largo_usuario = entrada_largo.get()
    costo_operario_usuario = entrada_costo_hora_operarios.get()

    # Verificar que los valores de los campos sean válidos
    if not costo_pack_usuario:
        messagebox.showerror("Errore", "Per favore, inserisci il costo del pack.")  
        return None
    
    if not costo_maquina_usuario:
        messagebox.showerror("Errore", "Per favore, inserisci il costo della macchina.")  
        return None
    
    if not ancho_usuario:
        messagebox.showerror("Errore", "Per favore, inserisci l'ancho.")  
        return None

    if not costo_operario_usuario:
        messagebox.showerror("Errore", "Per favore, inserisci l'costo ore operari.")  
        return None
    
    if not largo_usuario:
        messagebox.showerror("Errore", "Per favore, inserisci il largo.")  
        return None

    # Convertir las entradas de usuario a valores numéricos válidos
    try:
        costo_pack_usuario = float(costo_pack_usuario)
        costo_maquina_usuario = float(costo_maquina_usuario)
        ancho_usuario = float(ancho_usuario)
        largo_usuario = float(largo_usuario)
        costo_operario_usuario = float(costo_operario_usuario)
    except ValueError:
        messagebox.showerror("Errore", "Per favore, inserisci valori numerici validi.")
        return None

    # Cálculos
    costo_kilogramo = datos_filtrados_dict[0]['Costo']    
    costo_gas = cantidad_gas_dict * costo_pack_usuario
    tiempo_corte_horas = tiempo_corte_dict.get("tiempo_total_horas", 0) 
    costo_maquina = tiempo_corte_horas * costo_maquina_usuario
    area_usuario = (ancho_usuario * largo_usuario)  
    area_usuario_m2 = area_usuario / 1_000_000  # Convertir de mm² a m²
    costo_peso = area_usuario_m2 * costo_kilogramo
    
    total = costo_gas + costo_maquina + costo_peso + costo_operario_usuario

    # Crear un diccionario con los resultados
    costos_dict = {
        "costo_gas": costo_gas,
        "costo_maquina": costo_maquina,
        "costo_peso": costo_peso,
        "total": total,
        "cantidad_gas_dict": cantidad_gas_dict,
        "tiempo_corte_horas": tiempo_corte_horas,
        "Costo_operario": costo_operario_usuario,
    }

    # Mostrar los resultados
    print(costos_dict)

    return costos_dict

def mostrar_resultados():
    # Llamar a calcular_costos_general para obtener los datos
    costos_dict = calcular_costos_general()
    
    if costos_dict is None:
        return  # Si no se generaron datos, salir
    
    # Limpiar la tabla antes de agregar nuevos resultados
    for row in treeview.get_children():
        treeview.delete(row)
    
    # Insertar los datos en la tabla
    treeview.insert("", "end", values=(
        f"{costos_dict['costo_gas']} EUR",
        f"{costos_dict['costo_maquina']} EUR",
        f"{costos_dict['costo_peso']} EUR",
        f"{costos_dict['total']} EUR",
        f"{costos_dict['cantidad_gas_dict']} m³",
        f"{costos_dict['tiempo_corte_horas']} h",
        f"{costos_dict['Costo_operario']} EUR/h"
    ))
 
def generar_informe():
    try:
        # Llamar a calcular_costos_general para obtener los datos
        costos_dict = calcular_costos_general()
        
        if costos_dict is None:
            return  # Si los datos no fueron recolectados o no son válidos, salir
        
        # Obtener el nombre del informe
        nombre_informe = entrada_nombre_informe.get()  # Obtener el nombre del informe
        
        if not nombre_informe:
            messagebox.showerror("Errore", "Per favore, inserisci un nome per il rapporto.")  
            return
        
        # Guardar los datos recolectados en un archivo ODS
        df_informe = pd.DataFrame([costos_dict])  # Convertir el diccionario a DataFrame (como una fila)
        df_informe.to_excel(f"{nombre_informe}.ods", index=False, engine="odf")  # Guardar como .ods
        
        messagebox.showinfo("Rapporto Generato", f"Rapporto generato con successo come '{nombre_informe}.ods'.")  # Mensaje de éxito

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}") 


# Configurazione finestra principale
ventana = tk.Tk()
ventana.title("App di Calcoli")
ventana.geometry("600x1000")  # Dimensione fissa
ventana.resizable(False, False)  # Disabilitare il ridimensionamento
ventana.config(bg="#f4f4f4")  # Colore di sfondo

# Font per i widget
fuente = ('Helvetica', 12)

# Frame principale
frame_principal = tk.Frame(ventana, bg="#f4f4f4", padx=20, pady=20)
frame_principal.pack(fill='none', expand=False)

# Entrata del file
entrada_archivo = tk.StringVar(value=archivo_cargado)  # Inizializza con il file predefinito
tk.Label(frame_principal, text="File Excel:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
tk.Entry(frame_principal, textvariable=entrada_archivo, width=20, state='readonly', font=fuente).pack(pady=5, fill="none")

# Bottone per caricare un altro file
tk.Button(frame_principal, text="Carica un altro file", command=cargar_archivo, font=fuente, bg="#4CAF50", fg="white", relief="raised", padx=10, pady=5).pack(pady=10, fill="none")

# Leggere i dati dal file caricato
df = pd.read_excel(archivo_cargado, engine="odf", sheet_name="date")

# Ottenere l'elenco dei materiali unici
materiales = df['Material'].dropna().unique().tolist()

# Titolo della sezione "Dati della Tabella Excel"
tk.Label(frame_principal, text="Dati della Tabella Excel", font=('Helvetica', 14, 'bold'), bg="#f4f4f4").pack(anchor="w", pady=10, fill="none")

# Selezione Materiale e Spessore
frame_material = tk.Frame(frame_principal, bg="#f4f4f4")
frame_material.pack(fill="none", pady=5)

frame_comboboxes = tk.Frame(frame_material, bg="#f4f4f4")
frame_comboboxes.pack(fill="none", pady=5)

tk.Label(frame_comboboxes, text="Seleziona Materiale:", font=fuente, bg="#f4f4f4").pack(anchor="w", pady=5, fill="none")
combo_material = ttk.Combobox(frame_comboboxes, values=materiales, width=20, font=fuente)
combo_material.pack(pady=5, fill="none")

tk.Label(frame_comboboxes, text="Seleziona Spessore:", font=fuente, bg="#f4f4f4").pack(anchor="w", pady=5, fill="none")
combo_espesor = ttk.Combobox(frame_comboboxes, width=20, font=fuente)
combo_espesor.pack(pady=5, fill="none")

# Dati della Pièce
tk.Label(frame_principal, text="Dati della Pièce", font=('Helvetica', 14, 'bold'), bg="#f4f4f4").pack(anchor="w", pady=10, fill="none")

# Campi per la pièce
frame_pieza = tk.Frame(frame_principal, bg="#f4f4f4")
frame_pieza.pack(fill="none", pady=5)

frame_columna_izquierda = tk.Frame(frame_pieza, bg="#f4f4f4")
frame_columna_izquierda.pack(side="left", padx=20, fill="none")

frame_columna_derecha = tk.Frame(frame_pieza, bg="#f4f4f4")
frame_columna_derecha.pack(side="left", padx=20, fill="none")

tk.Label(frame_columna_izquierda, text="Numero di fori:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
entrada_aujeros = tk.Entry(frame_columna_izquierda, width=20, font=fuente)
entrada_aujeros.pack(pady=5, fill="none")

tk.Label(frame_columna_izquierda, text="Perimetro mm:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
entrada_perimetro = tk.Entry(frame_columna_izquierda, width=20, font=fuente)
entrada_perimetro.pack(pady=5, fill="none")

tk.Label(frame_columna_derecha, text="lunghezza mm:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
entrada_largo = tk.Entry(frame_columna_derecha, width=20, font=fuente)
entrada_largo.pack(pady=5, fill="none")

tk.Label(frame_columna_derecha, text="larghezza mm:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
entrada_ancho = tk.Entry(frame_columna_derecha, width=20, font=fuente)
entrada_ancho.pack(pady=5, fill="none")

# Dati Generali
tk.Label(frame_principal, text="Dati Generali", font=('Helvetica', 14, 'bold'), bg="#f4f4f4").pack(anchor="w", pady=10, fill="none")

# Frame para los campos del paquete
frame_paquete = tk.Frame(frame_principal, bg="#f4f4f4")
frame_paquete.pack(fill="none", pady=5)

# Sub-frame para organizar dos columnas
frame_columna_izquierda = tk.Frame(frame_paquete, bg="#f4f4f4")
frame_columna_izquierda.pack(side="left", padx=20, fill="none")

frame_columna_derecha = tk.Frame(frame_paquete, bg="#f4f4f4")
frame_columna_derecha.pack(side="left", padx=20, fill="none")

# Columna izquierda (primeros dos campos)
tk.Label(frame_columna_izquierda, text="Quantità del pacco m3:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
neto_pack = tk.Entry(frame_columna_izquierda, width=20, font=fuente)
neto_pack.pack(pady=5, fill="none")

tk.Label(frame_columna_izquierda, text="Costo del pacco EU:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
entrada_costo_pack = tk.Entry(frame_columna_izquierda, width=20, font=fuente)
entrada_costo_pack.pack(pady=5, fill="none")

# Columna derecha (último campo + nuevo campo)
tk.Label(frame_columna_derecha, text="Costo della macchina EU:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
entrada_maquina = tk.Entry(frame_columna_derecha, width=20, font=fuente)
entrada_maquina.pack(pady=5, fill="none")

# Campo adicional para Costo Hora Operarios
tk.Label(frame_columna_derecha, text="Costo ora operari EU:", font=fuente, bg="#f4f4f4").pack(anchor="w", fill="none")
entrada_costo_hora_operarios = tk.Entry(frame_columna_derecha, width=20, font=fuente)
entrada_costo_hora_operarios.pack(pady=5, fill="none")

# Vinculare l'evento di selezione del materiale con l'aggiornamento degli spessori
combo_material.bind("<<ComboboxSelected>>", actualizar_espesores)

# Bottone per raccogliere i dati
tk.Button(frame_principal, text="Calcolare i Dati", command=mostrar_resultados, font=fuente, bg="#2196F3", fg="white", relief="raised", padx=10, pady=5).pack(pady=15, fill="none")

# Etiqueta para el nombre del informe
tk.Label(frame_principal, text="Nome del Rapporto:", font=('Helvetica', 12)).pack(anchor="w", pady=5)
entrada_nombre_informe = tk.Entry(frame_principal, font=('Helvetica', 12), width=30)
entrada_nombre_informe.pack(pady=5)

# Crear un Treeview para mostrar los resultados en forma de tabla
treeview = ttk.Treeview(frame_principal, columns=("Costo Gas", "Costo Maquina", "Costo Peso", "Total", "Cantidad Gas", "Tiempo Corte", "Costo Operario"), show="headings")

# Configurar las columnas
treeview.heading("Costo Gas", text="Costo Gas")
treeview.heading("Costo Maquina", text="Costo Maquina")
treeview.heading("Costo Peso", text="Costo Peso")
treeview.heading("Total", text="Total")
treeview.heading("Cantidad Gas", text="Cantidad Gas")
treeview.heading("Tiempo Corte", text="Tiempo Corte (Horas)")
treeview.heading("Costo Operario", text="Costo Operario")

# Configurar el tamaño de las columnas
treeview.column("Costo Gas", width=100)
treeview.column("Costo Maquina", width=100)
treeview.column("Costo Peso", width=100)
treeview.column("Total", width=100)
treeview.column("Cantidad Gas", width=100)
treeview.column("Tiempo Corte", width=100)
treeview.column("Costo Operario", width=100)

# Colocar el Treeview en la ventana
treeview.pack(pady=20, fill="both", expand=True)

# Iniziare la finestra
ventana.mainloop()





