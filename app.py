import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

# Ruta del archivo predeterminado
ARCHIVO = "date.ods"  # O "date.xlsx" si usas archivo Excel

# Variable global para almacenar la ruta del archivo
archivo_cargado = ARCHIVO

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
        datos_filtrados_dict = {} # Usamos la variable global para almacenar los datos como diccionario
        archivo = archivo_cargado  # Usamos la variable global para la ruta del archivo cargado
        material_usuario = combo_material.get()  # Obtener el material seleccionado
        espesor_usuario = combo_espesor.get()  # Obtener el espesor seleccionado
        
        # Verificar si el espesor está vacío
        if espesor_usuario:
            espesor_usuario = float(espesor_usuario)  # Convertir espesor a float
        else:
            espesor_usuario = None  # Si no se ingresa espesor, no filtramos por espesor

        # Leer el archivo ODS (si es un archivo LibreOffice) o Excel
        if archivo.endswith(".ods"):
            df = pd.read_excel(archivo, engine="odf", sheet_name="date")  # Leemos la hoja llamada "date"
        else:
            df = pd.read_excel(archivo, sheet_name="date")  # Leemos la hoja llamada "date" para Excel

        # Filtrar los datos por material y espesor si es necesario
        if material_usuario and espesor_usuario is not None:
            df_filtrado = df[(df['Material'] == material_usuario) & (df['Espesor'] == espesor_usuario)]
        else:
            df_filtrado = df  # Si no se filtra por material o espesor, usamos todos los datos

        # Verificar si se encontraron resultados
        if df_filtrado.empty:
            messagebox.showwarning("Nessun risultato", "Non sono stati trovati dati con i criteri selezionati.")  # Advertencia en italiano
            return None  # Retorna None si no se encuentran resultados
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
        return None  # Si no hay datos, no hacer cálculos

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
        cantidad_aujeros = float(cantidad_aujeros)    # Número de agujeros


        
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
        tiempo_aujeros = cantidad_aujeros * (tiempo_1 + tiempo_2)  # Tiempo adicional por los agujeros
        tiempo_total = tiempo_corte + tiempo_aujeros  # Tiempo total

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
            return None  # Si no se puede obtener datos, salimos

        # Calculamos tiempos
        tiempo_corte_dict = calcular_tiempo_corte()
        if not tiempo_corte_dict:
            return None  # Si no se pueden calcular los tiempos, salimos

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
        duracion_pack = datos_filtrados_dict[0]['Duracion']  # Usamos None si no existe la clave
        
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
        
        return consumo  # Retorna el valor calculado de consumo

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")
        return None


def calcular_costos_general():
    
    # Primero recolectamos los datos
    datos_filtrados_dict = recolectar_datos()
    
    # Calculamo tiempos corte
    tiempo_corte_dict = calcular_tiempo_corte()

    # Calculamo consumo gas
    tiempo_gas_dict = calcular_consumo_gas()


    costo_pack_usuario = entrada_costo_pack.get()
    costo_maquina_usuario = entrada_maquina.get()


    if not costo_pack_usuario:
        messagebox.showerror("Errore", "Per favore, inserisci il costo del pack.")  
        return None
    
    if not costo_maquina_usuario:
        messagebox.showerror("Errore", "Per favore, inserisci il costo del pack.")  
        return None
    


def generar_informe():
    try:
        # Verificar si ya se han recolectado los datos
        if not datos_filtrados_dict:
            messagebox.showerror("Errore", "Non sono stati raccolti dati. Per favore, raccogli prima i dati.")  
            return
        
        # Aquí podrías guardar los datos recolectados en un informe
        # Por ejemplo, como un archivo ODS o Excel:
        nombre_informe = entrada_nombre_informe.get()  # Obtener el nombre del informe
        
        if not nombre_informe:
            messagebox.showerror("Errore", "Per favore, inserisci un nome per il rapporto.")  
            return

        # Guardar los datos recolectados en un archivo ODS
        df_informe = pd.DataFrame(datos_filtrados_dict)  # Convertimos el diccionario a un DataFrame
        df_informe.to_excel(f"{nombre_informe}.ods", index=False, engine="odf")  # Guardamos como .ods
        
        messagebox.showinfo("Rapporto Generato", f"Rapporto generato con successo come '{nombre_informe}.ods'.")  # Mensaje de éxito en italiano

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")  


# Configurar la ventana principal
ventana = tk.Tk()
ventana.title("App di Calcoli")
ventana.geometry("400x600")

# Entrada de archivo (ahora el archivo puede ser cambiado)
entrada_archivo = tk.StringVar(value=archivo_cargado)  # Inicializamos con el archivo por defecto
tk.Label(ventana, text="File Excel:").pack()  # Traducido a italiano
tk.Entry(ventana, textvariable=entrada_archivo, width=40, state='readonly').pack()  # Solo lectura

# Botón para cargar otro archivo
tk.Button(ventana, text="Carica un altro file", command=cargar_archivo).pack()  # Traducido a italiano

# Leer los datos del archivo cargado
df = pd.read_excel(archivo_cargado, engine="odf", sheet_name="date")

# Obtener lista de materiales únicos
materiales = df['Material'].dropna().unique().tolist()

# Crear combobox para material
tk.Label(ventana, text="Seleziona Materiale:").pack()  # Traducido
combo_material = ttk.Combobox(ventana, values=materiales, width=40)
combo_material.pack()

# Crear combobox para espesor
tk.Label(ventana, text="Seleziona Spessore:").pack()  # Traducido
combo_espesor = ttk.Combobox(ventana, width=40)
combo_espesor.pack()

# Crear entrada de texto para los agujeros
tk.Label(ventana, text="Inserisci numero di fori:").pack()  # Traducido
entrada_aujeros = tk.Entry(ventana, width=40)
entrada_aujeros.pack()

# Crear entrada de texto para el perímetro
tk.Label(ventana, text="Inserisci Perimetro mm:").pack()  # Traducido
entrada_perimetro = tk.Entry(ventana, width=40)
entrada_perimetro.pack()

# Crear entrada de texto para el volumen del paquete
tk.Label(ventana, text="Inserisci quantità del pacco m3:").pack()  # Traducido
neto_pack = tk.Entry(ventana, width=40)
neto_pack.pack()

# Crear entrada de texto para el costo
tk.Label(ventana, text="Inserisci Costo del pacco EU:").pack()  # Traducido
entrada_costo_pack = tk.Entry(ventana, width=40)
entrada_costo_pack.pack()

# Crear entrada de texto para el costo
tk.Label(ventana, text="Inserisci Costo della maquina EU:").pack()  # Traducido
entrada_maquina = tk.Entry(ventana, width=40)
entrada_maquina.pack()

# Vincular el evento de selección de material a la actualización de espesores
combo_material.bind("<<ComboboxSelected>>", actualizar_espesores)

# Botón para recolectar los datos
tk.Button(ventana, text="Raccogli Dati", command=calcular_costos_general).pack()  # Traducido

# Crear entrada de texto para el nombre del informe
tk.Label(ventana, text="Nome del Rapporto:").pack()  # Traducido
entrada_nombre_informe = tk.Entry(ventana, width=40)
entrada_nombre_informe.pack()

# Botón para generar informe
tk.Button(ventana, text="Genera Rapporto", command=generar_informe).pack()  # Traducido

ventana.mainloop()