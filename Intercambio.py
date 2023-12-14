import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import pandas as pd
import random
import threading
import time

# Leer Excel una vez
nomina = pd.read_excel("Insumo Rifa HLS 2023.xlsx", sheet_name="Nomina")
nominaRandom = nomina.sample(frac=1).reset_index(drop=True)
regalos = pd.read_excel("Insumo Rifa HLS 2023.xlsx", sheet_name="Regalos")

# Variable global para el índice del ganador actual
indice_ganador_actual = 0

# DataFrame para almacenar los resultados
resultado = pd.DataFrame(columns=['IdTrabajador','NombreTrabajador', 'NombreRegalo', 'IdRegalo', 'ValorRegalo', 'DiasTrabajados'])

# Función para generar el DataFrame de resultados
def generarResultados():
    Filtrado()

    # Imprimir el DataFrame de resultados
    print(resultado)

def Filtrado():
    # Filtrado Nomina
    nominaBaja = nominaRandom[nominaRandom['Dias'] <= 365]
    nominaMedia = nominaRandom[(nominaRandom['Dias'] > 365) & (nominaRandom['Dias'] <= 730)]
    nominaAlta = nominaRandom[nominaRandom['Dias'] > 730]

    # Filtrado Regalos
    regalosBaja = regalos[regalos['Valor'] == 'Baja']
    regalosMedia = regalos[regalos['Valor'] == 'Media']
    regalosAlta = regalos[regalos['Valor'] == 'Alta']

    otorgarRegalos(nominaBaja, regalosBaja)
    otorgarRegalos(nominaMedia, regalosMedia)
    otorgarRegalos(nominaAlta, regalosAlta)

def otorgarRegalos(nomina, regalos):
    # Asignar regalos aleatorios.
    for index, row in nomina.iterrows():
        idNomina = row['id']
        nombreNomina = row['Nombre']
        diasNomina = row['Dias']

        # Verificar si hay regalos disponibles
        if len(regalos) == 0:
            return

        # Regalo aleatorio
        regaloSeleccionado = regalos.sample(n=1)

        idRegalo = regaloSeleccionado['id'].values[0]
        nombreRegalo = regaloSeleccionado['Regalo'].values[0]
        valorRegalo = regaloSeleccionado['Valor'].values[0]

        # Agregar resultados al DataFrame
        resultado.loc[len(resultado)] = [idNomina, nombreNomina, nombreRegalo, idRegalo, valorRegalo, diasNomina]

        # Eliminar regalo asignado del DataFrame original de regalos
        regalos = regalos[regalos['id'] != idRegalo]

# Llamada a la nueva función
generarResultados()

# Nueva función para generar el archivo Excel
def generarExcel():
    global resultado
    # Verificar si hay resultados para guardar
    if not resultado.empty:
        resultado.to_excel("resultados_rifa.xlsx", index=False)
        print("Archivo Excel generado exitosamente.")
    else:
        print("No hay resultados para generar el archivo Excel.")

# Función para mostrar el ganador actual y su regalo
def mostrarGanador():
    global indice_ganador_actual

    # Verificar si hay ganadores para mostrar
    if indice_ganador_actual < len(resultado):
        ganador_actual = resultado.loc[indice_ganador_actual, 'NombreTrabajador']
        regalo_actual = resultado.loc[indice_ganador_actual, 'NombreRegalo']
        total_ganadores = len(resultado)
        # Agregar el conteo al texto
        etiqueta_ganador.config(text=f"Ganador #{indice_ganador_actual + 1} de {total_ganadores}: {ganador_actual}")
        etiqueta_regalo.config(text=f"Regalo: {regalo_actual}")
        indice_ganador_actual += 1
    else:
        etiqueta_ganador.config(text="¡Todos los ganadores han sido mostrados \n HLS les desea una Feliz Navidad!")
        etiqueta_regalo.config(text="")

def ganadorAnterior():
    global indice_ganador_actual

    # Verificar si hay ganadores para mostrar
    if indice_ganador_actual > 0:
        indice_ganador_actual -= 1
        ganador_actual = resultado.loc[indice_ganador_actual, 'NombreTrabajador']
        regalo_actual = resultado.loc[indice_ganador_actual, 'NombreRegalo']
        total_ganadores = len(resultado)

        # Agregar el conteo al texto
        etiqueta_ganador.config(text=f"Ganador #{indice_ganador_actual + 1} de {total_ganadores}: {ganador_actual}")
        etiqueta_regalo.config(text=f"Regalo: {regalo_actual}")
    else:
        etiqueta_ganador.config(text="Estás en el inicio de la lista")
        etiqueta_regalo.config(text="")

def salir():
    root.destroy()

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Rifa Navideña de HLS Group")
root.attributes('-fullscreen',True)

def mostrarLabelCarga():
    labelCargaLeyenda = tk.Label(root, text="Escogiendo Ganadores...", font=("Helvetica", 20,),bg="#192440", fg="white")
    labelCargaLeyenda.pack(pady=5)
    
    labelCarga = tk.Label(root, text="", font=("Helvetica", 20,),bg="#192440", fg="white")
    labelCarga.pack(pady=5)
    
    # Lista de nombres para mostrar
    nombres = resultado['NombreTrabajador'].tolist()
    
    def mostrarSiguienteNombre(indice):
        if indice < len(nombres):
            labelCarga.config(text=nombres[indice])
            root.after(100, mostrarSiguienteNombre, indice + 1)
        else:
            labelCarga.pack_forget()

    # Iniciar el bucle de nombres
    mostrarSiguienteNombre(0)
    
    # Programar la destrucción del label después de 5 segundos
    root.after(5000, labelCargaLeyenda.destroy)
    # Programar la destrucción del label después de 5 segundos
    root.after(5000, labelCarga.destroy)


# Cambiar el color de fondo de la ventana principal a un tono de gris claro
root.configure(bg="#192440")

# Cargar el logo
logo_path = "logo.png"
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((400, 299))
logo_tk = ImageTk.PhotoImage(logo_image)

# Mostrar el logo en un widget Label sin borde y ajustar el tamaño del Label
logo_label = tk.Label(root, image=logo_tk, bg="#192440", bd=0)
logo_label.config(width=400, height=400)
logo_label.pack()

mostrarLabelCarga()

root.after(5000,mostrarGanador)

# Crear etiqueta para mostrar el ganador
etiqueta_ganador = tk.Label(root, text="", font=("Helvetica", 20), bg="#192440", fg="white")
etiqueta_ganador.pack(pady=5)

# Crear etiqueta para mostrar el regalo
etiqueta_regalo = tk.Label(root, text="", font=("Helvetica", 20), bg="#192440", fg="white")
etiqueta_regalo.pack(pady=5)

# Utilizar un estilo temático para los botones
style = ttk.Style()
style.configure("TButton", foreground="#006579", background="#006579", padding=(20, 10), font=("Helvetica", 12))
style.configure("TFrame", background="#192440")

# Contenedor para los botones con fondo de color
botones_frame = ttk.Frame(root, padding=(10, 10), style="TFrame")
botones_frame.pack()

# Botón para generar Excel
boton_generarExcel = ttk.Button(botones_frame, text="Generar Excel", command=generarExcel)
boton_generarExcel.pack(side=tk.LEFT, padx=10)

# Botón para reiniciar el índice
boton_reiniciarIndice = ttk.Button(botones_frame, text="Ganador Anterior", command=ganadorAnterior)
boton_reiniciarIndice.pack(side=tk.LEFT, padx=10)

# Botón para mostrar el ganador actual
boton_mostrarGanador = ttk.Button(botones_frame, text="Siguiente Ganador", command=mostrarGanador)
boton_mostrarGanador.pack(side=tk.LEFT, padx=10)

boton_salir = ttk.Button(botones_frame, text="Salir", command=salir)
boton_salir.pack(pady=10)

# Configurar el logo en la ventana principal
root.tk.call('wm', 'iconphoto', root._w, logo_tk)

# Lanzar la interfaz gráfica
root.mainloop()