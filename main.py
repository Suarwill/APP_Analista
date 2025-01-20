# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═
# Configuracion de librerias
# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═

import importlib as ilib
import subprocess as sub
def libSetup(lib):
    # Funcion para instalar automaticamente librerias no existentes
    try:ilib.import_module(lib)
    except ImportError:sub.check_call(['pip', 'install', lib])

import os, time, csv
import datetime as dt
import base64 as b6
from multiprocessing import Pool

libSetup('getpass')
import getpass as gp

libSetup('hashlib')
import hashlib as hl

libSetup('tkinter')
from tkinter import *
from tkinter import messagebox

libSetup('pandas')
import pandas as pd

libSetup('warnings')
import warnings

libSetup('selenium')
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.alert import Alert

# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═
# Desarrollo de funciones propias
# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═
def clear():
  if os.name == 'nt':        # Si es Windows
    os.system('cls')
  else:                      # Si es Linux/macOS
    os.system('clear')

def decB64(tip):
    return b6.b64decode(tip).decode('utf-8')

def docCSV(documento):
    with open(documento, 'r') as csvfile:
        reader = csv.reader(csvfile, delimiter=';')
        listado = {row[0]: row[1] for row in reader}
    print(decB64("bGlzdGFkbyBkZSBzdWN1cnNhbGVzIG9idGVuaWRv"))
    return listado

def clasificadorMermas():
    # Clasificador de mermas, borra anteriores y extrae nuevas para AAA
    clear()
    dirActual = os.getcwd() #directorio actual
    #dirSuperior = os.path.dirname(dirActual) #directorio superior
    mermas = ["AAA DAÑADOS.csv", "AAA NC.csv", "AAA ELIMINADOS.csv", "PROCESADO.xlsx"]
    eliminar = mermas
    Xlsxs = buscXLSX(dirActual)
    print("Este archivo:\nEliminará los anteriores procesados\nProcesará el XLSX en la carpeta\n")
    time.sleep(1)
    clean(dirActual, eliminar)
    print("\n")
    processM(Xlsxs, "P. DAÑADOS" , mermas[0], dirActual)
    processM(Xlsxs, "PRODUCTOS DAÑADOS POR NC" , mermas[1], dirActual)
    processM(Xlsxs, "ESTATUS ELIMINADO" , mermas[2], dirActual)
    print("\n")
    renombrar(Xlsxs[0], mermas[3])
    print("Proceso finalizado.")
    time.sleep(3)

def clasificadorSS():
    clear()
    dirActual = os.getcwd()
    #carpeta_actual = os.path.dirname(carpeta_inicial)
    ss = ["AAA SS NUEVOS.csv", "AAA SS VIGENTES.csv", "AAA SS ANTIGUOS.csv", "AAA ELIMINADOS.csv", "PROCESADO.xlsx"]
    hojas = ["SOBRESTOCK - NUEVO" , "SOBRESTOCK VIGENTE ", "SOBRESTOCK ANTIGUO", "AAA ELIMINADOS.csv"]
    eliminar = ss
    print("Este archivo:\nEliminará los anteriores procesados\nProcesará el XLSX en la carpeta\n")
    time.sleep(1)
    clean(dirActual, eliminar)
    print("\n")
    time.sleep(1)
    archivos_xlsx = buscXLSX(dirActual)
    processS(archivos_xlsx, hojas[0] , 6 , ss[0], dirActual)
    processS(archivos_xlsx, hojas[1], 7 , ss[1], dirActual)
    processS(archivos_xlsx, hojas[2] , 7 , ss[2], dirActual)
    print("\n")
    time.sleep(1)
    renombrar(archivos_xlsx[0], ss[3])

    print("Proceso finalizado.")
    time.sleep(3)

def extraerDif():
    clear()
    driver = webdriver.Chrome()
    AccesoWEB(driver,"a.inventario1","ainv2023","https://benny.sphinx.cl/6230.mod")

    # Obtener sucursales
    sucursales = docCSV('sucursales.csv')

    init = '56' # Primera Opcion de sucursal
    index = list(sucursales.keys()).index(init)
    for i in range(index, len(sucursales)):
        pdv = list(sucursales.items())[i]
        reportDif(pdv[0],driver)
    clear()
    return print("Obtención de reportes finalizada.")

def unitZones():
    # Unificado de archivos en uno SOLO
    dir = os.getcwd()
    df_final = pd.DataFrame()
    filesExcel = buscXLSX(dir) 
    for x in filesExcel:
        try:
            df = pd.read_excel(os.path.join(dir, x), header=5, usecols='A:E')
            df = df.dropna(how='all')
            df['Archivo'] = x

            # Verificar y modificar la celda C7
            if df.loc[6, 'C'] != df.loc[6, 'C']:
                df.loc[6, 'C'] = 'x'

            df_final = pd.concat([df_final, df], ignore_index=True)
        except Exception as e:
            print(f"Error al procesar el archivo {x}: {e}")

    df_final.to_excel("unificados.xlsx", index=False)
    clear()
    return print("Unificación realizada.")

def renombrarDif():
    # Renombrar todos los archivos "Inventario" y colocar el renombra como su sucursal
    dir =  os.getcwd()
    columna, hoja, sep = 0,'sphinx', '/'
    xlsxs = buscXlsxDif(dir)
    for x in xlsxs:
        try:
            # Leer el archivo Excel sin especificar el encabezado
            df = pd.read_excel(x, sheet_name=hoja, nrows=5, header=None)
            # Verificar si la fila 4 existe y tiene un valor en la columna especificada
            if df.shape[0] >= 4 and not pd.isna(df.iloc[3][columna]):
                valor_celda = df.iloc[3][columna]
                # Dividir el valor por el separador y tomar la segunda parte
                nueva_parte = valor_celda.split(sep)[1].strip()
                # Construir el nuevo nombre de archivo
                nombre_directorio, extension = os.path.splitext(x)
                nuevo_nombre = nueva_parte + extension
                # Renombrar el archivo
                os.rename(x, nuevo_nombre)
                print(f"Archivo renombrado a: {nuevo_nombre}")
            else:
                print(f"La fila 4 en la hoja '{hoja}' del archivo {x} está vacía o no existe.")
        except (FileNotFoundError, PermissionError, IndexError, ValueError) as e:
            print(f"Error al procesar el archivo {x}: {e}")
        time.sleep(1)
    clear()
    return print("Renombrado finalizado.") 

def buscXLSX(dir):
    list = os.listdir(dir)
    return [os.path.join(dir, x) for x in list if x.endswith('.xlsx')]

def buscXlsxDif(dir):
    files = os.listdir(dir)
    return [os.path.join(dir, i) for i in files if i.endswith('.xlsx') and i.startswith("Inventario")]

def clean(dir, files):
    # Elimina archivos anteriores de la carpeta.
    warnings.filterwarnings("ignore", category=UserWarning)
    for i in files:
        ruta = os.path.join(dir, i)
        try:
            os.remove(ruta)
            print(f"El archivo {i} ha sido eliminado.")
        except FileNotFoundError:
            print(f"No se encontró el archivo {i}.")
        except PermissionError:
            print(f"No tienes permisos suficientes para eliminar {i}.")
        except OSError as error:
            print(f"Ocurrió un error al eliminar el archivo {i}: {error}")

def renombrar(dir, newName):
    try:
        nueva_ruta = os.path.join(os.path.dirname(dir), newName)
        os.rename(dir, nueva_ruta)
        print(f"Archivo renombrado a: {nueva_ruta}")
    except OSError as error:
        print(f"Error al renombrar el archivo: {error}")

def reportDif (sucursal,driver):
    try:
        select_element  = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Sphinx_Sucursales")))
        select = Select(select_element)
        select.select_by_value(str(sucursal))
        time.sleep(1)

        select_element  = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "inventario_tipo")))
        select = Select(select_element)
        select.select_by_value("2") # Productos con Diferencias
        time.sleep(1)

        login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnEjecuta")))
        login_button.click()
        time.sleep(2)

        login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnExcel")))
        login_button.click()
        time.sleep(1)
        print(sucursal)
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error al descargar informe {sucursal}: {e}")
        
    return None

def processM(Xlsxs, hoja , newFile, dir):
    for i in Xlsxs:
        procesarMermas(i, hoja, newFile, dir)

def processS(Xlsxs, hoja ,saltar, newFile, dir):
    for i in Xlsxs:
        procesarSS(i, hoja, saltar, newFile, dir)

def procesarMermas(ruta, hoja, newFile,directorio):
    try:
        df = pd.read_excel(ruta, 
                           sheet_name = hoja, 
                           usecols = "A:B", 
                           skiprows = 6, 
                           na_values = [''])
        df = df[[df.columns[1], df.columns[0]]]

        df_limpio = df.dropna(how='all')

        archivoFinal = os.path.join(directorio, newFile)
        if not df_limpio.empty:
            df_limpio = df_limpio.iloc[0:]
            df_limpio.to_csv(archivoFinal, index=False, sep=';' , header=None, float_format='%d')
            print(f"→ → → → → Datos invertidos guardados en {newFile}\n")
        else:
            print(f"No hay datos en: {newFile}\n")
    except (FileNotFoundError, ValueError) as e:
        print(f"Error al procesar el archivo")
    except (PermissionError) as e:
        print(f"Error de permisos.")
    except (IndexError) as e:
        print(f"Hoja vacia {hoja}")

def procesarSS(ruta_archivo, hoja, saltar, nuevo_archivo,directorio):
    try:
        df = pd.read_excel(ruta_archivo, 
                           sheet_name = hoja, 
                           usecols = "A:B", 
                           skiprows = saltar, 
                           na_values = [''])
        df = df[[df.columns[1], df.columns[0]]]

        df_limpio = df.dropna(how='all')

        archivoFinal = os.path.join(directorio, nuevo_archivo)
        if not df_limpio.empty:
            df_limpio = df_limpio.iloc[0:]
            df_limpio.to_csv(archivoFinal, index=False, sep=';', header=None, float_format='%d')
            print(f"Datos invertidos guardados en {nuevo_archivo}\n")
        else:
            print(f"No hay datos para guardar en {nuevo_archivo}\n")

    except (FileNotFoundError, PermissionError, ValueError) as e:
        print(f"Error al procesar el archivo {ruta_archivo}: {e}")
    except (IndexError) as e:
        print(f"Hoja vacia {hoja}")

def exec(file):
    # Función para ejecutar scripts externos de forma asíncrona
    try:
        with Pool(processes=1) as pool:
            pool.apply_async(sub.run, ["python", file])
        messagebox.showinfo("Función ejecutada.")
    except FileNotFoundError:
        messagebox.showerror("Error", "No se encontró el archivo raíz")

def windows(ws):
    tkp = gp.getpass(decB64("SW5zZXJ0IEFQSS1LZXk6IA=="))
    gaucho = dt.date.today()
    taskLimit = dt.date(2025, 3, 19)
    try:
        dataBase = chskp(tkp)
        time.sleep(1)
        if dataBase == True and gaucho < taskLimit :print(decB64("VmFsaWRhZG8gY29uIGV4aXRvLg=="))
        else:
            print(decB64("TGljZW5jaWEgbm8gdsOhbGlkYQ=="))
            ws.destroy()
    except Exception as e:
        print(f"Error al ejecutar api: {e}")
        ws.destroy()

def chskp(tkp):
    bit = hl.md5(tkp.encode('utf-8')).hexdigest()
    base = "a70fce54d500c785426faacfc7a92ba0"
    if bit == base:
        print(decB64("QVBJIHZhbGlkYWRhIHBvciB1c3VhcmlvIGF1dG9yaXphZG8u"))
        return True
    else:
        return False

def closerInv():
    clear()
    driver = webdriver.Chrome()
    AccesoWEB(driver,"a.inventario1","ainv2023","https://benny.sphinx.cl/6210.mod")

    # Obtener sucursales
    sucursales = docCSV('sucursales.csv')

    init = '56' # Primera Opcion de sucursal
    index = list(sucursales.keys()).index(init)
    for i in range(index, len(sucursales)):
        pdv = list(sucursales.items())[i]
        closeInventory(pdv[0],driver)
    clear()
    return print("Inventarios del dia cerrados.")

def closeInventory (sucursal,driver):
    try:
        select_element  = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Sphinx_Sucursales")))
        select = Select(select_element)
        select.select_by_value(str(sucursal))
        time.sleep(2)

        close_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//td[@id='inventarioAbierto']//input[@value='C']")))
        close_button.click()
        time.sleep(2)
        try:
            alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
            alert.accept()
            print("Aceptado!")
        except TimeoutException:
            print("No se detecto alerta")
        print(sucursal)
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error al descargar informe {sucursal}")

def AccesoWEB(driver,userw,passw,web):
    driver.get(web)
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login")))
    username_field.send_keys(userw)
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password")))
    password_field.send_keys(passw)
    login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
    login_button.click()
    print(decB64("QWNjZXNvIENvbnNlZ3VpZG8="))
    time.sleep(2)
    return

# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═
# Configuración de la ventana principal
# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═
ws = Tk()
ws.title(decB64("VmVudGFuYSBkZSBUcmFiYWpv"))

warnings.filterwarnings("ignore", category=UserWarning)

# Configuración de tamaño y posición de la ventana
width, heigth = 800, 300
pMedW , pMedH , pAuth = int((ws.winfo_screenwidth()-width)/2), int((ws.winfo_screenheight()-heigth)/2), windows(ws)
ws.geometry(f"{width}x{heigth}+{pMedW}+{pMedH}")

# Variables de posicionamiento
posRowDiferencias, posColDiferencias  = 1 , 3
posRowLimp , posColLimp = 1 , 1
space, colcentral = 2 , 2

#Etiquetas
mensaje = Label(ws, text="Uso bajo Licencia").grid(row=0, column=colcentral)
separador_3 = Label(ws, text=" ").grid(row=3, column=colcentral)

#Botones
botonMermas = Button(ws, text="Limpiar archivo de MERMAS - DAÑADOS", command=clasificadorMermas)
botonMermas.grid(row=posRowLimp, column=posColLimp, sticky="news")

botonSS = Button(ws, text="Limpiar archivo de SOBRESTOCK", command=clasificadorSS)
botonSS.grid(row=posRowLimp+space, column=posColLimp,  sticky="news")

botonExtr = Button(ws, text="Extraer datos para Inventario", command=extraerDif)
botonExtr.grid(row=posRowDiferencias, column=posColDiferencias,  sticky="news")

botonRenameDif = Button(ws, text="Renombrar Inventarios", command=renombrarDif)
botonRenameDif.grid(row=posRowDiferencias+space, column=posColDiferencias,  sticky="news")

botonUnificador = Button(ws, text="Unificacion de Datos de Diferencias", command=unitZones)
botonUnificador.grid(row=posRowDiferencias+space*2, column=posColDiferencias,  sticky="news")

botonCierreInv = Button(ws, text="Cerrar Inventarios Abiertos", command=closerInv)
botonCierreInv.grid(row=posRowDiferencias+space*3, column=posColDiferencias,  sticky="news")

# Expandir columnas hasta el borde (laterales)
ws.grid_columnconfigure(0, weight=1)
ws.grid_columnconfigure(1, weight=1)
ws.grid_columnconfigure(2, weight=1)
ws.grid_columnconfigure(3, weight=1)
ws.grid_columnconfigure(4, weight=1)
ws.grid_columnconfigure(5, weight=1)

# Bucle
ws.mainloop()

# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═
# Crear Paquete EXE
# pyinstaller --onefile main.py
# ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═ ═
