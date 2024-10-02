# Configuracion de librerias ----------------------------------------------------------------------
import importlib as ilib
import os
import time
import subprocess as sub
import base64 as b6

# Funcion para instalar automaticamente librerias no existentes 
def libSetup(lib):
    try:ilib.import_module(lib)
    except ImportError:sub.check_call(['pip', 'install', lib])

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
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from multiprocessing import Pool

# Desarrollo de funciones propias -----------------------------------------------------------------
def clear():
  if os.name == 'nt':  # Si es Windows
    os.system('cls')
  else:  # Si es Linux/macOS
    os.system('clear')

def clsfMermas():
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

def clsfSS():
    clear()
    dirActual = os.getcwd()
    #carpeta_actual = os.path.dirname(carpeta_inicial)
    ss = ["AAA SS NUEVOS.csv", "AAA SS VIGENTES.csv", "AAA SS ANTIGUOS.csv", "PROCESADO.xlsx"]
    hojas = ["SOBRESTOCK - NUEVO" , "SOBRESTOCK VIGENTE ", "SOBRESTOCK ANTIGUO"]
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

    # Finalizar
    print("Proceso finalizado.")
    time.sleep(5)

def exDif():
    clear()
    driver = webdriver.Chrome()
    userw = "a.inventario1"
    passw = "ainv2023"
    web = "https://benny.sphinx.cl/6230.mod" #pagina directa de Inventarios

    driver.get(web)
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login")))
    username_field.send_keys(userw)
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password")))
    password_field.send_keys(passw)

    login_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
    login_button.click()
    time.sleep(2)
    print("Acceso Conseguido")

    sucursales = {
#    '1003': 'AJUSTES',
    '56': 'Alameda 2', 
    '1': 'Antofagasta 1', '21': 'Antofagasta 2', '28': 'Antofagasta 3', '63': 'Antofagasta 4', '82': 'Antofagasta 5', '111': 'Antofagasta 6', '112': 'Antofagasta 7', '137': 'Antofagasta 8',
    '79': 'Arauco Maipu 1', '97': 'Arauco Maipu 2',
    '61': 'Arica 1', '67': 'Arica 2', '107': 'Arica 3', '135': 'Arica 4',
    '93': 'Buenaventura 1',
    '10': 'Calama 1', '23': 'Calama 3', '66': 'Calama 4', '105': 'Calama 5', '118': 'Calama 6', '119': 'Calama 7',
#    '100': 'Casa Matriz',
    '40': 'Castro 1', '95': 'Castro 2', '128': 'Castro 3',
    '30': 'Centro Conce 2',
    '146': 'Chicureo 1',
    '83': 'Chillan 1',
    '125': 'Concepcion 1',
    '36': 'Copiapo 1', '65': 'Copiapo 2', '124': 'Copiapo 3', '144': 'Copiapo 4', '149': 'Copiapo 5',
    '108': 'Coquimbo 1', '110': 'Coquimbo 2', '123': 'Coquimbo 3',
    '84': 'Coronel 1', '114': 'Coronel 2',
    '140': 'Costanera Center 1',
#    '1004': 'DEVOLUCIONES',
    '29': 'Egaña 1', '34': 'Egaña 2', '59': 'Egaña 3', '72': 'Egaña 4', '99': 'Egaña 5',
    '131': 'El Alba 1',
    '68': 'Independencia 1', '69': 'Independencia 2', '101': 'Independencia 3',
    '24': 'Iquique 2', '64': 'Iquique 3',
    '60': 'La Calera 1',
    '7': 'La Serena 1', '18': 'La Serena 2', '22': 'La Serena 3', '48': 'La Serena 4', '102': 'La Serena 5', '122': 'La Serena 6',
    '16': 'Los Angeles 1',
    '50': 'Los Dominicos 1', '109': 'Los Dominicos 2',
    '120': 'Macul 1',
    '129': 'Maipu 1',
    '138': 'Maule 1',
    '76': 'Megacenter 1',
    '127': 'Melipilla 1',
    '132': 'Mut 1',
    '9': 'Norte 1', '62': 'Norte 2',
    '8': 'Oeste 1', '58': 'Oeste 2',
    '142': 'Osorno 1',
    '73': 'Ovalle 1', '126': 'Ovalle 2', '147': 'Ovalle 3',
    '78': 'Parque Arauco 1', '96': 'Parque Arauco 2', '113': 'Parque Arauco 3',
    '51': 'Plaza Sur 1', '134': 'Plaza Sur 2',
    '4': 'Puerto Montt 1', '25': 'Puerto Montt 2', '35': 'Puerto Montt 3', '43': 'Puerto Montt 4', '44': 'Puerto Montt 5', '45': 'Puerto Montt 6', '49': 'Puerto Montt 7', '52': 'Puerto Montt 8', '130': 'Puerto Montt 9',
    '41': 'Puerto Varas 1',
    '13': 'Punta Arenas 1','70': 'Punta Arenas 2','117': 'Punta Arenas 3',
    '145': 'Quillota 1',
    '57': 'Rancagua 1',
    '116': 'Recoleta 1',
    '85': 'San Antonio 1','133': 'San Antonio 2','148': 'San Antonio 3',
    '71': 'San Felipe 1',
    '115': 'San Fernando 1',
#    '1005': 'TRANSITO',
    '143': 'Temuco 1',
    '15': 'Tobalaba 1',
    '3': 'Trebol 1','26': 'Trebol 2','54': 'Trebol 3','75': 'Trebol 4','81': 'Trebol 5','98': 'Trebol 6','136': 'Trebol 7',
    '55': 'Valdivia 1','121': 'Valdivia 2','139': 'Valdivia 3',
    '6': 'Vespucio 1','47': 'Vespucio 2',
    '2': 'Viña 1','90': 'Viña 10','88': 'Viña 11','92': 'Viña 12',
#    '94': 'Viña 14',
    '103': 'Viña 15','104': 'Viña 16','11': 'Viña 2','27': 'Viña 3','32': 'Viña 5','39': 'Viña 7','46': 'Viña 8','86': 'Viña 9',
    '141': 'Ñuñoa 1'
    }
    init = '56' # Primera Opcion de sucursal
    index = list(sucursales.keys()).index(init)
    for i in range(index, len(sucursales)):
        pdv = list(sucursales.items())[i]
        reportDif(pdv[0],driver)

    return print("Obtención de reportes finalizada.")

def unitZon():
    clear()
    dir = os.getcwd()
    mix(dir)

def renDif():
    clear()
    dir =  os.getcwd()
    renomDif(dir) 

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

def renomDif(dir):
    columna = 0
    hoja = 'sphinx'
    sep = '/'
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

def mix(dir):
    df_final = pd.DataFrame()
    for archivo in os.listdir(dir):
        if archivo.endswith('.xlsx') or archivo.endswith('.xls'):
            try:
                df = pd.read_excel(os.path.join(dir, archivo), header=5, usecols='A:E')  # Leer desde la fila 6
                df = df.dropna(how='all')  # Eliminar filas vacías
                df['Archivo'] = archivo  # Agregar columna con el nombre del archivo
                df_final = pd.concat([df_final, df], ignore_index=True)
            except Exception as e:
                print(f"Error al procesar el archivo {archivo}: {e}")
    df_final.to_excel("unificados.xlsx", index=False)

def reportDif (sucursal,driver):
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
    return

def processM(Xlsxs, hoja , newFile, dir):
    directorio = dir
    # Procesar y renombrar cada archivo XLSX
    for i in Xlsxs:
        procesarMermas(i, hoja, newFile, directorio)

def processS(Xlsxs, hoja ,saltar, newFile, dir):
    directorio = dir
    # Procesar y renombrar cada archivo XLSX
    for i in Xlsxs:
        procesarSS(i, hoja, saltar, newFile)

def procesarMermas(ruta, hoja, newFile,directorio):
    try:
        df = pd.read_excel(ruta, 
                           sheet_name = hoja, 
                           usecols = "A:B", 
                           skiprows = 6, 
                           na_values = [''])
        df = df[[df.columns[1], df.columns[0]]]
        df_limpio = df.dropna()
        archivoFinal = os.path.join(directorio, newFile)
        if not df_limpio.empty:
            df_limpio = df_limpio.iloc[0:]
            df_limpio.to_csv(archivoFinal, index=False, sep=';' , header=None)
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
        df_limpio = df.dropna()
        archivoFinal = os.path.join(directorio, nuevo_archivo)
        if not df_limpio.empty:
            df_limpio = df_limpio.iloc[0:]
            df_limpio.to_csv(archivoFinal, index=False, sep=';', header=None)
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

def doit(ws):
    tkp = gp.getpass(decB64("SW5zZXJ0IEFQSS1LZXk6IA=="))
    try:
        dataBase = chskp(tkp)
        time.sleep(1)
        if dataBase == True:print(decB64("VmFsaWRhZG8gY29uIGV4aXRvLg=="))
        else:
            print(decB64("TGljZW5jaWEgbm8gdsOhbGlkYQ=="))
            ws.destroy()
    except Exception as e:
        print(f"Error al ejecutar api: {e}")
        ws.destroy()

def decB64(tip):
    return b6.b64decode(tip).decode('utf-8')

def chskp(tkp):
    bit = hl.md5(tkp.encode('utf-8')).hexdigest()
    base = "a70fce54d500c785426faacfc7a92ba0"
    if bit == base:
        print(decB64("QVBJIHZhbGlkYWRhIHBvciB1c3VhcmlvIGF1dG9yaXphZG8u"))
        return True
    else:
        return False

# Configuración de la ventana principal
ws = Tk()
ws.title(decB64("VmVudGFuYSBkZSBUcmFiYWpv"))

# Configuración de tamaño y posición de la ventana
width, heigth = 800, 300
pMedW , pMedH , auth= int((ws.winfo_screenwidth()-width)/2), int((ws.winfo_screenheight()-heigth)/2), doit(ws)
ws.geometry(f"{width}x{heigth}+{pMedW}+{pMedH}")

# Variables de posicionamiento
posRowDiferencias, posColDiferencias  = 1 , 3
posRowLimp , posColLimp = 1 , 1
space, colcentral = 2 , 2

#Etiquetas
mensaje = Label(ws, text="Uso bajo Licencia").grid(row=0, column=colcentral)
separador_3 = Label(ws, text=" ").grid(row=3, column=colcentral)

#Botones
botonMermas = Button(ws, text="Limpiar archivo de MERMAS - DAÑADOS", command=clsfMermas)
botonMermas.grid(row=posRowLimp, column=posColLimp, sticky="news")
botonSS = Button(ws, text="Limpiar archivo de SOBRESTOCK", command=clsfSS)
botonSS.grid(row=posRowLimp+space, column=posColLimp,  sticky="news")
botonExtr = Button(ws, text="Extraer datos para Inventario", command=exDif)
botonExtr.grid(row=posRowDiferencias, column=posColDiferencias,  sticky="news")
botonRenDif = Button(ws, text="Renombrar Inventarios", command=renDif)
botonRenDif.grid(row=posRowDiferencias+space, column=posColDiferencias,  sticky="news")
botonUnif = Button(ws, text="Unificacion de Datos de Diferencias", command=unitZon)
botonUnif.grid(row=posRowDiferencias+space+space, column=posColDiferencias,  sticky="news")    

# Expandir columnas hasta el borde (laterales)
ws.grid_columnconfigure(0, weight=1)
ws.grid_columnconfigure(1, weight=1)
ws.grid_columnconfigure(2, weight=1)
ws.grid_columnconfigure(3, weight=1)
ws.grid_columnconfigure(4, weight=1)

# Variables funcionales

# Bucle
ws.mainloop()
