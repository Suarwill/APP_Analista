class Ventana:
    def __init__(self, titulo, width, height):
        self.ventana = Tk()
        load_dotenv(override=True)
        self.ventana.title(titulo)
        self.ventana.geometry(f"{width}x{height}+{int((self.ventana.winfo_screenwidth()-width)/2)}+{int((self.ventana.winfo_screenheight()-height)/2)}")

    def crearBoton(self, texto, comando, fila, columna, **kwargs):
        Button(self.ventana, text=texto, command=comando, **kwargs).grid(row=fila, column=columna, sticky="news")

    def crearEtiqueta(self, texto, fila, columna, **kwargs):
        Label(self.ventana, text=texto, **kwargs).grid(row=fila, column=columna, padx=5, pady=5)

    def crearEntradaTexto(self, fila, columna, width, height, **kwargs):
        Text(self.ventana, width=width, height=height, **kwargs).grid(row=fila, column=columna, padx=5, pady=5)

    def expandirColumnas(self, num_columnas):
        for x in range(num_columnas):
            self.ventana.grid_columnconfigure(x, weight=1)

    def destroy(self):
        self.ventana.destroy()

    def iniciar(self):
        self.ventana.mainloop()

class VentanaPrincipal(Ventana):
    def __init__(self):
        super().__init__("Principal", 300, 200)
        self.crearEtiqueta(" ", 0, 0)
        self.crearBoton("Archivos Excel", lambda: VentanaExcel(self.ventana), 1, 1, background="lightblue")
        self.crearBoton("Sphinx", lambda: VentanaSphinx(self.ventana), 2, 1, background="lightblue")
        self.crearEtiqueta(" ", 3, 2)
        self.crearBoton("Configuración", lambda: VentanaConfigurar(self.ventana), 4, 1, background="lightblue")
        self.crearEtiqueta(" ", 0, 2)
        self.expandirColumnas(3)

class VentanaExcel(Ventana):
    def __init__(self, ventana_padre):
        super().__init__("Secundaria", 700, 300)

        self.crearEtiqueta("Archivo:", 0, 0)
        self.crearBoton("Extraer SobreStocks", Excel.sobrestock(), 1, 1, background="lightblue")

class VentanaSphinx(Ventana):
    def __init__(self, ventana_padre):
        super().__init__("Secundaria", 500, 200)

        urlDif = "https://benny.sphinx.cl/6230.mod"
        urlInv = ""

        self.crearEtiqueta(" ", 0, 0)
        self.crearBoton("Extraer Diferencias", lambda: self.extraerDiferencias(urlDif), 1, 1, background="lightblue")
        self.crearBoton("Unificar Archivos", self.unificar, 1, 1, background="lightblue")
        self.crearEtiqueta(" ", 0, 2)

    def cerrarINV(url):
        web = paginaWeb(url)
        web.login("login","password","btnSubmit")
        listado = funciones.leerCSV("Sucursales.csv")
        for sucursal in listado:
            web.cerrarInventario(sucursal)
        web.quit()
        return  print("Inventarios del dia cerrados.")
    
    def extraerDiferencias(urlDif):
        funciones.clear()
        web = paginaWeb(urlDif)
        web.login("login","password","btnSubmit")
        listado = funciones.leerCSV("Sucursales.csv")
        for sucursal in listado:
            web.reporteDiferencias(sucursal)
        Excel.renombrarArchivos()
        web.quit()
        return print("Documentos extraidos")

    def unificar():
        return Excel.unificar()

class VentanaConfigurar(Ventana):
    def __init__(self, ventana_padre):
        super().__init__("Configuraciones",400, 200)
        load_dotenv(override=True)

        userLabel = Label(self.ventana, text="Usuario: ")
        userLabel.grid(row=0, column=1, padx=5, pady=5)
        userDato = Entry(self.ventana, width=30)
        userDato.grid(row=0, column=2, padx=5, pady=5)
        passLabel = Label(self.ventana, text="Contraseña: ")
        passLabel.grid(row=1, column=1, padx=5, pady=5)
        passDato = Entry(self.ventana, width=30)
        passDato.grid(row=1, column=2, padx=5, pady=5)

        user = funciones.codec(os.getenv("USERNAME"),False)
        userDato.insert(0,user)
        pasw = funciones.codec(os.getenv("PASSWORD"),False)
        passDato.insert(0,pasw)

        self.crearEtiqueta(" ", 2, 0)
        self.crearEtiqueta(" ", 0, 0)
        self.crearBoton("Guardar", lambda : self.guardar(self), 3, 2, background="lightblue")
        self.crearBoton("Cerrar", self.destroy, 3, 1, background="lightblue")

        self.expandirColumnas(5)
        self.iniciar()

    def guardar(self):
        user =      self.userDato.get("1.0", "end-1c")
        password =  self.passDato.get("1.0", "end-1c")

        if os.path.exists('.env'):
            set_key(".env", "USERNAME", funciones.codec(user))
            set_key(".env", "PASSWORD", funciones.codec(password))
        print("Archivos actualizados con éxito.")
        return 

class paginaWeb:
    def __init__(self, url):
        options = Options()
        load_dotenv(override=True)
        chrome_profile_path = os.path.expandvars(os.getenv("PERFIL_CHROME"))
        options.add_argument(f"user-data-dir={chrome_profile_path}")
        options.add_argument("--disable-notifications")

        self.driver = webdriver.Chrome(options=options)
        self.username = funciones.codec(os.getenv("USERNAME"),False)
        self.password = funciones.codec(os.getenv("PASSWORD"),False)
        self.url = url

    def login(self,NAMEBoxUsuario,IDBoxPassword,IDBotonLogin):
        self.driver.get(self.url)
        usuario = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, NAMEBoxUsuario)))
        usuario.send_keys(self.username)
        contrasena = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, IDBoxPassword)))
        contrasena.send_keys(self.password)
        botonLogin = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, IDBotonLogin)))
        botonLogin.click()
        print(funciones.decB64("QWNjZXNvIENvbnNlZ3VpZG8="))
        return time.sleep(2)

    def cerrarInventario(self, sucursal):
        try:
            select_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "Sphinx_Sucursales")))
            select = Select(select_element)
            select.select_by_value(str(sucursal))
            time.sleep(2)
            botonCerrar = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//td[@id='inventarioAbierto']//input[@value='C']")))
            botonCerrar.click()
            time.sleep(2)
            try:
                alert = WebDriverWait(self.driver, 5).until(EC.alert_is_present())
                alert.accept()
                print("Aceptado!")
            except TimeoutException:
                print("No se detecto alerta")
                print(sucursal)
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Error al cerrar inventario {sucursal}")

    def reporteDiferencias(self, sucursal):
        try:
            select_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "Sphinx_Sucursales")))
            select = Select(select_element)
            select.select_by_value(str(sucursal))
            time.sleep(1)

            select_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, "inventario_tipo")))
            select = Select(select_element)
            select.select_by_value("2") # Productos con Diferencias
            time.sleep(1)

            botonEjecutar = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "btnEjecuta")))
            botonEjecutar.click()
            time.sleep(2)

            botonExcel = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "btnExcel")))
            botonExcel.click()
            print(sucursal)
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Error al descargar informe {sucursal}: {e}")
            return time.sleep(1)
  
    def quit(self):
        self.driver.quit()

class funciones:
    def __init__(self):
        pass

    def decB64(texto):
        return b6.b64decode(texto).decode('utf-8')
    
    def codec(w, cif=True):
        x , i = "" , 1
        for c in w:
            y = ord(c)
            if cif : nV = (y+i)%256
            else: nV = (y-i)%256
            i += 1
            nV = max(0, min(nV, 0x10FFFF))
            nC = chr(nV)
            x += nC
        return x
    
    def buscarArchivos(directorio,nombreInicial,tipoArchivo):
        listado = os.listdir(directorio)
        return [os.path.join(dir, archivo) for archivo in listado if archivo.endswith(tipoArchivo) and archivo.startswith(nombreInicial)]

    def borrarArchivos(directorio, listaDeArchivos):
        warnings.filterwarnings("ignore", category=UserWarning)
        for x in listaDeArchivos:
            ruta = os.path.join(directorio, x)
            try:
                os.remove(ruta)
                print(f"El archivo {x} ha sido eliminado.")
            except FileNotFoundError:
                print(f"No se encontró el archivo {x}.")
            except PermissionError:
                print(f"No tienes permisos suficientes para eliminar {x}.")
            except OSError as error:
                print(f"Ocurrió un error al eliminar el archivo {x}: {error}")
        return print("Archivos eliminados.")

    def ejecutarAsincrono(file):
        try:
            with Pool(processes=1) as pool:
                pool.apply_async(sub.run, ["python", file])
            messagebox.showinfo("Función ejecutada.")
        except FileNotFoundError:
            messagebox.showerror("Error", "No se encontró el archivo raíz")

    def leerCSV(documento):
        with open(documento, 'r') as csvfile:
            reader = csv.reader(csvfile)
            listado = {row[0]: row[1] for row in reader}
        print("lista creada")
        return listado

    def clear():
        if os.name == 'nt':   os.system('cls')
        else:                 os.system('clear')
        return None

    def creacionEntorno():
        if not os.path.exists(".env"):
            with open(".env", "w") as env_file:
                chrome = "%APPDATA%/Google/Chrome"
                env_file.write("USERNAME=\n")
                env_file.write("PASSWORD=\n")
                env_file.write(f"PERFIL_CHROME={chrome}")
                env_file.close()

                user = input("Favor ingrese su usuario: ")
                password = input("Favor ingrese su contraseña: ")
                set_key(".env", "USERNAME", user)
                set_key(".env", "PASSWORD", password)
                
        if not os.path.exists("Sucursales.csv"):
            with open("Sucursales.csv", "w", newline="") as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(["ID_sucursal,Nombre_sucursal"])
        return print("entorno creado!")

    def carpetaDescargas():
        carpeta = os.getcwd()
        return carpeta
    
class Excel:
    def __init__():
        pass

    def renombrarArchivos():
        # Renombrar todos los archivos "Inventario" y colocar el renombra como su sucursal
        dir =  funciones.carpetaDescargas()
        columna, hoja, sep = 0,'sphinx', '/'
        xlsxs = funciones.buscarArchivos(dir,"Inventario",".xlsx")
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
                    dir, extension = os.path.splitext(x)
                    nuevo_nombre = nueva_parte + extension
                    # Renombrar el archivo
                    os.rename(x, nuevo_nombre)
                    print(f"Archivo renombrado a: {nuevo_nombre}")
                else:
                    print(f"La fila 4 en la hoja '{hoja}' del archivo {x} está vacía o no existe.")
            except (FileNotFoundError, PermissionError, IndexError, ValueError) as e:
                print(f"Error al procesar el archivo {x}: {e}")
        funciones.clear()
        return print("Renombrado finalizado.") 

    def unificar():
        # Unificado de archivos en uno SOLO
        dir = funciones.carpetaDescargas()
        df_final = pd.DataFrame()
        archivos = funciones.buscarArchivos(dir,"Inventario",".xlsx") 
        for x in archivos:
            try:
                df = pd.read_excel(os.path.join(dir, x), header=5, usecols='A:E')
                df = df.dropna(how='all')          
                if df.empty:
                    df = pd.DataFrame(columns=['Codigo', 'Nombre', 'Marca', 'Stock', 'Cantidad'])
                    df.loc[0] = [0,0,0,0,0]
                nombre_archivo = os.path.basename(x).rsplit('.', 1)[0]
                df['Archivo'] = nombre_archivo
                df_final = pd.concat([df_final, df], ignore_index=True)
            except Exception as e:
                print(f"Error al procesar el archivo {x}: {e}")
        df_final.to_excel("unificados.xlsx", index=False)
        funciones.clear()
        return print("Unificación realizada.")

    def sobrestock():
        directorio = funciones.carpetaDescargas()
        archivos = funciones.buscarArchivos(directorio,"",".xlsx")
        ss = ["AAA SS NUEVOS.csv", "AAA SS VIGENTES.csv", "AAA SS ANTIGUOS.csv", "AAA ELIMINADOS.csv", "PROCESADO.xlsx"]
        funciones.borrarArchivos(directorio, ss)
        hojas = ["SOBRESTOCK - NUEVO" , "SOBRESTOCK VIGENTE ", "SOBRESTOCK ANTIGUO", "AAA ELIMINADOS.csv"]
        for archivo in archivos:
            for hoja,nuevo_archivo in zip(hojas,ss):
                if hoja == "SOBRESTOCK - NUEVO": saltar = 6 
                else: saltar = 7
                try:
                    df = pd.read_excel(archivo, 
                                    sheet_name = hoja, 
                                    usecols = "A:B", 
                                    skiprows = saltar, 
                                    na_values = [''])
                    df = df[[df.columns[1], df.columns[0]]]

                    df_limpio = df.dropna(how='all')

                    if not df_limpio.empty:
                        df_limpio = df_limpio.iloc[0:]
                        df_limpio.to_csv(os.path.join(directorio, nuevo_archivo), index=False, sep=';', header=None, float_format='%d')
                        print(f"Datos invertidos guardados en {nuevo_archivo}\n")
                    else:
                        print(f"No hay datos para guardar en {nuevo_archivo}\n")

                except (FileNotFoundError, PermissionError, ValueError) as e:
                    print(f"Error al procesar el archivo {archivo}: {e}")
                except (IndexError) as e:
                    print(f"Hoja vacia {hoja}")
        funciones.renombrarArchivos(directorio,"",".xlsx")
        return print("Proceso finalizado.")

if __name__ == "__main__":
    import importlib as ilib
    import subprocess as sub
    def libSetup(lib):
        # Funcion para instalar automaticamente librerias no existentes
        try:ilib.import_module(lib)
        except ImportError:sub.check_call(['pip', 'install', lib])
        return

    import os, time, csv, json
    import datetime as dt
    import base64 as b6
    from multiprocessing import Pool
    libSetup('tkinter')
    from tkinter import *
    from tkinter import messagebox
    libSetup('warnings')
    import warnings
    libSetup('getpass')
    import getpass as gp
    libSetup('pandas')
    import pandas as pd
    libSetup('python-dotenv')
    from dotenv import load_dotenv, set_key
    libSetup('selenium')
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
    from selenium.webdriver.common.alert import Alert

    ventanaPrincipal = VentanaPrincipal()
    funciones.creacionEntorno()
    ventanaPrincipal.iniciar()