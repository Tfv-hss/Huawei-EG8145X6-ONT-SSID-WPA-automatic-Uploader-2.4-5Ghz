import os
import threading
import time
import tkinter as tk
from tkinter import ttk
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium import webdriver
import pandas as pd
from tkinter import messagebox
import sys
import io
import msoffcrypto


ONT_IP = "http://192.168.1.1"  ######### Cambiar IP de la ONT/Router aquí #########
USERNAME = "   "  ######### En este campo se añade un usuario fijo para todos los routers/ONTs ############
PASSWORD = "   "  ######### En este campo se añade una contraseña fija para todos los routers/ONTs #########

def start_process(browser, status_callback, location, ssid=None, password=None, band="2.4G"):


    try:
        
        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        xlsx_file = f"Prueba{location}.xlsx"  ######## Aqui donde estan las XXXXX se cambian a una parte fija que tengan todos los archivos de excel, por ejemplo: "CuentasDe"
        xlsx_path = os.path.join(base_dir, xlsx_file)


        if not os.path.exists(xlsx_path):
            status_callback(f"Archivo Excel no encontrado: {xlsx_file}")
            return
        else:
            status_callback(f"Archivo Excel detectado: {xlsx_file}")


        status_callback("Iniciando navegador...")

        if browser == "Chrome":
            driver = webdriver.Chrome()
        elif browser == "Edge":
            service = EdgeService()
            options = webdriver.EdgeOptions()
            options.add_argument("--start-maximized")
            driver = webdriver.Edge(service=service, options=options)
        elif browser == "Firefox":
            service = FirefoxService()
            options = webdriver.FirefoxOptions()
            options.add_argument("--start-maximized")
            driver = webdriver.Firefox(service=service, options=options)
        else:
            status_callback("❌ Navegador no soportado.")
            return

        wait = WebDriverWait(driver, 15)
        driver.get(ONT_IP)

        status_callback("Esperando carga de la página...")
        time.sleep(0)
        wait.until(EC.presence_of_element_located((By.ID, "txt_Username")))

        #Login vía JavaScript
        script = '''
            let u = document.getElementById("txt_Username");
            let p = document.getElementById("txt_Password");
            u.focus(); u.value = arguments[0];
            u.dispatchEvent(new Event("input", { bubbles: true }));
            p.focus(); p.value = arguments[1];
            p.dispatchEvent(new Event("input", { bubbles: true }));
            setTimeout(() => { LoginSubmit("loginbutton"); }, 500);
        '''
        driver.execute_script(script, USERNAME, PASSWORD)
        status_callback("✅ Login enviado por JS con LoginSubmit()")

        time.sleep(0)

        try:
            short_wait = WebDriverWait(driver, 3)
            exit_btn = short_wait.until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='Exit']"))
            )
            exit_btn.click()
            status_callback("Asistente inicial detectado y cerrado con 'Exit'")
            time.sleep(0)
        except:
            status_callback("✅ No se detectó asistente inicial, continuando rápido...")

        #Esperar a que la página cargue completamente
        status_callback("Navegando al menú de carga...")
        advanced_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[normalize-space()='Advanced']")))
        advanced_btn.click()
        time.sleep(0)

         # Click en el menú principal "WLAN"
        wlan_menu = wait.until(EC.element_to_be_clickable((By.ID, "name_wlanconfig")))
        wlan_menu.click()
        time.sleep(1)

        # Click en el submenú "2.4G Basic Network Settings"
        # Click dinámico en el submenú según banda
        band_id = "wlan2basic" if band == "2.4G" else "wlan5basic"
        band_label = "2.4G" if band == "2.4G" else "5G"

        wlan_basic = wait.until(EC.element_to_be_clickable((By.ID, band_id)))
        wlan_basic.click()
        status_callback(f"📡 Página de configuración {band_label} cargada.")

        try:
            driver.switch_to.frame(driver.find_element(By.ID, "menuIframe"))
        except:
            status_callback("⚠️ No se encontró iframe 'menuIframe', continuando sin cambiar de frame.")

    except Exception as e:
        status_callback(f"❌ Error: {str(e)}")

     # Reemplazar SSID y WPA Key automáticamente
    ssid_input = wait.until(EC.presence_of_element_located((By.ID, "wlSsid")))
    ssid_input.clear()
    ssid_input.send_keys(ssid)

 

        # ✅ Mostrar campo de texto y escribir la contraseña correctamente
    try:
        hide_checkbox = driver.find_element(By.ID, "hideWpaPsk")
        if hide_checkbox.is_selected():
            hide_checkbox.click()
            time.sleep(0.5)  # Dejar que se actualice el campo visible
    except:
        status_callback("⚠️ No se encontró checkbox para 'Hide', continuando igual.")

    password_input = wait.until(EC.element_to_be_clickable((By.ID, "wlWpaPsk")))
    driver.execute_script("arguments[0].value = '';", password_input)  # Limpieza robusta
    password_input.send_keys(password)

    status_callback(f"✅ SSID y contraseña escritos: SSID={ssid}, WPA={password.replace(chr(10), '').replace(chr(13), '')}")
    # 🔔 Mostrar popup para que el técnico verifique visualmente los valores 
    messagebox.showinfo(
    "Verifica antes de aplicar",
    f"Verifica que los siguientes valores sean correctos: \nSSID: {ssid} \nContraseña: {password} \n\nPulsa 'Apply' en la ONT si todo está bien."
)






##################################### INTERFAZ GRAFICA NO TOCAR ######################################

class ONTUploaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ONT SSID/WPA automatic Uploader 2.4 - 5 Ghz / by Tfv-hss")
        self.root.geometry("600x400")
        self.root.minsize(500, 300)
        root.resizable(False, False)


        for i in range(4):  # Expandir todas las columnas
         self.root.columnconfigure(i, weight=1)

        self.root.rowconfigure(3, weight=1)  # Expandir fila donde está la consola


        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)

        self.browser = tk.StringVar(value="Chrome")
        self.location = tk.StringVar(value=" ")
        self.band = tk.StringVar(value="2.4G")



        self.label_browser = tk.Label(root, text="Navegador:")
        self.label_browser.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))

        self.combo_browser = ttk.Combobox(
            root,
            textvariable=self.browser,
            values=["Chrome", "Edge", "Firefox"],
            state="readonly"
        )
        self.combo_browser.grid(row=1, column=0, padx=10, sticky="ew")

        self.label_location = tk.Label(root, text="Nº de excel:")
        self.label_location.grid(row=0, column=1, sticky="w", padx=10, pady=(10, 0))

        self.combo_location = ttk.Combobox(
           root,
           textvariable=self.location,
           values=[" 1", "2", "3",], ### Aqui añadiremos valores que se completaran con las XXXXXX que hemos cambiado antes
           state="readonly"
        )
        self.combo_location.grid(row=1, column=1, padx=10, sticky="ew")

        self.label_band = tk.Label(root, text="Banda WiFi:")
        self.label_band.grid(row=0, column=3, sticky="w", padx=10, pady=(10, 0))

        self.combo_band = ttk.Combobox(
        root,
        textvariable=self.band,
        values=["2.4G", "5G"],
        state="readonly"
   )
        self.combo_band.grid(row=1, column=3, padx=10, sticky="ew")


        self.label_codcliente = tk.Label(root, text="Código del Excel:")
        self.label_codcliente.grid(row=0, column=2, sticky="w", padx=10, pady=(10, 0))

        self.codcliente = tk.StringVar()
        self.entry_codcliente = tk.Entry(root, textvariable=self.codcliente)
        self.entry_codcliente.grid(row=1, column=2, padx=10, sticky="ew")

        self.btn_upload = tk.Button(
            root,
            text="● INICIAR",
            bg="green",
            fg="white",
            font=("Arial", 14, "bold"),
            command=self.validate_and_start
        )
        self.btn_upload.grid(row=2, column=0, pady=10)

        self.text_status = tk.Text(root, wrap="none")
        self.text_status.grid(row=3, column=0, columnspan=4, padx=(10,0), pady=10, sticky="nsew")

        scroll = tk.Scrollbar(root, command=self.text_status.yview)
        scroll.grid(row=3, column=4, sticky="ns", padx=(0,10), pady=10)
        self.text_status.config(yscrollcommand=scroll.set)

    def update_status(self, message):
        self.text_status.insert(tk.END, message + "\n")
        self.text_status.see(tk.END)

    def run_upload(self, ssid=None, password=None):
     threading.Thread(
        target=start_process,
        args=(self.browser.get(), self.update_status, self.location.get(), ssid, password, self.band.get()),
        daemon=True
    ).start()


    def validate_and_start(self):
        codcliente_input = self.codcliente.get().strip()
        if not codcliente_input:
         messagebox.showerror("Error", "Debes introducir un código de Excel antes de continuar.")
         return

        location = self.location.get()
        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        xlsx_path = os.path.join(base_dir, f"Prueba{location}.xlsx")


        if not os.path.exists(xlsx_path):
            self.update_status(f"❌ Archivo no encontrado: {xlsx_path}")
            return

       
        try:
            decrypted = io.BytesIO()
            with open(xlsx_path, "rb") as f:
              office_file = msoffcrypto.OfficeFile(f)
              office_file.load_key(password="1234") ################## ESTA CONTRASEÑA DEBE SER CAMBIADA POR LA NUESTRA ACTUAL##############

              office_file.decrypt(decrypted)

            df = pd.read_excel(decrypted, engine="openpyxl", dtype=str).fillna("")

            coincidencias = df[df.iloc[:, 1].apply(lambda x: str(x).strip()) == codcliente_input.strip()]


            if coincidencias.empty:
                messagebox.showerror("Error", f"❌ Código de cliente '{codcliente_input}' no encontrado en el archivo.")
                return

            row = coincidencias.iloc[0]



################################################ AQUI SE CAMBIA DE DONDE SE VAN A SELECCIONAR LOS VALORES RECOGIDOS DEL EXCEL #####################################################

            dir = row[1]   # Columna A
            nombre = row[2]  # Columna B
            apellidos = row[3]  # Columna C
            SSID = row[4]  # Columna D
            Contrasena = row[5]  # Columna E

####################################################################################################################################################################################
            confirm = messagebox.askyesno(
                "Confirmar datos",
                f"Casilla A: {dir} \nCasilla B: {nombre}\nCasilla C: {apellidos} \nCasilla D:  {SSID}\nCasilla E: {Contrasena}    \n\nDatos encontrados en el Excel.\n\n¿Deseas continuar con la carga?"
            )
##################################################### Esta parte del codigo muestra los valores recogidos en sus casillas para mostrar que se han encontrado correctamente #################
            if confirm:
               self.run_upload(SSID, Contrasena)
            else:
                self.update_status("❌ Operación cancelada por el usuario.")

        except Exception as e:
            messagebox.showerror("Error", f"⚠️ Error leyendo el Excel:\n{e}")




if __name__ == "__main__":
    root = tk.Tk()
    app = ONTUploaderGUI(root)
    root.mainloop()
