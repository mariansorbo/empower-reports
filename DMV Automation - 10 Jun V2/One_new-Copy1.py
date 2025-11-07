#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pywinauto import Application, Desktop
import time
import os
import config_runtime
import config

usuario=config_runtime.usuario

# Ruta a Power BI Desktop
powerbi_exe_path = config.powerbi_exe_path
# Carpeta donde está el archivo
base_report_directory = config.base_report_directory
# Nombre del reporte .pbix
reporte = config_runtime.reporte
carpeta_destino = config.step1_extracciones_dmv


print("Analizando el reporte : "+ reporte + " | Usuario: "+ usuario)


# In[2]:


nombre_reporte = os.path.splitext(reporte)[0]  # Esto es más seguro que rstrip(".pbix")
queries = [
    {
        "name": "columnas",
        "sql": f"""SELECT '{nombre_reporte}' AS Reporte, [ID], [TableID], [ExplicitName], [IsHidden], [ColumnStorageID], [Type], [SourceColumn], [Expression], [ModifiedTime], [StructureModifiedTime], [DisplayFolder] 
                   FROM $SYSTEM.TMSCHEMA_COLUMNS ;"""
    },
    {
        "name": "metricas",
        "sql": f"""SELECT '{nombre_reporte}' AS Reporte, [ID], [TableID], [Name], [DataType], [Expression], [IsHidden], [ModifiedTime], [StructureModifiedTime], [DisplayFolder] 
                   FROM $SYSTEM.TMSCHEMA_MEASURES ;"""
    },
    {
        "name": "tablas",
        "sql": f"""SELECT '{nombre_reporte}' AS Reporte, [ID], [Name], [IsHidden], [TableStorageID], [ModifiedTime], [StructureModifiedTime] 
                   FROM $SYSTEM.TMSCHEMA_TABLES ;"""
    },
    {
        "name": "partitions",
        "sql": f"""SELECT '{nombre_reporte}' AS Reporte, * 
                   FROM $SYSTEM.TMSCHEMA_PARTITIONS ;"""
    },
    {
        "name": "relaciones",
        "sql": f"""SELECT '{nombre_reporte}' AS Reporte, [ID], [FromTableID], [FromColumnID], [FromCardinality], [ToTableID], [ToColumnID], [ToCardinality], [ModifiedTime] 
                   FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS ;"""
    }
]


# In[4]:


# —————— Construir y lanzar Power BI (robusto) ——————
from pathlib import Path

def find_powerbi_exe():
    # 1) Si la config apunta a un exe válido, úsalo
    p = Path(config.powerbi_exe_path)
    if p.is_file():
        return str(p)

    # 2) Rutas típicas “clásicas”
    candidates = [
        r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe",
        r"C:\Program Files\Microsoft Power BI Desktop RS\bin\PBIDesktop.exe",
    ]
    # 3) Versión Microsoft Store (WindowsApps)
    wa = Path(r"C:\Program Files\WindowsApps")
    if wa.exists():
        try:
            for d in wa.iterdir():
                name = d.name.lower()
                if name.startswith("microsoft.microsoftpowerbidesktop") and name.endswith("_8wekyb3d8bbwe"):
                    exe = d / "bin" / "PBIDesktop.exe"
                    if exe.exists():
                        candidates.append(str(exe))
        except PermissionError:
            pass  # a veces no deja listar todo; seguimos con lo que tengamos

    for c in candidates:
        if Path(c).is_file():
            return c
    return None

safe_print("🟡 Lanzando Power BI...")

move_pointer_to_center()  # evita fail-safe

pbi_exe = find_powerbi_exe()
try:
    if pbi_exe:
        command_line = f'"{pbi_exe}" "{pbit_file_path}"'
        app = Application(backend="uia").start(command_line)
        safe_print(f"🟢 Power BI lanzado (EXE): {pbi_exe}")
    else:
        # Fallback: abrir el .pbit con la app asociada (Power BI)
        os.startfile(pbit_file_path)   # usa la asociación del sistema
        safe_print("🟢 Power BI lanzado por asociación de archivo (.pbit).")
except Exception as e:
    safe_print(f"❌ Error al lanzar Power BI: {e}")
    sys.exit(2)

# —————— Esperar ventana y maximizar (hasta 120s) ——————
win = wait_powerbi_window(timeout_sec=120)
if not win:
    safe_print("⚠️ No se encontró la ventana principal de Power BI tras 120s. Abortando para evitar clicks a ciegas.")
    sys.exit(2)

time.sleep(1.0)



# In[4]:


DAX_PATH = config.DAX_PATH

if not os.path.exists(DAX_PATH):
    raise FileNotFoundError("Ruta incorrecta a DAX Studio")

# 🚀 Abrir DAX Studio
app = Application(backend="uia").start(DAX_PATH)
print("[✓] DAX Studio lanzado, esperando ventana...")


# 🕒 Esperar a que aparezca una ventana con título que contenga 'DAX Studio'
window = None
for _ in range(20):
    time.sleep(2)
    windows = Desktop(backend="uia").windows()
    for w in windows:
        if "DAX Studio" in w.window_text():
            window = w
            break
    if window:
        break

if not window:
    raise Exception("No se encontró la ventana de DAX Studio")

# 🪟 Asegurar que esté visible y usable
window.set_focus()
print(f"[✓] Ventana detectada: '{window.window_text()}'")
#win.child_window(title="Connect", auto_id="Connect", control_type="Button").click_input()



# In[ ]:


# 🔍 Buscar el control que contenga 'DaxStudio.UI.ViewModels.AutoSaveRecoveryDialogViewModel' en su representación
recovery_window = None

children = window.descendants()
for i, c in enumerate(children):
    if "DaxStudio.UI.ViewModels.AutoSaveRecoveryDialogViewModel" in repr(c):
        recovery_window = c
        print(f"[✓] Ventana encontrada en índice {i}: '{c.window_text()}'")
        break

# Si no se encontró la ventana, continuar sin romper
if not recovery_window:
    print("ℹ️ No se encontró la ventana de recuperación. Continuando sin listar controles ni presionar botones.")
else:
    try:
        # 🪟 Foco y listado de controles
        recovery_window.set_focus()
        print(f"[✓] Conectado a ventana de recuperación: '{recovery_window.window_text()}'")

        print("🔍 Controles en la ventana de recuperación:")
        recovery_children = recovery_window.descendants()
        for i, c in enumerate(recovery_children):
            try:
                print(f"[{i}] {c.element_info.control_type} | '{c.window_text()}' | {c.element_info.class_name}")
            except Exception as e:
                print(f"[{i}] Error al acceder al control: {e}")

        # 🔎 Buscar y hacer clic en el botón "Cancel"
        cancel_button = None
        for ctrl in recovery_children:
            if ctrl.element_info.control_type == "Button" and ctrl.window_text() == "Cancel":
                cancel_button = ctrl
                break

        if cancel_button:
            cancel_button.click_input()
            print("🛑 Botón 'Cancel' presionado con éxito.")
        else:
            print("⚠️ No se encontró el botón 'Cancel' en la ventana de recuperación.")

    except Exception as e:
        print(f"❌ Error durante el manejo de la ventana de recuperación: {e}")



import pyautogui
import time
from pywinauto import Desktop

# 🪟 Esperar y conectar con ventana principal de DAX Studio
window = None
for _ in range(20):
    time.sleep(2)
    windows = Desktop(backend="uia").windows()
    for w in windows:
        if "DAX Studio" in w.window_text():
            window = w
            break
    if window:
        break

if not window:
    raise Exception("❌ No se encontró la ventana de DAX Studio")

window.set_focus()
print(f"[✓] Ventana principal detectada: '{window.window_text()}'")

# 🔍 Buscar ventana de conexión con radio buttons
print("[...] Buscando ventana de conexión dentro de DAX Studio...")
window.maximize()
dialog_found = None
for dlg in Desktop(backend="uia").windows():
    try:
        radios = dlg.descendants(control_type="RadioButton")
        if radios:
            dialog_found = dlg
            break
    except:
        continue

if not dialog_found:
    pyautogui.click(1895, 1615)  # Fallback visual si quedó algo colgado
    time.sleep(0.3)
    raise Exception("❌ No se encontró el diálogo de conexión (radio buttons no detectados).")

# 🟢 Seleccionar primer radio button (ej: 'Power BI / SSDT Model')
radios = dialog_found.descendants(control_type="RadioButton")
radios[0].select()
print(f"[✓] Radio button seleccionado: '{radios[0].window_text()}'")
time.sleep(1)

# 🔵 Buscar y hacer click en botón "Connect"
connect_button = None
for btn in dialog_found.descendants(control_type="Button"):
    if btn.window_text().strip().lower() == "connect":
        connect_button = btn
        break

if connect_button:
    connect_button.click_input()
    print("[✓] Click en botón 'Connect' exitoso.")
else:
    raise Exception("❌ Botón 'Connect' no encontrado.")


# In[ ]:


time.sleep(5)
#------------------------------------------------------------------
import pyautogui

# 1. Click en botón DMV
pyautogui.click(169, 164)
time.sleep(0.3)

# 2. Click en menú desplegable "Results"
pyautogui.click(320, 112)
time.sleep(0.3)

# 3. Click en opción "Static" (Excel)
pyautogui.click(385, 371)
time.sleep(0.4)
pyautogui.click(392, 365)
time.sleep(0.3)

# In[ ]:


import pyperclip
import pyautogui
import time



for query in queries:
    print(f"\n🔁 Ejecutando query: {query['name']}")

    # 1. Click en editor de DAX
    pyautogui.click(1232, 278)  # O usá el control [161] si estás con pywinauto
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')

    # 2. Pegar el query
    pyperclip.copy(query['sql'])
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.3)

    # 3. Ejecutar query
    pyautogui.click(32, 67)  # Botón "Run"
    print("[✓] Query ejecutado.")
    time.sleep(1.5)

    # 4. Manejar ventana 'Save As'
    print("[📁] Esperando ventana 'Save As'...")
    save_as_window = None

    for child in window.descendants():
        if child.element_info.control_type == "Window" and "Save As" in child.window_text():
            save_as_window = child
            break

    if not save_as_window:
        print("❌ No se abrió la ventana 'Save As'.")
        continue

    try:
        save_as_window.set_focus()
        print(f"[✓] Ventana 'Save As' detectada.")




        # 4.1 Escribir nombre del archivo (usando pyperclip para evitar errores con '+')
        nombreArchivo = f"{usuario}+{nombre_reporte}+{query['name']}.xlsx"
        pyperclip.copy(nombreArchivo)
        #name_input = save_as_window.descendants()[113]
        #name_input.set_focus()
        #name_input.type_keys("^a{BACKSPACE}", with_spaces=True)
        pyautogui.hotkey("ctrl", "v")
        print(f"[✓] Nombre de archivo: {nombreArchivo}")

        # 4.2 Cambiar carpeta destino
        pyperclip.copy(carpeta_destino)
        pyautogui.hotkey("ctrl", "l")
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("enter")
        time.sleep(1.2)
        print(f"[✓] Carpeta destino: {carpeta_destino}")

        
        # 4.4 Click en botón Save
        pyautogui.click(751, 604)  # O usá el control [161] si estás con pywinauto
        #save_button = save_as_window.descendants()[123]
        #save_button.click_input()
        #print("[✓] Click en 'Save' exitoso.")
        time.sleep(2)

        # 4.5 Confirmar reemplazo si aparece
        updated_children = window.descendants()
        yes_button = next(
            (c for c in updated_children if c.element_info.control_type == "Button" and c.window_text().strip().lower() == "yes"),
            None
        )
        if yes_button:
            yes_button.click_input()
            print("[↪] Confirmación de reemplazo aceptada.")

        # 5. Verificar si se cerró correctamente
        time.sleep(1)
        final_check = any(
            "Save As" in c.window_text() and c.element_info.control_type == "Window"
            for c in window.descendants()
        )
        if not final_check:
            print(f"[✅] Archivo '{nombreArchivo}' guardado correctamente.")
        else:
            print(f"[❌] Falló el guardado de '{nombreArchivo}'.")

    except Exception as e:
        print(f"❌ Error durante manejo de 'Save As': {e}")


# In[5]:


# In[ ]:


from pywinauto import Application, Desktop
import pyautogui
import time

try:
    # Conectarse a Power BI si está abierto
    app = Application(backend="uia").connect(path="PBIDesktop.exe")

    # Buscar ventana principal
    main_window = app.top_window()
    main_window.set_focus()

    # Intentar cerrar (2 veces si es necesario)
    main_window.close()
    time.sleep(0.5)
    main_window.close()
    print("[✓] Power BI cerrado correctamente.")

    # Esperar a que aparezca el diálogo de guardado
    time.sleep(2)

    # Click en "Don't Save" - ajustar coordenadas si hiciera falta
    pyautogui.click(740, 412)
    print("[↪] Click en 'Don't Save'")
    time.sleep(0.8)

    # ventana de recuperación, click en "si, quiero verlos mas adelante"
    pyautogui.click(471, 371)
    print("[↪] Click en 'Cancel' de recuperación")
    time.sleep(0.1)
    # Click en "Cancel" de la ventana de recuperación,
    pyautogui.click(789, 436)
    print("[↪] Click en 'Cancel' de recuperación")
    time.sleep(0.3)

#X=740, Y=412
#🖱️ Click detectado en: X=471, Y=371
#🖱️ Click detectado en: X=789, Y=436
except Exception as e:
    print(f"❌ Error durante el cierre de Power BI: {e}")



# In[35]:





# In[37]:


print(windows.descendants())


# In[ ]:




