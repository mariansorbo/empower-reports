#!/usr/bin/env python
# coding: utf-8

# # 📘 Documentación Técnica – Publicación Automatizada de Reporte en Power BI
# 🧠 Objetivo del Script
# Este script automatiza completamente el proceso de publicación de un archivo .pbix en Power BI Service, partiendo desde su apertura local con Power BI Desktop, refrescando los datos y finalizando con la publicación y la copia del link público del reporte publicado.
# 
# 📁 Estructura y Funcionamiento
# 1. Inicialización y Apertura
# Se utiliza pyautogui para emular clicks y escritura.
# 
# El script inicia haciendo clic en la barra de búsqueda de Windows y escribiendo "power" para lanzar Power BI Desktop.
# 
# Luego, selecciona y abre la Plantilla Empower BI, que se espera esté visible en la lista de recientes de Power BI.
# 
# 2. Publicación del Reporte
# Una vez abierta la plantilla, el script:
# 
# Hace clic en el botón "Publicar".
# 
# Interactúa con el menú desplegable del workspace.
# 
# Selecciona el workspace deseado.
# 
# Confirma la publicación.
# 
# 📍 Nota técnica: El proceso se basa en coordenadas pregrabadas. Requiere una resolución de pantalla estable y sin interrupciones.
# 
# -------------------------------------------------
# 
# 🌐 Recuperación del Link de Reporte
# El script continúa:
# 
# Abriendo el reporte publicado (desde el menú emergente).
# 
# Haciendo clic en la barra de dirección del navegador.
# 
# Copiando la URL pública usando el portapapeles.
# 
# Imprimiendo ese link como salida para el usuario.
# 
# -------------------------------------------------
# 
# 🗂️ Origen de los Datos
# Los datos que alimentan la plantilla:
# 
# Fueron previamente refrescados y guardados en OneDrive y Google Drive por un script anterior.
# 
# La plantilla de Power BI está configurada para apuntar a ese archivo .xlsx vía OneDrive, permitiendo su actualización automática al abrir.
# 
# -------------------------------------------------
# 
# 
# 🛠️ Dependencias y Herramientas
# pyautogui: Para automatizar clicks, escritura y movimientos.
# 
# time: Para controlar los delays entre pasos (espera de carga).
# 
# pynput.mouse: Para capturar coordenadas de click en sesiones de entrenamiento manual del script.
# 
# -------------------------------------------------
# 
# 
# 🧩 Requisitos Previos
# El archivo .pbix debe estar ubicado en:
# C:\Users\Administrator\OneDrive\Empower BI Archivos\Plantilla Empower BI.pbix
# 
# El archivo Excel con los datos debe haberse generado y subido correctamente antes de ejecutar este script.
# 
# Power BI Desktop debe estar instalado, accesible y su interfaz debe mantenerse estable durante el proceso.
# 
# 

# In[1]:

import pyautogui
from pywinauto.application import Application
import time
import config
import config_runtime


ruta_base = config.step2_archivos_consolidados
usuario = config_runtime.usuario

# Rutas desde config
powerbi_exe_path = config.powerbi_exe_path
pbix_file_path = config.plantilla
command_line = f'"{powerbi_exe_path}" "{pbix_file_path}"'

# Lanzar Power BI
app = Application(backend="uia").start(command_line)
time.sleep(20)  # Dar tiempo suficiente para abrir (ajustar si es necesario)

# Conectarse a la ventana principal
main_window = app.window(title_re=".*Plantilla Empower BI.*")
main_window.wait("visible", timeout=30)
main_window.maximize()

# Recorrer los controles de forma segura
print("Controles encontrados:")
try:
    children = main_window.descendants()
    for i, c in enumerate(children):
        try:
            print(f"[{i}] {c.element_info.control_type} | '{c.window_text()}' | {c.element_info.class_name}")
        except Exception as e:
            print(f"[{i}] Error al acceder al control: {e}")
except Exception as e:
    print(f"Error al obtener los controles: {e}")


# In[2]:


from pywinauto import Application
import time
import config

# 1. Conectarse a Power BI (puede requerir ajuste de título)
app = Application(backend="uia").connect(path= config.powerbi_exe_path)

# 2. Listar todas las ventanas abiertas de esa app
windows = app.windows()

for i, win in enumerate(windows):
    print(f"\n[{i}] Title: {win.window_text()} | Class: {win.element_info.class_name}")


# In[3]:


# Conectarse a la ventana principal
main_window = app.window(title_re=".*Plantilla Empower BI.*")
main_window.wait("visible", timeout=30)

# Recorrer los controles de forma segura
print("Controles encontrados:")
try:
    children = main_window.descendants()
    for i, c in enumerate(children):
        try:
            print(f"[{i}] {c.element_info.control_type} | '{c.window_text()}' | {c.element_info.class_name}")
        except Exception as e:
            print(f"[{i}] Error al acceder al control: {e}")
except Exception as e:
    print(f"Error al obtener los controles: {e}")


# In[4]:


time.sleep(1)
# Buscar todos los botones "Transform data"
buttons = main_window.descendants(control_type="Button")
transform_buttons = [b for b in buttons if b.window_text() == "Transform data"]

print(f"\nEncontrados {len(transform_buttons)} botones con título 'Transform data':\n")
for i, btn in enumerate(transform_buttons):
    print(f"[{i}] Class: {btn.element_info.class_name}, Rect: {btn.rectangle()}, Handle: {btn.handle}")

# Elegir uno (probá el 0 primero, o inspeccioná según posición visual)
btn_index = 0
btn = transform_buttons[btn_index]
btn.click_input()
time.sleep(1)  # esperar que se abra el menú

# Intentar encontrar y clickear "Data source settings"
menu_items = main_window.descendants(control_type="MenuItem")
data_source = next((m for m in menu_items if m.window_text() == "Data source settings"), None)

if data_source:
    data_source.click_input()
    print("✅ 'Data source settings' clickeado.")
else:
    print("❌ No se encontró 'Data source settings'.")


# In[5]:


# Enumerar todas las ventanas abiertas de Power BI
all_windows = app.windows()

for i, w in enumerate(all_windows):
    print(f"[{i}] Title: '{w.window_text()}' | Class: {w.element_info.class_name}")


# In[6]:


from pywinauto.controls.uiawrapper import UIAWrapper

popup_candidates = []
for i, ctrl in enumerate(main_window.descendants()):
    try:
        if ctrl.element_info.control_type in ["Window", "Pane", "Dialog"]:
            print(f"[{i}] {ctrl.element_info.control_type} | '{ctrl.window_text()}' | {ctrl.element_info.class_name}")
            popup_candidates.append(ctrl)
    except Exception as e:
        print(f"[{i}] Error: {e}")


# In[7]:


ds_window = popup_candidates[0]
ds_window


# In[8]:


print("Controles en la ventana 'Data source settings':")
for i, ctrl in enumerate(ds_window.descendants()):
    try:
        print(f"[{i}] {ctrl.element_info.control_type} | '{ctrl.window_text()}' | {ctrl.element_info.class_name}")
    except Exception as e:
        print(f"[{i}] Error: {e}")


# In[9]:


# Buscar entre los hijos del ds_window el botón con ese texto
time.sleep(2)
change_source_btn = next(
    (ctrl for ctrl in ds_window.descendants()
     if ctrl.element_info.control_type == "Button" and ctrl.window_text() == "Change Source..."),
    None
)

if change_source_btn:
    change_source_btn.click_input()
    print("✅ Click en 'Change Source...' realizado.")
else:
    print("❌ No se encontró el botón 'Change Source...'.")


# In[10]:


from pywinauto.controls.uiawrapper import UIAWrapper

popup_candidates = []
for i, ctrl in enumerate(main_window.descendants()):
    try:
        if ctrl.element_info.control_type in ["Window", "Pane", "Dialog"]:
            print(f"[{i}] {ctrl.element_info.control_type} | '{ctrl.window_text()}' | {ctrl.element_info.class_name}")
            popup_candidates.append(ctrl)
    except Exception as e:
        print(f"[{i}] Error: {e}")


# In[11]:


# Conectar a ventana "EXCEL WORKBOOK"
excel_popup = popup_candidates[1]
time.sleep(3.5)
# Listar todos sus controles
print("Controles en la ventana 'Excel Workbook':")
for i, ctrl in enumerate(excel_popup.descendants()):
    try:
        print(f"[{i}] {ctrl.element_info.control_type} | '{ctrl.window_text()}' | {ctrl.element_info.class_name}")
    except Exception as e:
        print(f"[{i}] Error al acceder al control: {e}")


# In[12]:
import pyautogui
import pyperclip

# Asegurarse que haya barra final si no la tiene
if not ruta_base.endswith("\\"):
    ruta_base += "\\"

# Ruta final compuesta
ruta_completa = ruta_base + f"{usuario} - Consolidado.xlsx"
pyautogui.click(372, 565)
time.sleep(2)

pyautogui.click(534, 356)
pyperclip.copy(ruta_completa)
pyautogui.hotkey('ctrl', 'v')
time.sleep(2)

#pyautogui.click(878, 454)


# 2. Obtener el campo de texto y setear la nueva ruta
edit_field = excel_popup.descendants()[16]
#edit_field.set_edit_text(ruta_completa)
print(f"✅ Ruta escrita: {ruta_completa}")

# 3. Click en el botón OK
ok_button = excel_popup.descendants()[22]
ok_button.click_input()
print("✅ Click en botón 'OK'.")
time.sleep(3)


# In[13]:


# Asegurate de que ds_window siga siendo accesible (ventana "Data source settings")
close_btn = ds_window.descendants()[35]
close_btn.click_input()
print("✅ Ventana 'Data source settings' cerrada.")
time.sleep(3)


# In[14]:


from pywinauto import Application, Desktop
import time

# Conectarse a Power BI si está abierto
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

# Buscar el botón "Refresh"
refresh_btn = None

for ctrl in main_window.descendants():
    if (
        ctrl.element_info.control_type == "Button"
        and ctrl.window_text().strip().lower() == "refresh"
    ):
        # Opcional: verificar que el class_name sea el que esperás
        if "ms-button" in ctrl.element_info.class_name.lower():
            refresh_btn = ctrl
            break

if refresh_btn:
    refresh_btn.click_input()
    print("✅ Botón 'Refresh' clickeado correctamente.")
else:
    print("❌ No se encontró el botón 'Refresh'.")

time.sleep(10)


# In[15]:


from pywinauto import Application, Desktop
import time

# Conectar a Power BI
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

# Buscar y hacer clic en el tab 'File'
file_tab = None

for ctrl in main_window.descendants():
    if (
        ctrl.element_info.control_type == "TabItem"
        and ctrl.window_text().strip().lower() == "file"
        and "ms-button" in ctrl.element_info.class_name.lower()
    ):
        file_tab = ctrl
        break

if file_tab:
    file_tab.click_input()
    print("✅ Tab 'File' clickeado.")
else:
    print("❌ No se encontró el tab 'File'.")
time.sleep(0.3)


# In[16]:


from pywinauto import Application, Desktop
import time

# Conectarse a Power BI
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

# Buscar el TabItem 'Save as'
save_as_tab = None

for ctrl in main_window.descendants():
    if (
        ctrl.element_info.control_type == "TabItem"
        and ctrl.window_text().strip().lower() == "save as"
        and "tabheader" in ctrl.element_info.class_name.lower()
    ):
        save_as_tab = ctrl
        break

if save_as_tab:
    save_as_tab.click_input()
    print("✅ Tab 'Save as' clickeado correctamente.")
else:
    print("❌ No se encontró el tab 'Save as'.")
time.sleep(0.3)


# In[17]:


from pywinauto import Application, Desktop
import time

# Conectar a Power BI
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

# Esperar a que se abra el panel luego de "Save as"
time.sleep(1.5)

# Buscar botón 'Browse this device'
browse_button = None

for ctrl in main_window.descendants():
    if (
        ctrl.element_info.control_type == "Button"
        and ctrl.window_text().strip().lower() == "browse this device"
        and "option" in ctrl.element_info.class_name.lower()
    ):
        browse_button = ctrl
        break

if browse_button:
    browse_button.click_input()
    print("✅ Botón 'Browse this device' clickeado correctamente.")
else:
    print("❌ No se encontró el botón 'Browse this device'.")

time.sleep(0.3)


# In[18]:


import pyautogui
import pyperclip
import time
from pywinauto import Application, Desktop
import config
import config_runtime

usuario = config_runtime.usuario
carpeta_destino = config.step4_entregables

# Conectarse a Power BI
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

# Esperar que aparezca ventana 'Save As'
print("[📁] Esperando ventana 'Save As' (como hijo)...")
save_as_window = None

# Reintentar por hasta 5 segundos
for _ in range(10):
    for child in main_window.children():
        if (
            child.element_info.control_type == "Window"
            and child.window_text().strip().lower() == "save as"
            and child.element_info.class_name == "#32770"
        ):
            save_as_window = app.window(handle=child.handle)  # 👈 Esta línea lo soluciona
            break
    if save_as_window:
        break
    time.sleep(0.5)

if not save_as_window:
    print("❌ No se abrió la ventana 'Save As'.")
else:
    try:
        save_as_window.set_focus()
        print("[✓] Ventana 'Save As' detectada.")

        # 🔤 1. Escribir nombre del archivo
        nombre_archivo = f"{usuario} - Empower BI"
        pyperclip.copy(nombre_archivo)
        name_input = save_as_window.child_window(auto_id="1001", control_type="Edit")
        name_input.set_focus()
        name_input.type_keys("^a{BACKSPACE}", with_spaces=True)
        pyautogui.hotkey("ctrl", "v")
        print(f"[✓] Nombre de archivo seteado: {nombre_archivo}")

        # 📂 2. Cambiar carpeta destino
        pyperclip.copy(carpeta_destino)
        pyautogui.hotkey("ctrl", "l")  # activar path input
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("enter")
        time.sleep(1.2)
        print(f"[✓] Carpeta destino seteada: {carpeta_destino}")

        # 🔁 3. Refocus a ventana 'Save As' por si redibuja
        # 🔁 3. Refocus a ventana 'Save As' por si redibuja
        for child in main_window.children():
            if (
                child.window_text().strip().lower() == "save as"
                and child.element_info.control_type == "Window"
                and child.element_info.class_name == "#32770"
            ):
                save_as_window = app.window(handle=child.handle)
                save_as_window.set_focus()
                break

        # 💾 4. Click en botón 'Save'
        save_button = save_as_window.child_window(title="Save", control_type="Button")
        save_button.click_input()
        print("[✓] Click en 'Save' exitoso.")
        time.sleep(2)

        # 🔁 5. Confirmar reemplazo si aparece
        for c in main_window.descendants():
            if c.window_text().strip().lower() == "yes" and c.element_info.control_type == "Button":
                c.click_input()
                print("[↪] Confirmación de reemplazo aceptada.")
                break

        # ✅ Verificación final
        time.sleep(1)
        still_open = any(
            child.window_text().strip().lower() == "save as" and child.element_info.control_type == "Window"
            for child in main_window.children()
        )
        if not still_open:
            print(f"[✅] Archivo '{nombre_archivo}' guardado correctamente.")
        else:
            print(f"[❌] Falló el guardado de '{nombre_archivo}'.")

    except Exception as e:
        print(f"❌ Error al manejar 'Save As': {e}")
time.sleep(0.3)


# In[19]:


from pywinauto import Application, Desktop
import time

# Conectarse a Power BI si está abierto
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

# Buscar el botón "Publish"
publish_btn = None

for ctrl in main_window.descendants():
    if (
        ctrl.element_info.control_type == "Button"
        and ctrl.window_text().strip().lower() == "publish"
    ):
        if "ms-button" in ctrl.element_info.class_name.lower():
            publish_btn = ctrl
            break

if publish_btn:
    publish_btn.click_input()
    print("✅ Botón 'Publish' clickeado correctamente.")
else:
    print("❌ No se encontró el botón 'Publish'.")

time.sleep(10)


# In[20]:


from pywinauto import Application
import time

# Conectarse al proceso Power BI
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

print("🔍 Buscando ventanas hijas del proceso Power BI...")
time.sleep(4.5)  # Esperar un poco a que aparezcan

# Listar hijos
children = main_window.children()

if not children:
    print("❌ No se encontraron ventanas hijas.")
else:
    for i, child in enumerate(children):
        print(f"[{i}] {child.friendly_class_name()} | '{child.window_text()}' | class: {child.element_info.class_name}")

    # Buscar la ventana tipo Dialog
    dialog = None
    for child in children:
        if child.friendly_class_name() == "Dialog":
            dialog = child
            break

    if dialog:
        print("\n✅ Dialog encontrado. Listando todos los descendants...\n")
        for i, desc in enumerate(dialog.descendants()):
            print(f"[{i}] {desc.friendly_class_name()} | '{desc.window_text()}' | class: {desc.element_info.class_name}")
    else:
        print("❌ No se encontró ninguna ventana tipo Dialog.")


# In[21]:


from pywinauto import Application, Desktop
import time

# Conectarse a Power BI
app = Application(backend="uia").connect(path="PBIDesktop.exe")
main_window = app.top_window()
main_window.set_focus()

print("🔍 Buscando ventanas hijas del proceso Power BI...")
time.sleep(2)  # esperar a que se renderice

# Buscar Dialog
dialog = None
for child in main_window.children():
    if child.friendly_class_name() == "Dialog":
        dialog = child
        break

if not dialog:
    print("❌ No se encontró ninguna ventana tipo Dialog.")
else:
    print("✅ Ventana Dialog detectada.")

    try:
        # Buscar y clickear el ListItem 'dgh'
        list_item = next(
            (c for c in dialog.descendants()
             if c.friendly_class_name() == "ListItem" and c.window_text().strip().lower() == "dgh"),
            None
        )
        if list_item:
            list_item.click_input()
            print("✅ Workspace 'dgh' seleccionado.")
        else:
            print("❌ No se encontró el ListItem 'dgh'.")

        time.sleep(0.5)

        # Buscar y clickear el botón 'Select'
        select_btn = next(
            (c for c in dialog.descendants()
             if c.friendly_class_name() == "Button" and c.window_text().strip().lower() == "select"),
            None
        )
        if select_btn:
            select_btn.click_input()
            print("✅ Botón 'Select' clickeado.")
        else:
            print("❌ No se encontró el botón 'Select'.")

        # Esperar posible pop-up de reemplazo
        print("⏳ Esperando posible ventana de reemplazo...")
        time.sleep(2)
        for i in range(10):
            for d in main_window.children():
                if d.friendly_class_name() == "Dialog":
                    replace_btn = next(
                        (c for c in d.descendants()
                         if c.friendly_class_name() == "Button" and c.window_text().strip().lower() == "replace"),
                        None
                    )
                    if replace_btn:
                        replace_btn.click_input()
                        print("✅ Botón 'Replace' clickeado.")
                        break
            else:
                time.sleep(0.5)
                continue
            break

        # Esperar mensaje de éxito y click en el link
        print("⏳ Esperando mensaje 'Success!' y link para abrir en Power BI...")
        for _ in range(30):
            for d in main_window.children():
                if d.friendly_class_name() == "Dialog":
                    success_link = next(
                        (c for c in d.descendants()
                         if c.window_text().strip().lower().startswith("open")
                         and "in power bi" in c.window_text().lower()),
                        None
                    )
                    if success_link:
                        success_link.click_input()
                        print(f"✅ Link '{success_link.window_text()}' clickeado.")
                        break
            else:
                time.sleep(0.5)
                continue
            break
        else:
            print("⚠️ No se detectó el link de apertura tras la publicación.")

    except Exception as e:
        print(f"❌ Error durante la automatización: {e}")


# In[ ]:




