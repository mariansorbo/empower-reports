#!/usr/bin/env python
# coding: utf-8

# In[16]:


import config
import config_runtime
from datetime import datetime
# Configuración
path_carpeta = config.step1_extracciones_dmv 
usuario_objetivo = config_runtime.usuario
reporte_objetivo=config_runtime.reporte


# In[ ]:


import warnings
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')


# In[17]:


import os
import re
from collections import defaultdict

def detectar_reportes(path_carpeta, usuario):
    archivos = os.listdir(path_carpeta)
    patron = re.compile(rf"^{usuario} \+ (.+?) \+ ")
    reportes = defaultdict(list)

    for archivo in archivos:
        if archivo.endswith(".xlsx"):
            match = patron.match(archivo)
            if match:
                nombre_reporte = match.group(1).strip()
                reportes[nombre_reporte].append(archivo)

    return dict(reportes)


# In[18]:


usuario = config_runtime.usuario

reportes = detectar_reportes(path_carpeta, usuario)
for nombre, archivos in reportes.items():
    print(f"\n📊 Reporte: {nombre}")
    for a in archivos:
        print("  -", a)


# In[21]:


import os
import pandas as pd
import re
from collections import defaultdict
import config
import config_runtime

# Configuración
path_carpeta = config.step1_extracciones_dmv
usuario_objetivo = config_runtime.usuario.strip().lower()

# Tipos válidos normalizados
tipos_validos = [
    "columnas",
    "metricas",
    "partitions",
    "tablas",
    "relaciones",
    "analisis de dependencias"
]

# Inicializar contenedor por tipo
dfs_por_tipo = defaultdict(list)

# Recorrer archivos
for archivo in os.listdir(path_carpeta):
    if not archivo.endswith(".xlsx"):
        continue

    match = re.match(r"(.*?)\s*\+\s*(.*?)\s*\+\s*(.*?)\.xlsx", archivo)
    if not match:
        continue

    usuario, reporte, tipo = match.groups()
    usuario_norm = usuario.strip().lower()
    tipo_norm = tipo.strip().lower()

    if usuario_norm != usuario_objetivo:
        continue

    if tipo_norm not in tipos_validos:
        continue

    try:
        filepath = os.path.join(path_carpeta, archivo)
        df = pd.read_excel(filepath)
        dfs_por_tipo[tipo_norm].append(df)
        print(f"[✓] Cargado archivo de tipo '{tipo_norm}': {archivo} ({df.shape[0]} filas)")
    except Exception as e:
        print(f"[!] Error leyendo '{archivo}': {e}")

# Combinar todos los archivos por tipo
dfs_finales = {}
for tipo, lista_dfs in dfs_por_tipo.items():
    if lista_dfs:
        df_total = pd.concat(lista_dfs, ignore_index=True)
        dfs_finales[tipo] = df_total
        print(f"[📊] Unificado df_{tipo}: {df_total.shape[0]} filas")

        # Guardar como variable global si querés
        globals()[f"df_{tipo.replace(' ', '_')}"] = df_total


# In[22]:


import import_ipynb
import Funciones  # without .ipynb

#llamo a la funcion para unificar columnas y metricas
df_unificado = Funciones.estandarizar_y_unir(df_columnas, df_metricas)  # call your function

#guardo el nuevo archivo "Columnas y MEtricas.xlsx"
df_unificado.head()


# In[23]:


df_tablas["Reporte + TableID"] = df_tablas["Reporte"] + " - " + df_tablas["ID"].astype(str)
df_tablas.head()
df_partitions["Reporte + TableID"] = df_partitions["Reporte"] + " - " + df_partitions["TableID"].astype(str)
df_partitions.head()
df_relaciones["Reporte + From TableID"]=df_relaciones["Reporte"]+ " - " +df_relaciones["FromTableID"].astype(str)
df_relaciones["Reporte + To TableID"]=df_relaciones["Reporte"]+ " - " +df_relaciones["ToTableID"].astype(str)
df_relaciones.head()
df_unificado["Reporte + TableID"] = df_unificado["Reporte"] + " - " + df_unificado["TableID"].astype(str)
df_unificado.head()


df_unificado


# '''5. 📊 Consolidación de los datos en un solo Excel
# Todos los DataFrames (columnas, métricas, tablas, relaciones y particiones) son escritos en un único archivo Excel con múltiples hojas (sheet_name).
# 
# Este archivo se guarda en:
# 
# 📁 OneDrive:
# C:\Users\Administrator\OneDrive\Empower BI Archivos\Mariano + Reporte + (fecha).xlsx

# In[24]:


#Configuración
# Ruta de carpeta y nombre de archivo
carpeta_destino = config.step2_archivos_consolidados
fecha = datetime.now().strftime("%Y-%m-%d")
nombre_archivo = f"{usuario} - Consolidado.xlsx"


# In[25]:


#ONE DRIVE

import pandas as pd
import os
from datetime import datetime


ruta_final = os.path.join(carpeta_destino, nombre_archivo)

# Lista de tus DataFrames
dfs = {
    "Columnas y Métricas": df_unificado,
    "Partitions": df_partitions,
    "Tablas": df_tablas,
    "Relaciones": df_relaciones,
    "Analisis de Dependencias":df_analisis_de_dependencias
}

# Guardar como Excel con múltiples hojas
with pd.ExcelWriter(ruta_final, engine='xlsxwriter') as writer:
    for nombre_hoja, df in dfs.items():
        df.to_excel(writer, sheet_name=nombre_hoja[:31], index=False)  # 31 = máx. de caracteres en nombre de hoja

print(f"[✓] Archivo Excel guardado exitosamente en: {ruta_final}")


# In[26]:


df_analisis_de_dependencias


# In[ ]:





# In[ ]:




