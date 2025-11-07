#!/usr/bin/env python
# coding: utf-8

# # Este script sirve para obtener un análisis de dependencias entre columnas y métricas.
# 
# 
# # **INPUT** (Se obtiene de la Documentación de Reportes, o de la Documentación de Único Reporte, en una solapa con el mismo nombre):
# # '0 - Diccionario Columnas y Metricas - Reportes.csv'  
# 
# ------------------------
# 
# # **OUTPUT** (Se usa en la Documentación de Reportes):
# # 2 - Analisis de Dependencias Intercolumnares - Reportes
# 
# 
# 

# In[1]:


import config_runtime
import config
usuario= config_runtime.usuario
reporte=config_runtime.reporte
path_carpeta =config.step1_extracciones_dmv 


# In[2]:


import os
import pandas as pd
import re
import import_ipynb
import Funciones  # debe contener la función estandarizar_y_unir


# Normalizar
usuario_norm = usuario.strip().lower()
reporte_norm = reporte.strip().lower()

# Inicializar
df_columnas = pd.DataFrame()
df_metricas = pd.DataFrame()
df_tablas = pd.DataFrame()

# Buscar archivos
for archivo in os.listdir(path_carpeta):
    if not archivo.endswith(".xlsx"):
        continue

    match = re.match(r"(.*?)\s*\+\s*(.*?)\s*\+\s*(.*?)\.xlsx", archivo)
    if not match:
        continue

    archivo_usuario, archivo_reporte, tipo = match.groups()
    tipo = tipo.strip().lower()

    if (archivo_usuario.strip().lower() == usuario_norm and
        archivo_reporte.strip().lower() == reporte_norm):

        filepath = os.path.join(path_carpeta, archivo)

        if tipo == "columnas":
            df_columnas = pd.read_excel(filepath)
            print(f"[OK] Cargado archivo de columnas: {archivo}")
        elif tipo == "metricas":
            df_metricas = pd.read_excel(filepath)
            print(f"[OK] Cargado archivo de métricas: {archivo}")
        elif tipo == "tablas":
            df_tablas = pd.read_excel(filepath)
            print(f"[OK] Cargado archivo de tablas: {archivo}")

# Unificar columnas + métricas
df_unificado = Funciones.estandarizar_y_unir(df_columnas, df_metricas)




# In[3]:


df_unificado


# In[4]:


# ✅ quiero traer el campo TableName, tengo Crear columna auxiliar en ambos DataFrames
df_unificado["join_key"] = df_unificado["Reporte"] + df_unificado["TableID"].astype(str)
df_tablas["join_key"] = df_tablas["Reporte"] + df_tablas["ID"].astype(str)

# ✅ Usar .map() para traer el campo 'Name' desde df_mariano_reporte_tablas
df_unificado["Tabla Nombre"] = df_unificado["join_key"].map(
    df_tablas.set_index("join_key")["Name"]
)


df_unificado


# In[5]:


# Definir función para extraer dependencias
def extraer_columnas(expresion, tabla_actual):
    # Buscar patrones de tipo dimension[columna] o [columna] en las expresiones
    patrones = re.findall(r'(\w*)\[(\w+)\]', expresion)
    columnas_referenciadas = []
    columnas_referenciadas_2=[]

    for dimension, columna in patrones:
        if dimension:  # Si hay un prefijo de dimensión
            columnas_referenciadas.append(f"{dimension}.{columna}")
        else:  # Si no hay prefijo de dimensión, pertenece a la tabla actual
            columnas_referenciadas.append(f"{tabla_actual}.{columna}")

    for dimension, columna in patrones:
        if dimension:  # Si hay un prefijo de dimensión
            columnas_referenciadas_2.append(f"{dimension}[{columna}]")
        else:  # Si no hay prefijo de dimensión, pertenece a la tabla actual
            columnas_referenciadas_2.append(f"{tabla_actual}[{columna}]")

    return columnas_referenciadas,columnas_referenciadas_2


# In[6]:


# Construir diccionarios de dependencias y expresiones
dependencias = {}
dependencias_2 = {}  # Para las referencias con el formato de corchetes
expresiones = {}

for index, row in df_unificado.iterrows():
    reporte = row['Reporte']
    tabla = row['Tabla Nombre']
    columna = row['Name']
    tabla_columna = f"{tabla}|{columna}"
    print(  reporte,tabla,columna,tabla_columna)
    if reporte not in dependencias:
        dependencias[reporte] = {}
        dependencias_2[reporte] = {}  # Inicializar para dependencias_2
        expresiones[reporte] = {}

    if pd.notna(row['Expression']):
        columnas_referenciadas, columnas_referenciadas_2 = extraer_columnas(row['Expression'], tabla)

        if tabla_columna not in dependencias[reporte]:
            dependencias[reporte][tabla_columna] = []
            dependencias_2[reporte][tabla_columna] = []  # Para dependencias_2

        dependencias[reporte][tabla_columna].extend(columnas_referenciadas)
        dependencias_2[reporte][tabla_columna].extend(columnas_referenciadas_2)  # Para dependencias_2
        expresiones[reporte][tabla_columna] = row['Expression']
    else:
        expresiones[reporte][tabla_columna] = None  # Marca para columnas no calculadas
        dependencias[reporte][tabla_columna] = []  # No tiene dependencias si es directa
        dependencias_2[reporte][tabla_columna] = []  # No tiene dependencias si es directa


# In[7]:


# Crear el DataFrame final con las columnas solicitadas
data = []

for reporte in dependencias:
    for tabla_columna in dependencias[reporte]:
        try:
            # Safely split into tabla and columna
            tabla, columna = tabla_columna.split("|", maxsplit=1)  # Use maxsplit to avoid extra splits
        except ValueError:
            # Handle cases where the split doesn't produce exactly two parts
            print(f"Warning: Unexpected format in 'tabla_columna': {tabla_columna}")
            tabla, columna = tabla_columna, None  # Assign defaults

        expresion = expresiones[reporte].get(tabla_columna)
        referencias = dependencias[reporte][tabla_columna]
        referencias_2 = dependencias_2[reporte].get(tabla_columna, [])  # Handle missing keys in dependencias_2

        # Determine if it is final
        es_final = 'Si' if expresion is None else 'No'

        # Add the row to the new DataFrame
        data.append({
            'reporte': reporte,
            'Tabla Nombre': tabla,
            'Columna Nombre': columna,
            'Expresión': expresion,
            'Referencias': referencias,
            'Referencias 2': ",".join(referencias_2),  # Join the list as a single string
            'Es Final': es_final
        })

# Crear el DataFrame
df_resultado = pd.DataFrame(data)

# Mostrar los primeros 20 registros para verificar
df_resultado.head(20)


# In[8]:


from collections.abc import Iterable

# Función que devuelve la expresión de la columna a partir del reporte, tabla y columna
def ObtenerExpresion(df_resultado, reporte, tabla, columna):
    # Filtrar el DataFrame para obtener la expresión de la columna actual
    fila_actual = df_resultado[
        (df_resultado['reporte'] == reporte) &
        (df_resultado['Tabla Nombre'] == tabla) &
        (df_resultado['Columna Nombre'] == columna)
    ]

    # Si se encuentra la fila, devolver la expresión, si no, devolver un mensaje indicando que no se encontró
    if not fila_actual.empty:
        return fila_actual.iloc[0]['Expresión']
    else:
        return "Expresión no encontrada para la columna especificada."

# Función que devuelve un array de referencias a partir del reporte, tabla y columna
def obtener_referencias(df_resultado, reporte, tabla, columna):
    # Filtrar el DataFrame para obtener las dependencias de la columna actual
    fila_actual = df_resultado[
        (df_resultado['reporte'] == reporte) &
        (df_resultado['Tabla Nombre'] == tabla) &
        (df_resultado['Columna Nombre'] == columna)
    ]

    # Si se encuentra la fila, devolver las referencias, si no, devolver un array vacío
    if not fila_actual.empty:
        return list(set(fila_actual.iloc[0]['Referencias']))
    else:
        return None

# Función que devuelve las referencias en formato "dimension[columna]"
def obtener_referencias_2(df_resultado, reporte, tabla, columna):
    # Filtrar el DataFrame para obtener las dependencias formateadas en "Referencias 2"
    fila_actual = df_resultado[
        (df_resultado['reporte'] == reporte) &
        (df_resultado['Tabla Nombre'] == tabla) &
        (df_resultado['Columna Nombre'] == columna)
    ]

    # Si se encuentra la fila, devolver las referencias 2, si no, devolver un array vacío
    if not fila_actual.empty:
        return list(set(fila_actual.iloc[0]['Referencias 2']))
    else:
        return None

# Función recursiva para analizar las dependencias entre columnas y devolver el análisis y las columnas finales
def Recursividad(reporte, tabla, columna, nivel=1, visitados=None, columnasFinales=None, analisis="", lista_columnas_finales=None):
    sangria="          "
    if nivel >= 10:
        return analisis  # Retorna el análisis si el nivel es muy alto


    # Inicializar los parámetros opcionales si son None
    if visitados is None:
        visitados = set()  # Para controlar los nodos ya visitados y evitar bucles
    if columnasFinales is None:
        columnasFinales = set()  # Lista de columnas finales (que no tienen dependencias)
    if lista_columnas_finales is None:
        lista_columnas_finales = []  # Lista para acumular las columnas finales

    # Definir la clave única para la columna actual (tabla.columna)
    clave_columna = f"{tabla}[{columna}]"

    # Caso base: Si la columna ya fue visitada o si es una columna final, terminar la recursividad
    if clave_columna in visitados or clave_columna in columnasFinales:
        return analisis

    # Marcar la columna como visitada
    visitados.add(clave_columna)

    # Obtener su expresión o lógica
    expresion_actual = ObtenerExpresion(df_resultado, reporte, tabla, columna)

    # Obtener las dependencias (referencias a otras columnas) de la columna actual
    dependencias_actuales = obtener_referencias(df_resultado, reporte, tabla, columna)
    dependencias_actuales_2 = obtener_referencias_2(df_resultado, reporte, tabla, columna)  # Obtener referencias 2

    # Agregar información de la columna actual al análisis
    if dependencias_actuales is not None and isinstance(expresion_actual, Iterable):
        try:
            analisis +=      f"Nivel {nivel}:\n"
            analisis +=   f"{clave_columna}=\n"
            analisis +=   f"{expresion_actual}\n"
            analisis +=    f"--Columnas Referenciadas : {', '.join(dependencias_actuales)}\n\n"
        except Exception as e:
            analisis += f"Error procesando {clave_columna}: {str(e)}\n"
            columnasFinales.add(clave_columna)
            lista_columnas_finales.append(clave_columna)
    else:
        # Columnas sin dependencias (finales)
        columnasFinales.add(clave_columna)
        lista_columnas_finales.append(clave_columna)

    # Recorrer las dependencias de la columna y llamar recursivamente
    for referencia in dependencias_actuales or []:
        # Verificar si la referencia sigue el formato "tabla.columna"
        if '.' in referencia:
            tabla_ref, columna_ref = referencia.split('.')
            analisis = Recursividad(reporte, tabla_ref, columna_ref, nivel + 1, visitados, columnasFinales, analisis, lista_columnas_finales)
        else:
            analisis += f"Error: referencia '{referencia}' no tiene formato 'tabla.columna'.\n"

    return analisis

# Ejemplo de uso:
reporte_ejemplo = 'COC and COI'
tabla_ejemplo = 'HR Headcount'
columna_ejemplo = '# Trámites 2_'

# Inicializar string de análisis
analisis_resultado = ""
columnas_finales_lista = []

# Llamar a la función recursiva
analisis_resultado = Recursividad(reporte_ejemplo, tabla_ejemplo, columna_ejemplo, 1, set(), None, analisis_resultado, columnas_finales_lista)

# Concatenar las columnas finales directas de bajada al final del análisis
if columnas_finales_lista:
    analisis_resultado += "\nColumnas Base (Directas de bajada):\n" + ', '.join(set(columnas_finales_lista))

# Mostrar el resultado del análisis
print("Análisis Jerárquico de dependencias:")
print(analisis_resultado)



# In[9]:


df_resultado.head()


# In[10]:


import time  # Importar el módulo para medir el tiempo

# Función que realiza el análisis y devuelve el string del análisis con las columnas finales concatenadas
def generar_analisis_completo(row):
    reporte = row['reporte']
    tabla = row['Tabla Nombre']
    columna = row['Columna Nombre']

    # Inicializar variables
    analisis_resultado = ""
    columnas_finales_lista = []

    # Llamar a la función recursiva
    analisis_resultado = Recursividad(reporte, tabla, columna, 1, set(), None, analisis_resultado, columnas_finales_lista)

    # Concatenar las columnas finales al final del análisis
    if columnas_finales_lista:
        analisis_resultado += "\nColumnas Base (Directas de bajada):\n" + ', '.join(set(columnas_finales_lista))

    return analisis_resultado

# Medir el tiempo de ejecución
start_time = time.time()  # Tiempo de inicio

# Aplicar la función a cada fila del DataFrame y agregar la nueva columna "Análisis Completo"
df_resultado['Análisis Completo'] = df_resultado.apply(
    lambda row: generar_analisis_completo(row), axis=1
)

end_time = time.time()  # Tiempo de finalización

# Calcular el tiempo de ejecución
execution_time = end_time - start_time

# Mostrar el DataFrame con la nueva columna "Análisis Completo"
print(df_resultado[['reporte', 'Tabla Nombre', 'Columna Nombre', 'Análisis Completo', 'Referencias 2']])

# Mostrar el tiempo de ejecución
print(f"Tiempo de ejecución: {execution_time:.2f} segundos")


# In[11]:


import os
import pandas as pd


# Carpeta destino y nombre del archivo
carpeta_destino = config.step3_analisis_dependencias
os.makedirs(carpeta_destino, exist_ok=True)

nombre_archivo = f"{usuario} + {reporte} + analisis de dependencias.xlsx"
ruta_final = os.path.join(carpeta_destino, nombre_archivo)

# Crear archivo nuevo con hoja 'Dependency Analysis'
with pd.ExcelWriter(ruta_final, engine='openpyxl', mode='w') as writer:
    df_resultado.to_excel(writer, sheet_name="Dependency Analysis", index=False)

print(f"[CREADO] Archivo: {ruta_final}")


# In[12]:


df_resultado


# In[ ]:




