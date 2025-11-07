# === Base paths ===
BASE_PATH_MARIANO = r"C:\Users\Mariano\Documents\DMV Automation - 10 Jun"
BASE_PATH_USER    = r"C:\Users\Mariano\Documents"

# Ejecutable de Power BI
powerbi_exe_path = r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"

# Ejecutable de DAX Studio
DAX_PATH = r"C:\Program Files\DAX Studio\DaxStudio.exe"

# Plantilla Empower BI
plantilla = fr"{BASE_PATH_MARIANO}\Plantilla Empower BI.pbix"

# Carpeta base donde están los reportes .pbix de entrada
base_report_directory = fr"{BASE_PATH_MARIANO}\Input - Reportes"

#-----------------------------------------------------------------------------------------------------

# Rutas a cada script .py
script_one_path     = fr"{BASE_PATH_MARIANO}\One_Extraccion_DMVs_V3.py"
script_two_a_path   = fr"{BASE_PATH_MARIANO}\Two_A_Analisis_Dependencias.py"
script_two_b_path   = fr"{BASE_PATH_MARIANO}\Two_B_Consolidacion.py"
script_three_path   = fr"{BASE_PATH_MARIANO}\Three_Cargar_plantilla_publicar.py"

#-----------------------------------------------------------------------------------------------------

# === Etapas del proceso ===
step1_extracciones_dmv     = fr"{BASE_PATH_MARIANO}\DMV Files"
step2_archivos_consolidados = fr"{BASE_PATH_MARIANO}\Archivos Consolidados"
step3_analisis_dependencias = step1_extracciones_dmv  # mismo que step1, si aplica
step4_entregables           = fr"{BASE_PATH_USER}"
