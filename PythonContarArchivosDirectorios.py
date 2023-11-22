#PythonContarArchivosDirectorios.py
# CuadraturaProvisoria

import pandas as pd
import os
import openpyxl

#
#Estructura de las Carpetas en REVISORES  por cada parcialidad.
#
#├───ESTRUCTURA DE CARPETAS DE CADA PARCIALIDAD
#│   ├───DOCUMENTOS APROBADOS
#│   │   ├───REV LETRA
#│   │   │   ├───01 PDF
#│   │   │   └───02 EDITABLE
#│   │   └───REV NUMERO
#│   │       ├───01 PDF
#│   │       └───02 EDITABLE
#│   └───DOCUMENTOS VIGENTES
#│       ├───01 PDF
#│       │   ├───CON OBSERVACIONES
#│       │   └───SIN OBSERVACIONES
#│       └───02 EDITABLE
#│           ├───CON OBSERVACIONES
#│           └───SIN OBSERVACIONES
#
         
# Lista de carpetas y subcarpetas
estructura = [
            'DOCUMENTOS APROBADOS/REV LETRA/01 PDF',
            'DOCUMENTOS APROBADOS/REV LETRA/02 EDITABLE',
            'DOCUMENTOS APROBADOS/REV NUMERO/01 PDF',
            'DOCUMENTOS APROBADOS/REV NUMERO/02 EDITABLE',
            'DOCUMENTOS VIGENTES/01 PDF/CON OBSERVACIONES',
            'DOCUMENTOS VIGENTES/01 PDF/SIN OBSERVACIONES',
            'DOCUMENTOS VIGENTES/02 EDITABLE/CON OBSERVACIONES',
            'DOCUMENTOS VIGENTES/02 EDITABLE/SIN OBSERVACIONES',
             ]

total_aprob_num_pdf = 0
total_aprob_num_editable = 0
total_aprob_letra_pdf = 0
total_aprob_letra_editable = 0
total_vig_pdf = 0
total_vig_editable = 0

# Función para contar archivos en una carpeta y sus subcarpetas
def contar_archivos(ruta):
    contador = 0
    for raiz, directorios, archivos in os.walk(ruta):
        for archivo in archivos:
            contador += 1
    return contador



# Ruta base donde se deben verificar los subdirectorios
ruta_base = 'R:\\01 PARCIALIDADES\\'   

# Nombre del archivo de log
archivo_log = ruta_base + '0000-00 ADMINISTRACION\\LOG\\log_CuadraturaProvisoria.txt'

# Planilla con la lista de parcialidades
archivo_excel = ruta_base + 'Listado de Parcialidades.xlsx'

# Carga el archivo Excel en un DataFrame Hoja de Parcialidades.
df = pd.read_excel(archivo_excel, sheet_name='PARCIALIDADES')

# Filtra el DataFrame para considerar solo parcialidades a 'PROCESAR' igual a 'S'
df_parcialidades = df[df['PROCESAR'] == 'S']

# Abre el archivo de log en modo de escritura
with open(archivo_log, 'w') as log_file:

    # Itera a través de cada parcialidad y la procesa
    for parcialidad in df_parcialidades['PARCIALIDAD']:

        #******* 
        #******* RECORRER TODAS LAS PARCIALIDADES CONTANDO LOS ARCHIVOS DE LA SIGUIENTE ESTRUCTURA de las Carpetas en REVISORES  por cada parcialidad.
        #******* 

        parcialidad_0_7_10 = parcialidad[0:7]
        if parcialidad_0_7_10   == '0029-14':
            parcialidad_0_7_10 = parcialidad[0:10]
        elif parcialidad_0_7_10 == '032ESO-':
            parcialidad_0_7_10 = parcialidad[0:9]
        elif parcialidad_0_7_10 == '032ESP-':
            parcialidad_0_7_10 = parcialidad[0:9]
      
        #******* Abrir Planilla CONTROL DOCUMENTOS ING DEF Pxxxx-xx con las 8 hojas para traspasar a BAT
        archivo_parcialidad = ruta_base + parcialidad + '\\CONTROL DOCUMENTOS ING DEF P' + parcialidad_0_7_10 + '.xlsx'
 
        if not os.path.exists(archivo_parcialidad):
              log_file.write(f'Parcialidad: {parcialidad_0_7_10} SIN ARCHIVO DE INGENIERIA {archivo_parcialidad}\n')
        else:
                print(f'Procesando Parcialidad: {parcialidad_0_7_10} ARCHIVO:  {archivo_parcialidad}')

                #******* Cargar HOJA: ÚLTIMA VERSIÓN del archivo Excel en un DataFrame (DF_xxxx)
                
                valor_uv = ''
                valor_rl = ''
                valor_rn = ''

                workbook = openpyxl.load_workbook(archivo_parcialidad, data_only=True, read_only=True)
                xl = pd.ExcelFile(archivo_parcialidad)
                nombres_hojas = xl.sheet_names

                if 'ÚLTIMA VERSIÓN' in nombres_hojas:
                    sheet_ultima_version = workbook['ÚLTIMA VERSIÓN']
                    # Accede al valor de las celda 'Z4'
                    valor_uv = sheet_ultima_version['Z4'].value
                else:
                    print(f'La hoja ÚLTIMA VERSIÓN no existe en {archivo_excel}.')  
                    log_file.write(f'La hoja ÚLTIMA VERSIÓN no existe en {archivo_excel}\n')
                if 'REV LETRA APRO' in nombres_hojas:
                    sheet_rev_letra_apro = workbook['REV LETRA APRO']
                    # Accede al valor de las celda 'Z7'
                    valor_rl = sheet_rev_letra_apro['Z7'].value
                else:
                    print(f'La hoja REV LETRA APRO no existe en {archivo_excel}.')  
                    log_file.write(f'La hoja REV LETRA APRO no existe en {archivo_excel}\n')
                if 'REV NUM APRO' in nombres_hojas:
                    sheet_rev_num_apro = workbook['REV NUM APRO']
                    # Accede al valor de las celda 'W5'
                    valor_rn = sheet_rev_num_apro['W5'].value
                else:
                    print(f'La hoja REV NUM APRO no existe en {archivo_excel}.')  
                    log_file.write(f'La hoja REV NUM APRO no existe en {archivo_excel}\n')

                ruta_subdirectorio = os.path.join(ruta_base, parcialidad)
                # Iterar a través de la estructura y crear carpetas si no existen
                for carpeta in estructura:
                    ruta_carpeta = os.path.join(ruta_subdirectorio, carpeta)
                    elementos = os.listdir(ruta_carpeta)
                    archivos = [elemento for elemento in elementos if os.path.isfile(os.path.join(ruta_carpeta, elemento))]
                    cantidad_de_archivos = len(archivos)
           
                    if carpeta  =='DOCUMENTOS APROBADOS/REV LETRA/01 PDF':                      
                        total_aprob_letra_pdf += cantidad_de_archivos
                    if carpeta  =='DOCUMENTOS APROBADOS/REV LETRA/02 EDITABLE':                       
                        total_aprob_letra_editable += cantidad_de_archivos
                    if carpeta == 'DOCUMENTOS APROBADOS/REV NUMERO/01 PDF':                         
                        total_aprob_num_pdf += cantidad_de_archivos
                    if carpeta == 'DOCUMENTOS APROBADOS/REV NUMERO/02 EDITABLE':                         
                        total_aprob_num_editable += cantidad_de_archivos
                    if carpeta == 'DOCUMENTOS VIGENTES/01 PDF/CON OBSERVACIONES':                         
                        total_vig_pdf += cantidad_de_archivos
                    if carpeta == 'DOCUMENTOS VIGENTES/01 PDF/SIN OBSERVACIONES':                         
                        total_vig_pdf += cantidad_de_archivos
                    if carpeta == 'DOCUMENTOS VIGENTES/02 EDITABLE/CON OBSERVACIONES':                         
                        total_vig_editable += cantidad_de_archivos
                    if carpeta == 'DOCUMENTOS VIGENTES/02 EDITABLE/SIN OBSERVACIONES':                         
                        total_vig_editable += cantidad_de_archivos
                    
                    print(f'{parcialidad_0_7_10:20}\t{carpeta:50}\t{cantidad_de_archivos}\t (uv:{valor_uv})\t (rl:{valor_rl})\t (rn:{valor_rn})')             
                    log_file.write(f'{parcialidad_0_7_10:20}\t{carpeta:50}\t{cantidad_de_archivos}\t (uv:{valor_uv})\t (rl:{valor_rl})\t (rn:{valor_rn})\n')
                log_file.write(f'\n')
log_file.close

with open(archivo_log, 'a') as log_file:

    # Graba los totales en LOG
    log_file.write(f'total_aprob_num_pdf: {total_aprob_num_pdf}\n')
    log_file.write(f'total_aprob_num_editable: {total_aprob_num_editable}\n')
    log_file.write(f'total_aprob_letra_pdf: {total_aprob_letra_pdf}\n')
    log_file.write(f'total_aprob_letra_editable: {total_aprob_letra_editable}\n')
    log_file.write(f'total_vig_pdf: {total_vig_pdf}\n')
    log_file.write(f'total_vig_editable: {total_vig_editable}\n')
                    
print("Contabilizacion finalizada. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\LOG en el archivo de log_CuadraturaProvisoria.")

# Imprime los totales por pantalla
print(f'total_aprob_num_pdf: {total_aprob_num_pdf}')
print(f'total_aprob_num_editable: {total_aprob_num_editable}')
print(f'total_aprob_letra_pdf: {total_aprob_letra_pdf}')
print(f'total_aprob_letra_editable: {total_aprob_letra_editable}')
print(f'total_vig_pdf: {total_vig_pdf}')
print(f'total_vig_editable: {total_vig_editable}')

log_file.close
