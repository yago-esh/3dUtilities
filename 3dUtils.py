import os
import shutil
import openpyxl
from datetime import datetime

debug = True;

datos_archivos = []

def copy_folders_with_Gcode(source_dir, usb_drive):
    # Lista para almacenar los datos
    for root, dirs, files in os.walk(source_dir):
        for directory in dirs:
            if "Gcode" in directory:
                source_path = os.path.join(root, directory)
                destination_path = os.path.join(usb_drive, directory)
                copy_folder(source_path, destination_path)
                    

def copy_folder(source_path, destination_path):
    if not os.path.exists(destination_path):
        copy_new_folder(source_path, destination_path)
    else:
        copy_existing_folder(source_path, destination_path)

def copy_new_folder(source_path, destination_path):
    shutil.copytree(source_path, destination_path)
    print(f"Copiado: \t{os.path.basename(source_path)}")
    rellenar_fichero_con_nombres(source_path)
    
def copy_existing_folder(source_path, destination_path):
    removeUnused(source_path, destination_path)
    for item in os.listdir(source_path):
        source_item = os.path.join(source_path, item)
        destination_item = os.path.join(destination_path, item)

        if os.path.isfile(source_item):
            copy_file(source_item, destination_item)
        else:
            copy_folder(source_item, destination_item)

def copy_file(source_item, destination_item):
    if os.path.exists(destination_item):
        copy_existing_file(source_item, destination_item)
    else:
       copy_new_file(source_item, destination_item) 

def copy_new_file(source_item, destination_item):
    shutil.copy2(source_item, destination_item)
    print(f"Copiado: \t{os.path.basename(destination_item)}")
    if "Done" in destination_item:
        item = os.path.dirname(destination_item)
        parent_dir = os.path.dirname(item)
        source_item = os.path.join(parent_dir, item)
        if os.path.exists(source_item):
            os.remove(source_item)
            print(f"Borrado: \t\t{os.path.basename(source_item)}")
    else:
        addFormatedDataToCSV(source_item, destination_item)

def copy_existing_file(source_item, destination_item):
    source_mtime = os.path.getmtime(source_item)
    dest_mtime = os.path.getmtime(destination_item)

    if source_mtime > dest_mtime:
        shutil.copy2(source_item, destination_item)
        print(f"Sobrescrito: \t\t{os.path.basename(source_item)}")
        addFormatedDataToCSV(source_item)

def rellenar_fichero_con_nombres(source_path):
    for root, dirs, files in os.walk(source_path):
        for file in files:
            addFormatedDataToCSV(file, source_path)

def removeUnused(source_path, destination_path):
     for item in os.listdir(destination_path):
        source_item = os.path.join(source_path, item)
        destination_item = os.path.join(destination_path, item)
        if os.path.exists(destination_item) and os.path.isfile(destination_item):
                  if not os.path.exists(source_item):
                    os.remove(destination_item)
                    print(f"Borrado por obsoleto: \t{getEasyPath(destination_item)}")
     
def getEasyPath (path):
    nombre_archivo = os.path.basename(path)
    directorio_superior = os.path.basename(os.path.dirname(path))
    return directorio_superior + "\\" + nombre_archivo 

def addFormatedDataToCSV(source_item, nombrePadre):
    proyect_name = os.path.basename(os.path.dirname(nombrePadre))
    time_to_print = os.path.basename(source_item)[:5]
    file_name = os.path.basename(source_item)[6:]
    datos_archivos.append([proyect_name, time_to_print, file_name])

def writeXlsx():
    # Abre el archivo Excel
    workbook = openpyxl.load_workbook('ejemplo.xlsx')

    # Selecciona la hoja en la que deseas agregar datos
    sheet = workbook['Resultado']  # O puedes seleccionar una hoja específica por nombre
    
    primera_fila_vacia = get_place_to_write(sheet)

    # Agrega los nuevos datos al final de la hoja
    numArchivos = 0
    for row in datos_archivos:
        print(f"Copiado: \t{row}")
        numArchivos+= 1
        # Itera sobre los elementos de la fila y escribe en celdas específicas
        for col, value in enumerate(row, start=1):
            sheet.cell(row=primera_fila_vacia, column=col, value=value)
        primera_fila_vacia += 1  # Avanza a la siguiente fila
        if primera_fila_vacia == 51:
            fecha_actual = datetime.now().strftime("%Y%m%d")
            sheet.title = f"Resultado_{fecha_actual}"
            plantilla = workbook['Plantilla']
            sheet = workbook.copy_worksheet(plantilla)
            sheet.title = "Resultado"
            primera_fila_vacia = 2
    print(f"Archivos copiados: {numArchivos}")
    # Guarda el archivo Excel con los nuevos datos
    workbook.save('ejemplo.xlsx')

    # Cierra el archivo Excel
    workbook.close()

def get_place_to_write(sheet):
    # Encuentra la primera fila vacía en la columna A
    columna_A = sheet['A']
    primera_fila_vacia = None

    for fila, celda in enumerate(columna_A, start=1):
        if not celda.value:
            return fila

if __name__ == "__main__":
    source_directory = "C:\\Users\\yager\\Documents\\Impresion3D"  # Cambia esta ruta a tu directorio de origen
    usb_drive = "I:"  # Cambia esta letra de unidad a tu USB
    if(debug):
        usb_drive = "C:\\Prueba"  # Cambia esta letra de unidad a tu USB
        shutil.rmtree(usb_drive)
        copy_folders_with_Gcode(source_directory, usb_drive)
        writeXlsx()
    else:
        if os.path.exists(usb_drive):
            copy_folders_with_Gcode(source_directory, usb_drive)
            writeXlsx()
            input("PROGRAMA FINALIZADO")
        else: 
            input("FALTA USB")