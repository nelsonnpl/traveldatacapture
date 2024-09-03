import os
import pandas as pd
import unicodedata

def main():
    # Paso 1: Pedir la ruta del archivo de Excel
    excel_file_path = input("Introduce la ruta (path) del archivo Excel: ")
    
    # Eliminar posibles comillas dobles si el usuario las incluye
    excel_file_path = excel_file_path.strip('"')
    
    # Verificar si el archivo existe
    if not os.path.exists(excel_file_path):
        print("No encontramos el archivo :(, por favor verifica la ruta e intenta de nuevo.")
        return  # Salir de la función si no se encuentra el archivo
    
    try:
        # Intentar leer el archivo Excel
        df = pd.read_excel(excel_file_path)
        print("¡Hemos leído el archivo con éxito! :)")
        
    except Exception as e:
        print("Algo salió mal al leer el archivo, por favor intenta de nuevo.")
        print(e)
        return  # Salir de la función en caso de error

    # Paso 2: Preguntar cómo queremos llamar a la carpeta (nombre del expediente)
    folder_name = input("Introduce el nombre de la carpeta que deseas crear: ")

    try:
        # Crear la carpeta con el nombre indicado
        os.makedirs(folder_name, exist_ok=True)
        print(f"La carpeta '{folder_name}' ha sido creada con éxito.")
        
        # Crear el nuevo nombre para el archivo basado en el nombre de la carpeta
        new_file_name = f"TDC_{folder_name}_raw.xlsx"
        
        # Obtener la ruta completa del nuevo archivo dentro de la carpeta creada
        new_file_path = os.path.join(folder_name, new_file_name)
        
        # Mover y renombrar el archivo Excel a la nueva carpeta
        os.rename(excel_file_path, new_file_path)
        print(f"El archivo ha sido movido y renombrado a '{new_file_path}'.")
        
    except Exception as e:
        print(f"Hubo un error al crear la carpeta o mover el archivo: {e}")
        return  # Salir de la función en caso de error

    # Paso 3: Procesar las columnas seleccionadas y crear un nuevo archivo Excel
    try:
        # Leer de nuevo el archivo ya movido y renombrado
        df = pd.read_excel(new_file_path)
        
        # Extraer columnas O, P, Q, R, S (índices 14, 15, 16, 17 y 18)
        selected_columns = df.iloc[:, [14, 15, 16, 17, 18]]
        
        # Renombrar las columnas
        selected_columns.columns = ["name", "Apellidos", "Nacimiento", "Documento", "NDoc"]
        
        # Eliminar tildes o acentos de las columnas name y Apellidos
        selected_columns["name"] = selected_columns["name"].apply(remove_accents)
        selected_columns["Apellidos"] = selected_columns["Apellidos"].apply(remove_accents)
        
        # Guardar en un nuevo archivo Excel llamado 'Amadeus.xlsx'
        amadeus_file_path = os.path.join(folder_name, "Amadeus.xlsx")
        selected_columns.to_excel(amadeus_file_path, index=False)
        print(f"El nuevo archivo 'Amadeus.xlsx' se ha guardado en: {amadeus_file_path}")
        
        # Generar archivo .txt con formato NM1[apellidos]/[name];
        create_txt_file(selected_columns, folder_name)
        
    except Exception as e:
        print(f"Hubo un error al procesar las columnas o guardar el nuevo archivo: {e}")

def remove_accents(input_str):
    # Eliminar acentos o tildes del string
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd_form if not unicodedata.combining(c)])

def create_txt_file(df, folder_name):
    try:
        # Crear el archivo .txt en la carpeta especificada
        txt_file_path = os.path.join(folder_name, "NM1-Amadeus.txt")
        
        with open(txt_file_path, "w") as file:
            current_line = ""
            for index, row in df.iterrows():
                # Crear la cadena en el formato NM1[apellidos]/[name];
                name_str = f"NM1{row['Apellidos'].upper()}/{row['name'].upper()};"
                
                # Verificar si la longitud de la línea actual más la nueva cadena excede los 182 caracteres
                if len(current_line) + len(name_str) > 182:
                    # Escribir la línea actual en el archivo
                    file.write(current_line + "\n")
                    # Añadir un separador
                    file.write("-----------------------------\n")
                    # Iniciar una nueva línea con el nuevo nombre
                    current_line = name_str
                else:
                    # Añadir el nombre a la línea actual
                    current_line += name_str
            
            # Escribir cualquier contenido restante en la última línea
            if current_line:
                file.write(current_line + "\n")
                file.write("-----------------------------\n")
        
        print(f"El archivo 'NM1-Amadeus.txt' ha sido generado en: {txt_file_path}")
    
    except Exception as e:
        print(f"Hubo un error al generar el archivo .txt: {e}")

if __name__ == "__main__":
    main()
