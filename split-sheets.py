import pandas as pd
import os

def formatear_fechas(df):
    """
    Recorre un DataFrame y formatea todas las columnas de tipo datetime.
    """
    for col in df.columns:
        # Comprobar si la columna es de tipo fecha/hora
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            print(f"Formateando columna de fecha: '{col}'...")
            # Aplicar el formato. .dt es el accesor para propiedades de datetime.
            # strftime convierte la fecha/hora a un string con el formato dado.
            df[col] = df[col].dt.strftime('%d-%m-%Y %H-%M')
    return df

def dividir_xlsx_en_archivos(ruta_archivo_entrada, carpeta_salida, filas_por_archivo):
    """
    Divide un archivo XLSX en varios archivos, manteniendo el encabezado y formateando fechas.

    Args:
        ruta_archivo_entrada (str): La ruta al archivo XLSX de entrada.
        carpeta_salida (str): La carpeta donde se guardarán los archivos divididos.
        filas_por_archivo (int): El número máximo de filas de datos por archivo.
    """
    try:
        # 1. Cargar el archivo Excel
        print(f"Cargando archivo '{ruta_archivo_entrada}'...")
        df = pd.read_excel(ruta_archivo_entrada)
        print("Archivo cargado con éxito.")

        # 2. Formatear las columnas de fecha/hora
        df = formatear_fechas(df)

        # 3. Crear la carpeta de salida si no existe
        os.makedirs(carpeta_salida, exist_ok=True)
        # Obtenemos el nombre base del archivo para usarlo en los nombres de salida
        prefijo_salida = os.path.splitext(os.path.basename(ruta_archivo_entrada))[0]


        # 4. Dividir y guardar los archivos
        num_archivos = (len(df) - 1) // filas_por_archivo + 1
        print(f"\nEl archivo se dividirá en {num_archivos} partes.")

        for i in range(num_archivos):
            inicio = i * filas_por_archivo
            fin = inicio + filas_por_archivo
            df_trozo = df.iloc[inicio:fin]

            # Generar el nombre y la ruta del archivo de salida
            nombre_archivo_salida = f"{prefijo_salida}_parte_{i + 1}.xlsx"
            ruta_archivo_salida = os.path.join(carpeta_salida, nombre_archivo_salida)

            # Guardar el trozo en un nuevo archivo Excel
            df_trozo.to_excel(ruta_archivo_salida, index=False, header=True, engine='openpyxl')

            print(f"Archivo '{ruta_archivo_salida}' creado con éxito.")

        print(f"\n¡Proceso completado! Se han creado {num_archivos} archivos en la carpeta '{carpeta_salida}'.")

    except FileNotFoundError:
        print(f"Error: No se encontró el archivo de entrada en '{ruta_archivo_entrada}'")
    except Exception as e:
        print(f"Ha ocurrido un error inesperado: {e}")

if __name__ == '__main__':
    # --- Configuración ---
    # Coloca aquí la ruta de tu archivo de entrada
    archivo_entrada = 'tu_archivo_grande.xlsx'
    
    # Nombre de la carpeta donde se guardarán los archivos divididos
    carpeta_salida = 'archivos_divididos'
    # --- Fin de la Configuración ---

    # Bucle para preguntar al usuario la cantidad de filas
    while True:
        try:
            filas_str = input("Introduce la cantidad de líneas por archivo en las que quieres dividir: ")
            filas = int(filas_str)
            if filas > 0:
                break  # Salir del bucle si la entrada es un número positivo
            else:
                print("Por favor, introduce un número mayor que cero.")
        except ValueError:
            print("Entrada no válida. Por favor, introduce solo un número entero.")

    # Llamar a la función principal con los parámetros configurados
    dividir_xlsx_en_archivos(archivo_entrada, carpeta_salida, filas)
