import pandas as pd
import os

def dividir_xlsx_en_archivos(ruta_archivo_entrada, prefijo_salida, filas_por_archivo=2500):
    """
    Divide un archivo XLSX en varios archivos XLSX, manteniendo el encabezado en cada uno.

    Args:
        ruta_archivo_entrada (str): La ruta al archivo XLSX de entrada.
        prefijo_salida (str): El prefijo para los nombres de los archivos de salida.
        filas_por_archivo (int): El número máximo de filas de datos por archivo (sin contar el encabezado).
    """
    try:
        # Cargar la primera hoja del archivo Excel en un DataFrame de pandas
        df = pd.read_excel(ruta_archivo_entrada)

        # Calcular el número total de archivos que se necesitarán
        num_archivos = (len(df) - 1) // filas_por_archivo + 1

        # Iterar a través de los trozos del DataFrame para crear cada archivo
        for i in range(num_archivos):
            # Calcular los índices de inicio y fin para el trozo actual
            inicio = i * filas_por_archivo
            fin = inicio + filas_por_archivo

            # Obtener el trozo de datos
            df_trozo = df.iloc[inicio:fin]

            # Generar el nombre del archivo de salida
            ruta_archivo_salida = f"{prefijo_salida}_{i + 1}.xlsx"

            # Escribir el trozo en un nuevo archivo Excel
            df_trozo.to_excel(ruta_archivo_salida, index=False, header=True, engine='openpyxl')

            print(f"Archivo '{ruta_archivo_salida}' creado con éxito.")

        print(f"\n¡Proceso completado! Se han creado {num_archivos} archivos a partir de '{ruta_archivo_entrada}'.")

    except FileNotFoundError:
        print(f"Error: No se encontró el archivo de entrada en '{ruta_archivo_entrada}'")
    except Exception as e:
        print(f"Ha ocurrido un error: {e}")

if __name__ == '__main__':
    # --- Configuración ---
    # Coloca aquí la ruta de tu archivo de entrada
    archivo_entrada = 'tu_archivo_grande.xlsx'

    # Coloca aquí el prefijo que deseas para los archivos de salida.
    # Por ejemplo, 'parte' resultará en 'parte_1.xlsx', 'parte_2.xlsx', etc.
    prefijo_archivo_salida = 'archivo_dividido'

    # Número de filas por archivo (sin contar el encabezado)
    filas = 2500
    # --- Fin de la Configuración ---

    dividir_xlsx_en_archivos(archivo_entrada, prefijo_archivo_salida, filas)
