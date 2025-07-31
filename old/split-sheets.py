import pandas as pd

def dividir_xlsx_en_hojas(ruta_archivo_entrada, ruta_archivo_salida, filas_por_hoja=2500):
    """
    Divide un archivo XLSX en varias hojas, manteniendo el encabezado en cada una.

    Args:
        ruta_archivo_entrada (str): La ruta al archivo XLSX de entrada.
        ruta_archivo_salida (str): La ruta para guardar el nuevo archivo XLSX con las hojas divididas.
        filas_por_hoja (int): El número máximo de filas de datos por hoja (sin contar el encabezado).
    """
    try:
        # Cargar la primera hoja del archivo Excel en un DataFrame de pandas
        df = pd.read_excel(ruta_archivo_entrada)

        # Obtener la fila de encabezado
        encabezado = df.columns.tolist()

        # Calcular el número total de hojas que se necesitarán
        num_hojas = (len(df) - 1) // filas_por_hoja + 1

        # Crear un objeto ExcelWriter para escribir en varias hojas
        with pd.ExcelWriter(ruta_archivo_salida, engine='openpyxl') as writer:
            # Iterar a través de los trozos del DataFrame
            for i in range(num_hojas):
                # Calcular los índices de inicio y fin para el trozo actual
                inicio = i * filas_por_hoja
                fin = inicio + filas_por_hoja

                # Obtener el trozo de datos
                df_trozo = df.iloc[inicio:fin]

                # Crear un nombre para la nueva hoja
                nombre_hoja = f'Hoja_{i + 1}'

                # Escribir el trozo en una nueva hoja, incluyendo el encabezado
                df_trozo.to_excel(writer, sheet_name=nombre_hoja, index=False, header=True)

        print(f"¡El archivo se ha dividido correctamente en {num_hojas} hojas en '{ruta_archivo_salida}'!")

    except FileNotFoundError:
        print(f"Error: No se encontró el archivo de entrada en '{ruta_archivo_entrada}'")
    except Exception as e:
        print(f"Ha ocurrido un error: {e}")

if __name__ == '__main__':
    # --- Configuración ---
    # Coloca aquí la ruta de tu archivo de entrada
    archivo_entrada = 'tu_archivo_grande.xlsx'

    # Coloca aquí el nombre que deseas para el archivo de salida
    archivo_salida = 'archivo_dividido.xlsx'

    # Número de filas por hoja (sin contar el encabezado)
    filas = 2500
    # --- Fin de la Configuración ---

    dividir_xlsx_en_hojas(archivo_entrada, archivo_salida, filas)
