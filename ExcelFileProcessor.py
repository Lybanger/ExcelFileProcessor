import os
import glob
import pandas as pd


def process_file(file_path):
    try:
        # Leer el archivo excluyendo la última fila
        df = pd.read_excel(file_path, skipfooter=1)

        # Eliminar columnas 'Unnamed: x'
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

        # Obtener información del nombre del archivo
        file_name = os.path.basename(file_path)
        empresa, tipo_nomina, tipo_periodo, periodo = file_name.split('_')[:3] + [file_name.split('_')[3].split('.')[0]]

        # Agregar las columnas nuevas
        df['Empresa'] = empresa
        df['Tipo de nomina'] = tipo_nomina
        df['Tipo de periodo'] = tipo_periodo
        df['Periodo'] = periodo

        return df
    except Exception as e:
        print(f"Error al procesar el archivo {file_path}: {e}")
        return None


try:
    origen = str("C:\\Users\\usuario\\Documents\\")
    pat = "*.xls"  # Cambia la extensión si tus archivos son .xls
    files_joined = os.path.join(origen, pat)
    # Crea una lista con los archivos unidos
    list_files = glob.glob(files_joined)

    # Procesar cada archivo y concatenar los resultados
    dfs = [process_file(file) for file in list_files]
    df = pd.concat(dfs, ignore_index=True, sort=False)
    # Guardar el DataFrame en un archivo .xls
    output_file = 'archivo.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Archivo guardado como {output_file}")
except Exception as e:
    print("Error:", e)










