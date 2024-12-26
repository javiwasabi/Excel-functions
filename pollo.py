import os
import pandas as pd
from tkinter import Tk, filedialog
from datetime import datetime, timedelta


def generar_excel_por_persona(atencion_path, control_path, output_path):
    try:
         # Datos adicionales que quieres agregar
        detalles_adicionales = [
            {'NOMBRE': 'BYHAMNY ALMONTE BATISTA', 'RUN': '254846147', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'SEPTIEMBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-09-2024', 'PERIODO_AL': '30-09-2024'},
            {'NOMBRE': 'NICOLAS GONZALEZ ZAVALA', 'RUN': '183856790', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'SEPTIEMBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-09-2024', 'PERIODO_AL': '30-09-2024'},
            {'NOMBRE': 'OMAR CARDENAS VILLACRES', 'RUN': '231890769', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'},
            {'NOMBRE': 'NICOLAS GONZALEZ ZAVALA', 'RUN': '183856790', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'},
            {'NOMBRE': 'DAYANA FERNANDEZ BETANCOURT', 'RUN': '266771185', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'},
            {'NOMBRE': 'MARTIN GANA OLIVARES', 'RUN': '197415177', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'},
            {'NOMBRE': 'FRANCO VEAS ZUÑIGA', 'RUN': '189259794', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'},
            {'NOMBRE': 'SIMON SEGUEL ESTAY', 'RUN': '191144082', 'CARGO': 'MEDICO ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'},
            {'NOMBRE': 'MARCELA WALTER GARRIDO', 'RUN': '202797431', 'CARGO': 'TENS ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'},
            {'NOMBRE': 'ANA ORTIZ GONZALEZ', 'RUN': '195707057', 'CARGO': 'TENS ATENCION NOCTURNA', 'MES': 'OCTUBRE', 'AÑO': '2024', 'PERIODO_DEL': '01-10-2024', 'PERIODO_AL': '31-10-2024'}
        ]

        # Cargar los datos
        atencion_data = pd.read_excel(atencion_path)
        control_sheets = pd.read_excel(control_path, sheet_name=None)

        # Normalizar columnas del archivo a.xls
        atencion_data.columns = atencion_data.columns.str.strip().str.upper()

        # Validar columnas necesarias
        if 'FECHA' not in atencion_data.columns or 'NOMBRE' not in atencion_data.columns or 'HORA' not in atencion_data.columns:
            raise ValueError("El archivo 'a.xls' debe contener las columnas: FECHA, NOMBRE, HORA.")

        # Convertir columnas al formato adecuado
        atencion_data['FECHA'] = pd.to_datetime(atencion_data['FECHA'], format='%d-%m-%Y', errors='coerce')
        atencion_data['HORA'] = pd.to_datetime(atencion_data['HORA'], format='%H:%M:%S', errors='coerce').dt.time

        # Verificar valores nulos
        atencion_data.dropna(subset=['FECHA', 'HORA'], inplace=True)

        # Ordenar datos
        atencion_data = atencion_data.sort_values(by=['FECHA', 'HORA'])

        # Obtener la lista única de personas
        nombres_unicos = atencion_data['NOMBRE'].unique()

        # Diccionario para convertir los días al español
        dias_en_espanol = {
            'Monday': 'Lunes',
            'Tuesday': 'Martes',
            'Wednesday': 'Miércoles',
            'Thursday': 'Jueves',
            'Friday': 'Viernes',
            'Saturday': 'Sábado',
            'Sunday': 'Domingo'
        }

        # Crear un ExcelWriter para guardar el archivo de salida
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            for nombre in nombres_unicos:
                # Filtrar datos por persona
                datos_persona = atencion_data[atencion_data['NOMBRE'] == nombre]

                # Resumir horas por día
                datos_resumidos = datos_persona.groupby('FECHA')['HORA'].apply(list).reset_index()

                # Crear columna de día en español
                datos_resumidos['DIA'] = datos_resumidos['FECHA'].dt.strftime('%A').map(dias_en_espanol)

                # Crear columnas de horas de entrada y salida
                datos_resumidos['HORA ENTRADA'] = datos_resumidos['HORA'].apply(lambda horas: horas[0].strftime('%H:%M:%S') if horas else None)
                datos_resumidos['HORA SALIDA'] = datos_resumidos['HORA'].apply(lambda horas: horas[-1].strftime('%H:%M:%S') if len(horas) > 1 else None)

                
                # Calcular horas extras y atraso
                def calcular_horas_extras_y_atraso(hora_entrada, hora_salida, dia):
                    # Configuración de horas estándar
                    horas_estandar = {
                        'Lunes': timedelta(hours=9),
                        'Martes': timedelta(hours=9),
                        'Miércoles': timedelta(hours=9),
                        'Jueves': timedelta(hours=9),
                        'Viernes': timedelta(hours=8),
                        'Sábado': timedelta(hours=4)
                    }

                    hora_inicio_jornada = datetime.strptime("09:00:00", "%H:%M:%S").time()

                    # Intervalos de horas extras
                    hora_extra_diurna_inicio = datetime.strptime("17:00:00", "%H:%M:%S").time()
                    hora_extra_diurna_fin = datetime.strptime("20:00:00", "%H:%M:%S").time()
                    hora_extra_nocturna_inicio = datetime.strptime("20:00:00", "%H:%M:%S").time()
                    hora_extra_nocturna_fin = datetime.strptime("22:00:00", "%H:%M:%S").time()

                    if not hora_entrada or not hora_salida:
                        return 0, 0, 0, 0  # Sin datos de horas

                    entrada = datetime.strptime(hora_entrada, "%H:%M:%S").time()
                    salida = datetime.strptime(hora_salida, "%H:%M:%S").time()

                    # Cálculo de atraso
                    atraso = max(timedelta(0), datetime.combine(datetime.min, entrada) - datetime.combine(datetime.min, hora_inicio_jornada))

                    # Cálculo de horas extra
                    horas_extra_diurna = 0
                    horas_extra_nocturna = 0

                    if salida > hora_extra_diurna_inicio:
                        if salida <= hora_extra_diurna_fin:
                            horas_extra_diurna += (datetime.combine(datetime.min, salida) - datetime.combine(datetime.min, hora_extra_diurna_inicio)).seconds
                        else:
                            horas_extra_diurna += (datetime.combine(datetime.min, hora_extra_diurna_fin) - datetime.combine(datetime.min, hora_extra_diurna_inicio)).seconds
                            if salida <= hora_extra_nocturna_fin:
                                horas_extra_nocturna += (datetime.combine(datetime.min, salida) - datetime.combine(datetime.min, hora_extra_nocturna_inicio)).seconds
                            else:
                                horas_extra_nocturna += (datetime.combine(datetime.min, hora_extra_nocturna_fin) - datetime.combine(datetime.min, hora_extra_nocturna_inicio)).seconds

                    # Calcular horas trabajadas y horas extra
                    jornada_estandar = horas_estandar.get(dia, timedelta(hours=0))

                    return atraso.total_seconds(), horas_extra_diurna, horas_extra_nocturna, jornada_estandar.total_seconds()

                # Aplicar la función de cálculo correctamente para cada fila
                resultados = datos_resumidos.apply(
                    lambda row: pd.Series(calcular_horas_extras_y_atraso(row['HORA ENTRADA'], row['HORA SALIDA'], row['DIA'])),
                    axis=1
                )

                # Asegurarse de que los resultados se asignen correctamente
                datos_resumidos[['ATRASO', 'HORA EXTRA DIURNA', 'HORA EXTRA NOCTURNA', 'HORAS TOTALES']] = resultados

                # Convertir los segundos en formato tiempo (hh:mm:ss)
                def convertir_a_tiempo(segundos):
                    horas = segundos // 3600
                    minutos = (segundos % 3600) // 60
                    segundos_restantes = segundos % 60
                    return f"{int(horas):02d}:{int(minutos):02d}:{int(segundos_restantes):02d}"

                # Convertir columnas a formato de tiempo
                datos_resumidos['ATRASO'] = datos_resumidos['ATRASO'].apply(convertir_a_tiempo)
                datos_resumidos['HORA EXTRA DIURNA'] = datos_resumidos['HORA EXTRA DIURNA'].apply(convertir_a_tiempo)
                datos_resumidos['HORA EXTRA NOCTURNA'] = datos_resumidos['HORA EXTRA NOCTURNA'].apply(convertir_a_tiempo)

                # Calcular horas totales trabajadas en formato hh:mm:ss
                def calcular_horas_totales(horas):
                    if len(horas) < 2:
                        return "00:00:00"
                    total_segundos = sum(
                        (datetime.combine(datetime.min, horas[i + 1]) - datetime.combine(datetime.min, horas[i])).seconds
                        for i in range(len(horas) - 1)
                    )
                    return convertir_a_tiempo(total_segundos)

                # Calcular las horas totales y agregarlas a la columna correspondiente
                datos_resumidos['HORAS TOTALES'] = datos_resumidos['HORA'].apply(calcular_horas_totales)
                # Agregar la columna INCONGRUENCIAS
                datos_resumidos['INCONGRUENCIAS'] = datos_resumidos.apply(
                    lambda row: "CHECKEAR ESTE DIA" if pd.isnull(row['HORA ENTRADA']) or pd.isnull(row['HORA SALIDA']) else None,
                    axis=1
                )


                # Eliminar columna original de lista de horas
                datos_resumidos.drop(columns=['HORA'], inplace=True)

                # Guardar en una hoja del archivo
                datos_resumidos.to_excel(writer, sheet_name=nombre[:31], index=False)

                # Ajustar columnas
                worksheet = writer.sheets[nombre[:31]]
                for col_num, col_name in enumerate(datos_resumidos.columns):
                    worksheet.set_column(col_num, col_num, max(len(col_name), 30))

        print(f"Archivo generado con éxito: {output_path}")

    except Exception as e:
        print(f"Ocurrió un error: {e}")


def seleccionar_archivo(titulo):
    """Abre un cuadro de diálogo para seleccionar un archivo."""
    root = Tk()
    root.withdraw()  # Oculta la ventana principal
    archivo_path = filedialog.askopenfilename(title=titulo, filetypes=[("Archivos Excel", "*.xls;*.xlsx")])
    return archivo_path


if __name__ == "__main__":
    # Pide al usuario seleccionar los archivos
    print("Selecciona el archivo de atención :")
    atencion_path = seleccionar_archivo("Selecciona el archivo de atención")

    print("Selecciona el archivo de control :")
    control_path = seleccionar_archivo("Selecciona el archivo de control")

    # Ruta de salida
    output_path = "Control planilla.xlsx"

    # Llama a la función para generar el archivo Excel
    generar_excel_por_persona(atencion_path, control_path, output_path)
