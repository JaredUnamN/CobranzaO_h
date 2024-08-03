import pandas as pd # type: ignore
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox # type: ignore
import customtkinter # type: ignore
import numpy as np # type: ignore
import os
import openpyxl # type: ignore
from openpyxl.utils.dataframe import dataframe_to_rows # type: ignore
from openpyxl.utils import get_column_letter # type: ignore
from tkinter import ttk
import webbrowser
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side # type: ignore
from reportlab.lib.pagesizes import inch, landscape, portrait # type: ignore 
from reportlab.lib import colors # type: ignore
from reportlab.lib.units import inch # type: ignore
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer # type: ignore
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle # type: ignore
from datetime import datetime
from PIL import Image, ImageTk # type: ignore

# Ajustar la opción para imprimir el DataFrame completo sin truncar
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

class ImportadorArchivos:
    def __init__(self, master):
        # Crear la aplicación
        self.master = master
        self.master.title("APP REPORTE_CARTERA BY JARED NÚÑEZ")

        # Centrar la ventana en la pantalla
        window_width = 600
        window_height = 400
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x_coordinate = (screen_width / 2) - (window_width / 2)
        y_coordinate = (screen_height / 2) - (window_height / 2)
        self.master.geometry(f"{window_width}x{window_height}+{int(x_coordinate)}+{int(y_coordinate)}")

        self.dataframes = []

        # Crear frame principal
        frame_principal = tk.Frame(self.master)
        frame_principal.pack(fill="both", expand=True)

        # Crear frame izquierdo
        frame_izquierdo = tk.Frame(frame_principal)
        frame_izquierdo.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")
        frame_izquierdo.config(width=250, height=300)  # Ajustar tamaño del frame izquierdo

        # Cargar y mostrar la imagen en el frame izquierdo
        image = Image.open("C:/Users/angel/OneDrive/Escritorio/Trabajo/ClienteSoniaBonilla/COBRANZA REPORTE/A mi forma/360_F_585864419_kgIYUcDQ0yiLOCo1aRjeu7kRxndcoitz (1).jpg")
        image = image.resize((480, 600), Image.LANCZOS)  # Ajustar el tamaño de la imagen
        self.photo = ImageTk.PhotoImage(image)
        label_image = tk.Label(frame_izquierdo, image=self.photo)
        label_image.pack()

        # Crear frame derecho
        frame_derecho = tk.Frame(frame_principal)
        frame_derecho.grid(row=0, column=1, padx=20, pady=10, sticky="nsew")


        # Etiqueta para mostrar la cantidad de archivos importados
        welcome_label = customtkinter.CTkLabel(frame_derecho, text="Generador de Cartera", font=("Roboto", 20, "bold"), text_color="#8e44ad")
        welcome_label.pack(pady=5)
        
        self.label_archivos_importados = tk.Label(frame_derecho, text="Archivos importados: 0", font=("Arial", 12))
        self.label_archivos_importados.pack(pady=10)

        # Botón para importar archivos
        self.button_importar = customtkinter.CTkButton(frame_derecho, text="Importar Archivos", command=self.importar_archivos, fg_color='#8e44ad', hover_color='#9b59b6', text_color='white', font=("Arial", 14, "bold"))
        self.button_importar.pack(pady=10)

        # Botón de ubicación del reporte
        ubicacion_button = customtkinter.CTkButton(frame_derecho, text="Ubicación del reporte", command=self.seleccionar_carpeta_reporte, fg_color='#3498db', hover_color='#2980b9', text_color='white', font=("Arial", 14, "bold"))
        ubicacion_button.pack(pady=10)

        # Etiqueta para la carpeta del reporte
        self.carpeta_reporte_label = tk.Label(frame_derecho, text="Carpeta del reporte:", font=("Arial", 12))
        self.carpeta_reporte_label.pack(pady=10)

        # Cuadro de texto para el nombre del reporte
        self.nombre_reporte_entry = customtkinter.CTkEntry(frame_derecho, width=220, placeholder_text='Ingrese el nombre para el reporte', border_color='#8e44ad', font=("Arial", 12, "italic"))
        self.nombre_reporte_entry.pack(pady=10)

        # Botón para generar el reporte
        self.generar_button = customtkinter.CTkButton(frame_derecho, text="Generar reporte", state=tk.DISABLED, fg_color='#e74c3c', hover_color='#c0392b', text_color='white', font=("Arial", 14, "bold"), command=self.generar_reporte)
        self.generar_button.pack(pady=10)

        # Botón de instrucciones
        instrucciones_button = customtkinter.CTkButton(frame_derecho, text="Instrucciones", fg_color='#f39c12', hover_color='#e67e22', text_color='white', command=self.mostrar_instrucciones, font=("Arial", 14, "bold"))
        instrucciones_button.pack(pady=10)

    def mostrar_instrucciones(self):
        instrucciones = (
            "Bienvenido a esta aplicación: Query to Excel report - OH Pedidos by Jared Núñez\n\n"
            "Instrucciones para usar el programa:\n"
            "1. Recuerda que todos los archivos a importar deben estar en el formato 'Libro de Excel' '.xlsx y formateados con la macro correspondiente'\n"
            "2. Al presionar el botón 'Importar archivos', selecciona el archivo a procesar\n"
            "3. Después de importar, se mostrarán los archivos importados para su revisión\n"
            "4. Selecciona el botón 'Ubicación del reporte' para elegir una carpeta donde guardar el reporte final\n"
            "5. Llena el cuadro de texto con el nombre para el reporte y presiona 'Generar reporte'\n\n"
            "Cualquier problema o duda, comunícate con soporte técnico vía correo electrónico\n\n"
            "Correo de soporte técnico: nunezjared082@gmail.com\n"
        )
        messagebox.showinfo("Instrucciones", instrucciones)


    def importar_archivos(self):
        archivos = filedialog.askopenfilenames(title="Selecciona archivos Excel", filetypes=[("Archivos Excel", "*.xlsx")])

        if not archivos:
            print("No se seleccionaron archivos.")
            return

        # Limpiar lista de DataFrames antes de importar nuevos archivos
        self.dataframes = []

        # Leer cada archivo y almacenar en la lista
        for archivo in archivos:
            df = pd.read_excel(archivo)
            self.dataframes.append(df)

        # Actualizar etiqueta con la cantidad de archivos importados
        self.label_archivos_importados.config(text=f"Archivos importados: {len(self.dataframes)}")
        self.label_archivos_importados.config(text="Archivos importados:\n" + "\n".join(archivos))

        # Habilitar el botón para realizar acciones después de importar
        self.generar_button.configure(state=tk.NORMAL)
        self.master.geometry("1000x400")

    def seleccionar_carpeta_reporte(self):
        carpeta_reporte = filedialog.askdirectory()
        self.carpeta_reporte_label.config(text="Carpeta del reporte:\n" + carpeta_reporte)
        self.carpeta_reporte = carpeta_reporte
        self.master.geometry("1000x400")

    
    def contactar_soporte_tecnico(self):
        # Abre el navegador web predeterminado con la URL de Gmail
        webbrowser.open("https://mail.google.com/")
    

    def generar_reporte(self):
        if hasattr(self, 'carpeta_reporte'):
            # Obtener carpeta de salida de reporte
            nombre_reporte = self.nombre_reporte_entry.get()

            if not nombre_reporte:
                print("Por favor, ingresa un nombre para el reporte.")
                return

            # Crear la ruta completa para guardar el archivo Excel en el mismo directorio
            output_file_path = os.path.join(self.carpeta_reporte, f"{nombre_reporte}.xlsx")
            print(f"Generando reporte en: {output_file_path}")

            # Verificar si hay DataFrames para consolidar
            if not self.dataframes:
                print("No hay archivos para consolidar. Importa archivos antes de generar el reporte.")
                return


            df_clasificador = pd.concat(self.dataframes, ignore_index=True)
            #definamos que la fecha sea del tipo datetime


            df_clasificador['Fecha de vencimiento'] = pd.to_datetime(df_clasificador['Fecha de vencimiento'])
             #definimos hoy
            fecha_actual = pd.Timestamp.now()
            #Condiciones para clasificar los documentos vencidos
            condiciones_documentovencido = (
                (df_clasificador['Tipo'] == 'RF') &
                (df_clasificador['Fecha de vencimiento'] < fecha_actual) &
                (~df_clasificador['Código de cliente'].isin(['C-CCC980828IW0','C-SRF-LIVERPOOL', 'C-SRFC-AMAZON', 'C-SRFC-LINIO', 'C-SRFC-MERLIBRE', 'C-SRFC-OH-SHOPI', 'C-SRFC-WALMART', 'C-SRFC-MATRIZ', 'C-BOUTIQUE O.H.'])) &
                ((df_clasificador['Nº documento'] < 10000000) |
                ((df_clasificador['Nº documento'] > 16000000) & (df_clasificador['Nº documento'] < 17000000)))
            )

            # Crear columna para el documento vencido
            df_clasificador['Documento Vencido'] = np.where(condiciones_documentovencido, df_clasificador['Saldo vencido'], np.nan)

            condiciones_burodecredito = (
                (df_clasificador['Código de cliente'].isin(['C-CCC980828IW0']))
            )

            # Crear columna para los de buro de credito
            df_clasificador['Buro de Credito'] = np.where(condiciones_burodecredito, df_clasificador['Saldo vencido'], np.nan)

            # Calcula los días desde la fecha de vencimiento hasta hoy
            df_clasificador['Días desde vencimiento'] = (fecha_actual - df_clasificador['Fecha de vencimiento']).dt.days

            # Condiciones para clasificar los documentos sin letra de cambio
            condiciones_sinletradecambio = (
                (df_clasificador['Código de cliente'].isin(['C-DLI931201MI9', 'C-CCI8111293TA', 'C-TSO991022PB6'])) &
                (df_clasificador['Tipo'] == 'RF') &
                (df_clasificador['Fecha de vencimiento'] > fecha_actual) &
                (df_clasificador['Nº documento'] > 1250) &  
                (df_clasificador['Nº documento'] < 10000000)
            )

            df_clasificador['Sin letra de cambio'] = np.where(condiciones_sinletradecambio, df_clasificador['Saldo vencido'], np.nan)


            condiciones_descporev= (
                (df_clasificador['Nº documento']>10000000) &
                (df_clasificador['Nº documento']<11000000)
            )

            df_clasificador['Descuentos por revisar/Registrar'] = np.where(condiciones_descporev, df_clasificador['Saldo vencido'], np.nan)

            condiciones_devoluciones = (
                (df_clasificador['Nº documento']>11000000) &
                (df_clasificador['Nº documento']<12000000)
            )

            df_clasificador['Devoluciones por recuperar - firme'] = np.where(condiciones_devoluciones, df_clasificador['Saldo vencido'], np.nan)

            condiciones_notasdec = (
                (df_clasificador['Tipo'] == 'RC') &
                (df_clasificador['Nº documento']<10000000) 
            )

            df_clasificador['Notas de crédito por aplicar/revisar'] = np.where(condiciones_notasdec, df_clasificador['Saldo vencido'], np.nan)

            condiciones_contado = (
                (df_clasificador['Tipo'] == 'RF') &
                (df_clasificador['Código de cliente'].isin(['C-SRF-LIVERPOOL', 'C-SRFC-AMAZON', 'C-SRFC-LINIO', 'C-SRFC-MERLIBRE', 'C-SRFC-OH-SHOPI', 'C-SRFC-WALMART', 'C-SRFC-MATRIZ', 'C-BOUTIQUE O.H.']))&
                (df_clasificador['Fecha de vencimiento'] < fecha_actual) 
            )

            df_clasificador['Clientes contado con saldo'] = np.where(condiciones_contado, df_clasificador['Saldo vencido'], np.nan)

            condiciones_pr = (
                (df_clasificador['Tipo'] == 'PR') &
                (df_clasificador['Saldo vencido']!=0)
            )

            df_clasificador['Cobros no aplicados'] = np.where(condiciones_pr, df_clasificador['Saldo vencido'], np.nan)

            condiciones_carteracorriente = (
                (df_clasificador['Tipo'] == "RF") &
                (df_clasificador['Fecha de vencimiento'] >= fecha_actual) &
                (
                    (df_clasificador['Nº documento'] < 10000000) & 
                    (~df_clasificador['Código de cliente'].isin(['C-DLI931201MI9', 'C-CCI8111293TA', 'C-TSO991022PB6'])) |
                    ((df_clasificador['Nº documento'] > 16000000) & (df_clasificador['Nº documento'] < 17000000))
                )
            )

            df_clasificador['Cartera corriente'] = np.where(condiciones_carteracorriente, df_clasificador['Saldo vencido'], np.nan)

            condiciones_todas = (
                condiciones_documentovencido |
                condiciones_burodecredito |
                condiciones_sinletradecambio |
                condiciones_descporev |
                condiciones_devoluciones |
                condiciones_notasdec |
                condiciones_contado |
                condiciones_carteracorriente |
                condiciones_pr
            )

            df_clasificador['Otras partidas'] = np.where(~condiciones_todas, df_clasificador['Saldo vencido'], np.nan)

            condiciones_total = (
                condiciones_carteracorriente 
            )

            df_clasificador['Total NO Corriente'] = np.where(~condiciones_total, df_clasificador['Saldo vencido'], np.nan)

            df_clasificador['Total Cartera']=df_clasificador['Saldo vencido']

            # Seleccionamos solo las columnas relevantes a partir de 'Documento Vencido'
            columnas_para_status = ['Documento Vencido', 'Buro de Credito', 'Sin letra de cambio',
                                    'Descuentos por revisar/Registrar', 'Devoluciones por recuperar - firme',
                                    'Notas de crédito por aplicar/revisar', 'Clientes contado con saldo', 'Cobros no aplicados',
                                    'Cartera corriente', 'Otras partidas']

            # Sumar los valores de cada columna
            df_totales_status = df_clasificador[columnas_para_status].sum().reset_index()
            df_totales_status.columns = ['Status', 'Total']

            # Verificar si 'Cartera corriente' está duplicado
            if df_totales_status[df_totales_status['Status'] == 'Cartera corriente'].shape[0] > 1:
                raise ValueError("La columna 'Cartera corriente' está duplicada en el DataFrame.")

            cartera_total = df_totales_status['Total'].sum()

            # Calcular el subtotal sin incluir 'Cartera corriente'
            subtotal = df_totales_status[df_totales_status['Status'] != 'Cartera corriente']['Total'].sum()

            # Crear filas de subtotal, cartera corriente y total
            fila_subtotal = pd.DataFrame([{'Status': 'Subtotal', 'Total': subtotal}])
            fila_cartera_corriente = pd.DataFrame([{'Status': 'Cartera corriente', 'Total': df_totales_status[df_totales_status['Status'] == 'Cartera corriente']['Total'].values[0]}])
            fila_total = pd.DataFrame([{'Status': 'Total', 'Total': cartera_total}])

            # Concatenar las filas adicionales
            df_totales_status = pd.concat([df_totales_status[df_totales_status['Status'] != 'Cartera corriente'], fila_subtotal, fila_cartera_corriente, fila_total], ignore_index=True)

            # Calcular los porcentajes
            df_totales_status['%'] = df_totales_status['Total'] / cartera_total
            df_totales_status['%'] = df_totales_status['%'].apply(lambda x: f"{x:.2%}")

            # Separar las filas especiales ('Total', 'Cartera corriente', 'Subtotal') del resto del DataFrame
            df_especiales = df_totales_status[df_totales_status['Status'].isin(['Total', 'Cartera corriente', 'Subtotal'])]
            df_resto = df_totales_status[~df_totales_status['Status'].isin(['Total', 'Cartera corriente', 'Subtotal'])]

            # Ordenar el resto del DataFrame por 'Total' de menor a mayor
            df_resto_ordenado = df_resto.sort_values('Total', ascending=False)

            # Concatenar las partes para obtener el DataFrame final
            df_totales_status_ordenado = pd.concat([df_resto_ordenado, df_especiales]).reset_index(drop=True)

            # Reordenar las filas para que Subtotal, Cartera corriente y Total estén al final
            orden = ['Documento Vencido', 'Buro de Credito', 'Sin letra de cambio', 'Descuentos por revisar/Registrar',
                    'Devoluciones por recuperar - firme', 'Notas de crédito por aplicar/revisar', 'Clientes contado con saldo', 
                    'Cobros no aplicados', 'Otras partidas', 'Subtotal', 'Cartera corriente', 'Total']
            df_totales_status['order'] = df_totales_status['Status'].apply(lambda x: orden.index(x))
            df_totales_status = df_totales_status.sort_values('order').drop('order', axis=1).reset_index(drop=True)



            columnas_a_sumar = ['Saldo vencido', 'Documento Vencido', 'Buro de Credito', 'Sin letra de cambio',
                                'Descuentos por revisar/Registrar', 'Devoluciones por recuperar - firme',
                                'Notas de crédito por aplicar/revisar', 'Clientes contado con saldo', 'Cobros no aplicados',
                                'Cartera corriente', 'Otras partidas', 'Total NO Corriente']

            for columna in columnas_a_sumar:
                df_clasificador[columna] = df_clasificador[columna].astype(float)

            # Agrupar por 'Cliente' y 'Nombre del cliente', y sumar las columnas especificadas
            df_clientes = df_clasificador.groupby(['Código de cliente','Nombre del cliente']).agg({col: 'sum' for col in columnas_a_sumar}).reset_index()

            # Eliminar la columna 'Días de vencimiento' si existe
            if 'Días de vencimiento' in df_clientes.columns:
                df_clientes.drop(columns=['Días de vencimiento'], inplace=True)
            # Calcular los totales de cada columna en 'columnas_a_sumar'
            totales = df_clientes[columnas_a_sumar].sum()

            # Identificar las columnas de tipo float
            columnas_float = [col for col in df_clientes.columns if df_clientes[col].dtype == 'float64']

            # Calcular los totales de cada columna de tipo float
            totales = df_clientes[columnas_float].sum()

            # Crear un diccionario con los totales, asegurándose de incluir valores nulos o strings vacíos para las columnas no sumables
            fila_totales = {col: '' for col in df_clientes.columns}  # Inicializar todas las columnas con strings vacíos o nulos
            fila_totales['Código de cliente'] = 'Total'  # Ajustar según sea necesario
            for columna in columnas_float:
                fila_totales[columna] = totales[columna]

            # Convertir el diccionario a DataFrame para poder usar _append con ignore_index=True
            df_fila_totales = pd.DataFrame([fila_totales])
            
            df_clientes = df_clientes.sort_values(by='Total NO Corriente', ascending=False)

            # Agregar la fila de totales al final de 'df_clientes'
            df_clientes = df_clientes._append(df_fila_totales, ignore_index=True)

            for columna in df_clientes.columns:
                df_clientes[columna] = df_clientes[columna].replace(0, np.nan)


            df_clasificador['Año de vencimiento'] = df_clasificador['Fecha de vencimiento'].dt.year

            # Agrupar por 'Cliente' y 'Nombre del cliente', y sumar los documentos vencidos por año
            df_documento_vencido_por_ano = df_clasificador.groupby(['Código de cliente', 'Nombre del cliente', 'Año de vencimiento']).agg({'Documento Vencido': 'sum'}).reset_index()

            # Crear un dataframe pivoteado para tener los años como columnas
            df_pivot = df_documento_vencido_por_ano.pivot(index=['Código de cliente', 'Nombre del cliente'], columns='Año de vencimiento', values='Documento Vencido').fillna(0)

            # Convertir índices 'Código de cliente' y 'Nombre del cliente' de nuevo en columnas
            df_pivot.reset_index(inplace=True)

            # Asegúrate de que las columnas numéricas sean de tipo float
            columnas_float_pivot = [col for col in df_pivot.columns if df_pivot[col].dtype in ['float64', 'int64']]

            # Suma de los totales por columnas
            totales_pivot = df_pivot[columnas_float_pivot].sum()
            fila_totales_pivot = {col: '' for col in df_pivot.columns}
            fila_totales_pivot['Código de cliente'] = 'Total'

            for columna in columnas_float_pivot:
                fila_totales_pivot[columna] = totales_pivot[columna]

            df_fila_totales_pivot = pd.DataFrame([fila_totales_pivot])
            df_pivot = df_pivot._append(df_fila_totales_pivot, ignore_index=True)

            # Suma de todas las columnas numéricas para el total de documentos vencidos
            df_pivot['Total Docs Vencidos'] = df_pivot[columnas_float_pivot].sum(axis=1)

            

            # Paso 1: Separar la fila de totales
            df_totales = df_pivot[df_pivot['Código de cliente'] == 'Total']
            df_sin_totales = df_pivot[df_pivot['Código de cliente'] != 'Total']

            df_sin_totales_filtrado = df_sin_totales[(df_sin_totales['Total Docs Vencidos'] != 0) & (~df_sin_totales['Total Docs Vencidos'].isna())]

            # Paso 2: Ordenar el DataFrame sin la fila de totales
            df_ordenado = df_sin_totales_filtrado.sort_values(by='Total Docs Vencidos', ascending=False)

            # Paso 3: Añadir la fila de totales al final del DataFrame ordenado
            df_final_pivot = pd.concat([df_ordenado, df_totales], ignore_index=True)            

            df_final_pivot = df_final_pivot.sort_values(by='Total Docs Vencidos', ascending=False)

            for columna in df_final_pivot.columns:
                df_final_pivot[columna] = df_final_pivot[columna].replace(0, np.nan)




            
#Escribir en excel-----------------------------------------------------------------------------------------
            with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
                df_clasificador['Fecha de vencimiento'] = df_clasificador['Fecha de vencimiento'].dt.strftime('%d/%m/%Y')
                df_clasificador['Fecha de contabilización'] = df_clasificador['Fecha de contabilización'].dt.strftime('%d/%m/%Y')

                df_clasificador.to_excel(writer, sheet_name='Clasificador', index=False, startrow=4)
                df_totales_status_ordenado.to_excel(writer, sheet_name='RESUMENES', index=False, startrow=4, startcol=1)

                final_clientes = len(df_clientes) + 25
                df_clientes.to_excel(writer, sheet_name='RESUMENES',index=False, startrow=25, startcol=1)
                df_final_pivot.to_excel(writer, sheet_name='RESUMENES', index=False, startrow=final_clientes+6, startcol=1)
                workbook = writer.book
                worksheet_clasificador = writer.sheets['Clasificador']
    
                inicio_pivot = final_clientes + 6

                currency_format = workbook.add_format({'num_format': '"$"#,##0.00'})


                header_clasificador_format = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#DBFDFF',  # Fondo azul claro en hexadecimal
                    'font_color': 'black',  # Letras negras
                    'border': 0,  # Bordeado de color negro
                    'align': 'center',
                    'valign' : 'vcenter',
                    'text_wrap':True
                })

                for col_num, value in enumerate(df_clasificador.columns.values):
                    worksheet_clasificador.write(4, col_num, value, header_clasificador_format)

                # Ajustar el ancho de las columnas automáticamente
                for col_num, col in enumerate(df_clasificador.columns):
                    max_len = max(
                        df_clasificador[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 2  # Ajuste adicional para el padding
                    adjusted_len =min(max_len,26)
                    worksheet_clasificador.set_column(col_num, col_num, adjusted_len, workbook.add_format({'text_wrap': True}))


                subtotal_format = workbook.add_format({
                    'bold': True,
                    'font_color': 'black',
                    'bg_color': '#E1E1E1',  # Fondo gris claro en hexadecimal
                    'align': 'center',
                    'valign': 'vcenter',
                    'num_format': '#,##0.00'
                })

                # Calcular subtotales y escribirlos en la fila 3
                row = 2  # Fila 3 en índice 0
                for col_num, column in enumerate(df_clasificador.columns):
                    if df_clasificador[column].dtype in [np.float64,]:  # Solo sumar columnas numéricas
                        subtotal = df_clasificador[column].sum(skipna=True)
                        worksheet_clasificador.write(row, col_num, subtotal, subtotal_format)
                    else:
                        worksheet_clasificador.write(row, col_num, "", subtotal_format)            
                # Formato para las fechas
                date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                number_format = workbook.add_format({'num_format': '$#,##0.00'})
                worksheet_clasificador.set_row(4, 58)
                worksheet_clasificador.set_column('F:F', 18)
                worksheet_clasificador.set_column('E:E', 18)
                worksheet_clasificador.set_column('I:I', 15,number_format)
                worksheet_clasificador.set_column('J:J', 12,number_format)
                worksheet_clasificador.set_column('K:K', 12)
                worksheet_clasificador.set_column('L:L', 15,number_format)
                worksheet_clasificador.set_column('M:M', 20,number_format)
                worksheet_clasificador.set_column('N:N', 21,number_format)
                worksheet_clasificador.set_column('O:O', 18,number_format)
                worksheet_clasificador.set_column('P:P', 13,number_format)
                worksheet_clasificador.set_column('Q:Q', 13,number_format)
                worksheet_clasificador.set_column('R:R', 13,number_format)
                worksheet_clasificador.set_column('S:S', 13,number_format)
                worksheet_clasificador.set_column('T:T', 15,number_format)
                # Formato para celdas especifícas
                worksheet_clasificador.write(0, 0, 'REPORTE CLASIFICADOR', workbook.add_format({'bold': True, 'font_size': 23}))
                worksheet_clasificador.write(1, 0, 'Fecha de generación:', workbook.add_format({'bold': True}))
                worksheet_clasificador.write(1, 1, pd.Timestamp.now(), date_format)
                worksheet_clasificador.set_row(0, 30)

                header_clasificador_saldo = workbook.add_format({
                    'bold': True,
                    'font_size': 14,
                    'bg_color': '#181818',#negro 
                    'font_color': 'white',  # Letras blancas
                    'border': 0,  # Bordeado de color negro
                    'align': 'center',
                    'valign' : 'vcenter',
                    'text_wrap':True
                })

                worksheet_clasificador.write('H5','Saldo', header_clasificador_saldo)


                header_clasificador_docvencido = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#C5E0E1',  # Fondo azul claro en hexadecimal
                    'font_color': 'black',  # Letras negras
                    'border': 0,  # Bordeado de color negro
                    'align': 'center',
                    'valign' : 'vcenter',
                    'text_wrap':True
                })

                worksheet_clasificador.write('I5','DOCUMENTO VENCIDO', header_clasificador_docvencido)
                
                header_buro_credito = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#B3D8D9',  # Fondo azul claro en hexadecimal
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('J5','BURO DE CREDITO', header_buro_credito)

                header_dias_vencimiento = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#BCE4E5',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('K5','DIAS VENCIDOS', header_dias_vencimiento)

                header_sin_letra_cambio = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#C3EEF0',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('L5','SIN LETRA DE CAMBIO', header_sin_letra_cambio)

                header_descuentos_revisar = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#ADD8E6',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('M5','DESCUENTOS POR REVISAR/REGISTRAR', header_descuentos_revisar)

                header_devoluciones_revisar = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#CEFDFF',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('N5','Devoluciones por recuperar - firme', header_devoluciones_revisar)

                header_notas_credito_revisar = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#F3F3F3',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('O5','NOTAS DE CRETIDO POR REVISAR/APLICAR', header_notas_credito_revisar)

                header_contado = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#DBDBDB',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('P5','Clientes contado con saldo', header_contado)

                header_cobros_no_aplicados = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#E1E0E0',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('Q5','COBROS NO APLICADOS', header_cobros_no_aplicados)

                header_cartera_corriente = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#ECECEC',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('R5','CARTERA CORRIENTE', header_buro_credito)

                header_otras_partidas = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#95EAE1',
                    'font_color': 'black',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('S5','OTRAS PARTIDAS', header_otras_partidas)

                header_total = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#238177',
                    'font_color': 'white',
                    'border': 0,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet_clasificador.write('T5','TOTAL', header_total)

                worksheet_resumenes = writer.sheets['RESUMENES']

                header_format = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#000000',  # Fondo negro
                    'font_color': 'FFFFFF',  # Letras blancas
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })

                for col_num, value in enumerate(df_totales_status.columns.values):
                    worksheet_resumenes.write(4, col_num + 1, value, header_format)


                # Formato de celdas
                subtotal_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#E1E1E1',  # Fondo gris claro
                    'num_format': '#,##0.00',
                    'align': 'center',
                    'valign': 'vcenter'
                })

                total_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#CCCCCC',  # Fondo gris
                    'num_format': '#,##0.00',
                    'align': 'center',
                    'valign': 'vcenter'
                })

                worksheet_resumenes.write(14, 1, 'SUBTOTAL', subtotal_format)
                valor_subtotal = df_totales_status.at[9, 'Total']
                worksheet_resumenes.write(14,2,valor_subtotal,subtotal_format)
                valor_total = df_totales_status.at[11, 'Total']
                worksheet_resumenes.write(16, 2, valor_total, total_format)
                valor_porcentaje = df_totales_status.at[9, '%']
                worksheet_resumenes.write(14, 3, valor_porcentaje, subtotal_format)
                valor_porcentaje2 = df_totales_status.at[11, '%']
                worksheet_resumenes.write(16, 3, valor_porcentaje2, total_format)
                worksheet_resumenes.write(16, 1, 'TOTAL', total_format)

                # Ajustar el ancho de las columnas automáticamente
                for col_num, col in enumerate(df_totales_status.columns):
                    max_len = max(
                        df_totales_status[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 2  # Ajuste adicional para el padding
                    worksheet_resumenes.set_column(col_num + 1, col_num + 1, max_len)

                # Aplicar formato a columnas de totales y porcentajes
                currency_format = workbook.add_format({'num_format': '"$"#,##0.00'})
                percentage_format = workbook.add_format({'num_format': '0.00%'})

                inicio_status=2
                finalstatus=2+len(df_totales_status)


                worksheet_resumenes.set_column('C:C', 37, currency_format)
                worksheet_resumenes.set_column('D:D', 12, percentage_format)

                # Ajustar el ancho de las columnas automáticamente
                for col_num, col in enumerate(df_clientes.columns):
                    max_len = max(
                        df_clientes[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 2  # Ajuste adicional para el padding
                    worksheet_resumenes.set_column(col_num + 1, col_num + 1, max_len)
                worksheet_resumenes.set_row(25, 30)


                
                worksheet_resumenes.set_column('F:F', 18, currency_format)
                worksheet_resumenes.set_column('G:G', 22, currency_format)
                worksheet_resumenes.set_column('H:H', 21, currency_format)
                worksheet_resumenes.set_column('I:I', 21, currency_format)
                worksheet_resumenes.set_column('C:Q', 21, currency_format)

                # Agregar títulos y encabezados
                worksheet_resumenes.write(0, 1, 'SERVICIOS INTERNACIONALES DE MEXICO SA DE CV', workbook.add_format({'bold': True, 'font_size': 14}))
                worksheet_resumenes.write(1, 1, 'STATUS GENERAL DE CARTERA', workbook.add_format({'bold': True, 'font_size': 14}))
                
                # Fecha
                date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                worksheet_resumenes.write(1, 3, pd.Timestamp.now(), date_format)



                worksheet_resumenes.write(21, 1, 'SERVICIOS INTERNACIONALES DE MEXICO SA DE CV', workbook.add_format({'bold': True, 'font_size': 14}))
                worksheet_resumenes.write(22, 1, 'INTEGRACION DE CARTERA POR CLIENTE', workbook.add_format({'bold': True, 'font_size': 14}))

                # Fecha
                date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                worksheet_resumenes.write(22, 3, pd.Timestamp.now(), date_format)

                header_format = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'bg_color': '#000000',  # Fondo negro
                    'font_color': 'FFFFFF',  # Letras blancas
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True
                })

                

                for col_num, value in enumerate(df_clientes.columns.values):
                    worksheet_resumenes.write(25, col_num + 1, value, header_format)


                worksheet_resumenes.write(final_clientes+3, 1, 'SERVICIOS INTERNACIONALES DE MEXICO SA DE CV', workbook.add_format({'bold': True, 'font_size': 14}))
                worksheet_resumenes.write(final_clientes+4, 1, 'CLIENTES CON DOCUMENTOS VENCIDOS', workbook.add_format({'bold': True, 'font_size': 14}))

                for col_num, value in enumerate(df_pivot.columns.values):
                    worksheet_resumenes.write(final_clientes+6, col_num + 1, value, header_format)

                # Fecha
                date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                worksheet_resumenes.write(final_clientes+4, 3, pd.Timestamp.now(), date_format)

                row_index_excel = final_clientes
                row_index=len(df_clientes)-1

                for col_num, column_name in enumerate(df_clientes.columns):
                    value = df_clientes.at[row_index, column_name]  # Ajusta el índice para acceder al DataFrame correctamente
                    worksheet_resumenes.write(row_index_excel, col_num+1, value, total_format)  # -1 porque las filas en xlsxwriter comienzan en 0

                #Dal formato a los totales de los reportes
                worksheet_resumenes.write(final_clientes, 1, 'TOTAL', total_format)
                
                # Escribir datos de df_final_pivot
                inicio_pivot_excel = inicio_pivot + len(df_final_pivot) - 1
                inicio_pivot = len(df_final_pivot) - 1

                for col_num, column_name in enumerate(df_final_pivot.columns):
                    value = df_final_pivot.at[inicio_pivot, column_name]
                    if pd.isna(value) or pd.isna(value):
                        value = 0  # Reemplazar NaN/Inf con 0 o cualquier otro valor por defecto
                    worksheet_resumenes.write(inicio_pivot_excel + 1, col_num + 1, value, total_format)

                worksheet_resumenes.set_column('B:B', 55) 
                worksheet_resumenes.set_column('C:C', 37, currency_format) 

            print(f"Reporte generado exitosamente en: {output_file_path}")
            dataframes_presentacion = [df_totales_status_ordenado,df_final_pivot,df_clientes]
            nombres = ['Status Cartera', 'Documentos vencidos', 'Status por cliente']

            output_file_path_pdf = output_file_path.replace('.xlsx', '.pdf')


            def formatear_moneda(valor):
                try:
                    if pd.isna(valor):  # Verifica si el valor es NaN antes de intentar formatearlo
                        return ''  # Retorna una cadena vacía para valores NaN
                    else:
                        return "${:,.2f}".format(valor)
                except (ValueError, TypeError) as e:
                    print(f"No se pudo convertir el valor: {valor}")
                    return valor

            # Paso 1: Aplicar la función de formato a las columnas numéricas
            for columna in columnas_float_pivot + ['Total Docs Vencidos']:
                df_final_pivot[columna] = df_final_pivot[columna].apply(formatear_moneda)



            elements = []
            max_col_width = 22 * inch  # Ancho máximo de columna
            row_height = 18  # Altura estimada de cada fila (en puntos)

            # Calcular el tamaño máximo necesario
            max_width = 0
            max_height = 0

            styles = getSampleStyleSheet()

            header_style = ParagraphStyle(
                'Header',
                fontName='Helvetica-Bold',
                fontSize=16,
                leading=14,
                alignment=1,  # Centramos el texto
                wordWrap='CJK',  # Habilitamos el ajuste de línea
                textColor=colors.white  # Texto blanco
            )

            contador_dedatas = 0

            for df, nombre in zip(dataframes_presentacion, nombres):
                # Eliminar valores NaN y formatear datos
                contador_dedatas += 1
                df = df.fillna('')
                formatted_data = [[Paragraph(str(col), header_style) for col in df.columns]]  # Asegurarse de que todos los valores sean cadenas

                for row in df.itertuples(index=False, name=None):
                    formatted_row = []
                    for val, col in zip(row, df.columns):
                        formatted_val = formatear_moneda(val) if isinstance(val, (int, float)) and isinstance(col, str) and '%' not in col else val
                        formatted_row.append(formatted_val)
                    formatted_data.append(formatted_row)

                # Ajustar anchos de columna basado en contenido con límite máximo
                col_widths = []
                for col in df.columns:
                    if contador_dedatas == 1:
                        # Incrementar el valor de max_len para el primer DataFrame
                        max_len = max(df[col].astype(str).apply(len).max(), len(str(col))) + 17  # +6 para más margen
                    elif contador_dedatas == 2:
                        max_len = max(df[col].astype(str).apply(len).max(), len(str(col))) + 3
                        if max_len < 10:
                            max_len = 12
                    elif contador_dedatas == 3:
                        max_len = max(df[col].astype(str).apply(len).max(), len(str(col))) +1  # +2 para un pequeño margen
                        if col in ['Descuentos por revisar/Registrar', 'Devoluciones por recuperar - firme', 'Notas de crédito por aplicar/revisar', 'Clientes contado con saldo']:
                            max_len = 24
                    col_width = min(max_len * 7, max_col_width)  # Limitar el ancho máximo
                    col_widths.append(col_width)

                # Calcular el tamaño de página requerido para este DataFrame
                table_width = sum(col_widths)
                table_height = row_height * (len(formatted_data) + 1)  # +1 para la fila del encabezado

                # Actualizar el tamaño máximo necesario
                if table_width > max_width:
                    max_width = table_width
                if table_height > max_height:
                    max_height = table_height

                # Crear la tabla del DataFrame
                table = Table(formatted_data, colWidths=col_widths)

                # Aplicar estilos generales
                general_style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.black),  # Fondo gris para los encabezados
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Texto blanco para los encabezados
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alineación centrada para los encabezados
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Fuente Helvetica-Bold para los encabezados
                    ('FONTSIZE', (0, 0), (-1, 0), 18),  # Tamaño de fuente para los encabezados
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Espacio inferior para los encabezados
                    ('TEXTCOLOR', (0, 1), (-1, -2), colors.black),  # Texto negro para el contenido
                    ('ALIGN', (0, 1), (-1, -2), 'CENTER'),  # Alineación centrada para el contenido
                    ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),  # Fuente Helvetica para el contenido
                    ('FONTSIZE', (0, 1), (-1, -2), 10),  # Tamaño de fuente para el contenido
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Cuadrícula negra
                    ('BACKGROUND', (0, -1), (-1, -1), colors.grey),  # Fondo gris para la última fila
                    ('TEXTCOLOR', (0, -1), (-1, -1), colors.whitesmoke),  # Texto blanco para la última fila
                    ('ALIGN', (0, -1), (-1, -1), 'CENTER'),  # Alineación centrada para la última fila
                    ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),  # Fuente Helvetica-Bold para la última fila
                    ('FONTSIZE', (0, -1), (-1, -1), 14),  # Tamaño de fuente para la última fila
                    ('BOTTOMPADDING', (0, -1), (-1, -1), 12),  # Espacio inferior para la última fila
                ])

                # Aplicar el estilo de la antepenúltima fila si es el primer DataFrame
                if contador_dedatas == 1 and len(formatted_data) > 2:
                    general_style.add('BACKGROUND', (0, -3), (-1, -3), colors.lightgrey)
                    general_style.add('FONTNAME', (0, -3), (-1, -3), 'Helvetica-Bold')
                    #actualiza el tamaño de la fuente
                    general_style.add('FONTSIZE', (0, 0), (-1, -1), 14)
                    general_style.add('BOTTOMPADDING', (0, 0), (-1, -1), 16)
                    general_style.add('TOPPADDING', (0, 0), (-1, -1), 16)
                elif contador_dedatas == 2 and len(formatted_data) > 2:
                    general_style.add('FONTNAME', (0, -3), (-1, -3), 'Helvetica-Bold')
                    general_style.add('FONTSIZE', (0, 0), (-1, -1), 12)
                    general_style.add('BOTTOMPADDING', (0, 0), (-1, -1), 10)
                    general_style.add('TOPPADDING', (0, 0), (-1, -1), 10)
                elif contador_dedatas == 3 and len(formatted_data) > 2:
                    general_style.add('FONTNAME', (0, -3), (-1, -3), 'Helvetica-Bold')
                    general_style.add('FONTSIZE', (0, 0), (-1, -1), 12)
                    general_style.add('BOTTOMPADDING', (0, 0), (-1, -1), 4)
                    general_style.add('TOPPADDING', (0, 0), (-1, -1), 4)
                table.setStyle(general_style)

                elements.append(Table([[nombre]], colWidths=[len(nombre) * 10]))
                elements[-1].setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.white),  # Fondo blanco
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),  # Texto negro
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Centrar texto
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),  # Fuente Helvetica-Bold
                    ('FONTSIZE', (0, 0), (-1, -1), 28),  # Tamaño de fuente más grande
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 25),  # Más espacio inferior
                ]))
                elements.append(table)
                elements.append(PageBreak())

            # Crear el PDF con el tamaño máximo
            pdf = SimpleDocTemplate(output_file_path_pdf, pagesize=(max_width, max_height + 300))

            # Título principal en la primera página
            style = styles['Title']
            style.fontSize = 35  # Increase font size to 50 points

            fecha_hoy = datetime.now().strftime("%d/%m/%Y")
            titulo_principal = f"SERVICIOS INTERNACIONALES S.A de CV<br/><br/><br/>REPORTE DE CARTERA A DIA DE: {fecha_hoy}"
            titulo_paragraph = Paragraph(titulo_principal, style)

            elements.insert(0, titulo_paragraph)
            elements.insert(1, Spacer(1, 60))  # Espacio entre el título y la tabla

            pdf.build(elements)
            print(f"PDF guardado en: {output_file_path_pdf}")
        else:
            print("Selecciona una carpeta para el reporte antes de generarlo.")





customtkinter.set_appearance_mode("system")
customtkinter.set_default_color_theme("green")
root = customtkinter.CTk()
app = ImportadorArchivos(root)
root.mainloop()
