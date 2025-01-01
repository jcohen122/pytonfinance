# -*- coding: utf-8 -*-
"""
Created on Sat Dec 14 09:39:11 2024

@author: jaaco
"""

import polygon
import pandas as pd
from polygon import RESTClient
from datetime import datetime, timedelta
import time

def obtener_datos_financieros(api_key):
    # Configura el cliente
    client = RESTClient("dTKnjQGs9siG5siVnRqZvn6Tnqe44Ppv")
    datos_tickers = []
    
    try:
        # Obtener todos los tickers
        tickers = client.get_snapshot_all(market_type="stocks", include_otc=False)
        
        for ticker in tickers:
            ticker_symbol = ticker.ticker
            print(f"Procesando: {ticker_symbol}")
            
            try:
                # Datos básicos del snapshot
                info_ticker = {
                    'Símbolo': ticker_symbol,
                    'Precio_Actual': ticker.session.close if hasattr(ticker, 'session') else None,
                    'Volumen': ticker.session.volume if hasattr(ticker, 'session') else None,
                    'Precio_Alto_Día': ticker.session.high if hasattr(ticker, 'session') else None,
                    'Precio_Bajo_Día': ticker.session.low if hasattr(ticker, 'session') else None,
                }
                
                # Obtener detalles del ticker
                try:
                    detalles = client.get_ticker_details(ticker_symbol)
                    if detalles:
                        info_ticker.update({
                            'Nombre': detalles.name,
                            'Descripción': detalles.description,
                            'Market_Cap': detalles.market_cap,
                            'País': detalles.locale,
                            'Moneda': detalles.currency_name,
                            'Empleados': detalles.total_employees,
                            'Sector': detalles.sic_description
                            
                        })
                except Exception as e:
                    print(f"Error en detalles para {ticker_symbol}: {e}")
                
                # Obtener cierre anterior
        
                
            except Exception as e:
                print(f"Error procesando {ticker_symbol}: {e}")
                continue
            
            time.sleep(0.12)  # Pausa para respetar límites de API
            
    except Exception as e:
        print(f"Error general: {e}")
        return None
    
    return datos_tickers

def guardar_excel(datos, nombre_archivo='datos_tickers.xlsx'):
    """Guarda los datos en un archivo Excel con formato"""
    if not datos:
        print("No hay datos para guardar")
        return
    
    # Crear DataFrame
    df = pd.DataFrame(datos)
    
    # Crear Excel writer
    writer = pd.ExcelWriter(nombre_archivo, engine='xlsxwriter')
    
    # Guardar datos
    df.to_excel(writer, index=False, sheet_name='Datos')
    
    # Obtener workbook y worksheet
    workbook = writer.book
    worksheet = writer.sheets['Datos']
    
    # Formato para encabezados
    formato_encabezado = workbook.add_format({
        'bold': True,
        'bg_color': '#D8E4BC',
        'border': 1,
        'text_wrap': True
    })
    
    # Formato para números
    formato_numero = workbook.add_format({
        'num_format': '#,##0.00',
        'border': 1
    })
    
    # Formato para volumen
    formato_volumen = workbook.add_format({
        'num_format': '#,##0',
        'border': 1
    })
    
    # Aplicar formatos
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, formato_encabezado)
        
        # Ajustar ancho de columnas
        max_length = max(
            df[value].astype(str).apply(len).max(),
            len(value)
        )
        worksheet.set_column(col_num, col_num, max_length + 2)
        
        # Aplicar formato según el tipo de columna
        if 'Volumen' in value:
            for row in range(1, len(df) + 1):
                cell_value = df.iloc[row-1][value]
                if pd.notnull(cell_value):
                    worksheet.write(row, col_num, cell_value, formato_volumen)
        elif any(text in value for text in ['Precio', 'Cap']):
            for row in range(1, len(df) + 1):
                cell_value = df.iloc[row-1][value]
                if pd.notnull(cell_value):
                    worksheet.write(row, col_num, cell_value, formato_numero)
    
    writer.close()
    print(f"Datos guardados en {nombre_archivo}")

def main():
    print("Iniciando extracción de datos...")
    datos = obtener_datos_financieros("dTKnjQGs9siG5siVnRqZvn6Tnqe44Ppv")
    
    if datos:
        guardar_excel(datos)
        print(f"Se procesaron {len(datos)} tickers exitosamente")
    else:
        print("No se encontraron datos para procesar")

if __name__ == "__main__":
    main()
