import pyodbc
import logging
from pyrfc import Connection
from datetime import datetime, date
import schedule
import time

logging.basicConfig(level=logging.ERROR)

def job():
    fecha_actual = datetime.now()

    #! Inicializar todas las variables 
    VBELN = None
    VBELN_VF = None
    GUIA_DESPACHO = None
    FECH_ENT_OP = None
    KWMENG_PV = None
    ORDEN_SALIDA = None
    LOTE_DESPACHADO = None
    OBSERVACION = None
    FECHARECEP = None
    LFDAT = None

    #! Configuración de la conexión MSSQL
    server = '10.8.0.51'
    database = 'prueba_xml_3'
    username = 'xxxxxx'
    password = 'xxxxxxxx'
    conn_sql_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    conn_sql = pyodbc.connect(conn_sql_str)

    #! Creación del cursor SQL
    cursor_sql = conn_sql.cursor()

    #! Consulta para la tabla "detalle"
    cursor_sql.execute("SELECT QtyItem, VlrCodigo FROM detalle WHERE Fecha_ingreso_xml >= ? AND Fecha_ingreso_xml < DATEADD(day, 1, ?)", fecha_actual.date(), fecha_actual.date())
    results = cursor_sql.fetchall()

    for row in results:
        LOTE_DESPACHADO = row.VlrCodigo
        KWMENG_PV = row.QtyItem

    #! Consulta para la tabla "factura"
    cursor_sql.execute("SELECT Folio, FE  FROM factura WHERE Fecha_ingreso_xml >= ? AND Fecha_ingreso_xml < DATEADD(day, 1, ?)", fecha_actual.date(), fecha_actual.date())
    results = cursor_sql.fetchall()

    for row in results:
        VBELN_VF = row.Folio
        FECH_ENT_OP = row.FE

        if isinstance(FECH_ENT_OP, str):
            FECH_ENT_OP = datetime.strptime(FECH_ENT_OP, '%Y-%m-%d')

    #! Consulta para la tabla "referencia"
    cursor_sql.execute("SELECT FolioRef, FchRef, TpoDocref FROM referencia WHERE TpoDocref = 802 AND Fecha_ingreso_xml >= ? AND Fecha_ingreso_xml < DATEADD(day, 1, ?)", fecha_actual.date(), fecha_actual.date())
    results = cursor_sql.fetchall()

    for row in results:
        VBELN = row.FolioRef
        if len(VBELN) == 9:
            VBELN = '0' + VBELN

    #! Consulta para la tabla "referencia 2"
    cursor_sql.execute("SELECT FolioRef, FchRef, TpoDocref FROM referencia WHERE TpoDocref = 52 AND Fecha_ingreso_xml >= ? AND Fecha_ingreso_xml < DATEADD(day, 1, ?)", fecha_actual.date(), fecha_actual.date())
    results = cursor_sql.fetchall()

    for row in results:
        GUIA_DESPACHO = row.FolioRef
        FECHARECEP = row.FchRef

        if isinstance(FECHARECEP, str):
            FECHARECEP = datetime.strptime(FECHARECEP, '%Y-%m-%d')

    #! Configuración de la conexión SAP
    sap_params = {
        "user": "xxxx",
        "passwd": "xxxxxx",
        "ashost": "10.8.0.76",
        "sysnr": "00",
        "client": "400",
    }

    try:
        with Connection(**sap_params) as conn:
            #! Llamada a la tabla "ZSD_DESP_PARAMS" para insertar los datos
            params = [
                {
                    'VBELN': VBELN if VBELN is not None else '',
                    'VBELN_VF': VBELN_VF if VBELN_VF is not None else '',
                    'GUIA_DESPACHO': GUIA_DESPACHO if GUIA_DESPACHO is not None else '',
                    'FECH_ENT_OP': FECH_ENT_OP.strftime('%Y%m%d') if FECH_ENT_OP is not None else None,
                    'KWMENG_PV': KWMENG_PV if KWMENG_PV is not None else 000.00,
                    'ORDEN_SALIDA': "",
                    'LOTE_DESPACHADO': LOTE_DESPACHADO if LOTE_DESPACHADO is not None else '',
                    'OBSERVACION': "XML",
                    'FECHARECEP': FECHARECEP.strftime('%Y%m%d') if FECHARECEP is not None else None,
                    'LFDAT': "",
                }
            ]

            #! Llamar al RFC para insertar los datos
            result = conn.call('ZSD_INTERFAZ_CARGA_DESP', TIPARAMS=params)

            #! Hacer algo con los resultados obtenidos, si es necesario
            #! Por ejemplo, imprimir un mensaje de éxito
            print(result)

    except ConnectionError as e:
        logging.error(f"Error al establecer la conexión con SAP: {str(e)}")
    except Exception as e:
        logging.error(f"Error desconocido: {str(e)}")

    # Cerrar la conexión SQL
    conn_sql.close()

def main():
    schedule.every(1).minutes.do(job)

    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == '__main__':
    main()
