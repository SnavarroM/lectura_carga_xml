import xml.dom.minidom as minidom
import pyodbc
import os
from datetime import datetime, time
import time

# Función para realizar el ciclo cada 60 segundos
def execute_cycle():
    # Establecer la conexión a la base de datos
    server = 'ESCRITORIO'
    database = 'prueba_xml_3'
    username = 'sa'
    password = 'caramelo7613'
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +server+';DATABASE='+database+';UID='+username+';PWD=' + password)
    cursor = conn.cursor()

    # Ruta de la carpeta donde se encuentran los archivos XML
    folder_path = (r"C:\Users\chilo\Downloads\ejecucion_xml-20230819T172426Z-001\ejecucion_xml\data")

    # Lista para almacenar los nombres de los archivos leídos
    read_files = []

    # Recorre todos los archivos en la carpeta
    for filename in os.listdir(folder_path):
        # Verifica si el archivo es un archivo XML y si no ha sido leído antes
        if filename.endswith(".xml") and filename not in read_files:
            # Agrega el nombre del archivo a la lista de archivos leídos
            read_files.append(filename)
            # Lee el archivo XML
            dom = minidom.parse(os.path.join(folder_path, filename))
            root = dom.documentElement       
            #!___________________________________________49 campos en factura______________________________________________________
            #! caratula
            rut_emisor = str(root.getElementsByTagName('RutEmisor')[0].firstChild.data)[:255]
            rut_envia = str(root.getElementsByTagName('RutEnvia')[0].firstChild.data)[:255]
            rut_receptor = str(root.getElementsByTagName('RutReceptor')[0].firstChild.data)[:255]
            fch_resol = str(root.getElementsByTagName('FchResol')[0].firstChild.data)[:255]
            nro_resol = str(root.getElementsByTagName('NroResol')[0].firstChild.data)[:255]

            #! SubToDTE
            tmst_firma_env = str(root.getElementsByTagName('TmstFirmaEnv')[0].firstChild.data)[:255]
            tpo_dte = str(root.getElementsByTagName('TpoDTE')[0].firstChild.data)[:255]
            nro_dte = str(root.getElementsByTagName('NroDTE')[0].firstChild.data)[:255]

            #! Id_Documento
            folio = str(root.getElementsByTagName('Folio')[0].firstChild.data)[:255]
            fecha_emision = str(root.getElementsByTagName('FchEmis')[0].firstChild.data)[:255]
            if root.getElementsByTagName('TermPagoGlosa'):
                term_pago_glosa = root.getElementsByTagName('TermPagoGlosa')[0].firstChild.data
            else:
                term_pago_glosa = None

            if root.getElementsByTagName('FchVenc'):
                fecha_vencimiento = str(root.getElementsByTagName('FchVenc')[0].firstChild.data)[:255]
            else:
                fecha_vencimiento = None

            #! _______________________________emisor___________________________________________________
            #?________________________________nuevo campos emisor________________________________________
            if root.getElementsByTagName('Telefono'):
                Telefono = str(root.getElementsByTagName('Telefono')[0].firstChild.data)[:255]
            else:
                Telefono = None
            

            if root.getElementsByTagName('CorreoEmisor'):
                CorreoEmisor= str(root.getElementsByTagName('CorreoEmisor')[0].firstChild.data)[:255]
            else:
                CorreoEmisor = None
            
            #?________________________________FIn nuevo campos emisor________________________________________



            razon_social = str(root.getElementsByTagName('RznSoc')[0].firstChild.data)[:255]
            if root.getElementsByTagName('GiroEmis'):
                giro_emisor = str(root.getElementsByTagName('GiroEmis')[0].firstChild.data)[:255]
            else:
                giro_emisor = None
            
            acteca = str(root.getElementsByTagName('Acteco')[0].firstChild.data)[:255]


            #! _________________________________________________________________________________________

            if root.getElementsByTagName('CdgSIISucur'):
                cod_sii_sucur = str(root.getElementsByTagName('CdgSIISucur')[0].firstChild.data)[:255]
            else:
                cod_sii_sucur = None

            

            direccion_origen = str(root.getElementsByTagName('DirOrigen')[0].firstChild.data)[:255]
            comuna_origen = str(root.getElementsByTagName('CmnaOrigen')[0].firstChild.data)[:255]
            ciudad_origen = str(root.getElementsByTagName('CiudadOrigen')[0].firstChild.data)[:255]

            if root.getElementsByTagName('CdgVendedor'):
                codigo_vendedor = str(root.getElementsByTagName('CdgVendedor')[0].firstChild.data)[:255]
            else:
                codigo_vendedor = None
            

            #! receptor
            if root.getElementsByTagName('CdgIntRecep'):
                CdgIntRecep = str(root.getElementsByTagName('CdgIntRecep')[0].firstChild.data)[:255]
            else:
                CdgIntRecep = None
            
            RznSocRecep = str(root.getElementsByTagName('RznSocRecep')[0].firstChild.data)[:255]
            GiroRecep = str(root.getElementsByTagName('GiroRecep')[0].firstChild.data)[:255]
            DirRecep = str(root.getElementsByTagName('DirRecep')[0].firstChild.data)[:255]
            CmnaRecep = str(root.getElementsByTagName('CmnaRecep')[0].firstChild.data)[:255]
            CiudadRecep = str(root.getElementsByTagName('CiudadRecep')[0].firstChild.data)[:255]

            #? ___________________________Contacto_____________________________________________
            if root.getElementsByTagName('Contacto'):
                Contacto = str(root.getElementsByTagName('Contacto')[0].firstChild.data)[:255]
            else:
                Contacto = None        

            if root.getElementsByTagName('CmnaPostal'):
                CmnaPostal = str(root.getElementsByTagName('CmnaPostal')[0].firstChild.data)[:255]
            else:
                CmnaPostal = None        

            if root.getElementsByTagName('CiudadPostal'):
                CiudadPostal = str(root.getElementsByTagName('CiudadPostal')[0].firstChild.data)[:255]
            else:
                CiudadPostal = None
            
            #?____________________________Fin Contacto__________________________________________

            #! Transporte
            if root.getElementsByTagName('DirDest'):
                DirDest = str(root.getElementsByTagName('DirDest')[0].firstChild.data)[:255]
            else:
                DirDest = None

            if root.getElementsByTagName('CmnaDest'):
                CmnaDest = str(root.getElementsByTagName('CmnaDest')[0].firstChild.data)[:255]
            else:
                CmnaDest = None

            if root.getElementsByTagName('CiudadDest'):
                CiudadDest = str(root.getElementsByTagName('CiudadDest')[0].firstChild.data)[:255]
            else:
                CiudadDest = None       
            

            #! Totales
            MntNeto = str(root.getElementsByTagName('MntNeto')[0].firstChild.data)[:255]

            if root.getElementsByTagName('MntExe'):
                MntExe = str(root.getElementsByTagName('MntExe')[0].firstChild.data)[:255]
            else:
                MntExe = None       
            
            
            TasaIVA = str(root.getElementsByTagName('TasaIVA')[0].firstChild.data)[:255]
            IVA = str(root.getElementsByTagName('IVA')[0].firstChild.data)[:255]
            MntTotal = str(root.getElementsByTagName('MntTotal')[0].firstChild.data)[:255]

            #! referencia
            if root.getElementsByTagName('NroLinRef'):
                NroLinRef = str(root.getElementsByTagName('NroLinRef')[0].firstChild.data)[:255]
            else:
                NroLinRef = None    

            if root.getElementsByTagName('TpoDocRef'):
                TpoDocRef = str(root.getElementsByTagName('TpoDocRef')[0].firstChild.data)[:255]
            else:
                TpoDocRef = None


            if root.getElementsByTagName('FolioRef'):
                FolioRef = str(root.getElementsByTagName('FolioRef')[0].firstChild.data)[:255]
            else:
                FolioRef = None    
            
            if root.getElementsByTagName('FchRef'):
                FchRef = str(root.getElementsByTagName('FchRef')[0].firstChild.data)[:255]
            else:
                FchRef = None 
            

            #! DD
            F = str(root.getElementsByTagName('F')[0].firstChild.data)[:255]
            FE = str(root.getElementsByTagName('FE')[0].firstChild.data)[:255]
            RR = str(root.getElementsByTagName('RR')[0].firstChild.data)[:255]
            RSR = str(root.getElementsByTagName('RSR')[0].firstChild.data)[:255]
            MNT = str(root.getElementsByTagName('MNT')[0].firstChild.data)[:255]
            IT1 = str(root.getElementsByTagName('IT1')[0].firstChild.data)[:255]

            #! DA
            RE = str(root.getElementsByTagName('RE')[1].firstChild.data)[:255]
            RS = str(root.getElementsByTagName('RS')[0].firstChild.data)[:255]
            TD = str(root.getElementsByTagName('TD')[1].firstChild.data)[:255]

            #! RNG
            D = str(root.getElementsByTagName('D')[0].firstChild.data)[:255]
            H = str(root.getElementsByTagName('H')[0].firstChild.data)[:255]

            

            #!_________________________________________fin 49 campos en factura_________________________________________________________

            # Concatenar los valores de rut_emisor, tpo_dte y folio para crear el valor de id_factura
            id_factura = rut_emisor + tpo_dte + folio 
            

            # Verificar si el valor de id_factura ya existe en la tabla "factura"
            cursor.execute("SELECT id_factura FROM factura WHERE id_factura = ?", id_factura)
            row = cursor.fetchone()

            if row:
                # Si el valor de id_factura ya existe, actualizar los valores en la tabla "factura"            

                cursor.execute("UPDATE factura SET rut_emisor = ?, rut_envia = ?, rut_receptor = ?, fch_resol = ?, nro_resol = ?, tmst_firma_env = ?, tpo_dte = ?, nro_dte = ?, folio = ?, fecha_emision = ?, term_pago_glosa = ?, fecha_vencimiento = ?, razon_social = ?, giro_emisor = ?, acteca = ?, cod_sii_sucur = ?, direccion_origen = ?, comuna_origen = ?, ciudad_origen = ?, codigo_vendedor = ?, CdgIntRecep = ?, RznSocRecep = ?, GiroRecep = ?, DirRecep = ?, CmnaRecep = ?, CiudadRecep = ?, DirDest = ?, CmnaDest = ?, CiudadDest = ?, MntNeto = ?, MntExe = ?, TasaIVA = ?, IVA = ?, MntTotal = ?, NroLinRef = ?, TpoDocRef = ?, FolioRef = ?, FchRef = ?, RE = ?, TD = ?, F = ?, FE = ?, RR = ?, RSR = ?, MNT = ?, IT1 = ?, RS = ?, D = ?, H = ?, Telefono = ?, CorreoEmisor = ?, Contacto = ?, CmnaPostal = ?, CiudadPostal = ? WHERE id_factura = ?",rut_emisor, rut_envia, rut_receptor, fch_resol, nro_resol, tmst_firma_env, tpo_dte, nro_dte, folio, fecha_emision, term_pago_glosa, fecha_vencimiento, razon_social, giro_emisor, acteca, cod_sii_sucur, direccion_origen, comuna_origen, ciudad_origen, codigo_vendedor, CdgIntRecep, RznSocRecep, GiroRecep, DirRecep, CmnaRecep, CiudadRecep, DirDest, CmnaDest, CiudadDest, MntNeto, MntExe, TasaIVA, IVA, MntTotal, NroLinRef, TpoDocRef, FolioRef, FchRef, RE, TD, F, FE, RR, RSR, MNT, IT1, RS, D, H, Telefono, CorreoEmisor, Contacto, CmnaPostal, CiudadPostal, id_factura)
                # Definir las variables para la sentencia "UPDATE" en la tabla "detalle"
                for documento in root.getElementsByTagName('Documento'):
                    id = documento.getAttribute('ID')
                    for detalle in documento.getElementsByTagName('Detalle'):    
                        
                        #!_________________________________________7 campos en detalle_________________________________________________________                         
                        NroLinDet = str(detalle.getElementsByTagName('NroLinDet')[0].firstChild.data)[:255]
                        NmbItem = str(detalle.getElementsByTagName('NmbItem')[0].firstChild.data)[:255]
                        DscItem = str(detalle.getElementsByTagName('DscItem')[0].firstChild.data)[:255]
                        QtyItem = str(detalle.getElementsByTagName('QtyItem')[0].firstChild.data)[:255]

                        if root.getElementsByTagName('UnmdItem'):
                            UnmdItem = str(detalle.getElementsByTagName('UnmdItem')[0].firstChild.data)[:255]
                        else:
                            UnmdItem = None 

                        

                        PrcItem = str(detalle.getElementsByTagName('PrcItem')[0].firstChild.data)[:255]
                        MontoItem = str(detalle.getElementsByTagName('MontoItem')[0].firstChild.data)[:255]
                        #?________________________________________campos nuevo Detalle____________________________________________
                        if root.getElementsByTagName('VlrCodigo'):
                            VlrCodigo = str(detalle.getElementsByTagName('VlrCodigo')[0].firstChild.data)[:255]
                        else:
                            VlrCodigo = None 
                        

                        if root.getElementsByTagName('TpoCodigo'):
                            TpoCodigo = str(detalle.getElementsByTagName('TpoCodigo')[0].firstChild.data)[:255]
                        else:
                            TpoCodigo = None 
                        

                        if root.getElementsByTagName('Qtyref'):
                            Qtyref = str(detalle.getElementsByTagName('Qtyref')[0].firstChild.data)[:255]
                        else:
                            Qtyref = None 
                        
                        if root.getElementsByTagName('UnmdRef'):
                            UnmdRef = str(detalle.getElementsByTagName('UnmdRef')[0].firstChild.data)[:255]
                        else:
                            UnmdRef = None
                        

                        if root.getElementsByTagName('FchVencim'):
                            FchVencim = str(detalle.getElementsByTagName('FchVencim')[0].firstChild.data)[:255]
                        else:
                            FchVencim = None
                        
                        
                        #?_________________________________________Fin Nuevo Detalle_________________________________    

                        #!_________________________________________fin 7 campos en detalle_________________________________________________________ 
                        # Actualizar los valores en la tabla "detalle"
                        cursor.execute("UPDATE detalle SET NroLinDet = ?, NmbItem = ?, DscItem = ?, QtyItem = ?, UnmdItem = ?, PrcItem = ?, MontoItem = ?, VlrCodigo =?, TpoCodigo = ?, Qtyref = ?, UnmdRef = ?, FchVencim = ? WHERE id_factura = ?", NroLinDet, NmbItem, DscItem, QtyItem, UnmdItem, PrcItem, MontoItem, VlrCodigo, TpoCodigo, Qtyref, UnmdRef,FchVencim, id_factura)
            
                #!_________________________________________________________ UPDATE 4 campos referencias ____________________________________________________________
                for documento in root.getElementsByTagName('Documento'):
                    id = documento.getAttribute('ID')
                    for detalle in documento.getElementsByTagName('Referencia'):    
                        
                                                
                        NroLinRef = str(detalle.getElementsByTagName('NroLinRef')[0].firstChild.data)[:255]
                        TpoDocRef = str(detalle.getElementsByTagName('TpoDocRef')[0].firstChild.data)[:255]
                        FolioRef = str(detalle.getElementsByTagName('FolioRef')[0].firstChild.data)[:255]
                        FchRef = str(detalle.getElementsByTagName('FchRef')[0].firstChild.data)[:255]

                        
                        # Actualizar los valores en la tabla "detalle"
                        cursor.execute("UPDATE referencia SET NroLinRef = ?, TpoDocRef = ?, FolioRef = ?, FchRef = ? WHERE id_factura = ?", NroLinRef, TpoDocRef, FolioRef, FchRef, id_factura)

                #!_________________________________________________________fin UPDATE 4 campos referencias ____________________________________________________________
            
            
            else:
                # Si el valor de id_factura no existe, insertar los valores en la tabla "factura"
                cursor.execute("INSERT INTO factura (id_factura, rut_emisor, rut_envia, rut_receptor, fch_resol, nro_resol, tmst_firma_env, tpo_dte, nro_dte, folio, fecha_emision, term_pago_glosa, fecha_vencimiento, razon_social, giro_emisor, acteca, cod_sii_sucur, direccion_origen, comuna_origen, ciudad_origen, codigo_vendedor, CdgIntRecep, RznSocRecep, GiroRecep, DirRecep, CmnaRecep, CiudadRecep, DirDest, CmnaDest, CiudadDest, MntNeto, MntExe, TasaIVA, IVA, MntTotal, NroLinRef, TpoDocRef, FolioRef, FchRef, RE, TD, F, FE, RR, RSR, MNT, IT1, RS, D, H, Telefono, CorreoEmisor, Contacto, CmnaPostal, CiudadPostal) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",id_factura, rut_emisor, rut_envia, rut_receptor, fch_resol, nro_resol, tmst_firma_env, tpo_dte, nro_dte, folio, fecha_emision, term_pago_glosa, fecha_vencimiento, razon_social, giro_emisor, acteca, cod_sii_sucur, direccion_origen, comuna_origen, ciudad_origen, codigo_vendedor, CdgIntRecep, RznSocRecep, GiroRecep, DirRecep, CmnaRecep, CiudadRecep, DirDest, CmnaDest, CiudadDest, MntNeto, MntExe, TasaIVA, IVA, MntTotal, NroLinRef, TpoDocRef, FolioRef, FchRef, RE, TD, F, FE, RR, RSR, MNT, IT1, RS, D, H, Telefono, CorreoEmisor, Contacto, CmnaPostal, CiudadPostal)
                #!_______________________________ Insertar los valores en la tabla "detalle"__________________________________________________________
                for documento in root.getElementsByTagName('Documento'):
                    id = documento.getAttribute('ID')
                    for detalle in documento.getElementsByTagName('Detalle'):                             
                        NroLinDet = str(detalle.getElementsByTagName('NroLinDet')[0].firstChild.data)[:255]
                        NmbItem = str(detalle.getElementsByTagName('NmbItem')[0].firstChild.data)[:255]
                        DscItem = str(detalle.getElementsByTagName('DscItem')[0].firstChild.data)[:255]
                        QtyItem = str(detalle.getElementsByTagName('QtyItem')[0].firstChild.data)[:255]

                        if root.getElementsByTagName('UnmdItem'):
                            UnmdItem = str(detalle.getElementsByTagName('UnmdItem')[0].firstChild.data)[:255]
                        else:
                            UnmdItem = None

                        PrcItem = str(detalle.getElementsByTagName('PrcItem')[0].firstChild.data)[:255]
                        MontoItem = str(detalle.getElementsByTagName('MontoItem')[0].firstChild.data)[:255]

                        #?________________________________________campos nuevo Detalle____________________________________________
                        if root.getElementsByTagName('VlrCodigo'):
                            VlrCodigo = str(detalle.getElementsByTagName('VlrCodigo')[0].firstChild.data)[:255]
                        else:
                            VlrCodigo = None 
                        

                        if root.getElementsByTagName('TpoCodigo'):
                            TpoCodigo = str(detalle.getElementsByTagName('TpoCodigo')[0].firstChild.data)[:255]
                        else:
                            TpoCodigo = None 


                        if root.getElementsByTagName('Qtyref'):
                            Qtyref = str(detalle.getElementsByTagName('Qtyref')[0].firstChild.data)[:255]
                        else:
                            Qtyref = None
                        
                        if root.getElementsByTagName('UnmdRef'):
                            UnmdRef = str(detalle.getElementsByTagName('UnmdRef')[0].firstChild.data)[:255]
                        else:
                            UnmdRef = None
                        
                        if root.getElementsByTagName('FchVencim'):
                            FchVencim = str(detalle.getElementsByTagName('FchVencim')[0].firstChild.data)[:255]
                        else:
                            FchVencim = None
                        
                        
                        #?_________________________________________Fin Nuevo Detalle_________________________________    


                        cursor.execute("INSERT INTO detalle (id_factura, NroLinDet, NmbItem, DscItem, QtyItem, UnmdItem, PrcItem, MontoItem, VlrCodigo, TpoCodigo, Qtyref, UnmdRef, FchVencim)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",(id_factura, NroLinDet, NmbItem, DscItem, QtyItem, UnmdItem, PrcItem, MontoItem, VlrCodigo, TpoCodigo, Qtyref, UnmdRef, FchVencim))
                #!_______________________________FIN  Insertar los valores en la tabla "detalle"__________________________________________________________

                #!_______________________________ Insertar los valores en la tabla "referencia"__________________________________________________________

                for documento in root.getElementsByTagName('Documento'):
                    id = documento.getAttribute('ID')
                    for detalle in documento.getElementsByTagName('Referencia'):   
                        NroLinRef = str(detalle.getElementsByTagName('NroLinRef')[0].firstChild.data)[:255]
                        TpoDocRef = str(detalle.getElementsByTagName('TpoDocRef')[0].firstChild.data)[:255]
                        FolioRef = str(detalle.getElementsByTagName('FolioRef')[0].firstChild.data)[:255]
                        FchRef = str(detalle.getElementsByTagName('FchRef')[0].firstChild.data)[:255]

                        
                        # Actualizar los valores en la tabla "detalle"
                        cursor.execute("INSERT INTO referencia (id_factura, NroLinRef, TpoDocRef, FolioRef, FchRef)VALUES(?,?,?,?,?)", id_factura, NroLinRef, TpoDocRef, FolioRef, FchRef)

                #!_______________________________Fin Insertar los valores en la tabla "referencia"__________________________________________________________
                     # Eliminar el archivo leído de la carpeta "data"
            os.remove(os.path.join(folder_path, filename))

            # Insertar los valores en la tabla "referencia"
            conn.commit()

    # Cierra la conexión a la base de datos
    conn.close()       

# Ciclo principal que se repite cada 60 segundos
while True:
    execute_cycle()
    time.sleep(60)
            
           

