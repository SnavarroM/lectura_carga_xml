
# Script 1: Descarga de correos electrónicos y archivos XML
1. Importa las bibliotecas necesarias:  `time` ,  `os` ,  `email` ,  `datetime`  y  `Imbox`  de la biblioteca  `imbox` .

2. Establece la configuración de conexión IMAP con el servidor de correo de Gmail.

3. Define la carpeta de destino donde se guardarán los archivos descargados.

4. Verifica si la carpeta de destino existe. Si no existe, la crea.

5. Establece una variable de contador para mantener un seguimiento de los archivos descargados.

6. Inicia un bucle infinito.

7. Obtiene la fecha y hora actual.

8. Verifica si es la hora programada para eliminar los archivos descargados. En este caso, se establece a la medianoche y un minuto.

9. Si es la hora programada, elimina todos los archivos de la carpeta de destino.

10. Establece una conexión con el servidor IMAP.

11. Obtiene todos los correos electrónicos del día actual.

12. Itera sobre cada correo electrónico obtenido.

13. Verifica si el correo electrónico tiene archivos adjuntos.

14. Si hay archivos adjuntos, itera sobre cada archivo adjunto.

15. Obtiene el nombre del archivo adjunto y lo decodifica si es necesario.

16. Verifica si el archivo adjunto tiene la extensión .xml o .XML.

17. Genera un nombre único para el archivo XML utilizando la fecha actual y el contador autoincremental.

18. Guarda el archivo adjunto en la carpeta de destino con el nombre único.

19. Espera 60 segundos antes de repetir el proceso.

Este código básicamente descarga archivos XML adjuntos de los correos electrónicos recibidos en una cuenta de correo específica y los guarda en una carpeta de destino. También tiene una funcionalidad para eliminar los archivos descargados a una hora específica.


# Script 2: insertar datos xml to sqlserver

El código proporcionado realiza las siguientes tareas:

1. Importa los módulos necesarios:  `xml.dom.minidom`  como  `minidom` ,  `pyodbc` ,  `os` ,  `datetime`  y  `time` .
2. Define una función llamada  `execute_cycle()`  que realiza un ciclo cada 60 segundos.
3. Establece la conexión a una base de datos utilizando los parámetros de servidor, base de datos, nombre de usuario y contraseña proporcionados.
4. Define la ruta de la carpeta donde se encuentran los archivos XML.
5. Crea una lista vacía llamada  `read_files`  para almacenar los nombres de los archivos leídos.
6. Recorre todos los archivos en la carpeta especificada.
7. Verifica si el archivo es un archivo XML y si no ha sido leído antes.
8. Lee el archivo XML utilizando  `minidom.parse()` .
9. Obtiene los valores de los campos de la factura (rut_emisor, tpo_dte, folio) del archivo XML.
10. Inserta los valores en la tabla "factura" en la base de datos.
11. Obtiene el ID de la factura insertada.
12. Inserta los valores de referencia (NroLinRef, TpoDocRef, FolioRef, FchRef) en la tabla "referencia" en la base de datos.
13. Realiza un commit para guardar los cambios en la base de datos.
14. Cierra la conexión a la base de datos.
15. Define una función llamada  `is_1am()`  que verifica si es la 01:00 horas.
16. Inicia un ciclo principal que se repite indefinidamente.
17. Si es la 01:00 horas, elimina todos los archivos de la carpeta "data".
18. Si no es la 01:00 horas, ejecuta la función  `execute_cycle()` .
19. Espera 60 segundos antes de repetir el ciclo principal.




# Script 3: Carga mssql to sap

Este código en Python realiza las siguientes acciones secuencialmente:

1. Importa los módulos necesarios:  `pyodbc` ,  `logging` ,  `Connection`  de  `pyrfc` ,  `datetime` ,  `date` ,  `schedule`  y  `time` .
2. Configura el nivel de registro de errores para el módulo de registro ( `logging` ).
3. Define una función llamada  `job()` .
4. Dentro de la función  `job()` , inicializa todas las variables necesarias.
5. Configura la conexión a una base de datos MSSQL utilizando  `pyodbc` .
6. Crea un cursor SQL para ejecutar consultas en la base de datos.
7. Realiza una consulta en la tabla "detalle" y guarda los resultados en la variable  `results` .
8. Recorre los resultados de la consulta y asigna los valores correspondientes a las variables.
9. Realiza una consulta en la tabla "factura" y guarda los resultados en la variable  `results` .
10. Recorre los resultados de la consulta y asigna los valores correspondientes a las variables.
11. Verifica si el valor de  `FECH_ENT_OP`  es una cadena de texto y, de ser así, convierte la cadena en un objeto  `datetime` .
12. Realiza una consulta en la tabla "referencia" y guarda los resultados en la variable  `results` .
13. Recorre los resultados de la consulta y asigna los valores correspondientes a las variables.
14. Verifica si el valor de  `VBELN`  tiene una longitud de 9 caracteres y, de ser así, agrega un "0" al inicio.
15. Realiza una consulta en la tabla "referencia 2" y guarda los resultados en la variable  `results` .
16. Recorre los resultados de la consulta y asigna los valores correspondientes a las variables.
17. Verifica si el valor de  `FECHARECEP`  es una cadena de texto y, de ser así, convierte la cadena en un objeto  `datetime` .
18. Configura los parámetros de conexión a SAP.
19. Intenta establecer una conexión con SAP utilizando los parámetros configurados.
20. Llama a la tabla "ZSD_DESP_PARAMS" en SAP para insertar los datos utilizando los valores de las variables.
21. Cierra la conexión a la base de datos MSSQL.
22. Define una función llamada  `main()` .
23. Configura una tarea programada para ejecutar la función  `job()`  cada 1 minuto utilizando  `schedule` .
24. Ejecuta el bucle principal que verifica y ejecuta las tareas programadas.
25. Si el archivo es ejecutado directamente, llama a la función  `main()` .

Este código está diseñado para realizar consultas en una base de datos MSSQL, obtener resultados y luego insertar esos resultados en una tabla en SAP utilizando la librería  `pyrfc` . También utiliza el módulo  `schedule`  para programar la ejecución periódica de la función  `job()` .