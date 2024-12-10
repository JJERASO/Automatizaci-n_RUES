######################################################################################

##  BIENVENIDOS AL PROGRAMA DE CONSULTA DE INFORMACIÓN EMPRESARIAL DEL RUES v1.0  #####

######################################################################################

Esta solución permite consultar la información financiera del RUES (Ingresos
por actividad ordinaria y Activos Totales), además de la última
fecha de renovación y razón social de la empresa. Cabe resaltar que si se
necesita alguna información extra (y esta se encuentra dentro del RUES) puede 
modificar el código para que se pueda obtener.

### REQUISITOS #######################################################################

El programa solo requiere dos componentes principales, los cuales son:

	- Python (Lenguaje de programación)

		Para poder instalarlo se debe pedir permiso a T.I. El link de descarga
		es el siguiente:

			- https://www.python.org/ftp/python/3.13.0/python-3.13.0-amd64.exe

	- ChromeDriver (una aplicación para desarrolladores creada para google chrome
		la cual permite hacer la parte de la automatización)

		Para poder instalarla se debe pedir permiso a T.I. El link de descarga
		es el siguiente:

	 		- https://googlechromelabs.github.io/chrome-for-testing/

		RECOMENDACIÓN: SIEMPRE DESCARGAR LA VERSIÓN 'Stable'

LIBRERÍAS: Las librerías que usa el programa se descargan automaticamente en caso
de no tenerlas instaladas. Dichas librerías son:

	- Pandas (Manipulación y transformación de datos)
	- Beautifoul Soup (Web Scrapping)
	- Selenium (Automatización)
	- Pathlib (Generalizar rutas del pc)
	- Tkinter (Alertas de la aplicación)
	- Customtkinter (Aplicación)
	- Sys (operaciones del sistema)

### MODO DE USO #####################################################################

Para hacer uso del programa debe tener instalado ChromeDriver en la carpeta de
descargas de su PC.

PASOS:

	1. Ubique el archivo del programa en el escritorio

	2. Ejecute (dé doble click) el programa

	3. Espere que carguen las librerías: En caso de no tenerlas instaladas,
		este proceso se demorará un poco más.

	4. Al abrirse la ventana de la aplicación, ingrese sus credenciales con
		las que normalmente accede al RUES en los cajones respectivos
		de usuario y contraseña.

	5. Dé click en el botón de 'Seleccionar archivo', el cual le abrirá una
		ventana para escoger un archivo excel. Aquí podrá subir el archivo
		excel en donde esta la columna 'NIT' los cuales son los nits a los
		que desea hacer la consulta.

		IMPORTANTE: LA COLUMNA DONDE ESTÁN LOS NITS DEBE LLAMARSE ASÍ TAL
				CUAL OBLIGATORIAMNETE, ES DECIR 'NIT' (sin las comillas).

	6. Seleccione el año para el cual necesita la información. La aplicación
		por defecto trae el año 2024.

	7. Dé click en el botón de 'Buscar'. Se le abrirá una ventana donde podrá
		ver el proceso de automatización. POR FAVOR INTERRUMPA LO MENOS
		POSIBLE EL PROCESO.

	8. Terminado el proceso se cerrará la ventana del ChromeDriver y encontrará
		en su carpeta de Descargas la base de datos con la información
		proveniente del RUES para los nits específicados.

## INQUIETUDES Y COMENTARIOS #########################################################

Cualquier inquietud, comentario, oportunidad de mejora, etc, no dude en contactarse
conmigo; mi correo es: 
	
		jeraso@ccc.org.co

Espero lo disfruten y les ayude a optimizar sus tareas!

Cámara de Comercio de Cali
Analítica y Estudios Económicos
Octubre 31 de 2024
Juan José Eraso T.