# Programa
Javascript Plus!, un editor de texto.

# Autor
Luis Leonardo Nuñez Ibarra. Año 2005. email : leo.nunez@gmail.com. 

Chileno, casado , tengo 2 hijos. Aficionado a los videojuegos y el tenis de mesa. Mi primer computador fue un Talent MSX que me compro mi papa por alla por el año 1985. En el di mis primeros pasos jugando juegos como Galaga y PacMan y luego programando en MSX-BASIC. 

En la actualidad mi area de conocimiento esta referida a las tecnologias .NET con mas de 15 años de experiencia desarrollando varias paginas web usando asp.net con bases de datos sql server y Oracle. Integrador de tecnologias, desarrollo de servicios, aplicaciones de escritorio.

# Tipo de Proyecto
JavaScript Plus fue un editor de texto para javascript creado por mi por alla por el año 2004. El proyecto fue distribuido como shareware y hoy liberado para estudio y disponible para la comunidad. Fue escrito en Visual Basic 6.0 usando principalmente muchas librerias creadas por el sitio web http://www.vbaccelerator.com y adaptadas al proyecto.

# Historia
Quisiera compartir con uds esta historia que puede ser de motivacion para mas de alguno que quiera partir con una idea y no sabe como empezar. Hace muchos años atras trabajaba como recurso externo en la AFP Habitat del metro pedro de valdivia. Como era externo casi eramos "mierda" a diferencia de los que trabajan para la afp. En esa afp el piso de informatica estaba en el piso 10 y trabajaban con un lenguaje que habia inventado un tipo de ahi al que llamaban IUX.

Era un lenguaje como XML apoyado por javascript mas enredado que la mierda con el cual construian portales para la afp. En ese tiempo no sabia javascript y si habia que hacer alguna consulta de como hacer algo habia que preguntar a los "internos" que poco o nada nos pescaban. Mas encima teniamos internet bloqueado y el ambiente laboral no era muy amigable. 
Ante esta situacion me vi forzado por decirlo de alguna manera a desarrollar un editor propio para el lenguaje javascript en mis ratos libres y en la casa (alla por el año 2005) en visual basic 6. Resulta que el editor empezo a ser de interes en mis compañeros de area los cuales me dieron ideas y sugerencias de como mejorarlo.

A fines de ese año se me presento una oportunidad laboral en el banco de chile de la calle estado y me fui de esa mierda de lugar de trabajo. Continue trabajando en mis ratos libres y en las noches cuando ya todos en casa estaban acostados en ideas y mejoras para mi editor. Un conocido me sugirio que lo desarrollara en ingles, que subiera un portal propio (.cl) y que lo vendiera en formato "shareware" (pruebas antes de usar). El ya tenia un utilitario que vendia en ese formato y le iba bastante bien.

Para hacerla corta, converti toda la aplicacion a ingles, levante un .cl y averigue los canales de venta en USA para los desarrolladores de software shareware. La venta la canalize a traves de REGNOW el cual te juntaba cierto monto que tu podias configurar y te avisaba del pago, la comision por las ventas y el deposito del dinero desde USA a tu cuenta corriente nacional. Busque todos los portales de distribucion de software en ese formato y levante mi aplicacion, la descripcion, link de descarga , imagenes, el tipo de trial, valor del software, etc ..
Las versiones 1 y 2 fueron una mierda literalmente. 

Cero ventas hasta que un dia me llega un correo de un tipo de belgica el cual me comenta que el hacia testeos de programas y que veia que mi software tenia potencial pero que habia que corregir y mejorar muchas cosas. Si yo queria el podia darme su ayuda como beta tester sin costo alguno. Trabaje como 3 meses en reahacer toda la interfaz, correccion de errores, ideas y mejoras que el tipo me iba dando a modo de ir mejorando la aplicacion. Por la diferencia horaria con europa solo coincidiamos encierto horario nocturno de aca de chilito.

Liberada la version 3 de mi aplicacion, cual fue mi sopresa que al dia siguiente en la mañana tenia 4 ordenes de compra pendiente de procesar y yo no tenia siquiera algun algoritmo o algo para levantar el trial de la aplicacion. Asi que a la chilena genere una version full , una poca documentacion basica de como instalar y de como acceder al sector "full" de mi software. Luego vinieron varias versiones, mejoras, ideas nuevas y mas ventas. Mi software de nombre "JavaScript Plus!" lo vendia en 45 USD y llegue a ganar como $2.000.000 de pesos en ventas en 5 años. (Duro hasta el año 2010 mi sitio web).

Un poco larga la historia, pero como veran con esfuerzo, paciencia, constancia, perseverancia todo es posible.

Finalmente al dia de hoy aun lo ocupo para algunas cosas en particular en mi trabajo. Fue desarrollado integramente en VisuaL Basic 6.0 con las librerias .dll del sitio www.vbaccelerator.com el cual para mi fue uno de los mejores y mas avanzados sitios dedicados a VB de mi epoca. 

Espero te haya gustado mi historia.

# Término y fin del proyecto
El proyecto finalizo el año 2010 por falta de tiempo, bajas ventas y por el periodo de vida util de la aplicacion. Debo agradecer todo lo que aprendi con el, las muchas noches que me acoste tarde, el frio del invierno de ese año en especial 2005 y las incontables tazas de te que tome .... xD

# Shareware y canales de promoción
Para usar el canal de venta ocupe el que provee la empresa http://www.regnow.com. Tienes que configurar una cuenta, indicar los datos del deposito de la cuenta destino y solicitar al banco un numero de transaccion para autorizar depositos internacionales. Luego en tu cuenta de regnow te configuras cada cuanto quieres que te lleguen los depositos (en mi caso eran cada 200 USD). 

Regnow te da todo el canal de venta y procesamiento del pago. Tu solo vas recibiendo las transacciones realizadas. En esa epoca regnow me cobraba el 10% de cada venta.

Para los canales de promoción existen muchos y variados sitios web que te ofrecen promocionar tu producto de muchas maneras. Algunos gratis y otros mejoran tu posicion de búsqueda haciendo algun pago. Algun editor revisa tu software y lo valora con "estrellitas" segun corresponda. En su epoca yo busque muchos portales de distribución de software y subia la informacion. 

La subida de la información se realizaba usando la aplicacion PAD la cual te permitia configurar varios parametros comunes en los portales de distribución de software o bien tenias que ingresar "a mano" todos los valores segun corresponda.

# Distribución y empaquetamiento de la aplicación
El proceso de instalación se realizaba usando la aplicacion Inno Setup Script Wizard el cual generabas todo el script y proceso de instalación de todos los archivos de la aplicación. Luego la aplicacion "compila" tu proyecto en un archivo setup.exe el cual va realizando todos los tipicos pasos tradicionales de un instalador de software.

# Proceso de validación y trial de la aplicación
Para el proceso de la validación del trial de la aplicacion el proyecto tiene un flag dentro de las opciones de compilación condicional. Este parámetro se llama LITE. Si tiene el valor 1 al momento de compilar entonces era la version trial y se permitia usar hasta 30 veces la aplicación. Pasado ese numero se invitaba al usuario a comprar la aplicación y se bloqueba el uso de esta. El metodo para evitar posibles hackeos o crackeos usaba el siguiente truco :

- La aplicacion ejecutable se "firmaba" con un pequeño programa escrito en visual basic que agrega una firma "adicional" al archivo ejecutable. Luego en el proceso de validacion se validaban estos bytes extras a modo de evitar alguna alteracion en el archivo.

- Luego el proceso de ejecución en su primera vez instala 10 archivos en el directorio windows/system del sistema y luego via api windows le cambia la fecha de creacion. Los nombres eran como de archivos de sistema a modo de no generar sospechas. El proceso en su carga validaba por la existencia de esos 10 archivos. Si alguno no existia entonces era un posible intento de hackeo/crackeo a la aplicacion y esta no partia.

- Para la versión de pago se le solicitaba al usuario agregar crear un archivo llamado "reguser.ini" en el cual simplemente tenia el valor del nombre del usuario. Luego la aplicacion detecta de forma interna cuando es registrada este archivo y ademas el instalador del ejecutable "registrado" venia un archivo adicional llamado "licencia.dat" el cual contenia codificado en base64 el poema de pablo neruda "Muere Lentamente". Si la lectura del archivo coincidia con el poema "codificado" entonces era una versión valida de lo contrario no era valido para su ejecución.

# Componentes del proyecto
El proyecto esta construido usando varias tecnologias de la epoca. Destaco las principales :

- Componentes ActiveX OCX para la interfaz de usuario.
- Librerias ActiveX DLL las cuales proveen varias funcionalidades de apoyo al proyecto.
- Archivo de ayuda .hlp 
- Instalador para la aplicación.
- Packetes de librerias las cuales se instalan descomprimiendo archivos .zip
- Llamadas a la API de windows.

# Funcionalidades generales de la aplicación
- Editor de texto multi archivos. (MDI)
- Construido en Intellisense. (usando componente activex codesence que era un muy buen componente para editar texto creado por la empresa http://www.winmain.com. Ya no existe ...
- Explorador de funciones.
- Capacidad de ampliar las funcionalidades usando plugins (Se desarrollaban en visual basic 6.0 y la idea era expandir las funcionalidades de la aplicación)
- Validacion de archivos javascript usando JSLINT
- Validacion de archivos .css usando Tidy
- Varias guias de diferentes versiones de javascript
- Tiene un tutorial de javascript incorporado.
- Manejo de FTP propio.
- Definiciones de los atributos de las propiedades de HTML y CSS en archivos .INI 

# Utilitarios Anexos (Algunos construidos tambien por mi ...)
- Easy Query. Gestor para conectarse a distintas bases de datos via ODBC. (Otra herramienta mas que nacio de la necesidad)
