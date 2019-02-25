# Codigo Fuente Servidor Argentum Online.
![AO Logo](https://github.com/ao-libre/ao-server/raw/master/Logo.jpg)

![Screenshot](https://preview.ibb.co/cojUf0/Screen-Shot-2018-11-04-at-6-53-52-PM.png)


## Wiki Desarrollo Argentum Online
[Manual para entender el codigo de Argentum Online](http://es.dao.wikia.com/wiki/Wiki_Desarrollo_Argentum_Online).

## Montar Servidor:
[Guia para montar mi propio servidor:](https://www.reddit.com/r/argentumonlineoficial/comments/9dow3q/como_montar_mi_propio_servidor/).


## Archivos Definicion:
[ArmadurasFaccionarias.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/ArmadurasFaccionarias.dat):
En este archivo se especifican a que índice de objeto corresponden las diferentes armaduras faccionarias.

[ArmadurasHerrero.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/ArmadurasHerrero.dat):
En este archivo se especifica el índice de las armaduras a la venta por el herrero.

[ArmasHerrero.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/ArmasHerrero.dat):
En este archivo se especifica el índice de las armaduras a la venta por el herrero.

[Balance.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Balance.dat):
En este archivo se encuentra la configuración del balance de las clases y los grupos.

[BanIps.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/BanIps.dat):
En este archivo se registran las IP's bloqueadas por el servidor o un Game-Master.

[Ciudades.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Ciudades.dat):
En este archivo se especifican las coordenadas de las ciudades del juego.

[Head.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Head.dat):
En este archivo se especifican la cantidad de cabezas que posee disponible cada raza en el juego.

[Hechizos.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Hechizos.dat):
En este archivo se especifica toda la información de los hechizos disponibles en el juego.

[Help.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Help.dat):
En este archivo se especifican los mensajes de ayuda en el juego.

[Invokar.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Invokar.dat):
Este comando solo está disponible para los Game-Masters. En este archivo se especifican los NPC's disponibles para invocar.

[Map.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Map.dat):
En este archivo se especifica la ubicación de la carpeta Maps y la cantidad de mapas que se cargarán en el servidor.

[Motd.ini](https://github.com/ao-libre/ao-server/blob/master/Dat/Motd.ini):
En este archivo se pueden escribir textos que apareceran a los usuarios al conectarse al servidor.

[NPCs.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/NPCs.dat):
En este archivo se especifica toda la información de los NPC's o Non-Playing-Characters del juego.

[NombresInvalidos.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/NombresInvalidos.txt):
En este archivo se especifican los nombres de personajes que no se pueden usar en el juego al crear uno.

[ObjCarpintero.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/ObjCarpintero.dat):
En este archivo se especifican los objetos que pueden crearse con la habilidad de carpinteria.

[Pretorianos.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/Pretorianos.dat):
En este archivo se especifica el índice de los NPC's que pertenecen al Clan Pretoriano.

[apuestas.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/apuestas.dat):
En este archivo se registra la informacion de las jugadas del sistema de apuestas.

[bkNPCs.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/bkNPCs.dat):
En este archivo se especifica la información de los NPC's cuando el servidor se inicia desde el BackUp.

[obj.dat](https://github.com/ao-libre/ao-server/blob/master/Dat/obj.dat):
En este archivo se especifica toda la información de los objetos disponibles en el juego.



# FAQs:

#### Error - Librerias faltantes (missing .dll)
En la carpeta `Librerias` estan todas las librerias necesarias para iniciar el server sin errores, copiar el contenido de la carpeta en `c:/Windows`

#### Error - Al abrir el proyecto en Visual Basic 6 no puede cargar todas las dependencias:
Este es un error comun que les suele pasar a varias personas, esto es debido que el EOL del archivo esta corrupto.
Visual Basic 6 lee el .vbp en CLRF, hay varias formas de solucionarlo:

Opcion a:
Con Notepad++ cambiar el EOL del archivo a CLRF

Opcion b:
Abrir un editor de texto y reemplazar todos los `'\n'` por `'\r\n'`


#### Server.ini - Summary:
Sumario explicando cada una de los valores utilizados en el archivo de configuracion [Server.ini.](Server.ini)

Summary explaining how to use each value in the configuration file [Server.ini.](Server.ini)

https://www.reddit.com/r/argentumonlineoficial/comments/9v4dln/serverini_sumario_explicando_parametros/

#### Autoupdates:

El programa al iniciar comparara la actual version del programa que se encuentra en `server.ini` en el parámetro `VersionTagRelease` con la ultima version que se encuentra en el [Endpoint Github Releases](https://api.github.com/repos/ao-libre/ao-server/releases/latest). En caso de ser diferente, se ejecuta nuestro programa `ao-autoupdate` para poder hacer el update.

Para mas información sobre este proceso:

[Funcion para comparar versiones](https://github.com/ao-libre/ao-server/blob/master/Codigo/frmCargando.frm#L137)

[Codigo fuente ao-autoupdate](https://github.com/ao-libre/ao-autoupdate)

Codigo fuente utilizado como base: http://www.gs-zone.org/temas/cliente-y-servidor-13-3-dx8-v1.95611/
