# Codigo Fuente Servidor Argentum Online.
![AO Logo](https://github.com/ao-libre/ao-server/raw/master/Logo.jpg)

![Screenshot](https://preview.ibb.co/cojUf0/Screen-Shot-2018-11-04-at-6-53-52-PM.png)

Mas informacion:

* http://www.gs-zone.org/temas/cliente-y-servidor-13-3-dx8-v1.95611/

## Wiki Desarrollo Argentum Online
[Manual para entender el codigo de Argentum Online](http://es.dao.wikia.com/wiki/Wiki_Desarrollo_Argentum_Online).

## Montar Servidor:
[Guia para montar mi propio servidor:](https://www.reddit.com/r/argentumonlineoficial/comments/9dow3q/como_montar_mi_propio_servidor/).


## Archivos Definicion:
[Motd.ini](https://github.com/ao-libre/ao-server/blob/master/Dat/Motd.ini)
En este archivo se pueden escribir textos que apareceran a los usuarios al conectarse al servidor.


# FAQs:

#### Error - Librerias faltantes (missing .dll)
En la carpeta `Librerias` estan todas las librerias necesarias para iniciar el server sin errores, copiar el contenido de la carpeta en `c:/Windows`


#### Server.ini - Summary:
Sumario explicando cada una de los valores utilizados en el archivo de configuracion [Server.ini.](Server.ini)

Summary explaining how to use each value in the configuration file [Server.ini.](Server.ini)

https://www.reddit.com/r/argentumonlineoficial/comments/9v4dln/serverini_sumario_explicando_parametros/

#### Autoupdates:

El programa al iniciar comparara la actual version del programa que se encuentra en `server.ini` en el parámetro `Version` con la ultima version que se encuentra en el [Endpoint Github Releases](https://api.github.com/repos/ao-libre/ao-server/releases/latest). En caso de ser diferente, se ejecuta nuestro programa `ao-autoupdate` para poder hacer el update.

Para mas información sobre este proceso:

[Funcion para comparar versiones](https://github.com/ao-libre/ao-server/blob/master/Codigo/frmCargando.frm#L137)

[Codigo fuente ao-autoupdate](https://github.com/ao-libre/ao-autoupdate)
