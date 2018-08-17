Attribute VB_Name = "History"
Option Explicit

'Alejoooooo yo quiero un sangucheeee

'9-5-2003 - v41
'1)Arreglado un bug en el q se podia construir objetos que
'no aparecian en la lista de herreria.

'17-4-2003 - v40
'1) agregados los comandos /BANIP y /UNBANIP. Sirven para
'banear por Ip a alguien ;)
'  Sintaxis c/ ejemplos:
'    /BANIP 1.1.1.1
'    /BANIP juanito
'    /UNBANIP 1.1.1.1

'9-4-2003 - v38
'1) Ahora el manejo de sockets se realiza x medio de la
'api de Winsock. Funca muy lindo :) pero CUIDADO CON
'INICIAR DESDE MODO DEBUG, DESDE EL ENTORNO DE VB, YA
'QUE CUANDO SE PONE 'CERRAR SERVIDOR' SE CIERRA VB6 TAMBIEN.
'OSEA, TRATEN DE NO HACER CAMBIOS EN EL CODIGO DESDE TIEMPO
'DE EJECUCION. SI QUIEREN HACERLOS, PONGAN 'STOP' ('DETENER')
'EN EL ENTORNO, GUARDEN, CIERREN Y ABRAN VB (PORQUE LOS
'SOCKETS NO SE CIERRAN Y X LO TANTO NO PUEDE ESCUCHAR EN EL
'PUERTO).
'2)Segun pablo el problema del SocketWrench es que el control
'no funciona bien y para cerrar hay que destruirlo y volverlo
'a crear (unload y load).
'3)Si quieren usar el control viejo en vez de la api (fijense
'que todavia esta en el form), vayan a Proyecto, Propiedades
'Gererar, Argumentos de compilacion y pongan UsarAPI en = 0
'NO LO HE PROBADO PERO DEBERIA ANDAR... :p

'3-4-2003 - v37
'1) La verdad no se que pasa. Hacen falta mas pruebas pero
'corremos riesgo de desbalanceo (extremista...)
'2) Acabo de resetear la funcion CLoseSocket a como esta
'en la version 22 (la ultima estable).
'3) cri cri

'1-4-203 - v36
'-------
'1) Estuve revisando el problema de los cuelgues. Al parecer
'es por un blucle infinito de ententos socket2_read. Ni idea
'que lo ocasiona.
'2) Eliminé el doevents de gametimer
'3) Ahora los socket además se cierran con .Disconnect.
'Espero que solucione el problema del bucle...
'4) Agregué un sistema para detectar los bucles esos con un
'contador. Cuando lo detecta se graba en el log de errores,
'cierra y limpia el socket y el slot

'History Log by CDT

'31-3-2003 - ver .35
'---------
'1) Reparé conteo de users (mal funcionamiento debido a la
'restauracion de CloseSocket [olvide NumUsers = NumUsers -1]
'2) Agrege aDos.RestarConeccion en PasarSegundo

'31-3-2003 - ver .34
'---------
'1) Saqué el Unload frmmain.socket2, creo que esta trayendo
'problemas...verifiqué y la forma de reutilizacion de socks
'es por UserList().flags.connid
'2) La funcion Cerrar_Usuario quedo solo para /salir y des_
'coneccion estando loggeado..devolví al estado original la
'func CloseSocket (bueh..un poko modifikada ta)

'30-3-2003 (estamos laburadores hoy eh) - ver .33
'---------
'1) Puse el autosalvado de pjes cada 30 minutos
'2) Cuando se gravaban pjes se cambiaba la cara si estabas
'muerto...solved it
'3) Los gms quedaban invisibles permanentemente con el /invisible
'4) Aplicados los 10 segundos en todos los casos..no solo si
'estas paralizado
'5) Identificador de /gmsg "nombre del gm> mensaje"

'Logaritmo de hitoria por Alejo (ya q todos dicen alguna boludex...)

'30-3-2003 - ver .32
'---------
'*Version 32 :)
'1) Agregados los comandos:
'   /CT mapa x y        Permite crear un teleport con
'                       con destino a mapa, x, y,
'                       posicionado un tile mas arriba
'                       que el dios.
'   /DT                 Destruir teleport de el ultimo
'                       click.

'History Log by CDT

'30-3-2003 (mismo dia, distinta hora :P)
'---------
'1) Errhand para el timer auditoria y que sea lo que dios quiera
'con respecto a este parche que ya cansa!, igualmente no creo
'que provoque un mal funcionamiento....tira que no existe el
'elemento en el array grgr i hate u
'2) Parche para mantener la invisibilidad al pasar de mapa

'30-3-2003
'---------
'1) Aplique un codigo para que cierre luego de 10 segundos el
'juego si estas paralizado..parece que hay problemas..no todo
'funciona bien..experimenté una caida, creo que el FINOK hacia
'que al desconectarse el cliente se llame de nuevo a closesocket
'espero que sea eso..habra que seguir experimentando :S estuve
'viendo el codigo y no encontre nada mal..no se :(
'2) Agrege el comando /GMSG para mensajeria entre gms (ToAdmin)
'3) Comando /REM para comentarios en los logs
'4) Agrege un boton que guarde todos los chars...ya que los
'cierres de server son aprovechados para clonar items
'5) Agrege la funcion que guarda los chars en un timer..lo mismo
'pero para prevenir caidas
'6) Estuve arreglando un poco los forms..habia algunos feitos ;)

'26-3-2003
'---------
'1) Apliqué el color Verde para los GMs (lease..lookattile)
'2) Los gms no muestran FXs cuando se mueven estando invisibles
'3) Me comí un pancho en la plaza deboto..que ricos son esos panchos! (??)

'History Log by Morgolock

'13-2-2003
'---------
'1) Modifiqué todas las llamadas a las funciones Mid, Left y
'Right por Mid$, Left$ y Right$ para que devuelvan strings
'en vez de variants. Se deberia ganar considerable velocidad.
'2) Quite el comando /GRABAR ya que generaba problemas con
'las mascotas y no era demasiado útil ya que los usuarios
'consiguen el mismo efecto saliendo y volviendo a entrar
'en el juego.
'3) Agregué el MOTD, el servidor levanta el mensaje del archivo
'motd.ini del directorio dat del servidor, les envia el motd
'a los usuarios cuando entran al juego.

'12-2-2003
'---------
'1) Limité a tres la máxima cantidad de mascotas
'2) A los newbies se les caen los objetos no newbies


