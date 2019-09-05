Attribute VB_Name = "mod_RetosFast"
Option Explicit
Private Type mapReto
    active As Boolean ' Mapa en uso?
    playerIndex1 As Integer ' Jugador 1
    playerIndex2 As Integer ' Jugador 2
    timeRemaining As Integer 'Tiempo restante para que termine el reto
    roundNumber As Byte ' Numero de ronda
    maxRounds As Byte
End Type
