Attribute VB_Name = "modInvisibles"
Option Explicit

' 0 = viejo
' 1 = nuevo
#Const MODO_INVISIBILIDAD = 0

' cambia el estado de invisibilidad a 1 o 0 dependiendo del modo: true o false
'
Public Sub PonerInvisible(ByVal UserIndex As Integer, ByVal estado As Boolean)
#If MODO_INVISIBILIDAD = 0 Then

UserList(UserIndex).flags.Invisible = IIf(estado, 1, 0)
UserList(UserIndex).flags.Oculto = IIf(estado, 1, 0)
UserList(UserIndex).Counters.Invisibilidad = 0
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & "," & IIf(estado, 1, 0))
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & "," & IIf(estado, 1, 0))
#If SeguridadAlkon Then
    End If
#End If

#Else

Dim EstadoActual As Boolean

' Está invisible ?
EstadoActual = (UserList(UserIndex).flags.Invisible = 1)

'If EstadoActual <> Modo Then
    If Modo = True Then
        ' Cuando se hace INVISIBLE se les envia a los
        ' clientes un Borrar Char
        UserList(UserIndex).flags.Invisible = 1
'        'Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "BP" & UserList(UserIndex).Char.CharIndex)
    Else
        
    End If
'End If

#End If
End Sub

