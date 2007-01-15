Attribute VB_Name = "modInvisibles"
Option Explicit

' 0 = viejo
' 1 = nuevo
#Const MODO_INVISIBILIDAD = 0

' cambia el estado de invisibilidad a 1 o 0 dependiendo del modo: true o false
'
Public Sub PonerInvisible(ByVal UserIndex As Integer, ByVal estado As Boolean)
#If MODO_INVISIBILIDAD = 0 Then

UserList(UserIndex).flags.invisible = IIf(estado, 1, 0)
UserList(UserIndex).flags.Oculto = IIf(estado, 1, 0)
UserList(UserIndex).Counters.Invisibilidad = 0

Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, Not estado))


#Else

Dim EstadoActual As Boolean

' Está invisible ?
EstadoActual = (UserList(UserIndex).flags.invisible = 1)

'If EstadoActual <> Modo Then
    If Modo = True Then
        ' Cuando se hace INVISIBLE se les envia a los
        ' clientes un Borrar Char
        UserList(UserIndex).flags.invisible = 1
'        'Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
        Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.Map, PrepareMessageCharacterRemove(UserList(UserIndex).Char.CharIndex))
    Else
        
    End If
'End If

#End If
End Sub

