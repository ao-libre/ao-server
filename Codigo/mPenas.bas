Attribute VB_Name = "mPenas"
'Almacenamos a los personajes encarcelados.

Option Explicit

Public UltimoCarcel As Byte
Public ArrayPenas() As Integer

Public Sub AgregarArrayCarcel(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .Counters.Pena > 0 Then Exit Sub
        
        UltimoCarcel = UltimoCarcel + 1
        ReDim ArrayPenas(1 To UltimoCarcel) As Integer
        
        ArrayPenas(UltimoCarcel) = UserIndex
        .flags.SlotCarcel = UltimoCarcel
        
    End With
End Sub
Public Sub QuitarArrayCarcel(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        
        UltimoCarcel = UltimoCarcel - 1
        
        If UserIndex = ArrayPenas(.flags.SlotCarcel) Then
            ArrayPenas(.flags.SlotCarcel) = -1
            .flags.SlotCarcel = 0
        End If
        
    End With
End Sub

