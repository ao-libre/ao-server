Attribute VB_Name = "Amigos"
Option Explicit

Public Sub ResetAmigos(ByVal Userindex As Integer)

Dim i As Integer

With UserList(Userindex)

    For i = 1 To MAXAMIGOS
        .Amigos(i).Nombre = vbNullString
        .Amigos(i).Ignorado = 0
        .Amigos(i).index = 0
    Next i

    .Quien = vbNullString

End With

End Sub

Public Function NoTieneEspacioAmigos(ByVal Userindex As Integer) As Boolean

Dim i     As Long
Dim Count As Byte

For i = 1 To MAXAMIGOS

    If LenB(UserList(Userindex).Amigos(i).Nombre) > 0 Then
        Count = Count + 1
    End If

Next i

If Count = MAXAMIGOS Then
    NoTieneEspacioAmigos = True
End If

End Function

Public Function BuscarSlotAmigoVacio(ByVal Userindex As Integer) As Byte

Dim i As Long

For i = 1 To MAXAMIGOS

    If LenB(UserList(Userindex).Amigos(i).Nombre) = 0 Then
        BuscarSlotAmigoVacio = i
        Exit Function
    End If

Next i

End Function

Public Function BuscarSlotAmigoName(ByVal Userindex As Integer, _
                                ByVal Nombre As String) As Boolean
Dim i As Long

For i = 1 To MAXAMIGOS

    If UCase$(UserList(Userindex).Amigos(i).Nombre) = UCase$(Nombre) Then
        BuscarSlotAmigoName = True
        Exit Function
    End If

Next i

End Function

Public Function BuscarSlotAmigoNameSlot(ByVal Userindex As Integer, _
                                    ByVal Nombre As String) As Byte
Dim i As Long

For i = 1 To MAXAMIGOS

    If UCase$(UserList(Userindex).Amigos(i).Nombre) = UCase$(Nombre) Then
        BuscarSlotAmigoNameSlot = i
        Exit Function
    End If

Next i

End Function

Public Sub BorrarAmigo(ByVal charName As String, ByVal Amigo As String)
Dim CharFile As String
Dim i        As Long
Dim Tiene    As Boolean
CharFile = CharPath & charName & ".chr"

If FileExist(CharFile) Then

    For i = 1 To MAXAMIGOS

        If UCase$(CStr(GetVar(CharFile, "AMIGOS", "NOMBRE" & i))) = UCase$(Amigo) Then
            Tiene = True
            Exit For
        End If

    Next i

    If Tiene Then
        'Lo borramos
        Call WriteVar(CharFile, "AMIGOS", "NOMBRE" & i, vbNullString)
        Call WriteVar(CharFile, "AMIGOS", "IGNORADO" & i, 0)
    End If

End If

End Sub

Public Function AgregarAmigo(ByVal Userindex As Integer, _
                                 ByVal Otro As Integer, _
                                 ByRef razon As String) As Boolean

With UserList(Userindex)

    If Otro = 0 Or Userindex = 0 Then
        razon = "Usuario Desconectado"
        AgregarAmigo = False
        Exit Function

    ElseIf Userindex = Otro Then
        razon = "Usuario Invalido"
        AgregarAmigo = False
        Exit Function

    ElseIf EsGm(Otro) = True Then
        razon = "No podes agregar a un Game Master como amigo."
        AgregarAmigo = False
        Exit Function

    ElseIf EsGm(Userindex) = True Then
        razon = "Los Game Masters no pueden agregar a usuarios como amigos."
        AgregarAmigo = False
        Exit Function

    ElseIf NoTieneEspacioAmigos(Userindex) = True Then
        razon = "No tienes mas espacio para poder agregar amigos."
        AgregarAmigo = False
        Exit Function

    ElseIf NoTieneEspacioAmigos(Otro) = True Then
        razon = "El otro usuario no tiene mas espacio para aceptar amigos."
        AgregarAmigo = False
        Exit Function

    ElseIf BuscarSlotAmigoName(Userindex, UserList(Otro).Name) = True Then
        razon = "Tu y " & UserList(Otro).Name & " ya son amigos."
        AgregarAmigo = False
        Exit Function

    End If

    AgregarAmigo = True

End With

End Function

Public Sub ActualizarSlotAmigo(ByVal Userindex As Integer, _
                           ByVal Slot As Byte, _
                           Optional ByVal Todo As Boolean = False)
Dim i As Long

With UserList(Userindex)

    If Todo Then

        For i = 1 To MAXAMIGOS
            Call WriteCargarListaDeAmigos(Userindex, i)
        Next i

    Else

        Call WriteCargarListaDeAmigos(Userindex, Slot)

    End If

End With

End Sub

Public Function ObtenerIndexLibre(ByVal Userindex As Integer) As Integer

Dim i As Long

For i = 1 To MAXAMIGOS

    If UserList(Userindex).Amigos(i).index <= 0 Then
        ObtenerIndexLibre = i
        Exit Function
    End If

Next i

End Function

Public Function ObtenerIndexUsuado(ByVal Userindex As Integer, _
                               ByVal Otro As Integer) As Integer
Dim i As Long

For i = 1 To MAXAMIGOS

    If UserList(Userindex).Amigos(i).index = Otro Then
        ObtenerIndexUsuado = i
        Exit Function
    End If

Next i

End Function

Public Sub ObtenerIndexAmigos(ByVal Userindex As Integer, ByVal Desconectar As Boolean)
Dim i    As Long
Dim Slot As Byte

With UserList(Userindex)

    If Desconectar = False Then

        For i = 1 To MAXAMIGOS

            If LenB(UserList(i).Name) > 0 Then

                If BuscarSlotAmigoName(Userindex, UserList(i).Name) Then

                    'Lo encontro y agregamos el index
                    Slot = ObtenerIndexLibre(Userindex)

                    'Por las dudas
                    If Slot > 0 Then .Amigos(Slot).index = i

                    If BuscarSlotAmigoName(i, .Name) Then

                        'Actualizamos la lista del otro
                        Slot = ObtenerIndexLibre(i)

                        If Slot > 0 Then

                            UserList(i).Amigos(Slot).index = Userindex

                            'Informamos al otro de nuestra presencia
                            Call WriteConsoleMsg(i, "Amigos> " & .Name & " se ha conectado", FontTypeNames.FONTTYPE_CONSEJO)

                        End If

                    End If

                End If

            End If

        Next i

    Else

        For i = 1 To MAXAMIGOS

            'Antes que nada
            If .Amigos(i).index > 0 Then

                Call WriteConsoleMsg(.Amigos(i).index, "Amigos> " & .Name & " se ha desconectado", FontTypeNames.FONTTYPE_CONSEJO)

                'Actualizamos la lista de index de los amigos
                Slot = ObtenerIndexUsuado(.Amigos(i).index, Userindex)

                If Slot > 0 Then UserList(.Amigos(i).index).Amigos(Slot).index = 0

            End If

        Next i

    End If

End With

End Sub

Public Sub HandleMsgAmigo(ByVal Userindex As Integer)

If UserList(Userindex).incomingData.Length < 3 Then
    Call Err.Raise(UserList(Userindex).incomingData.NotEnoughDataErrCode)
    Exit Sub
End If

On Error GoTo ErrHandler

With UserList(Userindex)

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(.incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim Mensaje As String
    Dim i       As Long

    Mensaje = buffer.ReadASCIIString()

    'If we got here then packet is complete, copy data back to original queue
    Call .incomingData.CopyBuffer(buffer)

    For i = 1 To MAXAMIGOS

        If .Amigos(i).index > 0 And .Amigos(i).index <> Userindex Then
            Call WriteConsoleMsg(.Amigos(i).index, "FMSG[" & .Name & "]: " & Mensaje, FontTypeNames.FONTTYPE_GM)
        End If

    Next i

    Call WriteConsoleMsg(Userindex, "FMSG[" & .Name & "]: " & Mensaje, FontTypeNames.FONTTYPE_GM)

End With

ErrHandler:

Dim Error As Long
    Error = Err.Number

On Error GoTo 0

'Destroy auxiliar buffer
Set buffer = Nothing

If Error <> 0 Then Call Err.Raise(Error)
End Sub

Public Sub HandleOnAmigo(ByVal Userindex As Integer)

With UserList(Userindex)

    'Remove packet ID
    Call .incomingData.ReadByte
    Dim list As String
    Dim i    As Long

    For i = 1 To MAXAMIGOS

        If .Amigos(i).index > 0 Then
            list = list & "[" & UserList(.Amigos(i).index).Name & "-" & MapInfo(UserList(.Amigos(i).index).Pos.Map).Name & "];"
        End If

    Next i

    If LenB(list) > 0 Then
        Call WriteConsoleMsg(Userindex, "Onlines: " & list, FontTypeNames.FONTTYPE_CONSEJO)
    Else
        Call WriteConsoleMsg(Userindex, "No tienes ningun amigo conectado.", FontTypeNames.FONTTYPE_GM)
    End If

End With

End Sub

Public Sub HandleAddAmigo(ByVal Userindex As Integer)

If UserList(Userindex).incomingData.Length < 3 Then
    Call Err.Raise(UserList(Userindex).incomingData.NotEnoughDataErrCode)
    Exit Sub
End If

On Error GoTo ErrHandler

With UserList(Userindex)

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(.incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName  As String
    Dim tUserName As String
    Dim caso      As Byte
    Dim razon     As String
    Dim tUser     As Integer
    Dim Slot      As Byte

    UserName = buffer.ReadASCIIString()
    caso = buffer.ReadByte
    tUser = NameIndex(UserName)

    'If we got here then packet is complete, copy data back to original queue
    Call .incomingData.CopyBuffer(buffer)

    'Mandar solicitudad de amistad
    If caso = 1 Then

        If AgregarAmigo(Userindex, tUser, razon) = True Then
            Call WriteConsoleMsg(Userindex, "Se ha enviado una solicitud de amistad a " & UserList(tUser).Name, FontTypeNames.FONTTYPE_CONSEJO)
            Call WriteConsoleMsg(tUser, UserList(Userindex).Name & " quiere ser tu amigo. Para aceptarlo usa el comando /FADD " & .Name, FontTypeNames.FONTTYPE_CONSEJO)
            UserList(tUser).Quien = .Name

        Else
            Call WriteConsoleMsg(Userindex, razon, FontTypeNames.FONTTYPE_CONSEJO)

        End If
        'Confirmar solicitudad de amistad

    ElseIf caso > 1 Then

        If AgregarAmigo(Userindex, tUser, razon) = True Then

            If LenB(.Quien) >= 3 Then

                If UCase$(.Quien) = UCase$(UserList(tUser).Name) Then

                    Slot = BuscarSlotAmigoVacio(Userindex)

                    .Amigos(Slot).Nombre = UserList(tUser).Name
                    .Amigos(Slot).Ignorado = 0

                    Call ActualizarSlotAmigo(Userindex, Slot)

                    Slot = BuscarSlotAmigoVacio(tUser)

                    UserList(tUser).Amigos(Slot).Nombre = .Name
                    UserList(tUser).Amigos(Slot).Ignorado = 0

                    Call ActualizarSlotAmigo(tUser, Slot)

                    Call WriteConsoleMsg(Userindex, UserList(tUser).Name & " agregado", FontTypeNames.FONTTYPE_DIOS)

                    Call WriteConsoleMsg(tUser, .Name & " agregado", FontTypeNames.FONTTYPE_DIOS)

                    Slot = ObtenerIndexLibre(Userindex)

                    If Slot > 0 Then
                        .Amigos(Slot).index = tUser
                    End If

                    Slot = ObtenerIndexLibre(tUser)

                    If Slot > 0 Then
                        UserList(tUser).Amigos(Slot).index = Userindex
                    End If

                    .Quien = vbNullString

                Else
                    Call WriteConsoleMsg(Userindex, "Solicitud de amistad invalida.", FontTypeNames.FONTTYPE_CONSEJO)

                End If

            End If

        Else
            Call WriteConsoleMsg(Userindex, razon, FontTypeNames.FONTTYPE_CONSEJO)

        End If

    End If

End With

ErrHandler:

Dim Error As Long
    Error = Err.Number

On Error GoTo 0

'Destroy auxiliar buffer
Set buffer = Nothing

If Error <> 0 Then Call Err.Raise(Error)

End Sub

Public Sub HandleDelAmigo(ByVal Userindex As Integer)

With UserList(Userindex)

    'Remove packet ID
    Call .incomingData.ReadByte

    Dim Slot     As Byte
    Dim tUser    As Integer
    Dim UserName As String

    Slot = .incomingData.ReadByte()

    If Slot <= 0 Or Slot > MAXAMIGOS Then Exit Sub

    'Por las duditas :P
    If LenB(.Amigos(Slot).Nombre) = 0 Then Exit Sub

    tUser = NameIndex(.Amigos(Slot).Nombre)
    UserName = .Amigos(Slot).Nombre

    Call WriteConsoleMsg(Userindex, .Amigos(Slot).Nombre & " ha sido borrado de la lista de amigos.", FontTypeNames.FONTTYPE_GMMSG)

    'reseteamos el slot
    .Amigos(Slot).Nombre = vbNullString
    .Amigos(Slot).Ignorado = 0
    Call ActualizarSlotAmigo(Userindex, Slot)

    If tUser > 0 Then

        'Puede pasar....
        If BuscarSlotAmigoName(tUser, .Name) Then

            Call WriteConsoleMsg(tUser, .Name & "te ha borrado de su lista de amigos.", FontTypeNames.FONTTYPE_GMMSG)

            Slot = BuscarSlotAmigoNameSlot(tUser, .Name)

            UserList(tUser).Amigos(Slot).Ignorado = 0
            UserList(tUser).Amigos(Slot).Nombre = vbNullString

            Call ActualizarSlotAmigo(tUser, Slot)

            Slot = ObtenerIndexUsuado(Userindex, tUser)

            If Slot > 0 Then
                .Amigos(Slot).index = 0
            End If

            Slot = ObtenerIndexUsuado(tUser, Userindex)

            If Slot > 0 Then
                UserList(tUser).Amigos(Slot).index = 0
            End If

        End If

    Else

        'verificamos desde el char
        Call BorrarAmigo(UserName, .Name)

    End If

End With

End Sub
