Attribute VB_Name = "Amigos"
Option Explicit

Public Function NoTieneEspacioAmigos(ByVal UserIndex As Integer) As Boolean
    
    Dim i     As Long
    Dim Count As Byte

    For i = 1 To MAXAMIGOS

        If Not UserList(UserIndex).Amigos(i).Nombre = "Nadies" Then
            Count = Count + 1
        End If
        
    Next i

    If Count = MAXAMIGOS Then
        NoTieneEspacioAmigos = True
    End If

End Function

Public Function BuscarSlotAmigoVacio(ByVal UserIndex As Integer) As Byte
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UserList(UserIndex).Amigos(i).Nombre = "Nadies" Then
            BuscarSlotAmigoVacio = i
            Exit Function
        End If
        
    Next i

End Function

Public Function BuscarSlotAmigoName(ByVal UserIndex As Integer, _
                                    ByVal Nombre As String) As Boolean
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UCase$(UserList(UserIndex).Amigos(i).Nombre) = UCase$(Nombre) Then
            BuscarSlotAmigoName = True
            Exit Function
        End If
    Next i

End Function

Public Function BuscarSlotAmigoNameSlot(ByVal UserIndex As Integer, _
                                        ByVal Nombre As String) As Byte
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UCase$(UserList(UserIndex).Amigos(i).Nombre) = UCase$(Nombre) Then
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
            Call WriteVar(CharFile, "AMIGOS", "NOMBRE" & i, "Nadies")
            Call WriteVar(CharFile, "AMIGOS", "IGNORADO" & i, 0)
        End If

    End If
    
End Sub

Public Function IntentarAgregarAmigo(ByVal UserIndex As Integer, _
                                     ByVal Otro As Integer, _
                                     ByRef razon As String) As Boolean

    With UserList(UserIndex)

        If Otro = 0 Or UserIndex = 0 Then
            razon = "Usuario Desconectado"
            IntentarAgregarAmigo = False
            Exit Function

        ElseIf UserIndex = Otro Then
            razon = "Usuario Invalido"
            IntentarAgregarAmigo = False
            Exit Function

        ElseIf EsGm(Otro) = True Then
            razon = "Usuario Desconectado"
            IntentarAgregarAmigo = False
            Exit Function

        ElseIf EsGm(UserIndex) = True Then
            razon = "Los Administradores no pueden agregar a usuarios"
            IntentarAgregarAmigo = False
            Exit Function

        ElseIf NoTieneEspacioAmigos(UserIndex) = True Then
            razon = "No tienes mas espacio para poder agregar amigos"
            IntentarAgregarAmigo = False
            Exit Function

        ElseIf NoTieneEspacioAmigos(Otro) = True Then
            razon = "El otro usuario no tiene mas espacio para aceptar Amigos"
            IntentarAgregarAmigo = False
            Exit Function

        ElseIf BuscarSlotAmigoName(UserIndex, UserList(Otro).Name) = True Then
            razon = "Tu y " & UserList(Otro).Name & "Ya son amigos"
            IntentarAgregarAmigo = False
            Exit Function
        
        End If

        IntentarAgregarAmigo = True
        
    End With
    
End Function

Public Sub ActualizarSlotAmigo(ByVal UserIndex As Integer, _
                               ByVal Slot As Byte, _
                               Optional ByVal Todo As Boolean = False)
    Dim i As Long

    With UserList(UserIndex)

        If Todo Then

            For i = 1 To MAXAMIGOS
                Call WriteCargarListaDeAmigos(UserIndex, i)
            Next i

        Else
        
            Call WriteCargarListaDeAmigos(UserIndex, Slot)
            
        End If
        
    End With
    
End Sub

Public Function ObtenerIndexLibre(ByVal UserIndex As Integer) As Integer
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UserList(UserIndex).Amigos(i).index <= 0 Then
            ObtenerIndexLibre = i
            Exit Function
        End If
        
    Next i

End Function

Public Function ObtenerIndexUsuado(ByVal UserIndex As Integer, _
                                   ByVal Otro As Integer) As Integer
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UserList(UserIndex).Amigos(i).index = Otro Then
            ObtenerIndexUsuado = i
            Exit Function
        End If
        
    Next i

End Function

Public Sub ObtenerIndexAmigos(ByVal UserIndex As Integer, ByVal Desconectar As Boolean)
    Dim i    As Long
    Dim Slot As Byte

    With UserList(UserIndex)

        If Desconectar = False Then

            For i = 1 To LastUser

                If LenB(UserList(i).Name) > 0 Then
                    
                    If BuscarSlotAmigoName(UserIndex, UserList(i).Name) Then
                        
                        'Lo encontro y agregamos el index
                        Slot = ObtenerIndexLibre(UserIndex)

                        'Por las dudas
                        If Slot > 0 Then .Amigos(Slot).index = i

                        If BuscarSlotAmigoName(i, .Name) Then
                            
                            'Actualizamos la lista del otro
                            Slot = ObtenerIndexLibre(i)

                            If Slot > 0 Then
                                
                                UserList(i).Amigos(Slot).index = UserIndex
                                
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
                    Slot = ObtenerIndexUsuado(.Amigos(i).index, UserIndex)

                    If Slot > 0 Then UserList(.Amigos(i).index).Amigos(Slot).index = 0
                    
                End If
                
            Next i

        End If
        
    End With
    
End Sub

Public Sub HandleMsgAmigo(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.Length < 3 Then
        Call Err.Raise(UserList(UserIndex).incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If

    On Error GoTo ErrHandler

    With UserList(UserIndex)
    
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
            
            If .Amigos(i).index > 0 Then
                Call WriteConsoleMsg(.Amigos(i).index, "FMSG[" & .Name & "]: " & Mensaje, FontTypeNames.FONTTYPE_GM)
            End If
            
        Next i

        Call WriteConsoleMsg(UserIndex, "FMSG[" & .Name & "]: " & Mensaje, FontTypeNames.FONTTYPE_GM)

    End With

ErrHandler:

    Dim Error As Long
        Error = Err.Number
    
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Call Err.Raise(Error)
End Sub

Public Sub HandleOnAmigo(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
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
            Call WriteConsoleMsg(UserIndex, "Onlines: " & list, FontTypeNames.FONTTYPE_CONSEJO)
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes ningun amigo conectado.", FontTypeNames.FONTTYPE_GM)
        End If
        
    End With
    
End Sub

Public Sub HandleAddAmigo(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.Length < 3 Then
        Call Err.Raise(UserList(UserIndex).incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
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
        
            If IntentarAgregarAmigo(UserIndex, tUser, razon) = True Then
                Call WriteConsoleMsg(UserIndex, "mandando solicitud de amistad a " & UserList(tUser).Name, FontTypeNames.FONTTYPE_CONSEJO)
                Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " quiere ser tu amigo.para aceptarlo usa el comando /FADD " & .Name, FontTypeNames.FONTTYPE_CONSEJO)
                UserList(tUser).Quien = .Name
            
            Else
                Call WriteConsoleMsg(UserIndex, razon, FontTypeNames.FONTTYPE_CONSEJO)
            
            End If
            'Confirmar solicitudad de amistad
        
        ElseIf caso > 1 Then

            If IntentarAgregarAmigo(UserIndex, tUser, razon) = True Then
                
                If LenB(.Quien) >= 3 Then
                    
                    If UCase$(.Quien) = UCase$(UserList(tUser).Name) Then
                        
                        Slot = BuscarSlotAmigoVacio(UserIndex)
                        
                        .Amigos(Slot).Nombre = UserList(tUser).Name
                        .Amigos(Slot).Ignorado = 0
                        
                        Call ActualizarSlotAmigo(UserIndex, Slot)
                        
                        Slot = BuscarSlotAmigoVacio(tUser)
                        
                        UserList(tUser).Amigos(Slot).Nombre = .Name
                        UserList(tUser).Amigos(Slot).Ignorado = 0
                        
                        Call ActualizarSlotAmigo(tUser, Slot)
                        
                        Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " agregado", FontTypeNames.FONTTYPE_DIOS)
                        
                        Call WriteConsoleMsg(tUser, .Name & " agregado", FontTypeNames.FONTTYPE_DIOS)
                        
                        Slot = ObtenerIndexLibre(UserIndex)

                        If Slot > 0 Then
                            .Amigos(Slot).index = tUser
                        End If
                        
                        Slot = ObtenerIndexLibre(tUser)

                        If Slot > 0 Then
                            UserList(tUser).Amigos(Slot).index = UserIndex
                        End If
                        
                        .Quien = vbNullString
                    
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solicitud de amistad invalida.", FontTypeNames.FONTTYPE_CONSEJO)
                    
                    End If
                
                End If
            
            Else
                Call WriteConsoleMsg(UserIndex, razon, FontTypeNames.FONTTYPE_CONSEJO)
            
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

Public Sub HandleDelAmigo(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot     As Byte
        Dim tUser    As Integer
        Dim UserName As String

        Slot = .incomingData.ReadByte()

        If Slot <= 0 Or Slot > MAXAMIGOS Then Exit Sub

        'Por las duditas :P
        If .Amigos(Slot).Nombre = "Nadies" Then Exit Sub

        tUser = NameIndex(.Amigos(Slot).Nombre)
        UserName = .Amigos(Slot).Nombre

        Call WriteConsoleMsg(UserIndex, .Amigos(Slot).Nombre & " ha sido borrado de la lista de amigos.", FontTypeNames.FONTTYPE_GMMSG)

        'reseteamos el slot
        .Amigos(Slot).Nombre = "Nadies"
        .Amigos(Slot).Ignorado = 0
        Call ActualizarSlotAmigo(UserIndex, Slot)
   
        If tUser > 0 Then

            'Puede pasar....
            If BuscarSlotAmigoName(tUser, .Name) Then
                
                Call WriteConsoleMsg(tUser, .Name & "te ha borrado de su lista de amigos.", FontTypeNames.FONTTYPE_GMMSG)
                
                Slot = BuscarSlotAmigoNameSlot(tUser, .Name)
                
                UserList(tUser).Amigos(Slot).Ignorado = 0
                UserList(tUser).Amigos(Slot).Nombre = "Nadies"
                
                Call ActualizarSlotAmigo(tUser, Slot)
                
                Slot = ObtenerIndexUsuado(UserIndex, tUser)

                If Slot > 0 Then
                    .Amigos(Slot).index = 0
                End If
                
                Slot = ObtenerIndexUsuado(tUser, UserIndex)

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
