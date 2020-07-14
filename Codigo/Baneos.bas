Attribute VB_Name = "Baneos"
Option Explicit

'*******************************************************************************************************
'           Modulo donde se gestiona el comportamiento de los distintas formas de baneo de personajes
'*******************************************************************************************************

'**********************************************************
'                   Baneo de Personajes
'**********************************************************
Public Function BANCheck(ByVal Name As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************
    If Not Database_Enabled Then
        BANCheck = BANCheckCharfile(Name)
    Else
        BANCheck = BANCheckDatabase(Name)

    End If

End Function

Public Sub UnBan(ByVal Name As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************
    
    If Not Database_Enabled Then
        Call UnBanCharfile(Name)
    Else
        Call UnBanDatabase(Name)

    End If
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")

End Sub

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, _
                        ByVal UserName As String, _
                        ByVal Reason As String)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/02/07
    '22/05/2010: Ya no se peude banear admins de mayor rango si estan online.
    '***************************************************

    Dim tUser     As Integer

    Dim UserPriv  As Byte

    Dim cantPenas As Byte

    Dim rank      As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)

        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_SERVER)
            
            If PersonajeExiste(UserName) Then
                UserPriv = UserDarPrivilegioLevel(UserName)
                
                If (UserPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                
                Else

                    If BANCheck(UserName) Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        Call SaveBan(UserName, Reason, .Name)
                        
                        If (UserPriv And rank) = (.flags.Privilegios And rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)

                        End If
                        
                        Call LogGM(.Name, "BAN a " & UserName)

                    End If

                End If

            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

            End If

        Else

            If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
            Else
            
                Call LogBan(tUser, bannerUserIndex, Reason)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
                
                'Ponemos el flag de ban a 1
                UserList(tUser).flags.Ban = 1
                
                If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                    .flags.Ban = 1
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                    Call CloseSocket(bannerUserIndex)

                End If
                
                Call LogGM(.Name, "BAN a " & UserName)
                
                Call SaveBan(UserName, Reason, .Name)
                
                Call CloseSocket(tUser)

            End If

        End If

    End With

End Sub

'**********************************************************
'                   Baneo de IP's
'**********************************************************

Public Sub BanIP(ByVal UserIndex As Integer, ByVal Razon As String)

    'Si esta offline...
    If UserIndex <= 0 Then Exit Sub
    
    'Si es Admin, Dios o Semi-Dios...
    If Not .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then Exit Sub
    
    With UserList(UserIndex)
        
        'Registramos la accion en los logs.
        Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)
                
        If BanIpBuscar(.IP) > 0 Then
            Call WriteConsoleMsg(UserIndex, "La IP " & .IP & " ya se encuentra en la lista de IP's restringidas.", FontTypeNames.FONTTYPE_INFO)
                
        Else
            
            Call BanIpAgrega(.IP)
      
            'Find every player with that ip and ban him!
            For i = 1 To LastUser

                If UserList(i).ConnIDValida Then
                    
                    If UserList(i).IP = .IP Then
                        Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & Razon)
                    End If

                End If

            Next i

        End If
    
    End With

End Sub

Public Sub BanIpAgrega(ByVal IP As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call BanIps.Add(IP)
    Call BanIpGuardar
    
    ' Agrego la regla al firewall para que bloquee la IP
    Call Shell("netsh.exe advfirewall firewall add rule name=""Baneo de IP " & IP & """ dir=in protocol=any action=block remoteip=" & IP)
    
End Sub

Public Function BanIpBuscar(ByVal IP As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Dale  As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1

    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> IP)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
        
    Else
        BanIpBuscar = LoopC - 1

    End If

End Function

Public Function BanIpQuita(ByVal IP As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim n As Long: n = BanIpBuscar(IP)

    If n > 0 Then
    
        Call BanIps.Remove(n)
        Call BanIpGuardar
        
        ' Agrego la regla al firewall para que borre la regla de la IP a desbanear.
        Call Shell("netsh.exe advfirewall firewall delete rule name=""Baneo de IP " & IP & """ dir=in protocol=any action=block remoteip=" & IP)
        
        BanIpQuita = True
        
    Else
        BanIpQuita = False

    End If

End Function

Public Sub BanIpGuardar()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ArchN        As Long
    Dim LoopC        As Long
    
    ArchN = FreeFile()
    Open App.Path & "\Dat\BanIps.dat" For Output As #ArchN
    
    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC
    
    Close #ArchN

End Sub

Public Sub BanIpCargar()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ArchN        As Long
    Dim Tmp          As String

    Set BanIps = New Collection
    
    ArchN = FreeFile()
    Open App.Path & "\Dat\BanIps.dat" For Input As #ArchN
        
        Do While Not EOF(ArchN)
            Line Input #ArchN, Tmp
            Call BanIps.Add(Tmp)
        Loop
    
    Close #ArchN

End Sub

'*******************************************************************************
'                   Baneo de numeros de serie del disco duro
'*******************************************************************************
