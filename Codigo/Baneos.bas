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
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Razon", "NO Razon")

End Sub

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, _
                        ByVal UserName As String, _
                        ByVal Razon As String)
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
                        Call LogBanFromName(UserName, bannerUserIndex, Razon)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        Call SaveBan(UserName, Razon, .Name)
                        
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
            
                Call LogBan(tUser, bannerUserIndex, Razon)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
                
                'Ponemos el flag de ban a 1
                UserList(tUser).flags.Ban = 1
                
                If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                    .flags.Ban = 1
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                    Call CloseSocket(bannerUserIndex)

                End If
                
                Call LogGM(.Name, "BAN a " & UserName)
                
                Call SaveBan(UserName, Razon, .Name)
                
                Call CloseSocket(tUser)

            End If

        End If

    End With

End Sub

'**********************************************************
'                   Baneo de IP's
'**********************************************************

Public Sub BanIP(ByVal bannerUserIndex As Integer, ByVal Nick As String, ByVal Razon As String)
    
    Dim TargetIndex As Integer: TargetIndex = NameIndex(Nick)
    
    'Si esta offline...
    If TargetIndex <= 0 Then
        Call WriteConsoleMsg(bannerUserIndex, "El usuario no se encuentra online.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'Si es Admin, Dios o Semi-Dios...
    If Not UserList(bannerUserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then Exit Sub

    With UserList(TargetIndex)
        
        'Registramos la accion en los logs.
        Call LogGM(UserList(bannerUserIndex).Name, "/BanIP " & .IP & " por " & Razon)
                
        If BanIpBuscar(.IP) > 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "La IP " & .IP & " ya se encuentra en la lista de IP's restringidas.", FontTypeNames.FONTTYPE_INFO)
                
        Else
            
            Call BanIpAgrega(.IP)
      
            'Find every player with that ip and ban him!
            Dim i As Long
            For i = 1 To LastUser

                If UserList(i).ConnIDValida Then
                    
                    If UserList(i).IP = .IP Then
                        Call BanCharacter(bannerUserIndex, UserList(i).Name, "IP POR " & Razon)
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

    Dim N As Long: N = BanIpBuscar(IP)

    If N > 0 Then
    
        Call BanIps.Remove(N)
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

Public Sub BanHD(ByVal bannerUserIndex As Integer, ByVal Nick As String, ByVal Razon As String)
    
    Dim TargetIndex As Integer: TargetIndex = NameIndex(Nick)
    
    'Si esta offline...
    If TargetIndex <= 0 Then
        Call WriteConsoleMsg(bannerUserIndex, "El usuario no se encuentra online.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    With UserList(TargetIndex)
   
        If LenB(.HD) > 0 Then
            
            'Registramos la accion en los logs.
            Call LogGM(UserList(bannerUserIndex).Name, "/Tolerancia0 " & .HD & " por " & Razon)
            
            If BuscarRegistroHD(.HD) > 0 Then
                Call WriteConsoleMsg(bannerUserIndex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                
            Else
            
                Call AgregarRegistroHD(.HD)
                
                Call WriteConsoleMsg(bannerUserIndex, "Has baneado el disco duro " & .HD & " del usuario " & .Name, FontTypeNames.FONTTYPE_INFO)
                
                Dim i As Long
                For i = 1 To LastUser
    
                    If UserList(i).ConnIDValida Then
                        
                        If UserList(i).HD = .HD Then
                            Call BanCharacter(bannerUserIndex, UserList(i).Name, "Ban de serial de disco duro.")
                        End If

                    End If

                Next i

            End If

        End If
    
    End With

End Sub

Public Function RemoverRegistroHD(ByVal HD As String) As Boolean '//Disco.
    
    On Error Resume Next
 
    Dim N As Long: N = BuscarRegistroHD(HD)
    
    If N > 0 Then
    
        Call BanHDs.Remove(N)
        
        Call RegistroBanHD
        
        RemoverRegistroHD = True
        
    Else
        RemoverRegistroHD = False
        
    End If
   
End Function

Public Sub AgregarRegistroHD(ByVal HD As String)
    
    Call BanHDs.Add(HD)
 
    Call RegistroBanHD

End Sub

Public Function BuscarRegistroHD(ByVal HD As String) As Long '//Disco.
    
    Dim Dale As Boolean: Dale = True
    Dim LoopC As Long: LoopC = 1

    Do While LoopC <= BanHDs.Count And Dale
        Dale = (BanHDs.Item(LoopC) <> HD)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BuscarRegistroHD = 0
    Else
        BuscarRegistroHD = LoopC - 1
    End If
    
End Function

Public Sub RegistroBanHD() '//Disco.
    
    Dim ArchN As Long
    Dim LoopC As Long
        
    ArchN = FreeFile()
    Open App.Path & "\Dat\BanHDs.dat" For Output As #ArchN
    
        For LoopC = 1 To BanHDs.Count
            Print #ArchN, BanHDs.Item(LoopC)
        Next LoopC
    
    Close #ArchN
    
End Sub

Public Sub BanHDCargar() '//Disco.
    
    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanHD As String
    
    Do While BanHDs.Count > 0
        Call BanHDs.Remove(1)
    Loop
    
    ArchN = FreeFile()
    Open App.Path & "\Dat\BanHDs.dat" For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        Call BanHDs.Add(Tmp)
    Loop
    
    Close #ArchN
    
End Sub
