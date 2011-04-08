Attribute VB_Name = "Admin"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public tInicioServer As Long
Public EstadisticasWeb As clsEstadisticasIPC

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public Const IntervaloParalizadoReducido As Integer = 37
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer '[Nacho]
Public IntervaloUserPuedeAtacar As Long
Public IntervaloGolpeUsar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloPuedeSerAtacado As Long
Public IntervaloAtacable As Long
Public IntervaloOwnedNpc As Long

'BALANCE

Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public MinutosGuardarUsuarios As Long
Public Puerto As Integer

Public BootDelBackUp As Byte
Public Lloviendo As Boolean
Public DeNoche As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    VersionOK = (Ver = ULTIMAVERSION)
End Function

Sub ReSpawnOrigPosNpcs()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim i As Integer
    Dim MiNPC As npc
       
    For i = 1 To LastNPC
       'OJO
       If Npclist(i).flags.NPCActive Then
            
            If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                    MiNPC = Npclist(i)
                    Call QuitarNPC(i)
                    Call ReSpawnNpc(MiNPC)
            End If
            
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
       End If
       
    Next i
    
End Sub

Sub WorldSave()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim loopX As Integer
    Dim hFile As Integer
    
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))
    
    #If SeguridadAlkon Then
        Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
    #End If
    
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
    Dim j As Integer, k As Integer
    
    For j = 1 To NumMaps
        If MapInfo(j).BackUp = 1 Then k = k + 1
    Next j
    
    FrmStat.ProgressBar1.min = 0
    FrmStat.ProgressBar1.max = k
    FrmStat.ProgressBar1.value = 0
    
    For loopX = 1 To NumMaps
        'DoEvents
        
        If MapInfo(loopX).BackUp = 1 Then
            Call GrabarMapa(loopX, App.Path & "\WorldBackUp\Mapa" & loopX)
            FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
        End If
    
    Next loopX
    
    FrmStat.Visible = False
    
    If FileExist(DatPath & "\bkNpcs.dat") Then Kill (DatPath & "bkNpcs.dat")
    
    hFile = FreeFile()
    
    Open DatPath & "\bkNpcs.dat" For Output As hFile
    
        For loopX = 1 To LastNPC
            If Npclist(loopX).flags.BackUp = 1 Then
                Call BackUPnPc(loopX, hFile)
            End If
        Next loopX
        
    Close hFile
    
    Call SaveForums
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER))
End Sub

Public Sub PurgarPenas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    UserList(UserIndex).Counters.Pena = Minutos
    
    
    Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
    
    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
    If UserList(UserIndex).flags.Traveling = 1 Then
        UserList(UserIndex).flags.Traveling = 0
        UserList(UserIndex).Counters.goHome = 0
        Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
    End If
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserName) & ".chr"
    End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'Unban the character
    Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    If MD5ClientesActivado = 1 Then
        For i = 0 To UBound(MD5s)
            If (md5formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function
            End If
        Next i
        MD5ok = False
    Else
        MD5ok = True
    End If

End Function

Public Sub MD5sCarga()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Integer
    
    MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))
    
    If MD5ClientesActivado = 1 Then
        ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
        For LoopC = 0 To UBound(MD5s)
            MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
            MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
        Next LoopC
    End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    BanIps.Add ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Dale As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1
    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> ip)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1
    End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim N As Long
    
    N = BanIpBuscar(ip)
    If N > 0 Then
        BanIps.Remove N
        BanIpGuardar
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

    Dim ArchivoBanIp As String
    Dim ArchN As Long
    Dim LoopC As Long
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
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

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanIp As String
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    Set BanIps = New Collection
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop
    
    Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Static Andando As Boolean
    Static Contador As Long
    Dim Tmp As Boolean
    
    Contador = Contador + 1
    
    If Contador >= 10 Then
        Contador = 0
        Tmp = EstadisticasWeb.EstadisticasAndando()
        
        If Andando = False And Tmp = True Then
            Call InicializaEstadisticas
        End If
        
        Andando = Tmp
    End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
'***************************************************
'Author: Unknown
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************

    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/02/07
'22/05/2010: Ya no se peude banear admins de mayor rango si estan online.
'***************************************************

    Dim tUser As Integer
    Dim UserPriv As Byte
    Dim cantPenas As Byte
    Dim rank As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)
            
            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                UserPriv = UserDarPrivilegioLevel(UserName)
                
                If (UserPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        'ponemos la pena
                        cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                        
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
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
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
                
                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                
                Call CloseSocket(tUser)
            End If
        End If
    End With
End Sub

