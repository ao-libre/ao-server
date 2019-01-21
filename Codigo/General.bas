Attribute VB_Name = "General"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Global LeerNPCs As clsIniManager

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************

Dim CuerpoDesnudo As Integer

With UserList(UserIndex)
    Select Case .Genero
        Case eGenero.Hombre
            Select Case .raza
                Case eRaza.Humano
                    CuerpoDesnudo = 21
                Case eRaza.Drow
                    CuerpoDesnudo = 32
                Case eRaza.Elfo
                    CuerpoDesnudo = 210
                Case eRaza.Gnomo
                    CuerpoDesnudo = 222
                Case eRaza.Enano
                    CuerpoDesnudo = 53
            End Select
        Case eGenero.Mujer
            Select Case .raza
                Case eRaza.Humano
                    CuerpoDesnudo = 39
                Case eRaza.Drow
                    CuerpoDesnudo = 40
                Case eRaza.Elfo
                    CuerpoDesnudo = 259
                Case eRaza.Gnomo
                    CuerpoDesnudo = 260
                Case eRaza.Enano
                    CuerpoDesnudo = 60
            End Select
    End Select
    
    If Mimetizado Then
        .CharMimetizado.body = CuerpoDesnudo
    Else
        .Char.body = CuerpoDesnudo
    End If
    
    .flags.Desnudo = 1
End With

End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
'***************************************************

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, X, Y, b)
    End If

End Sub


Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        With MapData(Map, X, Y)
            If ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
            (.Graphic(1) >= 5665 And .Graphic(1) <= 5680) Or _
            (.Graphic(1) >= 13547 And .Graphic(1) <= 13562)) And _
               .Graphic(2) = 0 Then
                    HayAgua = True
            Else
                    HayAgua = False
            End If
        End With
    Else
      HayAgua = False
    End If

End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
    If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False
        End If
    Else
      HayLava = False
    End If

End Function


Sub LimpiarMundo()
'***************************************************
'Author: Unknow
'Last Modification: 04/15/2008
'01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
'04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
'***************************************************
On Error GoTo ErrHandler

    Dim i As Integer
    Dim d As cGarbage
    Set d = New cGarbage
    
    For i = TrashCollector.Count To 1 Step -1
        Set d = TrashCollector(i)
        Call EraseObj(1, d.Map, d.X, d.Y)
        Call TrashCollector.Remove(i)
        Set d = Nothing
    Next i
    
    Call SecurityIp.IpSecurityMantenimientoLista
    
    Exit Sub

ErrHandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.description)
End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim k As Long
    Dim npcNames() As String
    
    ReDim npcNames(1 To UBound(SpawnList)) As String
    
    For k = 1 To UBound(SpawnList)
        npcNames(k) = SpawnList(k).NpcName
    Next k
    
    Call WriteSpawnList(UserIndex, npcNames())

End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

#If UsarQueSocket = 0 Then

    Obj.AddressFamily = AF_INET
    Obj.Protocol = IPPROTO_IP
    Obj.SocketType = SOCK_STREAM
    Obj.Binary = False
    Obj.Blocking = False
    Obj.BufferSize = 1024
    Obj.LocalPort = Port
    Obj.backlog = 5
    Obj.listen

#End If

End Sub

Sub Main()
'***************************************************
'Author: Unknown
'Last Modification: 15/03/2011
'15/03/2011: ZaMa - Modularice todo, para que quede mas claro.
'***************************************************

On Error Resume Next
    
    ChDir App.Path
    ChDrive App.Path
    
    Call LoadMotd
    Call BanIpCargar
    
    UltimoSlotLimpieza = -1
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
    ' Start loading..
    frmCargando.Show
    
    ' Constants & vars
    frmCargando.Label1(2).Caption = "Cargando constantes..."
    Call LoadConstants
    DoEvents
    
    ' Arrays
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call LoadArrays
    
    ' Server.ini & Apuestas.dat
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    Call LoadSini
    Call CargaApuestas
    
    ' Npcs.dat
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat

    ' Obj.dat
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    Call LoadOBJData
    
    ' Hechizos.dat
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
        
    ' Objetos de Herreria
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrer�a"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    
    ' Objetos de Capinteria
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpinter�a"
    Call LoadObjCarpintero
    
    ' Balance.dat
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance
    
    ' Armaduras faccionarias
    frmCargando.Label1(2).Caption = "Cargando ArmadurasFaccionarias.dat"
    Call LoadArmadurasFaccion
    
    ' Pretorianos
    frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
    Call LoadPretorianData

    ' Mapas
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    
    ' Home distance
    Call generateMatrix(MATRIX_INITIAL_MAP)
    
    ' Connections
    Call ResetUsersConnections
    
    ' Timers
    Call InitMainTimers
    
    ' Sockets
    Call SocketConfig
    
    ' End loading..
    Unload frmCargando
    
    'Log start time
    LogServerStartTime
    
    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    
    tInicioServer = GetTickCount() And &H7FFFFFFF
    Call InicializaEstadisticas

    Call MainLoop
End Sub

Private Sub LoadConstants()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Loads all constants and general parameters.
'*****************************************************************
On Error Resume Next
   
    LastBackup = Format(Now, "Short Time")
    Minutos = Format(Now, "Short Time")
    
    ' Paths
    IniPath = App.Path & "\"
    DatPath = App.Path & "\Dat\"
    CharPath = App.Path & "\Charfile\"
    
    ' Skills by level
    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100
    
    ' Races
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    
    ' Classes
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    ' Skills
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasi�n en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apu�alar) = "Apu�alar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    
    ' Attributes
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    
    ' Fishes
    ListaPeces(1) = PECES_POSIBLES.PESCADO1
    ListaPeces(2) = PECES_POSIBLES.PESCADO2
    ListaPeces(3) = PECES_POSIBLES.PESCADO3
    ListaPeces(4) = PECES_POSIBLES.PESCADO4

    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
    Set Ayuda = New cCola
    Set Denuncias = New cCola
    Denuncias.MaxLenght = MAX_DENOUNCES

    With Prision
        .Map = 66
        .X = 75
        .Y = 47
    End With
    
    With Libertad
        .Map = 66
        .X = 75
        .Y = 65
    End With

    MaxUsers = 0

    ' Initialize classes
    Set WSAPISock2Usr = New Collection
    Protocol.InitAuxiliarBuffer

    Set aClon = New clsAntiMassClon
    Set TrashCollector = New Collection

End Sub

Private Sub LoadArrays()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Loads all arrays
'*****************************************************************
On Error Resume Next
    ' Load Records
    Call LoadRecords
    ' Load guilds info
    Call LoadGuildsDB
    ' Load spawn list
    Call CargarSpawnList
    ' Load forbidden words
    Call CargarForbidenWords
End Sub

Private Sub ResetUsersConnections()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Resets Users Connections.
'*****************************************************************
On Error Resume Next

    Dim LoopC As Long
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
End Sub

Private Sub InitMainTimers()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Initializes Main Timers.
'*****************************************************************
On Error Resume Next

    With frmMain
        .AutoSave.Enabled = True
    End With
    
End Sub

Private Sub SocketConfig()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Sets socket config.
'*****************************************************************
On Error Resume Next

    Call SecurityIp.InitIpTables(1000)
    
#If UsarQueSocket = 1 Then
    
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")
    
#ElseIf UsarQueSocket = 0 Then
    
    frmCargando.Label1(2).Caption = "Configurando Sockets"
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Binary = False
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then
    
    frmMain.Serv.Iniciar Puerto
    
#ElseIf UsarQueSocket = 3 Then
    
    frmMain.TCPServ.Encolar True
    frmMain.TCPServ.IniciarTabla 1009
    frmMain.TCPServ.SetQueueLim 51200
    frmMain.TCPServ.Iniciar Puerto
    
#End If
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    
End Sub

Private Sub LogServerStartTime()
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 15/03/2011
'Logs Server Start Time.
'*****************************************************************
    Dim N As Integer
    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #N

End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************

    FileExist = LenB(dir$(File, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************

    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    frmMain.txtNumUsers.Text = NumUsers

End Sub


Public Sub LogCriticEvent(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoReal(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoCaos(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, desc
    Close #nfile

Exit Sub

ErrHandler:

End Sub


Public Sub LogIndex(ByVal index As Integer, ByVal desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub


Public Sub LogError(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogStatic(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

Exit Sub

ErrHandler:

End Sub

Public Sub LogTarea(desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile(1) ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & desc
    Close #nfile

Exit Sub

ErrHandler:


End Sub


Public Sub LogClanes(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub

Public Sub LogGM(Nombre As String, texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************�

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogAsesinato(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub
Public Sub LogHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogCheating(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Sub LogAntiCheat(texto As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, ""
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Arg As String
    Dim i As Integer
    
    
    For i = 1 To 33
    
    Arg = ReadField(i, cad, 44)
    
    If LenB(Arg) = 0 Then Exit Function
    
    Next i
    
    ValidInputNP = True

End Function


Sub Restart()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    
    Dim LoopC As Long
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next
    
    'Initialize statistics!!
    Call Statistics.Initialize
    
    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC
    
    ReDim UserList(1 To MaxUsers) As User
    
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    LastUser = 0
    NumUsers = 0
    
    Call FreeNPCs
    Call FreeCharIndexes
    
    Call LoadSini
    
    Call ResetForums
    Call LoadOBJData
    
    Call LoadMapData
    
    Call CargarHechizos

#If UsarQueSocket = 0 Then

    '*****************Setup socket
    frmMain.Socket1.AddressFamily = AF_INET
    frmMain.Socket1.Protocol = IPPROTO_IP
    frmMain.Socket1.SocketType = SOCK_STREAM
    frmMain.Socket1.Binary = False
    frmMain.Socket1.Blocking = False
    frmMain.Socket1.BufferSize = 1024
    
    frmMain.Socket2(0).AddressFamily = AF_INET
    frmMain.Socket2(0).Protocol = IPPROTO_IP
    frmMain.Socket2(0).SocketType = SOCK_STREAM
    frmMain.Socket2(0).Blocking = False
    frmMain.Socket2(0).BufferSize = 2048
    
    'Escucha
    frmMain.Socket1.LocalPort = val(Puerto)
    frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    
    'Log it
    Dim N As Integer
    N = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #N
    Print #N, Date & " " & time & " servidor reiniciado."
    Close #N
    
    'Ocultar
    
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If

  
End Sub


Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author: Unknown
'Last Modify Date: 15/11/2009
'15/11/2009: ZaMa - La lluvia no quita stamina en las arenas.
'23/11/2009: ZaMa - Optimizacion de codigo.
'**************************************************************

    With UserList(UserIndex)
        If MapInfo(.Pos.Map).Zona <> "DUNGEON" Then
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 1 And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 2 And _
               MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> 4 Then Intemperie = True
        Else
            Intemperie = False
        End If
    End With
    
    'En las arenas no te afecta la lluvia
    If IsArena(UserIndex) Then Intemperie = False
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    If UserList(UserIndex).flags.UserLogged Then
        If Intemperie(UserIndex) Then
            Dim modifi As Long
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 3)
            Call QuitarSta(UserIndex, modifi)
            Call FlushBuffer(UserIndex)
        End If
    End If
    
    Exit Sub
ErrHandler:
    LogError ("Error en EfectoLluvia")
End Sub

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        With UserList(UserIndex)
            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                   Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = _
                   Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia - 1
                   If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(.MascotasIndex(i), 0)
                End If
            End If
        End With
    Next i
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unkonwn
'Last Modification: 23/11/2009
'If user is naked and it's in a cold map, take health points from him
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    Dim modifi As Integer
    
    With UserList(UserIndex)
        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else
            If MapInfo(.Pos.Map).Terreno = eTerrain.terrain_nieve Then
                Call WriteConsoleMsg(UserIndex, "��Est�s muriendo de fr�o, abrigate o morir�s!!", FontTypeNames.FONTTYPE_INFO)
                modifi = Porcentaje(.Stats.MaxHp, 5)
                .Stats.MinHp = .Stats.MinHp - modifi
                
                If .Stats.MinHp < 1 Then
                    Call WriteConsoleMsg(UserIndex, "��Has muerto de fr�o!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = 0
                    Call UserDie(UserIndex)
                End If
                
                Call WriteUpdateHP(UserIndex)
            Else
                modifi = Porcentaje(.Stats.MaxSta, 5)
                Call QuitarSta(UserIndex, modifi)
                Call WriteUpdateSta(UserIndex)
            End If
            
            .Counters.Frio = 0
        End If
    End With
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 23/11/2009
'If user is standing on lava, take health points from him
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    With UserList(UserIndex)
            If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
                .Counters.Lava = .Counters.Lava + 1
            Else
                If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
                    Call WriteConsoleMsg(UserIndex, "��Quitate de la lava, te est�s quemando!!", FontTypeNames.FONTTYPE_INFO)
                    .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
                    
                    If .Stats.MinHp < 1 Then
                        Call WriteConsoleMsg(UserIndex, "��Has muerto quemado!!", FontTypeNames.FONTTYPE_INFO)
                        .Stats.MinHp = 0
                        Call UserDie(UserIndex)
                    End If
                    
                    Call WriteUpdateHP(UserIndex)
    
                End If
                
                .Counters.Lava = 0
            End If
    End With
End Sub

''
' Maneja  el efecto del estado atacable
'
' @param UserIndex  El index del usuario a ser afectado por el estado atacable
'

Public Sub EfectoEstadoAtacable(ByVal UserIndex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 18/09/2010 (ZaMa)
'18/09/2010: ZaMa - Ahora se activa el seguro cuando dejas de ser atacable.
'******************************************************

    ' Si ya paso el tiempo de penalizacion
    If Not IntervaloEstadoAtacable(UserIndex) Then
        ' Deja de poder ser atacado
        UserList(UserIndex).flags.AtacablePor = 0
        
        ' Activo el seguro si deja de estar atacable
        If Not UserList(UserIndex).flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn)
        End If
        
        ' Send nick normal
        Call RefreshCharStatus(UserIndex)
    End If
    
End Sub

''
' Maneja el tiempo de arrivo al hogar
'
' @param UserIndex  El index del usuario a ser afectado por el /hogar
'

Public Sub TravelingEffect(ByVal UserIndex As Integer)
'******************************************************
'Author: ZaMa
'Last Update: 01/06/2010 (ZaMa)
'******************************************************

    ' Si ya paso el tiempo de penalizacion
    If IntervaloGoHome(UserIndex) Then
        Call HomeArrival(UserIndex)
    End If

End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
'******************************************************
'Author: Unknown
'Last Update: 16/09/2010 (ZaMa)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'16/09/2010: ZaMa - Se recupera la apariencia de la barca correspondiente despues de terminado el mimetismo.
'******************************************************
    Dim Barco As ObjData
    
    With UserList(UserIndex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    Call ToggleBoatBody(UserIndex)
                Else
                    .Char.body = iFragataFantasmal
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco
                End If
            Else
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            End If
            
            With .Char
                Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
            End With
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Se fue el efecto del mimetismo, puede ser atacado por npcs
            .flags.Ignorado = False
        End If
    End With
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/09/2010 (ZaMa)
'16/09/2010: ZaMa - Al perder el invi cuando navegas, no se manda el mensaje de sacar invi (ya estas visible).
'***************************************************

    With UserList(UserIndex)
        If .Counters.Invisibilidad < IntervaloInvisible Then
            .Counters.Invisibilidad = .Counters.Invisibilidad + 1
        Else
            .Counters.Invisibilidad = RandomNumber(-100, 100) ' Invi variable :D
            .flags.invisible = 0
            If .flags.Oculto = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si navega ya esta visible..
                If Not .flags.Navegando = 1 Then
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
                
            End If
        End If
    End With

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With Npclist(NpcIndex)
        If .Contadores.Paralisis > 0 Then
            .Contadores.Paralisis = .Contadores.Paralisis - 1
        Else
            .flags.Paralizado = 0
            .flags.Inmovilizado = 0
        End If
    End With

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If .Counters.Ceguera > 0 Then
            .Counters.Ceguera = .Counters.Ceguera - 1
        Else
            If .flags.Ceguera = 1 Then
                .flags.Ceguera = 0
                Call WriteBlindNoMore(UserIndex)
            End If
            If .flags.Estupidez = 1 Then
                .flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)
            End If
        
        End If
    End With

End Sub


Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/12/2010
'02/12/2010: ZaMa - Now non-magic clases lose paralisis effect under certain circunstances.
'***************************************************

    With UserList(UserIndex)
    
        If .Counters.Paralisis > 0 Then
        
            Dim CasterIndex As Integer
            CasterIndex = .flags.ParalizedByIndex
        
            ' Only aplies to non-magic clases
            If .Stats.MaxMAN = 0 Then
                ' Paralized by user?
                If CasterIndex <> 0 Then
                
                    ' Close? => Remove Paralisis
                    If UserList(CasterIndex).Name <> .flags.ParalizedBy Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                        
                    ' Caster dead? => Remove Paralisis
                    ElseIf UserList(CasterIndex).flags.Muerto = 1 Then
                        Call RemoveParalisis(UserIndex)
                        Exit Sub
                    
                    ElseIf .Counters.Paralisis > IntervaloParalizadoReducido Then
                        ' Out of vision range? => Reduce paralisis counter
                        If Not InVisionRangeAndMap(UserIndex, UserList(CasterIndex).Pos) Then
                            ' Aprox. 1500 ms
                            .Counters.Paralisis = IntervaloParalizadoReducido
                            Exit Sub
                        End If
                    End If
                
                ' Npc?
                Else
                    CasterIndex = .flags.ParalizedByNpcIndex
                    
                    ' Paralized by npc?
                    If CasterIndex <> 0 Then
                    
                        If .Counters.Paralisis > IntervaloParalizadoReducido Then
                            ' Out of vision range? => Reduce paralisis counter
                            If Not InVisionRangeAndMap(UserIndex, Npclist(CasterIndex).Pos) Then
                                ' Aprox. 1500 ms
                                .Counters.Paralisis = IntervaloParalizadoReducido
                                Exit Sub
                            End If
                        End If
                    End If
                    
                End If
            End If
            
            .Counters.Paralisis = .Counters.Paralisis - 1

        Else
            Call RemoveParalisis(UserIndex)
        End If
    End With

End Sub

Public Sub RemoveParalisis(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Removes paralisis effect from user.
'***************************************************
    With UserList(UserIndex)
        .flags.Paralizado = 0
        .flags.Inmovilizado = 0
        .flags.ParalizedBy = vbNullString
        .flags.ParalizedByIndex = 0
        .flags.ParalizedByNpcIndex = 0
        .Counters.Paralisis = 0
        Call WriteParalizeOK(UserIndex)
    End With
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 2 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
        
        
        Dim massta As Integer
        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
            Else
                EnviarStats = True
                .Counters.STACounter = 0
                If .flags.Desnudo Then Exit Sub 'Desnudo no sube energ�a. (ToxicWaste)
               
                massta = RandomNumber(1, Porcentaje(.Stats.MaxSta, 5))
                .Stats.MinSta = .Stats.MinSta + massta
                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta
                End If
            End If
        End If
    End With
    
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer
    
    With UserList(UserIndex)
        If .Counters.Veneno < IntervaloVeneno Then
          .Counters.Veneno = .Counters.Veneno + 1
        Else
          Call WriteConsoleMsg(UserIndex, "Est�s envenenado, si no te curas morir�s.", FontTypeNames.FONTTYPE_VENENO)
          .Counters.Veneno = 0
          N = RandomNumber(1, 5)
          .Stats.MinHp = .Stats.MinHp - N
          If .Stats.MinHp < 1 Then Call UserDie(UserIndex)
          Call WriteUpdateHP(UserIndex)
        End If
    End With

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
'***************************************************
'Author: ??????
'Last Modification: 11/27/09 (Budi)
'Cuando se pierde el efecto de la poci�n updatea fz y agi (No me gusta que ambos atributos aunque se haya modificado solo uno, pero bueno :p)
'***************************************************
    With UserList(UserIndex)
        'Controla la duracion de las pociones
        If .flags.DuracionEfecto > 0 Then
           .flags.DuracionEfecto = .flags.DuracionEfecto - 1
           If .flags.DuracionEfecto = 0 Then
                .flags.TomoPocion = False
                .flags.TipoPocion = 0
                'volvemos los atributos al estado normal
                Dim loopX As Integer
                
                For loopX = 1 To NUMATRIBUTOS
                    .Stats.UserAtributos(loopX) = .Stats.UserAtributosBackUP(loopX)
                Next loopX
                
                Call WriteUpdateStrenghtAndDexterity(UserIndex)
           End If
        End If
    End With

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If Not .flags.Privilegios And PlayerType.User Then Exit Sub
        
        'Sed
        If .Stats.MinAGU > 0 Then
            If .Counters.AGUACounter < IntervaloSed Then
                .Counters.AGUACounter = .Counters.AGUACounter + 1
            Else
                .Counters.AGUACounter = 0
                .Stats.MinAGU = .Stats.MinAGU - 10
                
                If .Stats.MinAGU <= 0 Then
                    .Stats.MinAGU = 0
                    .flags.Sed = 1
                End If
                
                fenviarAyS = True
            End If
        End If
        
        'hambre
        If .Stats.MinHam > 0 Then
           If .Counters.COMCounter < IntervaloHambre Then
                .Counters.COMCounter = .Counters.COMCounter + 1
           Else
                .Counters.COMCounter = 0
                .Stats.MinHam = .Stats.MinHam - 10
                If .Stats.MinHam <= 0 Then
                       .Stats.MinHam = 0
                       .flags.Hambre = 1
                End If
                fenviarAyS = True
            End If
        End If
    End With

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 2 And _
           MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 4 Then Exit Sub
        
        Dim mashit As Integer
        'con el paso del tiempo va sanando....pero muy lentamente ;-)
        If .Stats.MinHp < .Stats.MaxHp Then
            If .Counters.HPCounter < Intervalo Then
                .Counters.HPCounter = .Counters.HPCounter + 1
            Else
                mashit = RandomNumber(2, Porcentaje(.Stats.MaxSta, 5))
                
                .Counters.HPCounter = 0
                .Stats.MinHp = .Stats.MinHp + mashit
                If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
                EnviarStats = True
            End If
        End If
    End With

End Sub

Public Sub CargaNpcsDat()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Set LeerNPCs = New clsIniManager
    Call LeerNPCs.Initialize(npcfile)
End Sub

Sub PasarSegundo()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    Dim i As Long
    
    'Limpieza del mundo
    If counterSV.Limpieza > 0 Then
        counterSV.Limpieza = counterSV.Limpieza - 1
        
        If counterSV.Limpieza < 6 Then Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza del mundo en " & counterSV.Limpieza & " segundos. ��Atentos!!", FontTypeNames.FONTTYPE_SERVER))
        
        If counterSV.Limpieza = 0 Then
            Call BorrarObjetosLimpieza
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Limpieza del mundo finalizada.", FontTypeNames.FONTTYPE_SERVER))
            UltimoSlotLimpieza = -1
        End If
    End If
    
    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                'Cerrar usuario
                If .Counters.Saliendo Then
                    .Counters.Salir = .Counters.Salir - 1
                    If .Counters.Salir <= 0 Then
                        Call WriteConsoleMsg(i, "Gracias por jugar Argentum Online", FontTypeNames.FONTTYPE_INFO)
                        Call WriteDisconnect(i)
                        Call FlushBuffer(i)
                        
                        Call CloseSocket(i)
                    End If
                End If
            
                If .Counters.Pena > 0 Then
                    'Restamos las penas del personaje
                    If UserList(ArrayPenas(i)).Counters.Pena > 0 Then
                        UserList(ArrayPenas(i)).Counters.Pena = UserList(ArrayPenas(i)).Counters.Pena - 1
                    
                        If UserList(ArrayPenas(i)).Counters.Pena < 1 Then
                            UserList(ArrayPenas(i)).Counters.Pena = 0
                            Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                            Call WriteConsoleMsg(i, "�Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                            Call QuitarArrayCarcel(i)
                            Call FlushBuffer(i)
                        End If
                    End If
                End If
                
                'Sacamos energ�a
                If Lloviendo Then _
                    Call EfectoLluvia(i)
                
                If Not .Pos.Map = 0 Then
                'Counter de piquete
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                            If .flags.Muerto = 0 Then
                                .Counters.PiqueteC = .Counters.PiqueteC + 1
                                Call WriteConsoleMsg(i, "���Est�s obstruyendo la v�a p�blica, mu�vete o ser�s encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                                
                                If .Counters.PiqueteC >= 5 Then
                                    .Counters.PiqueteC = 0
                                    Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                                End If
                                
                                Call FlushBuffer(i)
                            Else
                                .Counters.PiqueteC = 0
                            End If
                    Else
                        .Counters.PiqueteC = 0
                    End If
                End If
            End If
        End With
    Next i
Exit Sub

ErrHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'WorldSave
    Call ES.DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).Name) & ".chr", False)
        End If
    Next i
    
    'se guardan los seguimientos
    Call SaveRecords
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Ta As Long
    Ta = GetTickCount() And &H7FFFFFFF
    
    Set EstadisticasWeb = New clsEstadisticasIPC
    Call EstadisticasWeb.Inicializa(frmMain.hWnd)
    Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, RECORDusuarios)

End Sub

Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub

Public Sub ReproducirSonido(ByVal Destino As SendTarget, ByVal index As Integer, ByVal SoundIndex As Integer)
    Call SendData(Destino, index, PrepareMessagePlayWave(SoundIndex, UserList(index).Pos.X, UserList(index).Pos.Y))
End Sub
