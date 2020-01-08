Attribute VB_Name = "ES"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

#If False Then

    Dim X, Y, n, Map, Mapa, Email, max, Value As Variant

#End If

Public Sub CargarSpawnList()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Invokar.dat"

    Dim n As Integer, LoopC As Integer

    n = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador

    For LoopC = 1 To n
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
    If frmMain.Visible Then frmMain.txtStatus.Text = "Invokar.dat se cargo correctamente"
    
End Sub

Function EsAdmin(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)

End Function

Function EsDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)

End Function

Function EsSemiDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)

End Function

Function EsGmEspecial(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsGmEspecial = (val(Administradores.GetValue("Especial", Name)) = 1)

End Function

Function EsConsejero(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)

End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)

End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    'Returns true if char is administrative user.
    '***************************************************
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)

    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)

    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)

    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

End Function

Public Sub loadAdministrativeUsers()
    'Admines     => Admin
    'Dioses      => Dios
    'SemiDioses  => SemiDios
    'Especiales  => Especial
    'Consejeros  => Consejero
    'RoleMasters => RM
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Administradores/Dioses/Gms."

    'Si esta mierda tuviese array asociativos el codigo seria tan lindo.
    Dim buf  As Integer

    Dim i    As Long

    Dim Name As String
       
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager

    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(IniPath & "Server.ini")
       
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Admines"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next i
    
    ' Especiales
    buf = val(ServerIni.GetValue("INIT", "Especiales"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Especial", Name, "1")
        
    Next i
    
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next i
    
    Set ServerIni = Nothing

    If frmMain.Visible Then frmMain.txtStatus.Text =  Date & " " & time & " - Los Administradores/Dioses/Gms se han cargado correctamente."

End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
    '****************************************************
    'Author: ZaMa
    'Last Modification: 18/11/2010
    'Reads the user's charfile and retrieves its privs.
    '***************************************************

    Dim Privs As PlayerType

    If EsAdmin(UserName) Then
        Privs = PlayerType.Admin
        
    ElseIf EsDios(UserName) Then
        Privs = PlayerType.Dios

    ElseIf EsSemiDios(UserName) Then
        Privs = PlayerType.SemiDios
        
    ElseIf EsConsejero(UserName) Then
        Privs = PlayerType.Consejero
    
    Else
        Privs = PlayerType.User

    End If

    GetCharPrivs = Privs

End Function

Public Function TxtDimension(ByVal Name As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer, cad As String, Tam As Long

    n = FreeFile(1)
    Open Name For Input As #n
    Tam = 0

    Do While Not EOF(n)
        Tam = Tam + 1
        Line Input #n, cad
    Loop
    Close n
    TxtDimension = Tam

End Function

Public Sub CargarForbidenWords()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Nombres prohibidos (NombresInvalidos.txt)."

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))

    Dim n As Integer, i As Integer

    n = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #n
    
    For i = 1 To UBound(ForbidenNames)
        Line Input #n, ForbidenNames(i)
    Next i
    
    Close n

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - NombresInvalidos.txt han cargado con exito."

End Sub

Public Sub CargarHechizos()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '   NO USAR GetVar PARA LEER Hechizos.dat !!!!
    '
    'El que ose desafiar esta LEY, se las tendra que ver
    'con migo. Para leer Hechizos.dat se debera usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Hechizos."
    
    Dim Hechizo As Integer

    Dim Leer    As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos

        With Hechizos(Hechizo)
            '.Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            '.desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            '.PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            '.HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            '.TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            '.PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
            '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
            '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
            '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            'Barrin 30/9/03
            .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            
            .NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
            .StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))

        End With

    Next Hechizo
    
    Set Leer = Nothing

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los hechizos se han cargado con exito."
    
    Exit Sub

ErrHandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando archivo MOTD.INI."

    Dim i As Integer
    
    MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
    
    ReDim MOTD(1 To MaxLines)

    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = vbNullString
    Next i

    If frmMain.Visible Then frmMain.txtStatus.Text =  Date & " " & time & " - El archivo MOTD.INI fue cargado con exito"

End Sub

Public Sub DoBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los hechizos se han cargado con exito."

    haciendoBK = True
    
    ' Lo saco porque elimina elementales y mascotas - Maraxus
    ''''''''''''''lo pongo aca x sugernecia del yind
    'For i = 1 To LastNPC
    '    If Npclist(i).flags.NPCActive Then
    '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
    '            Call MuereNpc(i, 0)
    '        End If
    '    End If
    'Next i
    '''''''''''/'lo pongo aca x sugernecia del yind
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    Call WorldSave
    Call modGuilds.v_RutinaElecciones
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
    'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
    'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
    If ConexionAPI Then
        Call ApiEndpointBackupCharfiles
        Call ApiEndpointBackupCuentas
        Call ApiEndpointBackupLogs
        Call ApiEndpointSendWorldSaveMessageDiscord
    End If

    haciendoBK = False
    
    'Log
    On Error Resume Next

    Dim nfile As Integer

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - El WorldSave (backup) se hizo correctamente."

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time
    Close #nfile

End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2011
    '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
    '28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
    '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados )
    '***************************************************

    On Error Resume Next

    Dim FreeFileMap As Long

    Dim FreeFileInf As Long

    Dim Y           As Long

    Dim X           As Long

    Dim ByFlags     As Byte

    Dim LoopC       As Long

    Dim MapWriter   As clsByteBuffer

    Dim InfWriter   As clsByteBuffer

    Dim IniManager  As clsIniManager

    Dim NpcInvalido As Boolean
    
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"

    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"

    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(Map).MapVersion)
        
    Call MapWriter.putString(MiCabecera.Desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putInteger(.Graphic(1))
                
                For LoopC = 2 To 4

                    If .Graphic(LoopC) Then Call MapWriter.putInteger(.Graphic(LoopC))
                Next LoopC
                
                If .trigger Then Call MapWriter.putInteger(CInt(.trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.ObjIndex > 0 Then
                    If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        .ObjInfo.ObjIndex = 0
                        .ObjInfo.Amount = 0

                    End If

                End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs invalidos (Pretorianos, Mascotas, Invocados )
                If .NpcIndex Then
                    NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.Pretoriano) Or (Npclist(.NpcIndex).MaestroUser > 0)
                    
                    If Not NpcInvalido Then ByFlags = ByFlags Or 2

                End If
                
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)

                End If
                
                If .NpcIndex And Not NpcInvalido Then Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.putInteger(.ObjInfo.ObjIndex)
                    Call InfWriter.putInteger(.ObjInfo.Amount)

                End If
                
                NpcInvalido = False

            End With

        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(Map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & Map, "Name", .Name)
        Call IniManager.ChangeValue("Mapa" & Map, "MusicNum", .Music)
        Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
        Call IniManager.ChangeValue("Mapa" & Map, "OnDeathGoTo", .OnDeathGoTo.Map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y)
    
        Call IniManager.ChangeValue("Mapa" & Map, "Terreno", TerrainByteToString(.Terreno))
        Call IniManager.ChangeValue("Mapa" & Map, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & Map, "Restringir", RestrictByteToString(.Restringir))
        Call IniManager.ChangeValue("Mapa" & Map, "BackUp", str(.BackUp))
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")

        End If
        
        Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InvocarSinEfecto", .InvocarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "NoEncriptarMP", .NoEncriptarMP)
        Call IniManager.ChangeValue("Mapa" & Map, "RoboNpcsPermitido", .RoboNpcsPermitido)
    
        Call IniManager.DumpFile(MAPFILE & ".dat")

    End With
    
    Set IniManager = Nothing

End Sub

Sub LoadArmasHerreria()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc

End Sub

Sub LoadArmadurasHerreria()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc

End Sub

Sub LoadBalance()
    '***************************************************
    'Author: Unknown
    'Last Modification: 15/04/2010
    '15/04/2010: ZaMa - Agrego recompensas faccionarias.
    '***************************************************

    Dim i As Long

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando el archivo Balance.dat"
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES

        With ModClase(i)
            .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueWrestling = val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .DanoArmas = val(GetVar(DatPath & "Balance.dat", "MODDANOARMAS", ListaClases(i)))
            .DanoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDANOPROYECTILES", ListaClases(i)))
            .DanoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDANOWRESTLING", ListaClases(i)))
            .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))

        End With

    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS

        With ModRaza(i)
            .Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))

        End With

    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribucion de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i

    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
    
    ' Recompensas faccionarias
    For i = 1 To NUM_RANGOS_FACCION
        RecompensaFacciones(i - 1) = val(GetVar(DatPath & "Balance.dat", "RECOMPENSAFACCION", "Rango" & i))
    Next i
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo con exito el archivo Balance.dat"

End Sub

Sub LoadObjCarpintero()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To n) As Integer
    
    For lc = 1 To n
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc

End Sub

Sub LoadOBJData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    ' NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
    '
    'El que ose desafiar esta LEY, se las tendra que ver
    'con migo. Para leer desde el OBJ.DAT se debera usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    'Call LogTarea("Sub LoadOBJData")

    On Error GoTo ErrHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando base de datos de los objetos."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer

    Dim Leer   As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.Value = 0
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    'Llena la lista
    For Object = 1 To NumObjDatas

        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            
            'Pablo (ToxicWaste) Log de Objetos.
            .Log = val(Leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))

            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex

            End If
            
            .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
            
            .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
            Select Case .OBJType

                Case eOBJType.otarmadura
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                
                Case eOBJType.otescudo
                    .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otcasco
                    .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otWeapon
                    .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apunala = val(Leer.GetValue("OBJ" & Object, "Apunala"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                    .StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                    .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                    .WeaponRazaEnanaAnim = val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                
                Case eOBJType.otInstrumentos
                    .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otMinerales
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                    .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                Case eOBJType.otBarcos
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                Case eOBJType.otFlechas
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    
                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    
                Case eOBJType.otTeleport
                    .Radio = val(Leer.GetValue("OBJ" & Object, "Radio"))
                    
                Case eOBJType.otMochilas
                    .MochilaType = val(Leer.GetValue("OBJ" & Object, "MochilaType"))
                    
                Case eOBJType.otForos
                    Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))
                    
            End Select
            
            .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            
            .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
            .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
            .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
            .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
            .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
            .Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
            
            .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))

            If .Cerrada = 1 Then
                .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
                .Clave = val(Leer.GetValue("OBJ" & Object, "Clave"))

            End If
            
            'Puertas y llaves
            .Clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            .Acuchilla = val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            
            .Guante = val(Leer.GetValue("OBJ" & Object, "Guante"))
            
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            Dim i As Integer

            Dim n As Integer

            Dim S As String

            For i = 1 To NUMCLASES
                S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
                n = 1

                Do While LenB(S) > 0 And UCase$(ListaClases(n)) <> S
                    n = n + 1
                Loop
                .ClaseProhibida(i) = IIf(LenB(S) > 0, n, 0)
            Next i
            
            .DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
            .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            If .SkCarpinteria > 0 Then .Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
            .MaderaElfica = val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
            
            'Bebidas
            .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            .Upgrade = val(Leer.GetValue("OBJ" & Object, "Upgrade"))
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1

        End With

    Next Object
    
    Set Leer = Nothing
    
    ' Inicializo los foros faccionarios
    Call AddForum(FORO_CAOS_ID)
    Call AddForum(FORO_REAL_ID)

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo base de datos de los objetos. Operacion Realizada con exito."
    
    Exit Sub
ErrHandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description

End Sub

Sub LoadUserStats(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)

    '*************************************************
    'Author: Unknown
    'Last modified: 11/19/2009
    '11/19/2009: Pato - Load the EluSkills and ExpSkills
    '*************************************************
    Dim LoopC As Long

    With UserList(Userindex)
        With .Stats

            For LoopC = 1 To NUMATRIBUTOS
                .UserAtributos(LoopC) = CByte(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
                .UserAtributosBackUP(LoopC) = CByte(.UserAtributos(LoopC))
            Next LoopC
        
            For LoopC = 1 To NUMSKILLS
                .UserSkills(LoopC) = CByte(UserFile.GetValue("SKILLS", "SK" & LoopC))
                .EluSkills(LoopC) = CLng(UserFile.GetValue("SKILLS", "ELUSK" & LoopC))
                .ExpSkills(LoopC) = CLng(UserFile.GetValue("SKILLS", "EXPSK" & LoopC))
            Next LoopC
        
            For LoopC = 1 To MAXUSERHECHIZOS
                .UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
            Next LoopC
        
            .Gld = CLng(UserFile.GetValue("STATS", "GLD"))
            .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
        
            .MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
            .MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))
        
            .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
            .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
        
            .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
            .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
        
            .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
            .MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))
        
            .MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
            .MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))
        
            .MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
            .MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))
        
            .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))
        
            .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
            .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
            .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
        
            .UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
            .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

        End With
    
        With .flags

            If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
        
            If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then .Privilegios = .Privilegios Or PlayerType.ChaosCouncil

        End With

    End With

End Sub

Sub LoadUserReputacion(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)
    '***************************************************
    'Author: Unknown
    'Last Modification: Recox
    'Recox - Castie todo a long para que sea el mismo tipo de dato que en Declares
    '***************************************************

    With UserList(Userindex).Reputacion
        .AsesinoRep = CLng(UserFile.GetValue("REP", "Asesino"))
        .BandidoRep = CLng(UserFile.GetValue("REP", "Bandido"))
        .BurguesRep = CLng(UserFile.GetValue("REP", "Burguesia"))
        .LadronesRep = CLng(UserFile.GetValue("REP", "Ladrones"))
        .NobleRep = CLng(UserFile.GetValue("REP", "Nobles"))
        .PlebeRep = CLng(UserFile.GetValue("REP", "Plebe"))
        .Promedio = CLng(UserFile.GetValue("REP", "Promedio"))

    End With
    
End Sub

Sub LoadUserInit(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)

    '*************************************************
    'Author: Unknown
    'Last modified: 19/11/2019
    'Loads the Users RECORDs
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
    '19/11/2019 Recox - Casteo todas las propiedades a su tipo de dato en Declares para evitar errores
    '*************************************************
    Dim LoopC As Long

    Dim ln    As String
    
    With UserList(Userindex)
        With .Faccion
            .ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
            .FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
            .CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
            .CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
            .RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
            .RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
            .RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
            .RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
            .RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
            .RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
            .Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
            .NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
            .FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
            .MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
            .NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

        End With
        
        With .flags
            .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
            .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
            
            .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
            .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
            .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
            .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
            .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
            .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
            
            'Matrix
            .lastMap = val(UserFile.GetValue("FLAGS", "LastMap"))

        End With

        .Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
        .Counters.AsignedSkills = CByte(val(UserFile.GetValue("COUNTERS", "SkillsAsignados")))
        
        .Email = UserFile.GetValue("CONTACTO", "Email")
        
        .AccountHash = CStr(UserFile.GetValue("INIT", "AccountHash"))
        .Genero = CByte(UserFile.GetValue("INIT", "Genero"))
        .Clase = CByte(UserFile.GetValue("INIT", "Clase"))
        .raza = CByte(UserFile.GetValue("INIT", "Raza"))
        .Hogar = CByte(UserFile.GetValue("INIT", "Hogar"))
        .Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))
        
        With .OrigChar
            .Head = CInt(UserFile.GetValue("INIT", "Head"))
            .body = CInt(UserFile.GetValue("INIT", "Body"))
            .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
            .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
            .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
            .heading = eHeading.SOUTH

        End With
        
        #If ConUpTime Then
            .UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If

        .Desc = UserFile.GetValue("INIT", "Desc")
        
        .Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
        
        .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
        
        '[KEVIN]--------------------------------------------------------------------
        '***********************************************************************************
        .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

        'Lista de objetos del banco
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
            .BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
            .BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
        Next LoopC

        '------------------------------------------------------------------------------------
        '[/KEVIN]*****************************************************************************
        
        'Lista de objetos
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            .Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
        Next LoopC
        
        .Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
        .Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
        .Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
        .Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
        .Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
        .Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
        .Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
        .Invent.MochilaEqpSlot = CByte(UserFile.GetValue("Inventory", "MochilaSlot"))
        
        .NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))

        For LoopC = 1 To MAXMASCOTAS
            .MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
        Next LoopC
        
        ln = UserFile.GetValue("Guild", "GUILDINDEX")

        If IsNumeric(ln) Then
            .GuildIndex = CInt(ln)
        Else
            .GuildIndex = 0

        End If

    End With

End Sub

Function GetVar(ByVal File As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim sSpaces  As String ' This will hold the input that the program will retrieve

    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando backup."
    
    Dim Map       As Integer

    Dim tFileName As String
    
    On Error GoTo man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
        
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
        
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
        
    For Map = 1 To NumMaps

        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
                
            If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                tFileName = App.Path & MapPath & "Mapa" & Map

            End If

        Else
            tFileName = App.Path & MapPath & "Mapa" & Map

        End If
            
        Call CargarMapa(Map, tFileName)
            
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
    
    Exit Sub

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se termino de cargar el backup."

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando mapas..."
    
    Dim Map       As Integer

    Dim tFileName As String
    
    On Error GoTo man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
        
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
        
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
          
    For Map = 1 To NumMaps
            
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
            
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map
    
    Exit Sub

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargaron todos los mapas. Operacion Realizada con exito."

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByRef MAPFl As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/08/2010
    '10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
    '***************************************************

    On Error GoTo errh

    Dim hFile     As Integer

    Dim X         As Long

    Dim Y         As Long

    Dim ByFlags   As Byte

    Dim npcfile   As String

    Dim Leer      As clsIniManager

    Dim MapReader As clsByteBuffer

    Dim InfReader As clsByteBuffer

    Dim Buff()    As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set Leer = New clsIniManager
    
    npcfile = DatPath & "NPCs.dat"
    
    hFile = FreeFile

    Open MAPFl & ".map" For Binary As #hFile
    Seek hFile, 1

    ReDim Buff(LOF(hFile) - 1) As Byte
    
    Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'inf
    Open MAPFl & ".inf" For Binary As #hFile
    Seek hFile, 1

    ReDim Buff(LOF(hFile) - 1) As Byte
    
    Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    MapInfo(Map).MapVersion = MapReader.getInteger
    
    MiCabecera.Desc = MapReader.getString(Len(MiCabecera.Desc))
    MiCabecera.crc = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            With MapData(Map, X, Y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1

                .Graphic(1) = MapReader.getInteger

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getInteger

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getInteger

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getInteger

                'Trigger used?
                If ByFlags And 16 Then .trigger = MapReader.getInteger

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.Map = InfReader.getInteger
                    .TileExit.X = InfReader.getInteger
                    .TileExit.Y = InfReader.getInteger

                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                    .NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then

                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        If val(GetVar(npcfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
                            .NpcIndex = OpenNPC(.NpcIndex)
                            Npclist(.NpcIndex).Orig.Map = Map
                            Npclist(.NpcIndex).Orig.X = X
                            Npclist(.NpcIndex).Orig.Y = Y
                        Else
                            .NpcIndex = OpenNPC(.NpcIndex)

                        End If

                        Npclist(.NpcIndex).Pos.Map = Map
                        Npclist(.NpcIndex).Pos.X = X
                        Npclist(.NpcIndex).Pos.Y = Y

                        Call MakeNPCChar(True, 0, .NpcIndex, Map, X, Y)

                    End If

                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    .ObjInfo.ObjIndex = InfReader.getInteger
                    .ObjInfo.Amount = InfReader.getInteger

                End If

            End With

        Next X
    Next Y
    
    Call Leer.Initialize(MAPFl & ".dat")
    
    With MapInfo(Map)
        .Name = Leer.GetValue("Mapa" & Map, "Name")
        .Music = Leer.GetValue("Mapa" & Map, "MusicNum")
        .StartPos.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "StartPos"), Asc("-")))
        
        .OnDeathGoTo.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        
        .MagiaSinEfecto = val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
        .InviSinEfecto = val(Leer.GetValue("Mapa" & Map, "InviSinEfecto"))
        .ResuSinEfecto = val(Leer.GetValue("Mapa" & Map, "ResuSinEfecto"))
        .OcultarSinEfecto = val(Leer.GetValue("Mapa" & Map, "OcultarSinEfecto"))
        .InvocarSinEfecto = val(Leer.GetValue("Mapa" & Map, "InvocarSinEfecto"))
        
        .NoEncriptarMP = val(Leer.GetValue("Mapa" & Map, "NoEncriptarMP"))

        .RoboNpcsPermitido = val(Leer.GetValue("Mapa" & Map, "RoboNpcsPermitido"))
        
        If val(Leer.GetValue("Mapa" & Map, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False

        End If
        
        .Terreno = TerrainStringToByte(Leer.GetValue("Mapa" & Map, "Terreno"))
        .Zona = Leer.GetValue("Mapa" & Map, "Zona")
        .Restringir = RestrictStringToByte(Leer.GetValue("Mapa" & Map, "Restringir"))
        .BackUp = val(Leer.GetValue("Mapa" & Map, "BACKUP"))

    End With
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
    
    Erase Buff
    Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.description)

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing

End Sub

Sub LoadSini()
'***************************************************
'Author: Unknown
'Last Modification: 13/11/2019 (Recox)
'CHOTS: Database params
'Cucsifae: Agregados multiplicadores exp y oro
'CHOTS: Agregado multiplicador oficio
'CHOTS: Agregado min y max Dados
'Jopi: Uso de clsIniManager para cargar los valores.
'Recox: Cargamos si el centinela esta activo o no.
'***************************************************

    Dim Temporal As Long
    
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    
    If frmMain.Visible Then
        frmMain.txtStatus.Text = "Cargando info de inicio del server."
    End If
    
    Call Lector.Initialize(IniPath & "Server.ini")
    
    BootDelBackUp = CBool(val(Lector.GetValue("INIT", "IniciarDesdeBackUp")))
    
    'Misc
    Puerto = val(Lector.GetValue("INIT", "StartPort"))
    HideMe = CBool(Lector.GetValue("INIT", "Hide"))
    AllowMultiLogins = CBool(val(Lector.GetValue("INIT", "AllowMultiLogins")))
    IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
    
    'Lee la version correcta del cliente
    ULTIMAVERSION = Lector.GetValue("INIT", "VersionBuildCliente")
    
    STAT_MAXELV = val(Lector.GetValue("INIT", "NivelMaximo"))
    
    ExpMultiplier = val(Lector.GetValue("INIT", "ExpMulti"))
    OroMultiplier = val(Lector.GetValue("INIT", "OroMulti"))
    OficioMultiplier = val(Lector.GetValue("INIT", "OficioMulti"))
    DiceMinimum = val(Lector.GetValue("INIT", "MinDados"))
    DiceMaximum = val(Lector.GetValue("INIT", "MaxDados"))
    
    DropItemsAlMorir = CBool(Lector.GetValue("INIT", "DropItemsAlMorir"))

    'Esto es para ver si el centinela esta activo o no.
    isCentinelaActivated = CBool(val(Lector.GetValue("INIT", "CentinelaAuditoriaTrabajoActivo")))

    PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = val(Lector.GetValue("INIT", "ServerSoloGMs"))
    
    ArmaduraImperial1 = val(Lector.GetValue("INIT", "ArmaduraImperial1"))
    ArmaduraImperial2 = val(Lector.GetValue("INIT", "ArmaduraImperial2"))
    ArmaduraImperial3 = val(Lector.GetValue("INIT", "ArmaduraImperial3"))
    TunicaMagoImperial = val(Lector.GetValue("INIT", "TunicaMagoImperial"))
    TunicaMagoImperialEnanos = val(Lector.GetValue("INIT", "TunicaMagoImperialEnanos"))
    ArmaduraCaos1 = val(Lector.GetValue("INIT", "ArmaduraCaos1"))
    ArmaduraCaos2 = val(Lector.GetValue("INIT", "ArmaduraCaos2"))
    ArmaduraCaos3 = val(Lector.GetValue("INIT", "ArmaduraCaos3"))
    TunicaMagoCaos = val(Lector.GetValue("INIT", "TunicaMagoCaos"))
    TunicaMagoCaosEnanos = val(Lector.GetValue("INIT", "TunicaMagoCaosEnanos"))
    
    VestimentaImperialHumano = val(Lector.GetValue("INIT", "VestimentaImperialHumano"))
    VestimentaImperialEnano = val(Lector.GetValue("INIT", "VestimentaImperialEnano"))
    TunicaConspicuaHumano = val(Lector.GetValue("INIT", "TunicaConspicuaHumano"))
    TunicaConspicuaEnano = val(Lector.GetValue("INIT", "TunicaConspicuaEnano"))
    ArmaduraNobilisimaHumano = val(Lector.GetValue("INIT", "ArmaduraNobilisimaHumano"))
    ArmaduraNobilisimaEnano = val(Lector.GetValue("INIT", "ArmaduraNobilisimaEnano"))
    ArmaduraGranSacerdote = val(Lector.GetValue("INIT", "ArmaduraGranSacerdote"))
    
    VestimentaLegionHumano = val(Lector.GetValue("INIT", "VestimentaLegionHumano"))
    VestimentaLegionEnano = val(Lector.GetValue("INIT", "VestimentaLegionEnano"))
    TunicaLobregaHumano = val(Lector.GetValue("INIT", "TunicaLobregaHumano"))
    TunicaLobregaEnano = val(Lector.GetValue("INIT", "TunicaLobregaEnano"))
    TunicaEgregiaHumano = val(Lector.GetValue("INIT", "TunicaEgregiaHumano"))
    TunicaEgregiaEnano = val(Lector.GetValue("INIT", "TunicaEgregiaEnano"))
    SacerdoteDemoniaco = val(Lector.GetValue("INIT", "SacerdoteDemoniaco"))
    
    MAPA_PRETORIANO = val(Lector.GetValue("CLAN-PRETORIANO", "Mapa"))
    PRETORIANO_X = val(Lector.GetValue("CLAN-PRETORIANO", "X"))
    PRETORIANO_Y = val(Lector.GetValue("CLAN-PRETORIANO", "Y"))
    
    EnTesting = CBool(Lector.GetValue("INIT", "Testing"))
    
    ContadorAntiPiquete = val(Lector.GetValue("INIT", "ContadorAntiPiquete"))
    MinutosCarcelPiquete = val(Lector.GetValue("INIT", "MinutosCarcelPiquete"))

    'Usar Mundo personalizado / Use custom world
    UsarMundoPropio = CBool(Lector.GetValue("MUNDO", "UsarMundoPropio"))

    'Inventario Inicial
    InventarioUsarConfiguracionPersonalizada = CBool(val(Lector.GetValue("INVENTARIO", "InventarioUsarConfiguracionPersonalizada")))

    'Atributos Iniciales
    EstadisticasInicialesUsarConfiguracionPersonalizada = CBool(val(Lector.GetValue("ESTADISTICASINICIALESPJ", "Activado")))

    'Intervalos
    SanaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloSinDescansar"))
    StaminaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloSinDescansar"))
    SanaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloDescansar"))
    StaminaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloDescansar"))
    IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed"))
    IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre"))
    IntervaloVeneno = val(Lector.GetValue("INTERVALOS", "IntervaloVeneno"))
    IntervaloParalizado = val(Lector.GetValue("INTERVALOS", "IntervaloParalizado"))
    IntervaloInvisible = val(Lector.GetValue("INTERVALOS", "IntervaloInvisible"))
    IntervaloFrio = val(Lector.GetValue("INTERVALOS", "IntervaloFrio"))
    IntervaloWavFx = val(Lector.GetValue("INTERVALOS", "IntervaloWAVFX"))
    IntervaloInvocacion = val(Lector.GetValue("INTERVALOS", "IntervaloInvocacion"))
    IntervaloParaConexion = val(Lector.GetValue("INTERVALOS", "IntervaloParaConexion"))
    IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
    IntervaloUserPuedeTrabajar = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajo"))
    IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
    
    'TODO : Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    IntervaloPuedeSerAtacado = val(Lector.GetValue("INTERVALOS", "IntervaloPuedeSerAtacado"))
    IntervaloAtacable = val(Lector.GetValue("INTERVALOS", "IntervaloAtacable"))
    IntervaloOwnedNpc = val(Lector.GetValue("INTERVALOS", "IntervaloOwnedNpc"))

    MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))

    If MinutosWs < 60 Then MinutosWs = 180
    
    MinutosGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
    IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
    
    IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))
    
    '&&&&&&&&&&&&&&&&&&&&& SUERTE &&&&&&&&&&&&&&&&&&&&&&&
    DificultadPescar = val(Lector.GetValue("DIFICULTAD", "DificultadPescar"))
    DificultadTalar = val(Lector.GetValue("DIFICULTAD", "DificultadTalar"))
    DificultadMinar = val(Lector.GetValue("DIFICULTAD", "DificultadMinar"))
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
      
    RecordUsuariosOnline = val(Lector.GetValue("INIT", "Record"))

    ' HappyHour
    Dim lDayNumberTemp As Long
    Dim sDayName As String
    
    iniHappyHourActivado = CBool(val(Lector.GetValue("HAPPYHOUR", "Activado")))
    For lDayNumberTemp = 1 To 7
        sDayName = Lector.GetValue("HAPPYHOUR", "Dia" & lDayNumberTemp)
        HappyHourDays(lDayNumberTemp).Hour = val(ReadField(1, sDayName, 45)) ' GSZAO
        HappyHourDays(lDayNumberTemp).Multi = val(ReadField(2, sDayName, 45)) ' 0.13.5
        
        If HappyHourDays(lDayNumberTemp).Hour < 0 Or HappyHourDays(lDayNumberTemp).Hour > 23 Then
            HappyHourDays(lDayNumberTemp).Hour = 20 ' Hora de 0 a 23.
        End If
        
        If HappyHourDays(lDayNumberTemp).Multi < 0 Then
            HappyHourDays(lDayNumberTemp).Multi = 0
        End If
    Next

    'Conexion con la API hecha en Node.js
    'Mas info aqui: https://github.com/ao-libre/ao-api-server/
    ConexionAPI = CBool(Lector.GetValue("CONEXIONAPI", "Activado"))

    'CHOTS | Database
    Database_Enabled = CBool(val(Lector.GetValue("DATABASE", "Enabled")))
    Database_DataSource = Lector.GetValue("DATABASE", "DSN")
    Database_Host = Lector.GetValue("DATABASE", "Host")
    Database_Name = Lector.GetValue("DATABASE", "Name")
    Database_Username = Lector.GetValue("DATABASE", "Username")
    Database_Password = Lector.GetValue("DATABASE", "Password")
      
    'Max users
    Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User

    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agrego en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(Lector.GetValue("BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    Call Statistics.Initialize

    'En caso que usemos mundo propio, cargamos el mapa y la coordeanas donde se hara el spawn inicial'
    If UsarMundoPropio Then
        CustomSpawnMap.Map = Lector.GetValue("MUNDO", "Mapa")
        CustomSpawnMap.X = Lector.GetValue("MUNDO", "X")
        CustomSpawnMap.Y = Lector.GetValue("MUNDO", "Y")
    End If
    
    Set Lector = Nothing
    
    Set ConsultaPopular = New ConsultasPopulares
    Call ConsultaPopular.LoadData
    
    ' Admins
    Call loadAdministrativeUsers

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo la info de inicio del server (Sinfo.ini)"
    
End Sub

Sub CargarCiudades()
    
    '***************************************************
    'Author: Jopi
    'Last Modification: 15/05/2019 (Jopi)
    'Jopi: Uso de clsIniManager para cargar los valores.
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Ciudades.dat"
    
    Dim Lector As clsIniManager: Set Lector = New clsIniManager
    
    Call Lector.Initialize(DatPath & "Ciudades.dat")
        
        With Ullathorpe
            .Map = Lector.GetValue("Ullathorpe", "Mapa")
            .X = Lector.GetValue("Ullathorpe", "X")
            .Y = Lector.GetValue("Ullathorpe", "Y")
        End With
        
        With Nix
            .Map = Lector.GetValue("Nix", "Mapa")
            .X = Lector.GetValue("Nix", "X")
            .Y = Lector.GetValue("Nix", "Y")
        End With
        
        With Banderbill
            .Map = Lector.GetValue("Banderbill", "Mapa")
            .X = Lector.GetValue("Banderbill", "X")
            .Y = Lector.GetValue("Banderbill", "Y")
        End With

        
        With Lindos
            .Map = Lector.GetValue("Lindos", "Mapa")
            .X = Lector.GetValue("Lindos", "X")
            .Y = Lector.GetValue("Lindos", "Y")
        End With
        
        With Arghal
            .Map = Lector.GetValue("Arghal", "Mapa")
            .X = Lector.GetValue("Arghal", "X")
            .Y = Lector.GetValue("Arghal", "Y")
        End With
        
        With Arkhein
            .Map = Lector.GetValue("Arkhein", "Mapa")
            .X = Lector.GetValue("Arkhein", "X")
            .Y = Lector.GetValue("Arkhein", "Y")
        End With
        
        With Nemahuak
            .Map = Lector.GetValue("Nemahuak", "Mapa")
            .X = Lector.GetValue("Nemahuak", "X")
            .Y = Lector.GetValue("Nemahuak", "Y")
        End With

    Set Lector = Nothing
    
    Ciudades(eCiudad.cUllathorpe) = Ullathorpe
    Ciudades(eCiudad.cNix) = Nix
    Ciudades(eCiudad.cBanderbill) = Banderbill
    Ciudades(eCiudad.cLindos) = Lindos
    Ciudades(eCiudad.cArghal) = Arghal
    Ciudades(eCiudad.cArkhein) = Arkhein

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargaron las ciudades.dat"

End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Escribe VAR en un archivo
    '***************************************************

    writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUserToCharfile(ByVal Userindex As Integer, Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Unknown
    'Last modified: 10/10/2010 (Pato)
    'Saves the Users RECORDs
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '11/19/2009: Pato - Save the EluSkills and ExpSkills
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '10/10/2010: Pato - Saco el WriteVar e implemento la clase clsIniManager
    '18/09/2018: CHOTS - Nuevo nombre de la funcion, solo realiza el grabado
    '19/11/2019: Recox - Cambie el casteo de muchas propiedades, para evitar y arreglar errores
    '*************************************************

    On Error GoTo ErrorHandler

    Dim Manager  As clsIniManager

    Dim Existe   As Boolean

    Dim UserFile As String

    With UserList(Userindex)

        UserFile = CharPath & UCase$(.Name) & ".chr"
    
        Set Manager = New clsIniManager
    
        If FileExist(UserFile) Then
            Call Manager.Initialize(UserFile)
        
            If FileExist(UserFile & ".bk") Then Call Kill(UserFile & ".bk")
            Name UserFile As UserFile & ".bk"
        
            Existe = True

        End If
    
        Dim LoopC As Long
    
        Call Manager.ChangeValue("FLAGS", "Muerto", CByte(.flags.Muerto))
        Call Manager.ChangeValue("FLAGS", "Escondido", CByte(.flags.Escondido))
        Call Manager.ChangeValue("FLAGS", "Hambre", CByte(.flags.Hambre))
        Call Manager.ChangeValue("FLAGS", "Sed", CByte(.flags.Sed))
        Call Manager.ChangeValue("FLAGS", "Desnudo", CByte(.flags.Desnudo))
        Call Manager.ChangeValue("FLAGS", "Ban", CByte(.flags.Ban))
        Call Manager.ChangeValue("FLAGS", "Navegando", CByte(.flags.Navegando))
        Call Manager.ChangeValue("FLAGS", "Envenenado", CByte(.flags.Envenenado))
        Call Manager.ChangeValue("FLAGS", "Paralizado", CByte(.flags.Paralizado))
        'Matrix
        Call Manager.ChangeValue("FLAGS", "LastMap", CInt(.flags.lastMap))
    
        Call Manager.ChangeValue("CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
        Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
    
        Call Manager.ChangeValue("COUNTERS", "Pena", CLng(.Counters.Pena))
        Call Manager.ChangeValue("COUNTERS", "SkillsAsignados", CByte(.Counters.AsignedSkills))
    
        Call Manager.ChangeValue("FACCIONES", "EjercitoReal", CByte(.Faccion.ArmadaReal))
        Call Manager.ChangeValue("FACCIONES", "EjercitoCaos", CByte(.Faccion.FuerzasCaos))
        Call Manager.ChangeValue("FACCIONES", "CiudMatados", CLng(.Faccion.CiudadanosMatados))
        Call Manager.ChangeValue("FACCIONES", "CrimMatados", CLng(.Faccion.CriminalesMatados))
        Call Manager.ChangeValue("FACCIONES", "rArCaos", CByte(.Faccion.RecibioArmaduraCaos))
        Call Manager.ChangeValue("FACCIONES", "rArReal", CByte(.Faccion.RecibioArmaduraReal))
        Call Manager.ChangeValue("FACCIONES", "rExCaos", CByte(.Faccion.RecibioExpInicialCaos))
        Call Manager.ChangeValue("FACCIONES", "rExReal", CByte(.Faccion.RecibioExpInicialReal))
        Call Manager.ChangeValue("FACCIONES", "recCaos", CLng(.Faccion.RecompensasCaos))
        Call Manager.ChangeValue("FACCIONES", "recReal", CLng(.Faccion.RecompensasReal))
        Call Manager.ChangeValue("FACCIONES", "Reenlistadas", CByte(.Faccion.Reenlistadas))
        Call Manager.ChangeValue("FACCIONES", "NivelIngreso", CInt(.Faccion.NivelIngreso))
        Call Manager.ChangeValue("FACCIONES", "FechaIngreso", CStr(.Faccion.FechaIngreso))
        Call Manager.ChangeValue("FACCIONES", "MatadosIngreso", CInt(.Faccion.MatadosIngreso))
        Call Manager.ChangeValue("FACCIONES", "NextRecompensa", CInt(.Faccion.NextRecompensa))
    
        'Fueron modificados los atributos del usuario?
        If Not .flags.TomoPocion Then

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributos(LoopC)))
            Next LoopC

        Else

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributosBackUP(LoopC)))
            Next LoopC

        End If
    
        For LoopC = 1 To UBound(.Stats.UserSkills)
            Call Manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(.Stats.UserSkills(LoopC)))
            Call Manager.ChangeValue("SKILLS", "ELUSK" & LoopC, CStr(.Stats.EluSkills(LoopC)))
            Call Manager.ChangeValue("SKILLS", "EXPSK" & LoopC, CStr(.Stats.ExpSkills(LoopC)))
        Next LoopC
    
        Call Manager.ChangeValue("CONTACTO", "Email", CStr(.Email))
    
        Call Manager.ChangeValue("INIT", "AccountHash", CStr(.AccountHash))
        Call Manager.ChangeValue("INIT", "Genero", CByte(.Genero))
        Call Manager.ChangeValue("INIT", "Raza", CByte(.raza))
        Call Manager.ChangeValue("INIT", "Hogar", CByte(.Hogar))
        Call Manager.ChangeValue("INIT", "Clase", CByte(.Clase))
        Call Manager.ChangeValue("INIT", "Desc", CStr(.Desc))
    
        Call Manager.ChangeValue("INIT", "Heading", CByte(.Char.heading))
        Call Manager.ChangeValue("INIT", "Head", CInt(.OrigChar.Head))
    
        If .flags.Muerto = 0 Then
            If .Char.body <> 0 Then
                Call Manager.ChangeValue("INIT", "Body", CInt(.Char.body))

            End If

        End If
    
        Call Manager.ChangeValue("INIT", "Arma", CInt(.Char.WeaponAnim))
        Call Manager.ChangeValue("INIT", "Escudo", CInt(.Char.ShieldAnim))
        Call Manager.ChangeValue("INIT", "Casco", CInt(.Char.CascoAnim))
    
        #If ConUpTime Then
    
            If SaveTimeOnline Then

                Dim TempDate As Date

                TempDate = Now - .LogOnTime
                .LogOnTime = Now
                .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
                Call Manager.ChangeValue("INIT", "UpTime", CLng(.UpTime))

            End If

        #End If
    
        'First time around?
        If Manager.GetValue("INIT", "LastIP1") = vbNullString Then
            Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
            'Is it a different ip from last time?
        ElseIf .ip <> Left$(Manager.GetValue("INIT", "LastIP1"), InStr(1, Manager.GetValue("INIT", "LastIP1"), " ") - 1) Then

            Dim i As Integer

            For i = 5 To 2 Step -1
                Call Manager.ChangeValue("INIT", "LastIP" & i, Manager.GetValue("INIT", "LastIP" & CStr(i - 1)))
            Next i

            Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
            'Same ip, just update the date
        Else
            Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)

        End If
    
        Call Manager.ChangeValue("INIT", "Position", .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
    
        Call Manager.ChangeValue("STATS", "GLD", CLng(.Stats.Gld))
        Call Manager.ChangeValue("STATS", "BANCO", CLng(.Stats.Banco))
    
        Call Manager.ChangeValue("STATS", "MaxHP", CInt(.Stats.MaxHp))
        Call Manager.ChangeValue("STATS", "MinHP", CInt(.Stats.MinHp))
    
        Call Manager.ChangeValue("STATS", "MaxSTA", CInt(.Stats.MaxSta))
        Call Manager.ChangeValue("STATS", "MinSTA", CInt(.Stats.MinSta))
    
        Call Manager.ChangeValue("STATS", "MaxMAN", CInt(.Stats.MaxMAN))
        Call Manager.ChangeValue("STATS", "MinMAN", CInt(.Stats.MinMAN))
    
        Call Manager.ChangeValue("STATS", "MaxHIT", CInt(.Stats.MaxHIT))
        Call Manager.ChangeValue("STATS", "MinHIT", CInt(.Stats.MinHIT))
    
        Call Manager.ChangeValue("STATS", "MaxAGU", CByte(.Stats.MaxAGU))
        Call Manager.ChangeValue("STATS", "MinAGU", CByte(.Stats.MinAGU))
    
        Call Manager.ChangeValue("STATS", "MaxHAM", CByte(.Stats.MaxHam))
        Call Manager.ChangeValue("STATS", "MinHAM", CByte(.Stats.MinHam))
    
        Call Manager.ChangeValue("STATS", "SkillPtsLibres", CInt(.Stats.SkillPts))
    
        Call Manager.ChangeValue("STATS", "EXP", CDbl(.Stats.Exp))
        Call Manager.ChangeValue("STATS", "ELV", CByte(.Stats.ELV))
      
        Call Manager.ChangeValue("STATS", "ELU", CLng(.Stats.ELU))
    
        Call Manager.ChangeValue("MUERTES", "UserMuertes", CLng(.Stats.UsuariosMatados))
        Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CInt(.Stats.NPCsMuertos))
      
        '[KEVIN]----------------------------------------------------------------------------
        '*******************************************************************************************
        Call Manager.ChangeValue("BancoInventory", "CantidadItems", CInt(.BancoInvent.NroItems))

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            Call Manager.ChangeValue("BancoInventory", "Obj" & LoopC, .BancoInvent.Object(LoopC).ObjIndex & "-" & .BancoInvent.Object(LoopC).Amount)
        Next LoopC

        '*******************************************************************************************
        '[/KEVIN]-----------
      
        'Save Inv
        Call Manager.ChangeValue("Inventory", "CantidadItems", CInt(.Invent.NroItems))
    
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            Call Manager.ChangeValue("Inventory", "Obj" & LoopC, .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).Equipped)
        Next LoopC
    
        Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", CByte(.Invent.WeaponEqpSlot))
        Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", CByte(.Invent.ArmourEqpSlot))
        Call Manager.ChangeValue("Inventory", "CascoEqpSlot", CByte(.Invent.CascoEqpSlot))
        Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", CByte(.Invent.EscudoEqpSlot))
        Call Manager.ChangeValue("Inventory", "BarcoSlot", CByte(.Invent.BarcoSlot))
        Call Manager.ChangeValue("Inventory", "MunicionSlot", CByte(.Invent.MunicionEqpSlot))
        Call Manager.ChangeValue("Inventory", "AnilloSlot", CByte(.Invent.AnilloEqpSlot))
        Call Manager.ChangeValue("Inventory", "MochilaSlot", CByte(.Invent.MochilaEqpSlot))
    
        'Reputacion
        Call Manager.ChangeValue("REP", "Asesino", CLng(.Reputacion.AsesinoRep))
        Call Manager.ChangeValue("REP", "Bandido", CLng(.Reputacion.BandidoRep))
        Call Manager.ChangeValue("REP", "Burguesia", CLng(.Reputacion.BurguesRep))
        Call Manager.ChangeValue("REP", "Ladrones", CLng(.Reputacion.LadronesRep))
        Call Manager.ChangeValue("REP", "Nobles", CLng(.Reputacion.NobleRep))
        Call Manager.ChangeValue("REP", "Plebe", CLng(.Reputacion.PlebeRep))
        Call Manager.ChangeValue("REP", "Promedio", CLng(.Reputacion.Promedio))
    
        Dim cad As String
    
        For LoopC = 1 To MAXUSERHECHIZOS
            cad = .Stats.UserHechizos(LoopC)
            Call Manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
        Next
    
        Dim NroMascotas As Long

        NroMascotas = .NroMascotas
    
        For LoopC = 1 To MAXMASCOTAS

            ' Mascota valida?
            If .MascotasIndex(LoopC) > 0 Then

                ' Nos aseguramos que la criatura no fue invocada
                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    cad = .MascotasType(LoopC)
                Else 'Si fue invocada no la guardamos
                    cad = "0"
                    NroMascotas = NroMascotas - 1

                End If

                Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)
            Else
                cad = .MascotasType(LoopC)
                Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)

            End If
    
        Next
    
        Call Manager.ChangeValue("MASCOTAS", "NroMascotas", CInt(NroMascotas))
    
        'Devuelve el head de muerto
        If .flags.Muerto = 1 Then
            .Char.Head = iCabezaMuerto

        End If

    End With

    Call Manager.DumpFile(UserFile)

    Set Manager = Nothing

    If Existe Then Call Kill(UserFile & ".bk")

    Exit Sub

ErrorHandler:
    Call LogError("Error en SaveUserToCharfile: " & UserFile & " -- " & Err.Number & ": " & Err.description)

    Set Manager = Nothing

End Sub

Function criminal(ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim L As Long
    
    With UserList(Userindex).Reputacion
        L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
        L = L / 6
        criminal = (L < 0)

    End With

End Function

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/09/2010
    '10/09/2010 - Pato: Optimice el BackUp de NPCs
    '***************************************************

    Dim LoopC As Integer
    
    Print #hFile, "[NPC" & Npclist(NpcIndex).Numero & "]"
    
    With Npclist(NpcIndex)
        'General
        Print #hFile, "Name=" & .Name
        Print #hFile, "Desc=" & .Desc
        Print #hFile, "Head=" & val(.Char.Head)
        Print #hFile, "Body=" & val(.Char.body)
        Print #hFile, "Heading=" & val(.Char.heading)
        Print #hFile, "Movement=" & val(.Movement)
        Print #hFile, "Attackable=" & val(.Attackable)
        Print #hFile, "Comercia=" & val(.Comercia)
        Print #hFile, "TipoItems=" & val(.TipoItems)
        Print #hFile, "Hostil=" & val(.Hostile)
        Print #hFile, "GiveEXP=" & val(.GiveEXP)
        Print #hFile, "GiveGLD=" & val(.GiveGLD)
        Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
        Print #hFile, "NpcType=" & val(.NPCtype)
        
        'Stats
        Print #hFile, "Alineacion=" & val(.Stats.Alineacion)
        Print #hFile, "DEF=" & val(.Stats.def)
        Print #hFile, "MaxHit=" & val(.Stats.MaxHIT)
        Print #hFile, "MaxHp=" & val(.Stats.MaxHp)
        Print #hFile, "MinHit=" & val(.Stats.MinHIT)
        Print #hFile, "MinHp=" & val(.Stats.MinHp)
        
        'Flags
        Print #hFile, "ReSpawn=" & val(.flags.Respawn)
        Print #hFile, "BackUp=" & val(.flags.BackUp)
        Print #hFile, "Domable=" & val(.flags.Domable)
        
        'Inventario
        Print #hFile, "NroItems=" & val(.Invent.NroItems)

        If .Invent.NroItems > 0 Then

            For LoopC = 1 To .Invent.NroItems
                Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount
            Next LoopC

        End If
        
        Print #hFile, ""

    End With

End Sub

Sub CargarNpcBackUp(ByVal NpcIndex As Integer, ByVal NpcNumber As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Status
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando backup Npc"
    
    Dim npcfile As String
    
    'If NpcNumber > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
    npcfile = DatPath & "bkNPCs.dat"
    'End If
    
    With Npclist(NpcIndex)
    
        .Numero = NpcNumber
        .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
        .Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
        .Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
        .NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
        
        .Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
        .Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
        .Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
        
        .Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
        .Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
        .Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
        .GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
        
        .GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
        
        .InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
        
        .Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
        .Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
        .Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
        .Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
        .Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
        .Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
        
        Dim LoopC As Integer

        Dim ln    As String

        .Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

        If .Invent.NroItems > 0 Then

            For LoopC = 1 To MAX_INVENTORY_SLOTS
                ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
                .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
               
            Next LoopC

        Else

            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = 0
                .Invent.Object(LoopC).Amount = 0
            Next LoopC

        End If
        
        For LoopC = 1 To MAX_NPC_DROPS
            ln = GetVar(npcfile, "NPC" & NpcNumber, "Drop" & LoopC)
            .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
        Next LoopC
        
        .flags.NPCActive = True
        .flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
        .flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
        .flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
        .flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
        
        'Tipo de items con los que comercia
        .TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

    End With

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo bkNPCs.dat"

End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer

    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando apuestas.dat"

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo apuestas.dat"

End Sub

Public Sub generateMatrix(ByVal Mapa As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    Dim j As Integer
    
    ReDim distanceToCities(1 To NumMaps) As HomeDistance
    
    For j = 1 To NUMCIUDADES
        For i = 1 To NumMaps
            distanceToCities(i).distanceToCity(j) = -1
        Next i
    Next j
    
    For j = 1 To NUMCIUDADES
        For i = 1 To 4

            Select Case i

                Case eHeading.NORTH
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.NORTH), j, i, 0, 1)

                Case eHeading.EAST
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.EAST), j, i, 1, 0)

                Case eHeading.SOUTH
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.SOUTH), j, i, 0, 1)

                Case eHeading.WEST
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.WEST), j, i, -1, 0)

            End Select

        Next i
    Next j

End Sub

Public Sub setDistance(ByVal Mapa As Integer, _
                       ByVal city As Byte, _
                       ByVal side As Integer, _
                       Optional ByVal X As Integer = 0, _
                       Optional ByVal Y As Integer = 0)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i   As Integer

    Dim lim As Integer

    If Mapa <= 0 Or Mapa > NumMaps Then Exit Sub

    If distanceToCities(Mapa).distanceToCity(city) >= 0 Then Exit Sub

    If Mapa = Ciudades(city).Map Then
        distanceToCities(Mapa).distanceToCity(city) = 0
    Else
        distanceToCities(Mapa).distanceToCity(city) = Abs(X) + Abs(Y)

    End If

    For i = 1 To 4
        lim = getLimit(Mapa, i)

        If lim > 0 Then

            Select Case i

                Case eHeading.NORTH
                    Call setDistance(lim, city, i, X, Y + 1)

                Case eHeading.EAST
                    Call setDistance(lim, city, i, X + 1, Y)

                Case eHeading.SOUTH
                    Call setDistance(lim, city, i, X, Y - 1)

                Case eHeading.WEST
                    Call setDistance(lim, city, i, X - 1, Y)

            End Select

        End If

    Next i

End Sub

Public Function getLimit(ByVal Mapa As Integer, ByVal side As Byte) As Integer

    '***************************************************
    'Author: Budi
    'Last Modification: 31/01/2010
    'Retrieves the limit in the given side in the given map.
    'TODO: This should be set in the .inf map file.
    '***************************************************
    Dim X As Long

    Dim Y As Long

    If Mapa <= 0 Then Exit Function

    For X = 15 To 87
        For Y = 0 To 3

            Select Case side

                Case eHeading.NORTH
                    getLimit = MapData(Mapa, X, 7 + Y).TileExit.Map

                Case eHeading.EAST
                    getLimit = MapData(Mapa, 92 - Y, X).TileExit.Map

                Case eHeading.SOUTH
                    getLimit = MapData(Mapa, X, 94 - Y).TileExit.Map

                Case eHeading.WEST
                    getLimit = MapData(Mapa, 9 + Y, X).TileExit.Map

            End Select

            If getLimit > 0 Then Exit Function
        Next Y
    Next X

End Function

Public Sub LoadArmadurasFaccion()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando armaduras faccionarias"
    
    Dim ClassIndex    As Long
    
    Dim ArmaduraIndex As Integer
    
    For ClassIndex = 1 To NUMCLASES
    
        ' Defensa minima para armadas altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para armadas bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
    
        ' Defensa media para armadas altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para armadas bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
    
        ' Defensa alta para armadas altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para armadas bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos altos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos bajos
        ArmaduraIndex = val(GetVar(DatPath & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
    
    Next ClassIndex

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo ArmadurasFaccionarias.dat"

End Sub

Sub SendUserBovedaTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    'CHOTS: Lo movi a esta funcion porque tiene mas sentido
    '***************************************************

    On Error Resume Next

    Dim j        As Integer
    Dim CharFile As String, Tmp As String
    Dim ObjInd   As Long, ObjCant As Long

    CharFile = CharPath & charName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

        For j = 1 To MAX_BANCOINVENTORY_SLOTS
            Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserMiniStatsTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)

    '*************************************************
    'Author: Unknown
    'Last modified: 19/19/2018
    'Shows the users Stats when the user is offline.
    '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribucion de parametros.
    '19/09/2018 CHOTS - Movido a FileIO
    '*************************************************
    Dim CharFile      As String

    Dim Ban           As String

    Dim BanDetailPath As String
    
    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " usuarios matados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejercito real desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingreso en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legion oscura desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingreso en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejercito real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legion oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)

        End If
        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)

        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)

        End If

    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserInvTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    '19/09/2018 CHOTS - Movido a FileIO
    '***************************************************

    On Error Resume Next

    Dim j        As Long

    Dim CharFile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserOROTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    '19/09/2018 CHOTS - Movido a FileIO
    '***************************************************

    Dim CharFile As String
    
    On Error Resume Next

    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserStatsTxtCharfile(ByVal sendIndex As Integer, ByVal Nombre As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    '19/09/2018 CHOTS - Movido a FileIO
    '***************************************************

    If PersonajeExiste(Nombre) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energia: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
        #If ConUpTime Then

            Dim TempSecs As Long

            Dim TempStr  As String

            TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
            TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
            Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If
    
        Call WriteConsoleMsg(sendIndex, "Dados: " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT1") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT2") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT3") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT4") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT5"), FontTypeNames.FONTTYPE_INFO)

    End If

End Sub
