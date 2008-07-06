Attribute VB_Name = "ES"
'Argentum Online 0.11.6
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

Public Sub CargarSpawnList()
    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdmin(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Admines"))

For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Admines", "Admin" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsAdmin = True
        Exit Function
    End If
Next WizNum
EsAdmin = False

End Function

Function EsDios(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False

End Function

Function EsConsejero(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Function EsRolesMaster(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal name As String) As Long
Dim N As Integer, cad As String, Tam As Long
N = FreeFile(1)
Open name For Input As #N
Tam = 0
Do While Not EOF(N)
    Tam = Tam + 1
    Line Input #N, cad
Loop
Close N
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()

ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim N As Integer, i As Integer
N = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #N

For i = 1 To UBound(ForbidenNames)
    Line Input #N, ForbidenNames(i)
Next i

Close N

End Sub

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))

ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
'    Hechizos(Hechizo).Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    
    Hechizos(Hechizo).CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
    
'    Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
'    Hechizos(Hechizo).ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
    frmCargando.cargar.value = frmCargando.cargar.value + 1
    
    Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
Next Hechizo

Set Leer = Nothing
Exit Sub

Errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
Dim i As Integer

MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))

ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    MOTD(i).Formato = vbNullString
Next i

End Sub

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True
Dim i As Integer



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


Call LimpiarMundo
Call WorldSave
Call modGuilds.v_RutinaElecciones
Call ResetCentinelaInfo     'Reseteamos al centinela


Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & time
Close #nfile
End Sub

Public Sub GrabarMapa(ByVal map As Long, ByVal MAPFILE As String)
On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim TempInt As Integer
    Dim LoopC As Long
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
            
    Put FreeFileMap, , MapInfo(map).MapVersion
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(map, X, Y).Blocked Then ByFlags = ByFlags Or 1
                If MapData(map, X, Y).Graphic(2) Then ByFlags = ByFlags Or 2
                If MapData(map, X, Y).Graphic(3) Then ByFlags = ByFlags Or 4
                If MapData(map, X, Y).Graphic(4) Then ByFlags = ByFlags Or 8
                If MapData(map, X, Y).trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , MapData(map, X, Y).Graphic(1)
                
                For LoopC = 2 To 4
                    If MapData(map, X, Y).Graphic(LoopC) Then _
                        Put FreeFileMap, , MapData(map, X, Y).Graphic(LoopC)
                Next LoopC
                
                If MapData(map, X, Y).trigger Then _
                    Put FreeFileMap, , CInt(MapData(map, X, Y).trigger)
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
                   If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        MapData(map, X, Y).ObjInfo.ObjIndex = 0
                        MapData(map, X, Y).ObjInfo.amount = 0
                    End If
                End If
    
                If MapData(map, X, Y).TileExit.map Then ByFlags = ByFlags Or 1
                If MapData(map, X, Y).NpcIndex Then ByFlags = ByFlags Or 2
                If MapData(map, X, Y).ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(map, X, Y).TileExit.map Then
                    Put FreeFileInf, , MapData(map, X, Y).TileExit.map
                    Put FreeFileInf, , MapData(map, X, Y).TileExit.X
                    Put FreeFileInf, , MapData(map, X, Y).TileExit.Y
                End If
                
                If MapData(map, X, Y).NpcIndex Then _
                    Put FreeFileInf, , Npclist(MapData(map, X, Y).NpcIndex).Numero
                
                If MapData(map, X, Y).ObjInfo.ObjIndex Then
                    Put FreeFileInf, , MapData(map, X, Y).ObjInfo.ObjIndex
                    Put FreeFileInf, , MapData(map, X, Y).ObjInfo.amount
                End If
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf

    'write .dat file
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Name", MapInfo(map).name)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "MusicNum", MapInfo(map).Music)
    Call WriteVar(MAPFILE & ".dat", "mapa" & map, "MagiaSinefecto", MapInfo(map).MagiaSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "mapa" & map, "InviSinEfecto", MapInfo(map).InviSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "mapa" & map, "ResuSinEfecto", MapInfo(map).ResuSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "StartPos", MapInfo(map).StartPos.map & "-" & MapInfo(map).StartPos.X & "-" & MapInfo(map).StartPos.Y)
    

    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Terreno", MapInfo(map).Terreno)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Zona", MapInfo(map).Zona)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Restringir", MapInfo(map).Restringir)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "BackUp", str(MapInfo(map).BackUp))

    If MapInfo(map).Pk Then
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Pk", "0")
    Else
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Pk", "1")
    End If

End Sub
Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadBalance()
    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        ModClase(i).Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
        ModClase(i).AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
        ModClase(i).AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
        ModClase(i).DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
        ModClase(i).DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
        ModClase(i).DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
        ModClase(i).Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        ModRaza(i).Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
        ModRaza(i).Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
        ModRaza(i).Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
        ModRaza(i).Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
        ModRaza(i).Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
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
End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To N) As Integer

For lc = 1 To N
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub



Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).name = Leer.GetValue("OBJ" & Object, "Name")
    
    'Pablo (ToxicWaste) Log de Objetos.
    ObjData(Object).Log = val(Leer.GetValue("OBJ" & Object, "Log"))
    ObjData(Object).NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
    '07/09/07
    
    ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
    
    Select Case ObjData(Object).OBJType
        Case eOBJType.otArmadura
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
            'Pablo (ToxicWaste)
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
        Case eOBJType.otAnillo 'Pablo (ToxicWaste)
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
            
    End Select
    
    ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) / 2
    
    ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    ObjData(Object).RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
    ObjData(Object).RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
    ObjData(Object).RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    ObjData(Object).RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    
    ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
    
    'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
    Dim i As Integer
    Dim N As Integer
    Dim S As String
    For i = 1 To NUMCLASES
        S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
        N = 1
        Do While LenB(S) > 0 And UCase$(ListaClases(N)) <> S
            N = N + 1
        Loop
        ObjData(Object).ClaseProhibida(i) = IIf(LenB(S) > 0, N, 0)
    Next i
    
    ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
    frmCargando.cargar.value = frmCargando.cargar.value + 1
Next Object

Set Leer = Nothing

Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)

Dim LoopC As Long

For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next LoopC

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
Next LoopC

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
Next LoopC

UserList(UserIndex).Stats.GLD = CLng(UserFile.GetValue("STATS", "GLD"))
UserList(UserIndex).Stats.Banco = CLng(UserFile.GetValue("STATS", "BANCO"))

UserList(UserIndex).Stats.MaxHP = CInt(UserFile.GetValue("STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = CInt(UserFile.GetValue("STATS", "MinHP"))

UserList(UserIndex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))

UserList(UserIndex).Stats.MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))

UserList(UserIndex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
UserList(UserIndex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
UserList(UserIndex).Stats.ELV = CByte(UserFile.GetValue("STATS", "ELV"))


UserList(UserIndex).Stats.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil

If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)

UserList(UserIndex).Reputacion.AsesinoRep = val(UserFile.GetValue("REP", "Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = val(UserFile.GetValue("REP", "Bandido"))
UserList(UserIndex).Reputacion.BurguesRep = val(UserFile.GetValue("REP", "Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = val(UserFile.GetValue("REP", "Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = val(UserFile.GetValue("REP", "Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = val(UserFile.GetValue("REP", "Plebe"))
UserList(UserIndex).Reputacion.Promedio = val(UserFile.GetValue("REP", "Promedio"))

End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'*************************************************
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'*************************************************
Dim LoopC As Long
Dim ln As String

UserList(UserIndex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
UserList(UserIndex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
UserList(UserIndex).Faccion.NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
UserList(UserIndex).Faccion.FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
UserList(UserIndex).Faccion.MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
UserList(UserIndex).Faccion.NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

UserList(UserIndex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
UserList(UserIndex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))

UserList(UserIndex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
UserList(UserIndex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
UserList(UserIndex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
UserList(UserIndex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
UserList(UserIndex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
UserList(UserIndex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If


UserList(UserIndex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))

UserList(UserIndex).email = UserFile.GetValue("CONTACTO", "Email")

UserList(UserIndex).genero = UserFile.GetValue("INIT", "Genero")
UserList(UserIndex).clase = UserFile.GetValue("INIT", "Clase")
UserList(UserIndex).raza = UserFile.GetValue("INIT", "Raza")
UserList(UserIndex).Hogar = UserFile.GetValue("INIT", "Hogar")
UserList(UserIndex).Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))


UserList(UserIndex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
UserList(UserIndex).OrigChar.body = CInt(UserFile.GetValue("INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))

#If ConUpTime Then
    UserList(UserIndex).UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
#End If

UserList(UserIndex).OrigChar.heading = eHeading.SOUTH

If UserList(UserIndex).flags.Muerto = 0 Then
    UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else
    UserList(UserIndex).Char.body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco
End If


UserList(UserIndex).desc = UserFile.GetValue("INIT", "Desc")

UserList(UserIndex).Pos.map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

UserList(UserIndex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
    UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).BancoInvent.Object(LoopC).amount = CInt(ReadField(2, ln, 45))
Next LoopC
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).amount = CInt(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto anilo
UserList(UserIndex).Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
    UserList(UserIndex).Invent.AnilloEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AnilloEqpSlot).ObjIndex
End If

UserList(UserIndex).NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))
Dim NpcIndex As Integer
For LoopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
Next LoopC

ln = UserFile.GetValue("Guild", "GUILDINDEX")
If IsNumeric(ln) Then
    UserList(UserIndex).guildIndex = CInt(ln)
Else
    UserList(UserIndex).guildIndex = 0
End If

End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = vbNullString
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
    
    For map = 1 To NumMaps
        If val(GetVar(App.Path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & map
        Else
            tFileName = App.Path & MapPath & "Mapa" & map
        End If
        
        Call CargarMapa(map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & map
        Call CargarMapa(map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal map As Long, ByVal MAPFl As String)
On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim TempInt As Integer
      
    FreeFileMap = FreeFile
    
    Open MAPFl & ".map" For Binary As #FreeFileMap
    Seek FreeFileMap, 1
    
    FreeFileInf = FreeFile
    
    'inf
    Open MAPFl & ".inf" For Binary As #FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    Get #FreeFileMap, , MapInfo(map).MapVersion
    Get #FreeFileMap, , MiCabecera
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    
    'inf Header
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.dat file
            Get FreeFileMap, , ByFlags

            If ByFlags And 1 Then
                MapData(map, X, Y).Blocked = 1
            End If
            
            Get FreeFileMap, , MapData(map, X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get FreeFileMap, , TempInt
                MapData(map, X, Y).trigger = TempInt
            End If
            
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(map, X, Y).TileExit.map
                Get FreeFileInf, , MapData(map, X, Y).TileExit.X
                Get FreeFileInf, , MapData(map, X, Y).TileExit.Y
            End If
            
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(map, X, Y).NpcIndex
                
                If MapData(map, X, Y).NpcIndex > 0 Then
                    'If MapData(Map, X, Y).NpcIndex > 499 Then
                    '    npcfile = DatPath & "NPCs-HOSTILES.dat"
                    'Else
                        npcfile = DatPath & "NPCs.dat"
                    'End If

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(map, X, Y).NpcIndex = OpenNPC(MapData(map, X, Y).NpcIndex)
                        Npclist(MapData(map, X, Y).NpcIndex).Orig.map = map
                        Npclist(MapData(map, X, Y).NpcIndex).Orig.X = X
                        Npclist(MapData(map, X, Y).NpcIndex).Orig.Y = Y
                    Else
                        MapData(map, X, Y).NpcIndex = OpenNPC(MapData(map, X, Y).NpcIndex)
                    End If
                    
                    Npclist(MapData(map, X, Y).NpcIndex).Pos.map = map
                    Npclist(MapData(map, X, Y).NpcIndex).Pos.X = X
                    Npclist(MapData(map, X, Y).NpcIndex).Pos.Y = Y
                    
                    Call MakeNPCChar(True, 0, MapData(map, X, Y).NpcIndex, map, X, Y)
                End If
            End If
            
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(map, X, Y).ObjInfo.ObjIndex
                Get FreeFileInf, , MapData(map, X, Y).ObjInfo.amount
            End If
        Next X
    Next Y
    
    
    Close FreeFileMap
    Close FreeFileInf
    
    MapInfo(map).name = GetVar(MAPFl & ".dat", "Mapa" & map, "Name")
    MapInfo(map).Music = GetVar(MAPFl & ".dat", "Mapa" & map, "MusicNum")
    MapInfo(map).StartPos.map = val(ReadField(1, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
    MapInfo(map).StartPos.X = val(ReadField(2, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
    MapInfo(map).StartPos.Y = val(ReadField(3, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
    MapInfo(map).MagiaSinEfecto = val(GetVar(MAPFl & ".dat", "Mapa" & map, "MagiaSinEfecto"))
    MapInfo(map).InviSinEfecto = val(GetVar(MAPFl & ".dat", "Mapa" & map, "InviSinEfecto"))
    MapInfo(map).ResuSinEfecto = val(GetVar(MAPFl & ".dat", "Mapa" & map, "ResuSinEfecto"))
    MapInfo(map).NoEncriptarMP = val(GetVar(MAPFl & ".dat", "Mapa" & map, "NoEncriptarMP"))
    
    If val(GetVar(MAPFl & ".dat", "Mapa" & map, "Pk")) = 0 Then
        MapInfo(map).Pk = True
    Else
        MapInfo(map).Pk = False
    End If
    
    
    MapInfo(map).Terreno = GetVar(MAPFl & ".dat", "Mapa" & map, "Terreno")
    MapInfo(map).Zona = GetVar(MAPFl & ".dat", "Mapa" & map, "Zona")
    MapInfo(map).Restringir = GetVar(MAPFl & ".dat", "Mapa" & map, "Restringir")
    MapInfo(map).BackUp = val(GetVar(MAPFl & ".dat", "Mapa" & map, "BACKUP"))
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " - Pos: " & X & "," & Y & "." & Err.description)
End Sub

Sub LoadSini()

Dim Temporal As Long

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

'Misc
#If SeguridadAlkon Then

Call Security.SetServerIp(GetVar(IniPath & "Server.ini", "INIT", "ServerIp"))

#End If


Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
CamaraLenta = val(GetVar(IniPath & "Server.ini", "INIT", "CamaraLenta"))
ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))
ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))

VestimentaImperialHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialHumano"))
VestimentaImperialEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialEnano"))
TunicaConspicuaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaHumano"))
TunicaConspicuaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaEnano"))
ArmaduraNobilisimaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaHumano"))
ArmaduraNobilisimaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaEnano"))
ArmaduraGranSacerdote = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraGranSacerdote"))

VestimentaLegionHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionHumano"))
VestimentaLegionEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionEnano"))
TunicaLobregaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaHumano"))
TunicaLobregaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaEnano"))
TunicaEgregiaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaHumano"))
TunicaEgregiaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaEnano"))
SacerdoteDemoniaco = val(GetVar(IniPath & "Server.ini", "INIT", "SacerdoteDemoniaco"))

MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))

EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))
EncriptarProtocolosCriticos = val(GetVar(IniPath & "Server.ini", "INIT", "Encriptar"))

'Start pos
StartPos.map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.Text = IntervaloSed

IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

'TODO : Agregar estos intervalos al form!!!
IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))

frmMain.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))

IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))

'&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&

'Ressurect pos
ResPos.map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

'&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
'Se agregó en LoadBalance y en el Balance.dat
'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))

''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
Call Statistics.Initialize

Nix.map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

Lindos.map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

Arghal.map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")


Call MD5sCarga

Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Saves the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************

On Error GoTo Errhandler

Dim OldUserHead As Long


'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
'clase=0 es el error, porq el enum empieza de 1!!
If UserList(UserIndex).clase = 0 Or UserList(UserIndex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).name)
    Exit Sub
End If


If UserList(UserIndex).flags.Mimetizado = 1 Then
    UserList(UserIndex).Char.body = UserList(UserIndex).CharMimetizado.body
    UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
End If



If FileExist(UserFile, vbNormal) Then
    If UserList(UserIndex).flags.Muerto = 1 Then
        OldUserHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Head = GetVar(UserFile, "INIT", "Head")
    End If
'       Kill UserFile
End If

Dim LoopC As Integer


Call WriteVar(UserFile, "FLAGS", "Muerto", CStr(UserList(UserIndex).flags.Muerto))
Call WriteVar(UserFile, "FLAGS", "Escondido", CStr(UserList(UserIndex).flags.Escondido))
Call WriteVar(UserFile, "FLAGS", "Hambre", CStr(UserList(UserIndex).flags.Hambre))
Call WriteVar(UserFile, "FLAGS", "Sed", CStr(UserList(UserIndex).flags.Sed))
Call WriteVar(UserFile, "FLAGS", "Desnudo", CStr(UserList(UserIndex).flags.Desnudo))
Call WriteVar(UserFile, "FLAGS", "Ban", CStr(UserList(UserIndex).flags.Ban))
Call WriteVar(UserFile, "FLAGS", "Navegando", CStr(UserList(UserIndex).flags.Navegando))
Call WriteVar(UserFile, "FLAGS", "Envenenado", CStr(UserList(UserIndex).flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Paralizado", CStr(UserList(UserIndex).flags.Paralizado))

Call WriteVar(UserFile, "CONSEJO", "PERTENECE", IIf(UserList(UserIndex).flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", IIf(UserList(UserIndex).flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))


Call WriteVar(UserFile, "COUNTERS", "Pena", CStr(UserList(UserIndex).Counters.Pena))

Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", CStr(UserList(UserIndex).Faccion.ArmadaReal))
Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", CStr(UserList(UserIndex).Faccion.FuerzasCaos))
Call WriteVar(UserFile, "FACCIONES", "CiudMatados", CStr(UserList(UserIndex).Faccion.CiudadanosMatados))
Call WriteVar(UserFile, "FACCIONES", "CrimMatados", CStr(UserList(UserIndex).Faccion.CriminalesMatados))
Call WriteVar(UserFile, "FACCIONES", "rArCaos", CStr(UserList(UserIndex).Faccion.RecibioArmaduraCaos))
Call WriteVar(UserFile, "FACCIONES", "rArReal", CStr(UserList(UserIndex).Faccion.RecibioArmaduraReal))
Call WriteVar(UserFile, "FACCIONES", "rExCaos", CStr(UserList(UserIndex).Faccion.RecibioExpInicialCaos))
Call WriteVar(UserFile, "FACCIONES", "rExReal", CStr(UserList(UserIndex).Faccion.RecibioExpInicialReal))
Call WriteVar(UserFile, "FACCIONES", "recCaos", CStr(UserList(UserIndex).Faccion.RecompensasCaos))
Call WriteVar(UserFile, "FACCIONES", "recReal", CStr(UserList(UserIndex).Faccion.RecompensasReal))
Call WriteVar(UserFile, "FACCIONES", "Reenlistadas", CStr(UserList(UserIndex).Faccion.Reenlistadas))
Call WriteVar(UserFile, "FACCIONES", "NivelIngreso", CStr(UserList(UserIndex).Faccion.NivelIngreso))
Call WriteVar(UserFile, "FACCIONES", "FechaIngreso", UserList(UserIndex).Faccion.FechaIngreso)
Call WriteVar(UserFile, "FACCIONES", "MatadosIngreso", CStr(UserList(UserIndex).Faccion.MatadosIngreso))
Call WriteVar(UserFile, "FACCIONES", "NextRecompensa", CStr(UserList(UserIndex).Faccion.NextRecompensa))


'¿Fueron modificados los atributos del usuario?
If Not UserList(UserIndex).flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next
Else
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)))
    Next
End If

For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
    Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, CStr(UserList(UserIndex).Stats.UserSkills(LoopC)))
Next


Call WriteVar(UserFile, "CONTACTO", "Email", UserList(UserIndex).email)

Call WriteVar(UserFile, "INIT", "Genero", UserList(UserIndex).genero)
Call WriteVar(UserFile, "INIT", "Raza", UserList(UserIndex).raza)
Call WriteVar(UserFile, "INIT", "Hogar", UserList(UserIndex).Hogar)
Call WriteVar(UserFile, "INIT", "Clase", UserList(UserIndex).clase)
Call WriteVar(UserFile, "INIT", "Desc", UserList(UserIndex).desc)

Call WriteVar(UserFile, "INIT", "Heading", CStr(UserList(UserIndex).Char.heading))

Call WriteVar(UserFile, "INIT", "Head", CStr(UserList(UserIndex).OrigChar.Head))

If UserList(UserIndex).flags.Muerto = 0 Then
    Call WriteVar(UserFile, "INIT", "Body", CStr(UserList(UserIndex).Char.body))
End If

Call WriteVar(UserFile, "INIT", "Arma", CStr(UserList(UserIndex).Char.WeaponAnim))
Call WriteVar(UserFile, "INIT", "Escudo", CStr(UserList(UserIndex).Char.ShieldAnim))
Call WriteVar(UserFile, "INIT", "Casco", CStr(UserList(UserIndex).Char.CascoAnim))

#If ConUpTime Then
    Dim TempDate As Date
    TempDate = Now - UserList(UserIndex).LogOnTime
    UserList(UserIndex).LogOnTime = Now
    UserList(UserIndex).UpTime = UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
    UserList(UserIndex).UpTime = UserList(UserIndex).UpTime
    Call WriteVar(UserFile, "INIT", "UpTime", UserList(UserIndex).UpTime)
#End If

'First time around?
If GetVar(UserFile, "INIT", "LastIP1") = vbNullString Then
    Call WriteVar(UserFile, "INIT", "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & time)
'Is it a different ip from last time?
ElseIf UserList(UserIndex).ip <> Left$(GetVar(UserFile, "INIT", "LastIP1"), InStr(1, GetVar(UserFile, "INIT", "LastIP1"), " ") - 1) Then
    Dim i As Integer
    For i = 5 To 2 Step -1
        Call WriteVar(UserFile, "INIT", "LastIP" & i, GetVar(UserFile, "INIT", "LastIP" & CStr(i - 1)))
    Next i
    Call WriteVar(UserFile, "INIT", "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & time)
'Same ip, just update the date
Else
    Call WriteVar(UserFile, "INIT", "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & time)
End If



Call WriteVar(UserFile, "INIT", "Position", UserList(UserIndex).Pos.map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)


Call WriteVar(UserFile, "STATS", "GLD", CStr(UserList(UserIndex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "BANCO", CStr(UserList(UserIndex).Stats.Banco))

Call WriteVar(UserFile, "STATS", "MaxHP", CStr(UserList(UserIndex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", CStr(UserList(UserIndex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "MaxSTA", CStr(UserList(UserIndex).Stats.MaxSta))
Call WriteVar(UserFile, "STATS", "MinSTA", CStr(UserList(UserIndex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxMAN", CStr(UserList(UserIndex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", CStr(UserList(UserIndex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", CStr(UserList(UserIndex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", CStr(UserList(UserIndex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "MaxAGU", CStr(UserList(UserIndex).Stats.MaxAGU))
Call WriteVar(UserFile, "STATS", "MinAGU", CStr(UserList(UserIndex).Stats.MinAGU))

Call WriteVar(UserFile, "STATS", "MaxHAM", CStr(UserList(UserIndex).Stats.MaxHam))
Call WriteVar(UserFile, "STATS", "MinHAM", CStr(UserList(UserIndex).Stats.MinHam))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", CStr(UserList(UserIndex).Stats.SkillPts))
  
Call WriteVar(UserFile, "STATS", "EXP", CStr(UserList(UserIndex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", CStr(UserList(UserIndex).Stats.ELV))





Call WriteVar(UserFile, "STATS", "ELU", CStr(UserList(UserIndex).Stats.ELU))
Call WriteVar(UserFile, "MUERTES", "UserMuertes", CStr(UserList(UserIndex).Stats.UsuariosMatados))
'Call WriteVar(UserFile, "MUERTES", "CrimMuertes", CStr(UserList(UserIndex).Stats.CriminalesMatados))
Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", CStr(UserList(UserIndex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call WriteVar(UserFile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).amount)
Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped)
Next

Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", CStr(UserList(UserIndex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", CStr(UserList(UserIndex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", CStr(UserList(UserIndex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", CStr(UserList(UserIndex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Inventory", "BarcoSlot", CStr(UserList(UserIndex).Invent.BarcoSlot))
Call WriteVar(UserFile, "Inventory", "MunicionSlot", CStr(UserList(UserIndex).Invent.MunicionEqpSlot))
'/Nacho

Call WriteVar(UserFile, "Inventory", "AnilloSlot", CStr(UserList(UserIndex).Invent.AnilloEqpSlot))


'Reputacion
Call WriteVar(UserFile, "REP", "Asesino", CStr(UserList(UserIndex).Reputacion.AsesinoRep))
Call WriteVar(UserFile, "REP", "Bandido", CStr(UserList(UserIndex).Reputacion.BandidoRep))
Call WriteVar(UserFile, "REP", "Burguesia", CStr(UserList(UserIndex).Reputacion.BurguesRep))
Call WriteVar(UserFile, "REP", "Ladrones", CStr(UserList(UserIndex).Reputacion.LadronesRep))
Call WriteVar(UserFile, "REP", "Nobles", CStr(UserList(UserIndex).Reputacion.NobleRep))
Call WriteVar(UserFile, "REP", "Plebe", CStr(UserList(UserIndex).Reputacion.PlebeRep))

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Call WriteVar(UserFile, "REP", "Promedio", CStr(L))

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next

Dim NroMascotas As Long
NroMascotas = UserList(UserIndex).NroMascotas

For LoopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(UserIndex).MascotasIndex(LoopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(UserIndex).MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(UserIndex).MascotasType(LoopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            NroMascotas = NroMascotas - 1
        End If
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
    Else
        cad = UserList(UserIndex).MascotasType(LoopC)
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
    End If

Next

Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", CStr(NroMascotas))

'Devuelve el head de muerto
If UserList(UserIndex).flags.Muerto = 1 Then
    UserList(UserIndex).Char.Head = iCabezaMuerto
End If

Exit Sub

Errhandler:
Call LogError("Error en SaveUser")

End Sub

Function criminal(ByVal UserIndex As Integer) As Boolean

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
criminal = (L < 0)

End Function

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

'If NpcNumero > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados)) 'Que es ESTO?!!




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

Dim npcfile As String

'If NpcNumber > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).amount = 0
    Next LoopC
End If



Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "BannedBy", UserList(UserIndex).name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)


'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub
