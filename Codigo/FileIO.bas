Attribute VB_Name = "ES"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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
    Dim Leer As New clsLeerInis

    Leer.Abrir App.Path & "\Dat\Invokar.dat"

    Dim N As Integer, LoopC As Integer
    N = val(Leer.DarValor("INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(Leer.DarValor("LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = Leer.DarValor("LIST", "NN" & LoopC)
    Next LoopC
        
    Set Leer = Nothing
End Sub

Function EsAdmin(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
Dim Leer As New clsLeerInis

Leer.Abrir IniPath & "Server.ini"

NumWizs = val(Leer.DarValor("INIT", "Admines"))

For WizNum = 1 To NumWizs
    NomB = UCase$(Leer.DarValor("Admines", "Admin" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsAdmin = True
        Exit Function
    End If
Next WizNum
EsAdmin = False

Set Leer = Nothing
End Function

Function EsDios(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
Dim Leer As New clsLeerInis

Leer.Abrir IniPath & "Server.ini"

NumWizs = val(Leer.DarValor("INIT", "Dioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(Leer.DarValor("Dioses", "Dios" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
Set Leer = Nothing
End Function

Function EsSemiDios(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
Dim Leer As New clsLeerInis

Leer.Abrir IniPath & "Server.ini"

NumWizs = val(Leer.DarValor("INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(Leer.DarValor("SemiDioses", "SemiDios" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False
Set Leer = Nothing
End Function

Function EsConsejero(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
Dim Leer As New clsLeerInis

Leer.Abrir IniPath & "Server.ini"

NumWizs = val(Leer.DarValor("INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    NomB = UCase$(Leer.DarValor("Consejeros", "Consejero" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
Set Leer = Nothing
End Function

Function EsRolesMaster(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
Dim Leer As New clsLeerInis

Leer.Abrir IniPath & "Server.ini"

NumWizs = val(Leer.DarValor("INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(Leer.DarValor("RolesMasters", "RM" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
Set Leer = Nothing
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

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsLeerInis

Leer.Abrir DatPath & "Hechizos.dat"
'j = Val(Leer.DarValor(

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.DarValor("INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.Value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = Leer.DarValor("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).Desc = Leer.DarValor("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.DarValor("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.DarValor("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.DarValor("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.DarValor("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(Leer.DarValor("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = val(Leer.DarValor("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.DarValor("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(Leer.DarValor("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Resis = val(Leer.DarValor("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.DarValor("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.DarValor("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.DarValor("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.DarValor("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.DarValor("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.DarValor("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.DarValor("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(Leer.DarValor("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.DarValor("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.DarValor("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.DarValor("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.DarValor("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.DarValor("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.DarValor("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.DarValor("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.DarValor("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = val(Leer.DarValor("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.DarValor("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    
    Hechizos(Hechizo).CuraVeneno = val(Leer.DarValor("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.DarValor("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(Leer.DarValor("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.DarValor("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.DarValor("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.DarValor("Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = val(Leer.DarValor("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(Leer.DarValor("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(Leer.DarValor("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(Leer.DarValor("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).cant = val(Leer.DarValor("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = val(Leer.DarValor("hechizo" & Hechizo, "Mimetiza"))
    
    
    Hechizos(Hechizo).Materializa = val(Leer.DarValor("Hechizo" & Hechizo, "Materializa"))
    Hechizos(Hechizo).ItemIndex = val(Leer.DarValor("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.DarValor("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.DarValor("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = val(Leer.DarValor("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = val(Leer.DarValor("Hechizo" & Hechizo, "Target"))
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
    
    Hechizos(Hechizo).NeedStaff = val(Leer.DarValor("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(val(Leer.DarValor("Hechizo" & Hechizo, "StaffAffected")))
    
Next Hechizo

Set Leer = Nothing
Exit Sub

errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

Sub LoadMotd()
Dim i As Integer
Dim Leer As New clsLeerInis

Leer.Abrir App.Path & "\Dat\Motd.ini"

MaxLines = val(Leer.DarValor("INIT", "NumLines"))
ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = Leer.DarValor("Motd", "Line" & i)
    MOTD(i).Formato = ""
Next i

Set Leer = Nothing
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



Call SendData(SendTarget.ToAll, 0, 0, "BKW")


Call LimpiarMundo
Call WorldSave
Call modGuilds.v_RutinaElecciones
Call ResetCentinelaInfo     'Reseteamos al centinela


Call SendData(SendTarget.ToAll, 0, 0, "BKW")

'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
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
            
    Put FreeFileMap, , MapInfo(Map).MapVersion
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
                
                If MapData(Map, X, Y).Blocked Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).Graphic(2) Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).Graphic(3) Then ByFlags = ByFlags Or 4
                If MapData(Map, X, Y).Graphic(4) Then ByFlags = ByFlags Or 8
                If MapData(Map, X, Y).trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , MapData(Map, X, Y).Graphic(1)
                
                For LoopC = 2 To 4
                    If MapData(Map, X, Y).Graphic(LoopC) Then _
                        Put FreeFileMap, , MapData(Map, X, Y).Graphic(LoopC)
                Next LoopC
                
                If MapData(Map, X, Y).trigger Then _
                    Put FreeFileMap, , CInt(MapData(Map, X, Y).trigger)
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                   If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        MapData(Map, X, Y).OBJInfo.ObjIndex = 0
                        MapData(Map, X, Y).OBJInfo.Amount = 0
                    End If
                End If
    
                If MapData(Map, X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).NpcIndex Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(Map, X, Y).TileExit.Map Then
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.X
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Y
                End If
                
                If MapData(Map, X, Y).NpcIndex Then _
                    Put FreeFileInf, , Npclist(MapData(Map, X, Y).NpcIndex).Numero
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
                End If
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf

    'write .dat file
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Name", MapInfo(Map).name)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "MusicNum", MapInfo(Map).Music)
    Call WriteVar(MAPFILE & ".dat", "mapa" & Map, "MagiaSinefecto", MapInfo(Map).MagiaSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "StartPos", MapInfo(Map).StartPos.Map & "-" & MapInfo(Map).StartPos.X & "-" & MapInfo(Map).StartPos.Y)

    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Terreno", MapInfo(Map).Terreno)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Zona", MapInfo(Map).Zona)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Restringir", MapInfo(Map).Restringir)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "BackUp", str(MapInfo(Map).BackUp))

    If MapInfo(Map).Pk Then
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "0")
    Else
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "1")
    End If

End Sub
Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer
Dim Leer As New clsLeerInis

Leer.Abrir DatPath & "ArmasHerrero.dat"

N = val(Leer.DarValor("INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = val(Leer.DarValor("Arma" & lc, "Index"))
Next lc

Set Leer = Nothing
End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer
Dim Leer As New clsLeerInis

Leer.Abrir DatPath & "ArmadurasHerrero.dat"

N = val(Leer.DarValor("INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = val(Leer.DarValor("Armadura" & lc, "Index"))
Next lc

Set Leer = Nothing

End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lc As Integer
Dim Leer As New clsLeerInis

Leer.Abrir DatPath & "ObjCarpintero.dat"

N = val(Leer.DarValor("INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To N) As Integer

For lc = 1 To N
    ObjCarpintero(lc) = val(Leer.DarValor("Obj" & lc, "Index"))
Next lc

Set Leer = Nothing
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

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsLeerInis

Leer.Abrir DatPath & "Obj.dat"
'j = val(Leer.DarValor("INIT", "NumObjs"))  '

'obtiene el numero de obj
NumObjDatas = val(Leer.DarValor("INIT", "NumObjs"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.Value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).name = Leer.DarValor("OBJ" & Object, "Name")
    
    ObjData(Object).GrhIndex = val(Leer.DarValor("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = val(Leer.DarValor("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = val(Leer.DarValor("OBJ" & Object, "Newbie"))
    
    Select Case ObjData(Object).OBJType
        Case eOBJType.otArmadura
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = val(Leer.DarValor("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.DarValor("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = val(Leer.DarValor("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.DarValor("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = val(Leer.DarValor("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.DarValor("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = val(Leer.DarValor("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.DarValor("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.DarValor("OBJ" & Object, "Caos"))
        
        Case eOBJType.otHerramientas
            ObjData(Object).LingH = val(Leer.DarValor("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.DarValor("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.DarValor("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.DarValor("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = val(Leer.DarValor("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = val(Leer.DarValor("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = val(Leer.DarValor("OBJ" & Object, "SND3"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = val(Leer.DarValor("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = val(Leer.DarValor("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = val(Leer.DarValor("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = val(Leer.DarValor("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = val(Leer.DarValor("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = val(Leer.DarValor("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = val(Leer.DarValor("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = val(Leer.DarValor("OBJ" & Object, "DuracionEfecto"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = val(Leer.DarValor("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = val(Leer.DarValor("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.DarValor("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.DarValor("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.DarValor("OBJ" & Object, "Paraliza"))
    End Select
    
    ObjData(Object).Ropaje = val(Leer.DarValor("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.DarValor("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = val(Leer.DarValor("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = val(Leer.DarValor("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(Leer.DarValor("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.DarValor("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = val(Leer.DarValor("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.DarValor("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.DarValor("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.DarValor("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = val(Leer.DarValor("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.DarValor("OBJ" & Object, "MAXDEF"))
    
    ObjData(Object).RazaEnana = val(Leer.DarValor("OBJ" & Object, "RazaEnana"))
    
    ObjData(Object).Valor = val(Leer.DarValor("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(Leer.DarValor("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(Leer.DarValor("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = val(Leer.DarValor("OBJ" & Object, "Llave"))
        ObjData(Object).clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = val(Leer.DarValor("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.DarValor("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.DarValor("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.DarValor("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.DarValor("OBJ" & Object, "ID")
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = Leer.DarValor("OBJ" & Object, "CP" & i)
    Next i
    
    ObjData(Object).DefensaMagicaMax = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.DarValor("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = val(Leer.DarValor("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(Leer.DarValor("OBJ" & Object, "Madera"))
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.DarValor("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = val(Leer.DarValor("OBJ" & Object, "NoSeCae"))
    
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
Next Object

Set Leer = Nothing

Exit Sub

errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

Sub LoadUserStats(UserIndex As Integer, UserFile As String)

Dim LoopC As Integer
Dim Leer As New clsLeerInis

Leer.Abrir UserFile


For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = Leer.DarValor("ATRIBUTOS", "AT" & LoopC)
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = val(Leer.DarValor("SKILLS", "SK" & LoopC))
Next

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = val(Leer.DarValor("Hechizos", "H" & LoopC))
Next

UserList(UserIndex).Stats.GLD = val(Leer.DarValor("STATS", "GLD"))
UserList(UserIndex).Stats.Banco = val(Leer.DarValor("STATS", "BANCO"))

UserList(UserIndex).Stats.MET = val(Leer.DarValor("STATS", "MET"))
UserList(UserIndex).Stats.MaxHP = val(Leer.DarValor("STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = val(Leer.DarValor("STATS", "MinHP"))

UserList(UserIndex).Stats.FIT = val(Leer.DarValor("STATS", "FIT"))
UserList(UserIndex).Stats.MinSta = val(Leer.DarValor("STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = val(Leer.DarValor("STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = val(Leer.DarValor("STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = val(Leer.DarValor("STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = val(Leer.DarValor("STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = val(Leer.DarValor("STATS", "MinHIT"))

UserList(UserIndex).Stats.MaxAGU = val(Leer.DarValor("STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = val(Leer.DarValor("STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = val(Leer.DarValor("STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = val(Leer.DarValor("STATS", "MinHAM"))

UserList(UserIndex).Stats.SkillPts = val(Leer.DarValor("STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.Exp = val(Leer.DarValor("STATS", "EXP"))
UserList(UserIndex).Stats.ELU = val(Leer.DarValor("STATS", "ELU"))
UserList(UserIndex).Stats.ELV = val(Leer.DarValor("STATS", "ELV"))


UserList(UserIndex).Stats.UsuariosMatados = val(Leer.DarValor("MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.CriminalesMatados = val(Leer.DarValor("MUERTES", "CrimMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = val(Leer.DarValor("MUERTES", "NpcsMuertes"))

UserList(UserIndex).flags.PertAlCons = val(Leer.DarValor("CONSEJO", "PERTENECE"))
UserList(UserIndex).flags.PertAlConsCaos = val(Leer.DarValor("CONSEJO", "PERTENECECAOS"))

Set Leer = Nothing

End Sub

Sub LoadUserReputacion(UserIndex As Integer, UserFile As String)

Dim Leer As New clsLeerInis

Leer.Abrir UserFile

UserList(UserIndex).Reputacion.AsesinoRep = val(Leer.DarValor("REP", "Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = val(Leer.DarValor("REP", "Bandido"))
UserList(UserIndex).Reputacion.BurguesRep = val(Leer.DarValor("REP", "Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = val(Leer.DarValor("REP", "Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = val(Leer.DarValor("REP", "Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = val(Leer.DarValor("REP", "Plebe"))
UserList(UserIndex).Reputacion.Promedio = val(Leer.DarValor("REP", "Promedio"))

Set Leer = Nothing

End Sub


Sub LoadUserInit(UserIndex As Integer, UserFile As String)


Dim LoopC As Integer
Dim ln As String
Dim ln2 As String
Dim Cantidad As Long
Dim Leer As New clsLeerInis

Leer.Abrir UserFile

UserList(UserIndex).Faccion.ArmadaReal = val(Leer.DarValor("FACCIONES", "EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = val(Leer.DarValor("FACCIONES", "EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = val(Leer.DarValor("FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = val(Leer.DarValor("FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = val(Leer.DarValor("FACCIONES", "rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = val(Leer.DarValor("FACCIONES", "rArReal"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = val(Leer.DarValor("FACCIONES", "rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = val(Leer.DarValor("FACCIONES", "rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = val(Leer.DarValor("FACCIONES", "recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = val(Leer.DarValor("FACCIONES", "recReal"))
UserList(UserIndex).Faccion.Reenlistadas = val(Leer.DarValor("FACCIONES", "Reenlistadas"))

UserList(UserIndex).flags.Muerto = val(Leer.DarValor("FLAGS", "Muerto"))
UserList(UserIndex).flags.Escondido = val(Leer.DarValor("FLAGS", "Escondido"))

UserList(UserIndex).flags.Hambre = val(Leer.DarValor("FLAGS", "Hambre"))
UserList(UserIndex).flags.Sed = val(Leer.DarValor("FLAGS", "Sed"))
UserList(UserIndex).flags.Desnudo = val(Leer.DarValor("FLAGS", "Desnudo"))

UserList(UserIndex).flags.Envenenado = val(Leer.DarValor("FLAGS", "Envenenado"))
UserList(UserIndex).flags.Paralizado = val(Leer.DarValor("FLAGS", "Paralizado"))
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If
UserList(UserIndex).flags.Navegando = val(Leer.DarValor("FLAGS", "Navegando"))


UserList(UserIndex).Counters.Pena = val(Leer.DarValor("COUNTERS", "Pena"))

UserList(UserIndex).email = Leer.DarValor("CONTACTO", "Email")

UserList(UserIndex).Genero = Leer.DarValor("INIT", "Genero")
UserList(UserIndex).Clase = Leer.DarValor("INIT", "Clase")
UserList(UserIndex).Raza = Leer.DarValor("INIT", "Raza")
UserList(UserIndex).Hogar = Leer.DarValor("INIT", "Hogar")
UserList(UserIndex).Char.Heading = val(Leer.DarValor("INIT", "Heading"))


UserList(UserIndex).OrigChar.Head = val(Leer.DarValor("INIT", "Head"))
UserList(UserIndex).OrigChar.Body = val(Leer.DarValor("INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = val(Leer.DarValor("INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = val(Leer.DarValor("INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = val(Leer.DarValor("INIT", "Casco"))
UserList(UserIndex).OrigChar.Heading = eHeading.SOUTH

If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
End If


UserList(UserIndex).Desc = Leer.DarValor("INIT", "Desc")


UserList(UserIndex).Pos.Map = val(ReadField(1, Leer.DarValor("INIT", "Position"), 45))
UserList(UserIndex).Pos.X = val(ReadField(2, Leer.DarValor("INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = val(ReadField(3, Leer.DarValor("INIT", "Position"), 45))

UserList(UserIndex).Invent.NroItems = Leer.DarValor("Inventory", "CantidadItems")

Dim loopd As Integer

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = val(Leer.DarValor("BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    ln2 = Leer.DarValor("BancoInventory", "Obj" & loopd)
    UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex = val(ReadField(1, ln2, 45))
    UserList(UserIndex).BancoInvent.Object(loopd).Amount = val(ReadField(2, ln2, 45))
Next loopd
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = Leer.DarValor("Inventory", "Obj" & LoopC)
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = val(Leer.DarValor("Inventory", "WeaponEqpSlot"))
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = val(Leer.DarValor("Inventory", "ArmourEqpSlot"))
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = val(Leer.DarValor("Inventory", "EscudoEqpSlot"))
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = val(Leer.DarValor("Inventory", "CascoEqpSlot"))
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = val(Leer.DarValor("Inventory", "BarcoSlot"))
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = val(Leer.DarValor("Inventory", "MunicionSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto herramienta
UserList(UserIndex).Invent.HerramientaEqpSlot = val(Leer.DarValor("Inventory", "HerramientaSlot"))
If UserList(UserIndex).Invent.HerramientaEqpSlot > 0 Then
    UserList(UserIndex).Invent.HerramientaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.HerramientaEqpSlot).ObjIndex
End If

UserList(UserIndex).NroMacotas = 0

ln = Leer.DarValor("Guild", "GUILDINDEX")
If IsNumeric(ln) Then
    UserList(UserIndex).GuildIndex = CInt(ln)
Else
    UserList(UserIndex).GuildIndex = 0
End If

Set Leer = Nothing

End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.Min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
        Else
            tFileName = App.Path & MapPath & "Mapa" & Map
        End If
        
        Call CargarMapa(Map, tFileName)
        
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.Min = 0
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

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByVal MAPFl As String)
On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim TempInt As Integer
    Dim Leer As New clsLeerInis
    
    Leer.Abrir MAPFl & ".dat"
      
    FreeFileMap = FreeFile
    
    Open MAPFl & ".map" For Binary As #FreeFileMap
    Seek FreeFileMap, 1
    
    FreeFileInf = FreeFile
    
    'inf
    Open MAPFl & ".inf" For Binary As #FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    Get #FreeFileMap, , MapInfo(Map).MapVersion
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
                MapData(Map, X, Y).Blocked = 1
            End If
            
            Get FreeFileMap, , MapData(Map, X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get FreeFileMap, , TempInt
                MapData(Map, X, Y).trigger = TempInt
            End If
            
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.X
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Y
            End If
            
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(Map, X, Y).NpcIndex
                
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                    Else
                        npcfile = DatPath & "NPCs.dat"
                    End If

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                    Else
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    End If
                            
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                            
                    Call MakeNPCChar(SendTarget.ToMap, 0, 0, MapData(Map, X, Y).NpcIndex, 1, 1, 1)
                End If
            End If
            
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
            End If
        Next X
    Next Y
    
    
    Close FreeFileMap
    Close FreeFileInf
    
    MapInfo(Map).name = Leer.DarValor("Mapa" & Map, "Name")
    MapInfo(Map).Music = Leer.DarValor("Mapa" & Map, "MusicNum")
    MapInfo(Map).StartPos.Map = val(ReadField(1, Leer.DarValor("Mapa" & Map, "StartPos"), Asc("-")))
    MapInfo(Map).StartPos.X = val(ReadField(2, Leer.DarValor("Mapa" & Map, "StartPos"), Asc("-")))
    MapInfo(Map).StartPos.Y = val(ReadField(3, Leer.DarValor("Mapa" & Map, "StartPos"), Asc("-")))
    MapInfo(Map).MagiaSinEfecto = val(Leer.DarValor("Mapa" & Map, "MagiaSinEfecto"))
    MapInfo(Map).NoEncriptarMP = val(Leer.DarValor("Mapa" & Map, "NoEncriptarMP"))
    
    If val(Leer.DarValor("Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = Leer.DarValor("Mapa" & Map, "Terreno")
    MapInfo(Map).Zona = Leer.DarValor("Mapa" & Map, "Zona")
    MapInfo(Map).Restringir = Leer.DarValor("Mapa" & Map, "Restringir")
    MapInfo(Map).BackUp = val(Leer.DarValor("Mapa" & Map, "BACKUP"))
        
    Set Leer = Nothing
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & "." & Err.Description)
End Sub

Sub LoadSini()

Dim Temporal As Long
Dim Temporal1 As Long
Dim LoopC As Integer
Dim Leer As New clsLeerInis
Dim LeerCiudades As New clsLeerInis

Leer.Abrir IniPath & "Server.ini"
LeerCiudades.Abrir DatPath & "Ciudades.dat"

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

BootDelBackUp = val(Leer.DarValor("INIT", "IniciarDesdeBackUp"))

'Misc
CrcSubKey = val(Leer.DarValor("INIT", "CrcSubKey"))

ServerIp = Leer.DarValor("INIT", "ServerIp")
Temporal = InStr(1, ServerIp, ".")
Temporal1 = (Mid(ServerIp, 1, Temporal - 1) And &H7F) * 16777216
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + Mid(ServerIp, 1, Temporal - 1) * 65536
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))
Temporal = InStr(1, ServerIp, ".")
Temporal1 = Temporal1 + Mid(ServerIp, 1, Temporal - 1) * 256
ServerIp = Mid(ServerIp, Temporal + 1, Len(ServerIp))

MixedKey = (Temporal1 + ServerIp) Xor &H65F64B42

Puerto = val(Leer.DarValor("INIT", "StartPort"))
HideMe = val(Leer.DarValor("INIT", "Hide"))
AllowMultiLogins = val(Leer.DarValor("INIT", "AllowMultiLogins"))
IdleLimit = val(Leer.DarValor("INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = Leer.DarValor("INIT", "Version")

PuedeCrearPersonajes = val(Leer.DarValor("INIT", "PuedeCrearPersonajes"))
CamaraLenta = val(Leer.DarValor("INIT", "CamaraLenta"))
ServerSoloGMs = val(Leer.DarValor("init", "ServerSoloGMs"))

ArmaduraImperial1 = val(Leer.DarValor("INIT", "ArmaduraImperial1"))
ArmaduraImperial2 = val(Leer.DarValor("INIT", "ArmaduraImperial2"))
ArmaduraImperial3 = val(Leer.DarValor("INIT", "ArmaduraImperial3"))
TunicaMagoImperial = val(Leer.DarValor("INIT", "TunicaMagoImperial"))
TunicaMagoImperialEnanos = val(Leer.DarValor("INIT", "TunicaMagoImperialEnanos"))
ArmaduraCaos1 = val(Leer.DarValor("INIT", "ArmaduraCaos1"))
ArmaduraCaos2 = val(Leer.DarValor("INIT", "ArmaduraCaos2"))
ArmaduraCaos3 = val(Leer.DarValor("INIT", "ArmaduraCaos3"))
TunicaMagoCaos = val(Leer.DarValor("INIT", "TunicaMagoCaos"))
TunicaMagoCaosEnanos = val(Leer.DarValor("INIT", "TunicaMagoCaosEnanos"))

VestimentaImperialHumano = val(Leer.DarValor("INIT", "VestimentaImperialHumano"))
VestimentaImperialEnano = val(Leer.DarValor("INIT", "VestimentaImperialEnano"))
TunicaConspicuaHumano = val(Leer.DarValor("INIT", "TunicaConspicuaHumano"))
TunicaConspicuaEnano = val(Leer.DarValor("INIT", "TunicaConspicuaEnano"))
ArmaduraNobilisimaHumano = val(Leer.DarValor("INIT", "ArmaduraNobilisimaHumano"))
ArmaduraNobilisimaEnano = val(Leer.DarValor("INIT", "ArmaduraNobilisimaEnano"))
ArmaduraGranSacerdote = val(Leer.DarValor("INIT", "ArmaduraGranSacerdote"))

VestimentaLegionHumano = val(Leer.DarValor("INIT", "VestimentaLegionHumano"))
VestimentaLegionEnano = val(Leer.DarValor("INIT", "VestimentaLegionEnano"))
TunicaLobregaHumano = val(Leer.DarValor("INIT", "TunicaLobregaHumano"))
TunicaLobregaEnano = val(Leer.DarValor("INIT", "TunicaLobregaEnano"))
TunicaEgregiaHumano = val(Leer.DarValor("INIT", "TunicaEgregiaHumano"))
TunicaEgregiaEnano = val(Leer.DarValor("INIT", "TunicaEgregiaEnano"))
SacerdoteDemoniaco = val(Leer.DarValor("INIT", "SacerdoteDemoniaco"))

MAPA_PRETORIANO = val(Leer.DarValor("INIT", "MapaPretoriano"))

ClientsCommandsQueue = val(Leer.DarValor("INIT", "ClientsCommandsQueue"))
EnTesting = val(Leer.DarValor("INIT", "Testing"))
EncriptarProtocolosCriticos = val(Leer.DarValor("INIT", "Encriptar"))

'Start pos
StartPos.Map = val(ReadField(1, Leer.DarValor("INIT", "StartPos"), 45))
StartPos.X = val(ReadField(2, Leer.DarValor("INIT", "StartPos"), 45))
StartPos.Y = val(ReadField(3, Leer.DarValor("INIT", "StartPos"), 45))

'Intervalos
SanaIntervaloSinDescansar = val(Leer.DarValor("INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = val(Leer.DarValor("INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(Leer.DarValor("INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(Leer.DarValor("INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloSed = val(Leer.DarValor("INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.Text = IntervaloSed

IntervaloHambre = val(Leer.DarValor("INTERVALOS", "IntervaloHambre"))
FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(Leer.DarValor("INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(Leer.DarValor("INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = val(Leer.DarValor("INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(Leer.DarValor("INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = val(Leer.DarValor("INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = val(Leer.DarValor("INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(Leer.DarValor("INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(Leer.DarValor("INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = val(Leer.DarValor("INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = val(Leer.DarValor("INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = val(Leer.DarValor("INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(Leer.DarValor("INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

frmMain.tLluvia.Interval = val(Leer.DarValor("INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval

frmMain.CmdExec.Interval = val(Leer.DarValor("INTERVALOS", "IntervaloTimerExec"))
FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

MinutosWs = val(Leer.DarValor("INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = val(Leer.DarValor("INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(Leer.DarValor("INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = val(Leer.DarValor("INTERVALOS", "IntervaloFlechasCazadores"))

IntervaloAutoReiniciar = val(Leer.DarValor("INTERVALOS", "IntervaloAutoReiniciar"))


'Ressurect pos
ResPos.Map = val(ReadField(1, Leer.DarValor("INIT", "ResPos"), 45))
ResPos.X = val(ReadField(2, Leer.DarValor("INIT", "ResPos"), 45))
ResPos.Y = val(ReadField(3, Leer.DarValor("INIT", "ResPos"), 45))
  
recordusuarios = val(Leer.DarValor("INIT", "Record"))
  
'Max users
Temporal = val(Leer.DarValor("INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

Nix.Map = LeerCiudades.DarValor("NIX", "Mapa")
Nix.X = LeerCiudades.DarValor("NIX", "X")
Nix.Y = LeerCiudades.DarValor("NIX", "Y")

Ullathorpe.Map = LeerCiudades.DarValor("Ullathorpe", "Mapa")
Ullathorpe.X = LeerCiudades.DarValor("Ullathorpe", "X")
Ullathorpe.Y = LeerCiudades.DarValor("Ullathorpe", "Y")

Banderbill.Map = LeerCiudades.DarValor("Banderbill", "Mapa")
Banderbill.X = LeerCiudades.DarValor("Banderbill", "X")
Banderbill.Y = LeerCiudades.DarValor("Banderbill", "Y")

Lindos.Map = LeerCiudades.DarValor("Lindos", "Mapa")
Lindos.X = LeerCiudades.DarValor("Lindos", "X")
Lindos.Y = LeerCiudades.DarValor("Lindos", "Y")


Call MD5sCarga

Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

Set Leer = Nothing


End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)
On Error GoTo errhandler

Dim OldUserHead As Long


'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If UserList(UserIndex).Clase = "" Or UserList(UserIndex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).name)
    Exit Sub
End If


If UserList(UserIndex).flags.Mimetizado = 1 Then
    UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
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
        UserList(UserIndex).Char.Head = val(GetVar(UserFile, "INIT", "Head"))
       End If
'       Kill UserFile
End If

Dim LoopC As Integer


Call WriteVar(UserFile, "FLAGS", "Muerto", val(UserList(UserIndex).flags.Muerto))
Call WriteVar(UserFile, "FLAGS", "Escondido", val(UserList(UserIndex).flags.Escondido))
Call WriteVar(UserFile, "FLAGS", "Hambre", val(UserList(UserIndex).flags.Hambre))
Call WriteVar(UserFile, "FLAGS", "Sed", val(UserList(UserIndex).flags.Sed))
Call WriteVar(UserFile, "FLAGS", "Desnudo", val(UserList(UserIndex).flags.Desnudo))
Call WriteVar(UserFile, "FLAGS", "Ban", val(UserList(UserIndex).flags.Ban))
Call WriteVar(UserFile, "FLAGS", "Navegando", val(UserList(UserIndex).flags.Navegando))

Call WriteVar(UserFile, "FLAGS", "Envenenado", val(UserList(UserIndex).flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Paralizado", val(UserList(UserIndex).flags.Paralizado))

Call WriteVar(UserFile, "CONSEJO", "PERTENECE", UserList(UserIndex).flags.PertAlCons)
Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", UserList(UserIndex).flags.PertAlConsCaos)


Call WriteVar(UserFile, "COUNTERS", "Pena", val(UserList(UserIndex).Counters.Pena))

Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", val(UserList(UserIndex).Faccion.ArmadaReal))
Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", val(UserList(UserIndex).Faccion.FuerzasCaos))
Call WriteVar(UserFile, "FACCIONES", "CiudMatados", val(UserList(UserIndex).Faccion.CiudadanosMatados))
Call WriteVar(UserFile, "FACCIONES", "CrimMatados", val(UserList(UserIndex).Faccion.CriminalesMatados))
Call WriteVar(UserFile, "FACCIONES", "rArCaos", val(UserList(UserIndex).Faccion.RecibioArmaduraCaos))
Call WriteVar(UserFile, "FACCIONES", "rArReal", val(UserList(UserIndex).Faccion.RecibioArmaduraReal))
Call WriteVar(UserFile, "FACCIONES", "rExCaos", val(UserList(UserIndex).Faccion.RecibioExpInicialCaos))
Call WriteVar(UserFile, "FACCIONES", "rExReal", val(UserList(UserIndex).Faccion.RecibioExpInicialReal))
Call WriteVar(UserFile, "FACCIONES", "recCaos", val(UserList(UserIndex).Faccion.RecompensasCaos))
Call WriteVar(UserFile, "FACCIONES", "recReal", val(UserList(UserIndex).Faccion.RecompensasReal))
Call WriteVar(UserFile, "FACCIONES", "Reenlistadas", val(UserList(UserIndex).Faccion.Reenlistadas))

'¿Fueron modificados los atributos del usuario?
If Not UserList(UserIndex).flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next
Else
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)))
    Next
End If

For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
    Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, val(UserList(UserIndex).Stats.UserSkills(LoopC)))
Next


Call WriteVar(UserFile, "CONTACTO", "Email", UserList(UserIndex).email)

Call WriteVar(UserFile, "INIT", "Genero", UserList(UserIndex).Genero)
Call WriteVar(UserFile, "INIT", "Raza", UserList(UserIndex).Raza)
Call WriteVar(UserFile, "INIT", "Hogar", UserList(UserIndex).Hogar)
Call WriteVar(UserFile, "INIT", "Clase", UserList(UserIndex).Clase)
Call WriteVar(UserFile, "INIT", "Password", UserList(UserIndex).Password)
Call WriteVar(UserFile, "INIT", "Desc", UserList(UserIndex).Desc)

Call WriteVar(UserFile, "INIT", "Heading", str(UserList(UserIndex).Char.Heading))

Call WriteVar(UserFile, "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))

If UserList(UserIndex).flags.Muerto = 0 Then
    Call WriteVar(UserFile, "INIT", "Body", str(UserList(UserIndex).Char.Body))
End If

Call WriteVar(UserFile, "INIT", "Arma", str(UserList(UserIndex).Char.WeaponAnim))
Call WriteVar(UserFile, "INIT", "Escudo", str(UserList(UserIndex).Char.ShieldAnim))
Call WriteVar(UserFile, "INIT", "Casco", str(UserList(UserIndex).Char.CascoAnim))

Call WriteVar(UserFile, "INIT", "LastIP", UserList(UserIndex).ip)
Call WriteVar(UserFile, "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)


Call WriteVar(UserFile, "STATS", "GLD", str(UserList(UserIndex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "BANCO", str(UserList(UserIndex).Stats.Banco))

Call WriteVar(UserFile, "STATS", "MET", str(UserList(UserIndex).Stats.MET))
Call WriteVar(UserFile, "STATS", "MaxHP", str(UserList(UserIndex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", str(UserList(UserIndex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "FIT", str(UserList(UserIndex).Stats.FIT))
Call WriteVar(UserFile, "STATS", "MaxSTA", str(UserList(UserIndex).Stats.MaxSta))
Call WriteVar(UserFile, "STATS", "MinSTA", str(UserList(UserIndex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxMAN", str(UserList(UserIndex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", str(UserList(UserIndex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", str(UserList(UserIndex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", str(UserList(UserIndex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "MaxAGU", str(UserList(UserIndex).Stats.MaxAGU))
Call WriteVar(UserFile, "STATS", "MinAGU", str(UserList(UserIndex).Stats.MinAGU))

Call WriteVar(UserFile, "STATS", "MaxHAM", str(UserList(UserIndex).Stats.MaxHam))
Call WriteVar(UserFile, "STATS", "MinHAM", str(UserList(UserIndex).Stats.MinHam))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", str(UserList(UserIndex).Stats.SkillPts))
  
Call WriteVar(UserFile, "STATS", "EXP", str(UserList(UserIndex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", str(UserList(UserIndex).Stats.ELV))





Call WriteVar(UserFile, "STATS", "ELU", str(UserList(UserIndex).Stats.ELU))
Call WriteVar(UserFile, "MUERTES", "UserMuertes", val(UserList(UserIndex).Stats.UsuariosMatados))
Call WriteVar(UserFile, "MUERTES", "CrimMuertes", val(UserList(UserIndex).Stats.CriminalesMatados))
Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", val(UserList(UserIndex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call WriteVar(UserFile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount)
Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped)
Next

Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", str(UserList(UserIndex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", str(UserList(UserIndex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", str(UserList(UserIndex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", str(UserList(UserIndex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Inventory", "BarcoSlot", str(UserList(UserIndex).Invent.BarcoSlot))
Call WriteVar(UserFile, "Inventory", "MunicionSlot", str(UserList(UserIndex).Invent.MunicionEqpSlot))
Call WriteVar(UserFile, "Inventory", "HerramientaSlot", str(UserList(UserIndex).Invent.HerramientaEqpSlot))


'Reputacion
Call WriteVar(UserFile, "REP", "Asesino", val(UserList(UserIndex).Reputacion.AsesinoRep))
Call WriteVar(UserFile, "REP", "Bandido", val(UserList(UserIndex).Reputacion.BandidoRep))
Call WriteVar(UserFile, "REP", "Burguesia", val(UserList(UserIndex).Reputacion.BurguesRep))
Call WriteVar(UserFile, "REP", "Ladrones", val(UserList(UserIndex).Reputacion.LadronesRep))
Call WriteVar(UserFile, "REP", "Nobles", val(UserList(UserIndex).Reputacion.NobleRep))
Call WriteVar(UserFile, "REP", "Plebe", val(UserList(UserIndex).Reputacion.PlebeRep))

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Call WriteVar(UserFile, "REP", "Promedio", val(L))

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next

Dim NroMascotas As Long
NroMascotas = UserList(UserIndex).NroMacotas

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
    End If

Next

Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", str(NroMascotas))

'Devuelve el head de muerto
If UserList(UserIndex).flags.Muerto = 1 Then
    UserList(UserIndex).Char.Head = iCabezaMuerto
End If

Exit Sub

errhandler:
Call LogError("Error en SaveUser")

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Criminal = (L < 0)

End Function

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

If NpcNumero > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Inflacion", val(Npclist(NpcIndex).Inflacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados))




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

Dim npcfile As String
Dim Leer As New clsLeerInis

If NpcNumber > 499 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
        npcfile = DatPath & "bkNPCs.dat"
End If

Leer.Abrir npcfile

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).name = Leer.DarValor("NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = Leer.DarValor("NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(Leer.DarValor("NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(Leer.DarValor("NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(Leer.DarValor("NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(Leer.DarValor("NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(Leer.DarValor("NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(Leer.DarValor("NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(Leer.DarValor("NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(Leer.DarValor("NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(Leer.DarValor("NPC" & NpcNumber, "GiveEXP"))


Npclist(NpcIndex).GiveGLD = val(Leer.DarValor("NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).InvReSpawn = val(Leer.DarValor("NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHP = val(Leer.DarValor("NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(Leer.DarValor("NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(Leer.DarValor("NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(Leer.DarValor("NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(Leer.DarValor("NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(Leer.DarValor("NPC" & NpcNumber, "Alineacion"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(Leer.DarValor("NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = Leer.DarValor("NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
    Next LoopC
End If

Npclist(NpcIndex).Inflacion = val(Leer.DarValor("NPC" & NpcNumber, "Inflacion"))


Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = val(Leer.DarValor("NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(Leer.DarValor("NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(Leer.DarValor("NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(Leer.DarValor("NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(Leer.DarValor("NPC" & NpcNumber, "TipoItems"))

Set Leer = Nothing

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
