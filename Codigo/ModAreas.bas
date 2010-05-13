Attribute VB_Name = "ModAreas"
'**************************************************************
' ModAreas.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Original Idea by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
' Implemented by Lucio N. Tourrilhes (DuNga)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

' Modulo de envio por areas compatible con la versión 9.10.x ... By DuNga

Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    
    AreaReciveX As Integer
    AreaReciveY As Integer
    
    MinX As Integer '-!!!
    MinY As Integer '-!!!
    
    AreaID As Long
End Type

Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type

Public Const USER_NUEVO As Byte = 255

'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte

Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100) As Byte

Private AreasRecive(12) As Integer

Public ConnGroups() As ConnGroup

Public Sub InitAreas()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim loopX As Long

' Setup areas...
    For LoopC = 0 To 11
        AreasRecive(LoopC) = (2 ^ LoopC) Or IIf(LoopC <> 0, 2 ^ (LoopC - 1), 0) Or IIf(LoopC <> 11, 2 ^ (LoopC + 1), 0)
    Next LoopC
    
    For LoopC = 1 To 100
        PosToArea(LoopC) = LoopC \ 9
    Next LoopC
    
    For LoopC = 1 To 100
        For loopX = 1 To 100
            'Usamos 121 IDs de area para saber si pasasamos de area "más rápido"
            AreasInfo(LoopC, loopX) = (LoopC \ 9 + 1) * (loopX \ 9 + 1)
        Next loopX
    Next LoopC

'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
    
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For LoopC = 1 To NumMaps
        ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
        
        If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
        ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
    Next LoopC
End Sub

Public Sub AreasOptimizacion()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
'**************************************************************
    Dim LoopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
    
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(time) \ 3)) Then
        
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
        
        For LoopC = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
            
            ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, tCurDay & "-" & tCurHour))
            If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
            If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).NumUsers Then ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
        Next LoopC
        
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub

Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, Optional ByVal ButIndex As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 15/07/2009
'Es la función clave del sistema de areas... Es llamada al mover un user
'15/07/2009: ZaMa - Now it doesn't send an invisible admin char info
'**************************************************************
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long, Map As Long
    
    With UserList(UserIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        
        Map = .Pos.Map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex)
        
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
                
                '<<< User >>>
                If MapData(Map, X, Y).UserIndex Then
                    
                    TempInt = MapData(Map, X, Y).UserIndex
                    
                    If UserIndex <> TempInt Then
                        
                        ' Solo avisa al otro cliente si no es un admin invisible
                        If Not (UserList(TempInt).flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, UserIndex, TempInt, Map, X, Y)
                            
                            'Si el user estaba invisible le avisamos al nuevo cliente de eso
                            If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
                                If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                                    Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                                End If
                            End If
                        End If
                        
                        ' Solo avisa al otro cliente si no es un admin invisible
                        If Not (.flags.AdminInvisible = 1) Then
                            Call MakeUserChar(False, TempInt, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
                            
                            If .flags.invisible Or .flags.Oculto Then
                                If UserList(TempInt).flags.Privilegios And PlayerType.User Then
                                    Call WriteSetInvisible(TempInt, .Char.CharIndex, True)
                                End If
                            End If
                        End If
                        
                        Call FlushBuffer(TempInt)
                    
                    ElseIf Head = USER_NUEVO Then
                        If Not ButIndex Then
                            Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y)
                        End If
                    End If
                End If
                
                '<<< Npc >>>
                If MapData(Map, X, Y).NpcIndex Then
                    Call MakeNPCChar(False, UserIndex, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                End If
                 
                '<<< Item >>>
                If MapData(Map, X, Y).ObjInfo.ObjIndex Then
                    TempInt = MapData(Map, X, Y).ObjInfo.ObjIndex
                    If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                        Call WriteObjectCreate(UserIndex, ObjData(TempInt).GrhIndex, X, Y)
                        
                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, UserIndex, X, Y, MapData(Map, X, Y).Blocked)
                            Call Bloquear(False, UserIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                        End If
                    End If
                End If
            
            Next Y
        Next X
        
        'Precalculados :P
        TempInt = .Pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' Se llama cuando se mueve un Npc
'**************************************************************
    If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long
    
    With Npclist(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 9
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 35
            MinY = MinY + 27
            MaxX = MinX + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 9
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 35
            MinX = MinX + 27
            MaxY = MinY + 26
            .AreasInfo.MinX = CInt(MinX - 18)
            .AreasInfo.MinY = CInt(MinY)
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 9) - 1) * 9
            MaxY = MinY + 26
            
            MinX = ((.Pos.X \ 9) - 1) * 9
            MaxX = MinX + 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100

        
        'Actualizamos!!!
        If MapInfo(.Pos.Map).NumUsers <> 0 Then
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    If MapData(.Pos.Map, X, Y).UserIndex Then _
                        Call MakeNPCChar(False, MapData(.Pos.Map, X, Y).UserIndex, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)
                Next Y
            Next X
        End If
        
        'Precalculados :P
        TempInt = .Pos.X \ 9
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .Pos.Y \ 9
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub

Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim TempVal As Long
    Dim LoopC As Long
    
    'Search for the user
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(LoopC) = UserIndex Then Exit For
    Next LoopC
    
    'Char not found
    If LoopC > ConnGroups(Map).CountEntrys Then Exit Sub
    
    'Remove from old map
    ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
    TempVal = ConnGroups(Map).CountEntrys
    
    'Move list back
    For LoopC = LoopC To TempVal
        ConnGroups(Map).UserEntrys(LoopC) = ConnGroups(Map).UserEntrys(LoopC + 1)
    Next LoopC
    
    If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
    End If
End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, Optional ByVal ButIndex As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 04/01/2007
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'   - Now the method checks for repetead users instead of trusting parameters.
'   - If the character is new to the map, update it
'**************************************************************
    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long
    
    If Not MapaValido(Map) Then Exit Sub
    
    EsNuevo = True
    
    'Prevent adding repeated users
    For i = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i
    
    If EsNuevo Then
        'Update map and connection groups data
        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
        TempVal = ConnGroups(Map).CountEntrys
        
        If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
        End If
        
        ConnGroups(Map).UserEntrys(TempVal) = UserIndex
    End If
    
    With UserList(UserIndex)
        'Update user
        .AreasInfo.AreaID = 0
        
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
        .AreasInfo.AreaReciveX = 0
        .AreasInfo.AreaReciveY = 0
    End With
    
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, ButIndex)
End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    With Npclist(NpcIndex)
        .AreasInfo.AreaID = 0
        
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
        .AreasInfo.AreaReciveX = 0
        .AreasInfo.AreaReciveY = 0
    End With
    
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub
