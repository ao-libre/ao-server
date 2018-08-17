Attribute VB_Name = "PathFinding"
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

'#######################################################
'PathFinding Module
'Coded By Gulfas Morgolock
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'
'Ore is an excellent engine for introducing you not only
'to online game programming but also to general
'game programming. I am convinced that Aaron Perkings, creator
'of ORE, did a great work. He made possible that a lot of
'people enjoy for no fee games made with his engine, and
'for me, this is something great.
'
'I'd really like to contribute to this work, and all the
'projects of free ore-based MMORPGs that are on the net.
'
'I did some basic improvements on the AI of the NPCs, I
'added pathfinding, so now, the npcs are able to avoid
'obstacles. I believe that this improvement was essential
'for the engine.
'
'I'd like to see this as my contribution to ORE project,
'I hope that someone finds this source code useful.
'So, please feel free to do whatever you want with my
'pathfinging module.
'
'I'd really appreciate that if you find this source code
'useful you mention my nickname on the credits of your
'program. But there is no obligation ;).
'
'.........................................................
'Note:
'There is a little problem, ORE refers to map arrays in a
'different manner that my pathfinding routines. When I wrote
'these routines, I did it without thinking in ORE, so in my
'program I refer to maps in the usual way I do it.
'
'For example, suppose we have:
'Map(1 to Y,1 to X) as MapBlock
'I usually use the first coordinate as Y, and
'the second one as X.
'
'ORE refers to maps in converse way, for example:
'Map(1 to X,1 to Y) as MapBlock. As you can see the
'roles of first and second coordinates are different
'that my routines
'
'#######################################################


Option Explicit

Private Const ROWS As Integer = 100
Private Const COLUMS As Integer = 100
Private Const MAXINT As Integer = 1000

Private Type tIntermidiateWork
    DistV As Integer
    PrevV As tVertice
End Type

Private TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Limites = ((vcolu >= 1) And (vcolu <= COLUMS) And (vfila >= 1) And (vfila <= ROWS))
End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

With MapData(Map, row, Col)
    IsWalkable = ((.Blocked Or .NpcIndex) = 0)
    
    If .UserIndex <> 0 Then
         If .UserIndex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False
    End If
End With

End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef T() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim V As tVertice
    Dim j As Integer
    
    'Look to North
    j = vfila - 1
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
            With T(j, vcolu)
                'Nos aseguramos que no hay un camino más corto
                If .DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    .DistV = T(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu
                    V.Y = j
                    Call Push(V)
                End If
            End With
        End If
    End If
    
    j = vfila + 1
    'look to south
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
            With T(j, vcolu)
                'Nos aseguramos que no hay un camino más corto
                If .DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    .DistV = T(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu
                    V.Y = j
                    Call Push(V)
                End If
            End With
        End If
    End If
    
    j = vcolu - 1
    'look to west
    If Limites(vfila, j) Then
        If IsWalkable(MapIndex, vfila, j, NpcIndex) Then
            With T(vfila, j)
                'Nos aseguramos que no hay un camino más corto
                If .DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    .DistV = T(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = j
                    V.Y = vfila
                    Call Push(V)
                End If
            End With
        End If
    End If
    
    j = vcolu + 1
    'look to east
    If Limites(vfila, j) Then
        If IsWalkable(MapIndex, vfila, j, NpcIndex) Then
            With T(vfila, j)
                'Nos aseguramos que no hay un camino más corto
                If .DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    .DistV = T(vfila, vcolu).DistV + 1
                    .PrevV.X = vcolu
                    .PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = j
                    V.Y = vfila
                    Call Push(V)
                End If
            End With
        End If
    End If
   
End Sub

Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
'***************************************************
'Author: Unknown
'Last Modification: -
'This Sub seeks a path from the npclist(npcindex).pos
'to the location NPCList(NpcIndex).PFINFO.Target.
'The optional parameter MaxSteps is the maximum of steps
'allowed for the path.
'***************************************************

    Dim cur_npc_pos As tVertice
    Dim tar_npc_pos As tVertice
    Dim V As tVertice
    Dim NpcMap As Integer
    Dim steps As Integer
    
    With Npclist(NpcIndex)
        NpcMap = .Pos.Map
        
        cur_npc_pos.X = .Pos.Y
        cur_npc_pos.Y = .Pos.X
        
        tar_npc_pos.X = .PFINFO.Target.X '  UserList(.PFINFO.TargetUser).Pos.X
        tar_npc_pos.Y = .PFINFO.Target.Y '  UserList(.PFINFO.TargetUser).Pos.Y
        
        Call InitializeTable(TmpArray, cur_npc_pos)
        Call InitQueue
        
        'We add the first vertex to the Queue
        Call Push(cur_npc_pos)
        
        Do While (Not IsEmpty)
            If steps > MaxSteps Then Exit Do
            V = Pop
            If (V.X = tar_npc_pos.X) And (V.Y = tar_npc_pos.Y) Then Exit Do
            Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
        Loop
        
        Call MakePath(NpcIndex)
    End With
End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'Builds the path previously calculated
'***************************************************

    Dim Pasos As Integer
    Dim miV As tVertice
    Dim i As Integer
    
    With Npclist(NpcIndex)
        Pasos = TmpArray(.PFINFO.Target.Y, .PFINFO.Target.X).DistV
        .PFINFO.PathLenght = Pasos
        
        If Pasos = MAXINT Then
            'MsgBox "There is no path."
            .PFINFO.NoPath = True
            .PFINFO.PathLenght = 0
            Exit Sub
        End If
        
        ReDim .PFINFO.Path(1 To Pasos) As tVertice
        
        miV.X = .PFINFO.Target.X
        miV.Y = .PFINFO.Target.Y
        
        For i = Pasos To 1 Step -1
            .PFINFO.Path(i) = miV
            miV = TmpArray(miV.Y, miV.X).PrevV
        Next i
        
        .PFINFO.CurPos = 1
        .PFINFO.NoPath = False
    End With
   
End Sub

Private Sub InitializeTable(ByRef T() As tIntermidiateWork, ByRef S As tVertice, Optional ByVal MaxSteps As Integer = 30)
'***************************************************
'Author: Unknown
'Last Modification: -
'Initialize the array where we calculate the path
'***************************************************

Dim j As Integer, k As Integer
Const anymap = 1

For j = S.Y - MaxSteps To S.Y + MaxSteps
    For k = S.X - MaxSteps To S.X + MaxSteps
        If InMapBounds(anymap, j, k) Then
            With T(j, k)
                .DistV = MAXINT
                .PrevV.X = 0
                .PrevV.Y = 0
            End With
        End If
    Next k
Next j

T(S.Y, S.X).DistV = 0

End Sub
