Attribute VB_Name = "PathFinding"
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
Private Const Walkable As Integer = 0

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Dim TilePosX As Integer, TilePosY As Integer

Dim MyVert As tVertice
Dim MyFin As tVertice

Dim Iter As Integer

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS
End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
IsWalkable = MapData(Map, row, Col).Blocked = 0 And MapData(Map, row, Col).NpcIndex = 0

If MapData(Map, row, Col).UserIndex <> 0 Then
     If MapData(Map, row, Col).UserIndex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False
End If

End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef t() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
    Dim V As tVertice
    Dim j As Integer
    'Look to North
    j = vfila - 1
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                    'Nos aseguramos que no hay un camino más corto
                    If t(j, vcolu).DistV = MAXINT Then
                        'Actualizamos la tabla de calculos intermedios
                        t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
                        t(j, vcolu).PrevV.X = vcolu
                        t(j, vcolu).PrevV.Y = vfila
                        'Mete el vertice en la cola
                        V.X = vcolu
                        V.Y = j
                        Call Push(V)
                    End If
            End If
    End If
    j = vfila + 1
    'look to south
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(j, vcolu).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(j, vcolu).DistV = t(vfila, vcolu).DistV + 1
                    t(j, vcolu).PrevV.X = vcolu
                    t(j, vcolu).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu
                    V.Y = j
                    Call Push(V)
                End If
            End If
    End If
    'look to west
    If Limites(vfila, vcolu - 1) Then
            If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(vfila, vcolu - 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(vfila, vcolu - 1).DistV = t(vfila, vcolu).DistV + 1
                    t(vfila, vcolu - 1).PrevV.X = vcolu
                    t(vfila, vcolu - 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu - 1
                    V.Y = vfila
                    Call Push(V)
                End If
            End If
    End If
    'look to east
    If Limites(vfila, vcolu + 1) Then
            If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If t(vfila, vcolu + 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    t(vfila, vcolu + 1).DistV = t(vfila, vcolu).DistV + 1
                    t(vfila, vcolu + 1).PrevV.X = vcolu
                    t(vfila, vcolu + 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    V.X = vcolu + 1
                    V.Y = vfila
                    Call Push(V)
                End If
            End If
    End If
   
   
End Sub


Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
'############################################################
'This Sub seeks a path from the npclist(npcindex).pos
'to the location NPCList(NpcIndex).PFINFO.Target.
'The optional parameter MaxSteps is the maximum of steps
'allowed for the path.
'############################################################

Dim cur_npc_pos As tVertice
Dim tar_npc_pos As tVertice
Dim V As tVertice
Dim NpcMap As Integer
Dim steps As Integer

NpcMap = Npclist(NpcIndex).Pos.Map

steps = 0

cur_npc_pos.X = Npclist(NpcIndex).Pos.Y
cur_npc_pos.Y = Npclist(NpcIndex).Pos.X

tar_npc_pos.X = Npclist(NpcIndex).PFINFO.Target.X '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y

Call InitializeTable(TmpArray, cur_npc_pos)
Call InitQueue

'We add the first vertex to the Queue
Call Push(cur_npc_pos)

Do While (Not IsEmpty)
    If steps > MaxSteps Then Exit Do
    V = Pop
    If V.X = tar_npc_pos.X And V.Y = tar_npc_pos.Y Then Exit Do
    Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.X, NpcIndex)
Loop

Call MakePath(NpcIndex)

End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
'#######################################################
'Builds the path previously calculated
'#######################################################

Dim Pasos As Integer
Dim miV As tVertice
Dim i As Integer

Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.X).DistV
Npclist(NpcIndex).PFINFO.PathLenght = Pasos


If Pasos = MAXINT Then
    'MsgBox "There is no path."
    Npclist(NpcIndex).PFINFO.NoPath = True
    Npclist(NpcIndex).PFINFO.PathLenght = 0
    Exit Sub
End If

ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice

miV.X = Npclist(NpcIndex).PFINFO.Target.X
miV.Y = Npclist(NpcIndex).PFINFO.Target.Y

For i = Pasos To 1 Step -1
    Npclist(NpcIndex).PFINFO.Path(i) = miV
    miV = TmpArray(miV.Y, miV.X).PrevV
Next i

Npclist(NpcIndex).PFINFO.CurPos = 1
Npclist(NpcIndex).PFINFO.NoPath = False
   
End Sub

Private Sub InitializeTable(ByRef t() As tIntermidiateWork, ByRef s As tVertice, Optional ByVal MaxSteps As Integer = 30)
'#########################################################
'Initialize the array where we calculate the path
'#########################################################

Dim j As Integer, k As Integer
Const anymap = 1
For j = s.Y - MaxSteps To s.Y + MaxSteps
    For k = s.X - MaxSteps To s.X + MaxSteps
        If InMapBounds(anymap, j, k) Then
            t(j, k).Known = False
            t(j, k).DistV = MAXINT
            t(j, k).PrevV.X = 0
            t(j, k).PrevV.Y = 0
        End If
    Next
Next

t(s.Y, s.X).Known = False
t(s.Y, s.X).DistV = 0

End Sub

