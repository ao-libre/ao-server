Attribute VB_Name = "Acciones"
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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

'¿Rango Visión? (ToxicWaste)
If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
    Exit Sub
End If

'¿Posicion valida?
If InMapBounds(map, X, Y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
       
    If MapData(map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
        'Set the target NPC
        UserList(UserIndex).flags.TargetNPC = MapData(map, X, Y).NpcIndex
        
        If Npclist(MapData(map, X, Y).NpcIndex).Comercia = 1 Then
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it already in commerce mode??
            If UserList(UserIndex).flags.Comerciando Then
                Exit Sub
            End If
            
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Iniciamos la rutina pa' comerciar.
            Call IniciarComercioNPC(UserIndex)
        
        ElseIf Npclist(MapData(map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it already in commerce mode??
            If UserList(UserIndex).flags.Comerciando Then
                Exit Sub
            End If
            
            If Distancia(Npclist(MapData(map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'A depositar de una
            Call IniciarDeposito(UserIndex)
        
        ElseIf Npclist(MapData(map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or Npclist(MapData(map, X, Y).NpcIndex).NPCtype = eNPCType.ResucitadorNewbie Then
            If Distancia(UserList(UserIndex).Pos, Npclist(MapData(map, X, Y).NpcIndex).Pos) > 10 Then
                Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Revivimos si es necesario
            If UserList(UserIndex).flags.Muerto = 1 And (Npclist(MapData(map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                Call RevivirUsuario(UserIndex)
            End If
            
            If Npclist(MapData(map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                'curamos totalmente
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                Call WriteUpdateUserStats(UserIndex)
            End If
        End If
        
    '¿Es un obj?
    ElseIf MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, X, Y).ObjInfo.ObjIndex
        
        Select Case ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).OBJType
            
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(map, X, Y, UserIndex)
            Case eOBJType.otCarteles 'Es un cartel
                Call AccionParaCartel(map, X, Y, UserIndex)
            Case eOBJType.otForos 'Foro
                Call AccionParaForo(map, X, Y, UserIndex)
            Case eOBJType.otLeña    'Leña
                If MapData(map, X, Y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                    Call AccionParaRamita(map, X, Y, UserIndex)
                End If
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, X + 1, Y).ObjInfo.ObjIndex
        
        Select Case ObjData(MapData(map, X + 1, Y).ObjInfo.ObjIndex).OBJType
            
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(map, X + 1, Y, UserIndex)
            
        End Select
    ElseIf MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex

        Select Case ObjData(MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType
            
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(map, X + 1, Y + 1, UserIndex)
            
        End Select
    ElseIf MapData(map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, X, Y + 1).ObjInfo.ObjIndex

        Select Case ObjData(MapData(map, X, Y + 1).ObjInfo.ObjIndex).OBJType
            
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(map, X, Y + 1, UserIndex)
            
        End Select
    End If
End If

End Sub

Sub AccionParaForo(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.map = map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

'¿Hay mensajes?
Dim f As String, tit As String, men As String, BASE As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    BASE = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To num
        N = FreeFile
        f = BASE & i & ".for"
        Open f For Input Shared As #N
        Input #N, tit
        men = vbNullString
        auxcad = vbNullString
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call WriteAddForumMsg(UserIndex, tit, men)
        
    Next
End If
Call WriteShowForumForm(UserIndex)
End Sub


Sub AccionParaPuerta(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                    
                    'Desbloquea
                    MapData(map, X, Y).Blocked = 0
                    MapData(map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, map, X, Y, 0)
                    Call Bloquear(True, map, X - 1, Y, 0)
                    
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                                
                MapData(map, X, Y).Blocked = 1
                MapData(map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, map, X - 1, Y, 1)
                Call Bloquear(True, map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(map, X, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
    Call WriteShowSignal(UserIndex, MapData(map, X, Y).ObjInfo.ObjIndex)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer

Dim Pos As WorldPos
Pos.map = map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(map).Pk = False Then
    Call WriteConsoleMsg(UserIndex, "En zona segura no puedes hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(UserIndex).Pos.map).Zona <> Ciudad Then
        Obj.ObjIndex = FOGATA
        Obj.amount = 1
        
        Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
        Call MakeObj(Obj, map, X, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.map = map
        Fogatita.X = X
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
    Else
        Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
End If

Call SubirSkill(UserIndex, Supervivencia)

End Sub
