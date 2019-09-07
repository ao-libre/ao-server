Attribute VB_Name = "Acciones"
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

Sub Accion(ByVal Userindex As Integer, _
           ByVal Map As Integer, _
           ByVal x As Integer, _
           ByVal Y As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim tempIndex As Integer
    
    On Error Resume Next

    'Rango Vision? (ToxicWaste)
    If (Abs(UserList(Userindex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(Userindex).Pos.x - x) > RANGO_VISION_X) Then
        Exit Sub

    End If
    
    'Posicion valida?
    If InMapBounds(Map, x, Y) Then

        With UserList(Userindex)

            If MapData(Map, x, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(Map, x, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then

                    'Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                        Exit Sub

                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub

                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(Userindex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then

                    'Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                        Exit Sub

                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub

                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    'A depositar de una
                    Call IniciarDeposito(Userindex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then

                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex)) Then
                        Call RevivirUsuario(Userindex)

                    End If
                    
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex) Then
                        'TODO: ya hay una funcion que hace esto, la del comando /Curar habria que refactorizar
                        'curamos totalmente
                        .Stats.MinHp = .Stats.MaxHp
                        
                        Call WriteConsoleMsg(Userindex, "Te has curado!", FontTypeNames.FONTTYPE_INFO)
                        
                        If .flags.Envenenado = 1 Then
                            'curamos veneno
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(Userindex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)

                        End If
                        
                        Call WriteUpdateUserStats(Userindex)

                    End If

                End If
                
                'Es un obj?
            ElseIf MapData(Map, x, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, x, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, x, Y, Userindex)

                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(Map, x, Y, Userindex)

                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(Map, x, Y, Userindex)

                    Case eOBJType.otLena    'Lena

                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(Map, x, Y, Userindex)

                        End If

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(Map, x + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, x + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, x + 1, Y, Userindex)
                    
                End Select
            
            ElseIf MapData(Map, x + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, x + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, x + 1, Y + 1, Userindex)

                End Select
            
            ElseIf MapData(Map, x, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, x, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType

                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, x, Y + 1, Userindex)

                End Select

            End If

        End With

    End If

End Sub

Public Sub AccionParaForo(ByVal Map As Integer, _
                          ByVal x As Integer, _
                          ByVal Y As Integer, _
                          ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/01/2010
    '02/01/2010: ZaMa - Agrego foros faccionarios
    '***************************************************

    On Error Resume Next

    Dim Pos As WorldPos
    
    Pos.Map = Map
    Pos.x = x
    Pos.Y = Y
    
    If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    If SendPosts(Userindex, ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).ForoID) Then
        Call WriteShowForumForm(Userindex)

    End If
    
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, _
                     ByVal x As Integer, _
                     ByVal Y As Integer, _
                     ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    If Not (Distance(UserList(Userindex).Pos.x, UserList(Userindex).Pos.Y, x, Y) > 2) Then
        If ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).Llave = 0 Then
            If ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).Cerrada = 1 Then

                'Abre la puerta
                If ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, x, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(Map, x, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).GrhIndex, x, Y))
                    
                    'Desbloquea
                    MapData(Map, x, Y).Blocked = 0
                    MapData(Map, x - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, Map, x, Y, 0)
                    Call Bloquear(True, Map, x - 1, Y, 0)
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, x, Y))
                    
                Else
                    Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                'Cierra puerta
                MapData(Map, x, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(Map, x, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).GrhIndex, x, Y))
                                
                MapData(Map, x, Y).Blocked = 1
                MapData(Map, x - 1, Y).Blocked = 1
                
                Call Bloquear(True, Map, x - 1, Y, 1)
                Call Bloquear(True, Map, x, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, x, Y))

            End If
        
            UserList(Userindex).flags.TargetObj = MapData(Map, x, Y).ObjInfo.ObjIndex
        Else
            Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

        End If

    Else
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, _
                     ByVal x As Integer, _
                     ByVal Y As Integer, _
                     ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    If ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
        If Len(ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).texto) > 0 Then
            Call WriteShowSignal(Userindex, MapData(Map, x, Y).ObjInfo.ObjIndex)

        End If
  
    End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, _
                     ByVal x As Integer, _
                     ByVal Y As Integer, _
                     ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim Suerte             As Byte

    Dim exito              As Byte

    Dim obj                As obj

    Dim SkillSupervivencia As Byte

    Dim Pos                As WorldPos
    
    With Pos
        .Map = Map
        .X = X
        .Y = Y
    End With
    

    With UserList(Userindex)

        If Distancia(Pos, .Pos) > 2 Then
            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        If MapData(Map, x, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
            Call WriteConsoleMsg(Userindex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        SkillSupervivencia = .Stats.UserSkills(eSkill.Supervivencia)
    
        If SkillSupervivencia < 6 Then
            Suerte = 3
        
        ElseIf SkillSupervivencia <= 10 Then
            Suerte = 2
        
        Else
            Suerte = 1

        End If
    
        exito = RandomNumber(1, Suerte)
    
        If exito = 1 Then
            If MapInfo(.Pos.Map).Zona <> Ciudad Then
            
                With obj
                    .ObjIndex = FOGATA
                    .Amount = 1
                End With
            
                Call WriteConsoleMsg(Userindex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
            
                Call MakeObj(obj, Map, x, Y)
            
                Call mLimpieza.AgregarObjetoLimpieza(Pos)
            
                Call SubirSkill(Userindex, eSkill.Supervivencia, True)
            Else
                Call WriteConsoleMsg(Userindex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        Else
            Call WriteConsoleMsg(Userindex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
            Call SubirSkill(Userindex, eSkill.Supervivencia, False)

        End If

    End With

End Sub

Public Sub AccionParaSacerdote(ByVal UserIndex As Integer)

    '******************************
    'Adaptacion a 13.0: Kaneidra
    'Last Modification: 15/05/2012
    '******************************
    
    With UserList(UserIndex)
        
        ' Si esta muerto...
        If .flags.Muerto = 1 Then
            
            ' Lo resucitamos.
            Call RevivirUsuario(UserIndex)
            
            ' Restauramos su mana.
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteUpdateMana(UserIndex)
            
            ' Lo curamos.
            .Stats.MinHp = .Stats.MaxHp
            Call WriteUpdateHP(UserIndex)
            
            ' Le avisamos.
            Call WriteConsoleMsg(UserIndex, "El sacerdote te ha resucitado y curado.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        ' Si esta herido... lo curamos.
        If .Stats.MinHp < .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
            Call WriteUpdateHP(UserIndex)
            Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        ' Curamos su envenenamiento.
        If .flags.Envenenado = 1 Then .flags.Envenenado = 0
        
        ' Sacamos la maldicion.
        If .flags.Maldicion = 1 Then .flags.Maldicion = 0
        
        ' Sacamos la ceguera.
        If .flags.Ceguera = 1 Then .flags.Ceguera = 0

    End With
 
End Sub
