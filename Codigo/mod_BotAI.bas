Attribute VB_Name = "mod_BotAI"
Option Explicit

Private Type tRutaArea
    sX As Byte
    sY As Byte
    eX As Byte
    eY As Byte
End Type

Private Type tRutaMap
    x(1 To 100) As Byte
    y(1 To 100) As Byte
    numMap As Integer
End Type

Private Type tRuta
    AreaDestino As tRutaArea
    AreaOrigen As tRutaArea
    numMapsRuta As Integer
    mapsRuta(1 To 15) As tRutaMap
End Type

Private Enum SpellsEnum
    dardo = 2
    curagraves = 5
    Flecha = 6
    proyectil = 8
    paralizar = 9
    remo = 10
    invi = 14
    torme = 15
    desca = 23
    apoca = 25
End Enum

Private rutas(1 To 10) As tRuta

Private Function tieneSpell(ByVal botindex As Byte, ByVal spindex As Integer) As Boolean
    With Npclist(user_Bot(botindex).npcIndex)
        Dim i As Long
        
        For i = 1 To 35
            If .BotData.stats.UserHechizos(i) = spindex Then tieneSpell = True: Exit Function
        Next i
        
    End With
End Function

Private Function tieneMana(ByVal npcIndex As Integer, ByVal spindex As Integer, ByVal mana As Integer) As Boolean
    If Npclist(npcIndex).BotData.stats.MinMAN >= Hechizos(spindex).ManaRequerido Then
        tieneMana = True
    End If
End Function
Sub BotLanzaUnSpell(ByVal botindex As Byte, ByVal Userindex As Integer, ByVal nSpell As Integer)

    If UserList(Userindex).flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
    
    With Npclist(user_Bot(botindex).npcIndex)
        Dim i As Long
        
        K = RandomNumber(1, Npclist(npcIndex).flags.LanzaSpells)
        
        Call NpcLanzaSpellSobreUser(npcIndex, Userindex, Npclist(npcIndex).Spells(K))
        
    End With
End Sub

Private Sub EligeLanzaHechizos(ByVal botindex As Byte)
    'aqui lanza hechizos. Respetando el intervalo.
    With Npclist(user_Bot(botindex).npcIndex)
        If .Target <= 0 Then Exit Sub
        'Select Case .BotData.Clase
         If .BotData.Clase = eClass.Mage Or .BotData.Clase = eClass.Druid Then
            Call NpcLanzaUnSpell(user_Bot(botindex).npcIndex, .Target)
         ElseIf .BotData.Clase = eClass.Bard Or .BotData.Clase = eClass.Paladin Or .BotData.Clase = eClass.Cleric Or .BotData.Clase = eClass.Assasin Then
            
        End If
        'End Select
    End With
End Sub

Private Sub BotBuscaTarget(ByVal botindex As Byte)
    With Npclist(user_Bot(botindex).npcIndex)
        Dim y As Long, x As Long
        Dim ni As Integer
                    For y = .Pos.y - RANGO_VISION_NPC_Y To .Pos.y + RANGO_VISION_NPC_Y
                        For x = .Pos.x - RANGO_VISION_NPC_Y To .Pos.x + RANGO_VISION_NPC_Y
        
                            If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                                ni = MapData(.Pos.Map, x, y).Userindex
        
                                If ni > 0 Then _
                                    .Target = ni
                            End If
                        Next x
                    Next y
        End With
End Sub
Public Sub bot_Accion(ByVal botindex As Byte)
    'este sub es llamado con una frecuencia de 150 ms
    'Aqui se maneja el modo agite lo que es hechizos, golpes, pociones, trabajo, etc.
    Dim npci As Integer
    With user_Bot(botindex)
        If .online = False Then Exit Sub
        npci = .npcIndex
        Select Case .accion
            Case eBotAccion.quieto
                'nos fijamos si hay usuarios para agitarles
                .accion = eBotAccion.Agite
            Case eBotAccion.Agite
                'lanza hechizos a target
                'se remueve, se cura, toma pociones y lanza hechizos
                With Npclist(.npcIndex)
                    Call BotBuscaTarget(botindex)
                    If .flags.Paralizado > 0 Then
                        If RandomNumber(1, 4) > 2 Then
                            If .BotData.stats.MinMAN >= 300 Then
                                .BotData.stats.MinMAN = .BotData.stats.MinMAN - 300
                                .flags.Paralizado = 0
                                Call SendData(SendTarget.ToNPCArea, user_Bot(botindex).npcIndex, PrepareMessagePalabrasMagicas(10, .Char.CharIndex))
                                
                            End If
                        End If
                    End If
                    
                    Call EligeLanzaHechizos(botindex)
                    
                    If .BotData.stats.MinHp < .BotData.stats.MaxHp Then
                        .BotData.stats.MinHp = .BotData.stats.MinHp + 30
                        '**********aqui falta que le quite la pocion de su inventario
                      '  QuitarObjetos 38, 1, botindex, True
                    End If
                    
                    If .BotData.stats.MinMAN < .BotData.stats.MaxMAN Then
                        .BotData.stats.MinMAN = .BotData.stats.MinMAN + Porcentaje(.BotData.stats.MaxMAN, 4) + .BotData.Lvl / 2 + 40 / .BotData.Lvl
                        '**********aqui falta que le quite la pocion de su inventario
                      '  QuitarObjetos 38, 1, botindex, True
                    End If
                    ia_RandomMoveChar botindex, 0, False
                
                    '
                      '  .stats.MinMAN = .stats.MinMAN + Porcentaje(.stats.MaxMAN, 4) + .stats.ELV \ 2 + 40 / .stats.ELV
                End With
        End Select
    End With
End Sub
Function ia_LegalPos(ByVal x As Byte, ByVal y As Byte, ByVal botindex As Byte, Optional ByVal siguiendoUser As Integer = 0) As Boolean
 
' @designer     :  maTih.-
' @date         :  2012/02/01
' @modificated  :  Esta función ya no trabaja con la pos del npc si no que ahora usa los parámetros.
 
ia_LegalPos = False
 
With MapData(Npclist(user_Bot(botindex).npcIndex).Pos.Map, x, y)
 
     '¿Es un mapa valido?
    If (Npclist(user_Bot(botindex).npcIndex).Pos.Map <= 0 Or Npclist(user_Bot(botindex).npcIndex).Pos.Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then Exit Function
 
     'Tile bloqueado?
     If .Blocked <> 0 Then Exit Function
   
     'Hay un usuario?
     If .Userindex > 0 Then
        'Si no es un adminInvisible entonces nos vamos.
        If UserList(.Userindex).flags.AdminInvisible <> 1 Then Exit Function
    End If
 
    'Hay un NPC?
    If .npcIndex <> 0 Then Exit Function
     
    'Hay un bot?
    'If .BotIndex <> 0 Then Exit Function
    
    'Siguiendo Index?
    If siguiendoUser <> 0 Then
        'Válido para evitar el rango Y pero no su eje X.
        If Abs(y - UserList(siguiendoUser).Pos.y) > RANGO_VISION_Y Then Exit Function
   
        If Abs(x - UserList(siguiendoUser).Pos.x) > RANGO_VISION_X Then Exit Function
    End If
    
     ia_LegalPos = True
   
End With
 
End Function

Sub ia_RandomMoveChar(ByVal botindex As Byte, ByVal siguiendoIndex As Integer, ByRef HError As Boolean)
 
' @designer     :  maTih.-
' @date         :  2012/02/01
 
With user_Bot(botindex)
 
    Dim nRandom     As Byte
   
    '25% De probabilidades de moverse a
    'cualquiera de las cuatro direcciones.
   
    nRandom = RandomNumber(1, 4)
   
    Select Case nRandom
   
           Case 1
           
                If ia_LegalPos(Npclist(.npcIndex).Pos.x + 1, Npclist(.npcIndex).Pos.y, botindex, siguiendoIndex) = False Then HError = True: Exit Sub
                
                'Borro el BotIndex del tile anterior.
                
                 Call MoveNPCChar(user_Bot(botindex).npcIndex, eHeading.EAST)
                
           Case 2
           
                If ia_LegalPos(Npclist(.npcIndex).Pos.x - 1, Npclist(.npcIndex).Pos.y, botindex, siguiendoIndex) = False Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                 Call MoveNPCChar(user_Bot(botindex).npcIndex, eHeading.WEST)
           
           Case 3
           
                If ia_LegalPos(Npclist(.npcIndex).Pos.x, Npclist(.npcIndex).Pos.y + 1, botindex, siguiendoIndex) = False Then HError = True: Exit Sub
           
                'Borro el BotIndex del tile anterior.
                 Call MoveNPCChar(user_Bot(botindex).npcIndex, eHeading.SOUTH)
           
           Case 4
                
                If ia_LegalPos(Npclist(.npcIndex).Pos.x, Npclist(.npcIndex).Pos.y - 1, botindex, siguiendoIndex) = False Then HError = True: Exit Sub
                
                'Borro el BotIndex del tile anterior.
                'MapData(.Pos.Map, .Pos.X, .Pos.Y).npcIndex = 0
                Call MoveNPCChar(user_Bot(botindex).npcIndex, eHeading.NORTH)
              '  .Pos.Y = .Pos.Y - 1
   
    End Select
    
    'MapData(.Pos.Map, .Pos.X, .Pos.Y).npcIndex = user_Bot(BotIndex).npcIndex
    
End With
 
End Sub

Private Function getNextPos(ByRef Pos As WorldPos, ByVal ruta As Byte, ByVal mapViaje As Byte) As Position
    Dim x As Long, y As Long
    For x = -1 To x = 1
        For y = -1 To y = 1
            If rutas(ruta).mapsRuta(mapViaje).x(Pos.x + x) = 1 Then 'Se encontro la ruta, hay que ver la forma de saber de donde viene y
                                                                                                            ' hacia donde va.
                
            End If
        Next y
    Next x
End Function

Public Sub bot_Camina(ByVal botindex As Byte)
    'este sub es  llamado con la misma frecuencia que el Walk
    Dim acPos As WorldPos, nPos As Position, npci As Integer, cmap As Byte, cRuta As Byte
    npci = user_Bot(botindex).npcIndex
    With Npclist(npci)
        acPos = .Pos
        cmap = user_Bot(botindex).curMapViaje
        cRuta = user_Bot(botindex).curRutaViaje
        
        nPos = getNextPos(acPos, cRuta, cmap)
    End With
End Sub

Private Function BotFaccionPuedeAtacar(ByVal toBot As Boolean, ByVal botindex As Byte, _
                                                                 ByVal Userindex As Integer, ByVal botindextarget As Byte) As Boolean
                                                                 
    Dim EsCriminal As Boolean
    
    EsCriminal = criminal(botindex, True)
    
    If toBot Then 'target es bot
        If EsCriminal And botEsLegion(botindex) = False Then BotFaccionPuedeAtacar = True: Exit Function
        
        If EsCriminal = False And criminal(botindextarget, True) = True Then BotFaccionPuedeAtacar = True
        If EsCriminal = False And criminal(botindextarget, True) = False Then
            'SON LOS 2 CIUDAS
            If botEsArmada(botindex) = True Then Exit Function
            '*******************POR AHORA NO SE VUELVEN CRIMINALES LOS BOTS.
            'If RandomNumber(1, 5) = 1 Then '20% de posibilidad de querer volverse criminal _
                'ToogleToAtackable
                
        End If
        
        If botEsLegion(botindex) Then
            If Not botEsLegion(botindextarget) Then
                BotFaccionPuedeAtacar = True
            End If
        End If
    Else 'target es user
    
    End If
        
End Function

Private Function botEsLegion(ByVal botindex As Byte) As Boolean
    botEsLegion = (Npclist(user_Bot(botindex).npcIndex).BotData.faccion.FuerzasCaos = 1)
End Function

Private Function botEsArmada(ByVal botindex As Byte) As Boolean
    botEsArmada = (Npclist(user_Bot(botindex).npcIndex).BotData.faccion.ArmadaReal = 1)
End Function

Private Function buscarTargetBot(ByVal botindex As Byte, ByRef FindBot As Boolean) As Integer ' integer x si es user.

    'Aqui chequeamos si el usuario/bot con el que estaba peleando se escapo a mas de 3 tiles fuera del area de vision.
    ' y si no tiene otro usuario con el que pelear en su area de vision + 3 tiles
    Dim x As Long, y As Long, npci As Integer
    npci = user_Bot(botindex).npcIndex
    
    For x = -3 To RANGO_VISION_X + 3
        For y = -3 To RANGO_VISION_Y + 3
            If MapData(Npclist(npci).Pos.x + x, Npclist(npci).Pos.y + y).Userindex > 0 Then
                'encontro un usuario
                'CHORIZO DE CODIGO
                buscarTargetBot = MapData(Npclist(npci).Pos.x + x, Npclist(npci).Pos.y + y).Userindex
                Exit Function
            End If
            If MapData(Npclist(npci).Pos.x + x, Npclist(npci).Pos.y + y).npcIndex > 0 Then
                If Npclist(MapData(Npclist(npci).Pos.x + x, Npclist(npci).Pos.y + y).npcIndex).esBot = True Then
                    'encontro un bot
                    'CHORIZO DE CODIGO
                    FindBot = True
                    buscarTargetBot = Npclist(MapData(Npclist(npci).Pos.x + x, Npclist(npci).Pos.y + y).npcIndex).BotData.botindex
                    Exit Function
                End If
            End If
        Next y
    Next x
        
End Function

Public Sub bot_CheckAcciones()
    'Este sub es llamado con frecuencia de 1 segundo para verificar las acciones de los bots.
    Dim i As Long
    
    For i = 1 To maxBots
        With user_Bot(i)
            If .online = True Then
                Select Case .accion
                    Case eBotAccion.Viajando
                        
                    Case eBotAccion.Agite
                        'chequeamos si el contrincante con el que estaba atacando se fue del area de vision(Le damos una yapa de 3 tiles.)
                        'Si se fue, entonces volvemos al estado anterior.
                        '************* IMPORTANTE:: al agregar modo viaje, importante que al finalizar el agite siga viajando hacia su destino.
                    
                        
                       ' If checkAccionAgite(i) = False Then .accion = eBotAccion.quieto
                        
                End Select
            End If
        End With
    Next i
    
End Sub
