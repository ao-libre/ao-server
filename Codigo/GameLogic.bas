Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)

End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)

End Function

Public Function EsGm(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************

    EsGm = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))

End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByVal x As Integer, _
                        ByVal y As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 06/03/2010
    'Handles the Map passage of Users. Allows the existance
    'of exclusive maps for Newbies, Royal Army and Caos Legion members
    'and enables GMs to enter every map without restriction.
    'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
    ' 06/03/2010 : Now we have 5 attemps to not fall into a map change or another teleport while going into a teleport. (Marco)
    '***************************************************

    Dim nPos       As WorldPos

    Dim FxFlag     As Boolean

    Dim TelepRadio As Integer

    Dim DestPos    As WorldPos
    
    On Error GoTo errHandler

    'Controla las salidas
    If InMapBounds(Map, x, y) Then

        With MapData(Map, x, y)

            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
                TelepRadio = ObjData(.ObjInfo.ObjIndex).Radio

            End If
            
            If .TileExit.Map > 0 And .TileExit.Map <= NumMaps Then
                
                ' Es un teleport, entra en una posicion random, acorde al radio (si es 0, es pos fija)
                ' We have 5 attempts to not falling into another teleport or a map exit.. If we get to the fifth attemp,
                ' the teleport will act as if its radius = 0.
                If FxFlag And TelepRadio > 0 Then

                    Dim attemps As Long

                    Dim exitMap As Boolean

                    Do
                        DestPos.x = .TileExit.x + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.y = .TileExit.y + RandomNumber(TelepRadio * (-1), TelepRadio)
                        
                        attemps = attemps + 1
                        
                        exitMap = MapData(.TileExit.Map, DestPos.x, DestPos.y).TileExit.Map > 0 And MapData(.TileExit.Map, DestPos.x, DestPos.y).TileExit.Map <= NumMaps
                    Loop Until (attemps >= 5 Or exitMap = False)
                    
                    If attemps >= 5 Then
                        DestPos.x = .TileExit.x
                        DestPos.y = .TileExit.y

                    End If

                    ' Posicion fija
                Else
                    DestPos.x = .TileExit.x
                    DestPos.y = .TileExit.y

                End If
                
                DestPos.Map = .TileExit.Map
                
                If EsGm(UserIndex) Then
                    Call LogGM(UserList(UserIndex).Name, "Utilizo un teleport hacia el mapa " & DestPos.Map & " (" & DestPos.x & "," & DestPos.y & ")")

                End If
                
                ' Si es un mapa que no admite muertos
                If MapInfo(DestPos.Map).OnDeathGoTo.Map <> 0 Then

                    ' Si esta muerto no puede entrar
                    If UserList(UserIndex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Solo se permite entrar al mapa a los personajes vivos.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                        End If
                        
                        Exit Sub

                    End If

                End If
                
                'Es mapa de newbies?
                If MapInfo(DestPos.Map).Restringir = eRestrict.restrict_newbie Then

                    'El usuario es un newbie?
                    If EsNewbie(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.x, DestPos.y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.x, DestPos.y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.x <> 0 And nPos.y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                            End If

                        End If

                    Else 'No es newbie
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
        
                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, False)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_armada Then 'Es mapa de Armadas?

                    'El usuario es Armada?
                    If esArmada(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.x, DestPos.y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.x, DestPos.y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.x <> 0 And nPos.y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                            End If

                        End If

                    Else 'No es armada
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros del ejercito real.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_caos Then 'Es mapa de Caos?

                    'El usuario es Caos?
                    If esCaos(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.x, DestPos.y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.x, DestPos.y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.x <> 0 And nPos.y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                            End If

                        End If

                    Else 'No es caos
                        Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros de la legion oscura.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                        End If

                    End If

                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_faccion Then 'Es mapa de faccionarios?

                    'El usuario es Armada o Caos?
                    If esArmada(UserIndex) Or esCaos(UserIndex) Or EsGm(UserIndex) Then
                        If LegalPos(DestPos.Map, DestPos.x, DestPos.y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, DestPos.Map, DestPos.x, DestPos.y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)

                            If nPos.x <> 0 And nPos.y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                            End If

                        End If

                    Else 'No es Faccionario
                        Call WriteConsoleMsg(UserIndex, "Solo se permite entrar al mapa si eres miembro de alguna faccion.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                        End If

                    End If

                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.

                    If LegalPos(DestPos.Map, DestPos.x, DestPos.y, PuedeAtravesarAgua(UserIndex)) Then
                        Call WarpUserChar(UserIndex, DestPos.Map, DestPos.x, DestPos.y, FxFlag)
                    Else
                        Call ClosestLegalPos(DestPos, nPos)

                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, FxFlag)

                        End If

                    End If

                End If

                'Te fusite del mapa. La criatura ya no es mas tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(UserIndex).flags.AtacadoPorNpc

                If aN > 0 Then
                    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                    Npclist(aN).flags.AttackedBy = vbNullString

                End If
            
                aN = UserList(UserIndex).flags.NPCAtacado

                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString

                    End If

                End If

                UserList(UserIndex).flags.AtacadoPorNpc = 0
                UserList(UserIndex).flags.NPCAtacado = 0

            End If

        End With

    End If

    Exit Sub

errHandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)

End Sub

Function InRangoVision(ByVal UserIndex As Integer, _
                       ByVal x As Integer, _
                       ByVal y As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If x > UserList(UserIndex).Pos.x - MinXBorder And x < UserList(UserIndex).Pos.x + MinXBorder Then
        If y > UserList(UserIndex).Pos.y - MinYBorder And y < UserList(UserIndex).Pos.y + MinYBorder Then
            InRangoVision = True
            Exit Function

        End If

    End If

    InRangoVision = False

End Function

Public Function InVisionRangeAndMap(ByVal UserIndex As Integer, _
                                    ByRef OtherUserPos As WorldPos) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        ' Same map?
        If .Pos.Map <> OtherUserPos.Map Then Exit Function
    
        ' In x range?
        If OtherUserPos.x < .Pos.x - MinXBorder Or OtherUserPos.x > .Pos.x + MinXBorder Then Exit Function
        
        ' In y range?
        If OtherUserPos.y < .Pos.y - MinYBorder And OtherUserPos.y > .Pos.y + MinYBorder Then Exit Function

    End With

    InVisionRangeAndMap = True
    
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, _
                          x As Integer, _
                          y As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If x > Npclist(NpcIndex).Pos.x - MinXBorder And x < Npclist(NpcIndex).Pos.x + MinXBorder Then
        If y > Npclist(NpcIndex).Pos.y - MinYBorder And y < Npclist(NpcIndex).Pos.y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function

        End If

    End If

    InRangoVisionNPC = False

End Function

Function InMapBounds(ByVal Map As Integer, _
                     ByVal x As Integer, _
                     ByVal y As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If (Map <= 0 Or Map > NumMaps) Or x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True

    End If
    
End Function

Private Function RhombLegalPos(ByRef Pos As WorldPos, _
                               ByRef vX As Long, _
                               ByRef vY As Long, _
                               ByVal Distance As Long, _
                               Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False) As Boolean
    '***************************************************
    'Author: Marco Vanotti (Marco)
    'Last Modification: -
    ' walks all the perimeter of a rhomb of side  "distance + 1",
    ' which starts at Pos.x - Distance and Pos.y
    '***************************************************

    Dim i As Long
    
    vX = Pos.x - Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY - i
            RhombLegalPos = True
            Exit Function

        End If

    Next
    
    vX = Pos.x
    vY = Pos.y - Distance
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY + i
            RhombLegalPos = True
            Exit Function

        End If

    Next
    
    vX = Pos.x + Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY + i
            RhombLegalPos = True
            Exit Function

        End If

    Next
    
    vX = Pos.x
    vY = Pos.y + Distance
    
    For i = 0 To Distance - 1

        If (LegalPos(Pos.Map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY - i
            RhombLegalPos = True
            Exit Function

        End If

    Next
    
    RhombLegalPos = False
    
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, _
                                  ByRef vX As Long, _
                                  ByRef vY As Long, _
                                  ByVal Distance As Long, _
                                  ByVal ObjIndex As Integer, _
                                  ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, _
                                  ByVal PuedeTierra As Boolean) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    ' walks all the perimeter of a rhomb of side  "distance + 1",
    ' which starts at Pos.x - Distance and Pos.y
    ' and searchs for a valid position to drop items
    '***************************************************
    On Error GoTo errHandler

    Dim i           As Long

    Dim HayObj      As Boolean
    
    Dim x           As Integer

    Dim y           As Integer

    Dim MapObjIndex As Integer
    
    vX = Pos.x - Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1
        
        x = vX + i
        y = vY - i
        
        If (LegalPos(Pos.Map, x, y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, x, y, ObjIndex, ObjAmount) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function

            End If
            
        End If

    Next
    
    vX = Pos.x
    vY = Pos.y - Distance
    
    For i = 0 To Distance - 1
        
        x = vX + i
        y = vY + i
        
        If (LegalPos(Pos.Map, x, y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, x, y, ObjIndex, ObjAmount) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function

            End If

        End If

    Next
    
    vX = Pos.x + Distance
    vY = Pos.y
    
    For i = 0 To Distance - 1
        
        x = vX - i
        y = vY + i
    
        If (LegalPos(Pos.Map, x, y, PuedeAgua, PuedeTierra, True)) Then
        
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, x, y, ObjIndex, ObjAmount) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function

            End If

        End If

    Next
    
    vX = Pos.x
    vY = Pos.y + Distance
    
    For i = 0 To Distance - 1
        
        x = vX - i
        y = vY - i
    
        If (LegalPos(Pos.Map, x, y, PuedeAgua, PuedeTierra, True)) Then

            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, x, y, ObjIndex, ObjAmount) Then
                vX = x
                vY = y
                
                RhombLegalTilePos = True
                Exit Function

            End If

        End If

    Next
    
    RhombLegalTilePos = False
    
    Exit Function
    
errHandler:
    Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.description)

End Function

Public Function HayObjeto(ByVal Mapa As Integer, _
                          ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal ObjIndex As Integer, _
                          ByVal ObjAmount As Long) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    'Checks if there's space in a tile to add an itemAmount
    '***************************************************
    Dim MapObjIndex As Integer

    MapObjIndex = MapData(Mapa, x, y).ObjInfo.ObjIndex
            
    ' Hay un objeto tirado?
    If MapObjIndex <> 0 Then

        ' Es el mismo objeto?
        If MapObjIndex = ObjIndex Then
            ' La suma es menor a 10k?
            HayObjeto = (MapData(Mapa, x, y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
        Else
            HayObjeto = True

        End If

    Else
        HayObjeto = False

    End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, _
                    ByRef nPos As WorldPos, _
                    Optional PuedeAgua As Boolean = False, _
                    Optional PuedeTierra As Boolean = True, _
                    Optional ByVal CheckExitTile As Boolean = False)
    '*****************************************************************
    'Author: Unknown (original version)
    'Last Modification: 09/14/2010 (Marco)
    'History:
    ' - 01/24/2007 (ToxicWaste)
    'Encuentra la posicion legal mas cercana y la guarda en nPos
    '*****************************************************************

    Dim Found As Boolean

    Dim LoopC As Integer

    Dim tX    As Long

    Dim tY    As Long
    
    nPos = Pos
    tX = Pos.x
    tY = Pos.y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.Map, nPos.x, nPos.y, PuedeAgua, PuedeTierra, CheckExitTile) Then
        Found = True
    
        ' Busca en las demas posiciones, en forma de "rombo"
    Else

        While (Not Found) And LoopC <= 12

            If RhombLegalPos(Pos, tX, tY, LoopC, PuedeAgua, PuedeTierra, CheckExitTile) Then
                nPos.x = tX
                nPos.y = tY
                Found = True

            End If
        
            LoopC = LoopC + 1
        Wend
        
    End If
    
    If Not Found Then
        nPos.x = 0
        nPos.y = 0

    End If

End Sub

Public Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
    '***************************************************
    'Author: Unknown
    'Last Modification: 09/14/2010
    'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
    '*****************************************************************

    Call ClosestLegalPos(Pos, nPos, , , True)

End Sub

Function NameIndex(ByVal Name As String) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim UserIndex As Long
    
    'Nombre valido?
    If LenB(Name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
                                        
    Name = UCase$(Name)
                                        
    If InStrB(Name, "+") <> 0 Then
        Name = Replace(Name, "+", " ")
    End If
    
    UserIndex = 1
    
     Do Until StrComp(UCase$(UserList(UserIndex).Name), Name) = 0
        
        UserIndex = UserIndex + 1
        
        If UserIndex > MaxUsers Then
            NameIndex = 0
            Exit Function

        End If

    Loop
     
    NameIndex = UserIndex

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers

        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
                CheckForSameIP = True
                Exit Function

            End If

        End If

    Next LoopC
    
    CheckForSameIP = False

End Function

Function CheckForSameName(ByVal Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
                CheckForSameName = True
                Exit Function

            End If

        End If

    Next LoopC
    
    CheckForSameName = False

End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Toma una posicion y se mueve hacia donde esta perfilado
    '*****************************************************************

    Select Case Head

        Case eHeading.NORTH
            Pos.y = Pos.y - 1
        
        Case eHeading.SOUTH
            Pos.y = Pos.y + 1
        
        Case eHeading.EAST
            Pos.x = Pos.x + 1
        
        Case eHeading.WEST
            Pos.x = Pos.x - 1

    End Select

End Sub

Function LegalPos(ByVal Map As Integer, _
                  ByVal x As Integer, _
                  ByVal y As Integer, _
                  Optional ByVal PuedeAgua As Boolean = False, _
                  Optional ByVal PuedeTierra As Boolean = True, _
                  Optional ByVal CheckExitTile As Boolean = False) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Checks if the position is Legal.
    '***************************************************

    'Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
        LegalPos = False
    Else

        With MapData(Map, x, y)

            If PuedeAgua And PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.NpcIndex = 0) And (Not HayAgua(Map, x, y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = (.Blocked <> 1) And (.UserIndex = 0) And (.NpcIndex = 0) And (HayAgua(Map, x, y))
            Else
                LegalPos = False

            End If

        End With
        
        If CheckExitTile Then
            LegalPos = LegalPos And (MapData(Map, x, y).TileExit.Map = 0)

        End If
        
    End If

End Function

Function MoveToLegalPos(ByVal Map As Integer, _
                        ByVal x As Integer, _
                        ByVal y As Integer, _
                        Optional ByVal PuedeAgua As Boolean = False, _
                        Optional ByVal PuedeTierra As Boolean = True) As Boolean
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 13/07/2009
    'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
    '13/07/2009: ZaMa - Now it's also legal move where an invisible admin is.
    '***************************************************

    Dim UserIndex        As Integer

    Dim IsDeadChar       As Boolean

    Dim IsAdminInvisible As Boolean

    'Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
        MoveToLegalPos = False
    Else

        With MapData(Map, x, y)
            UserIndex = .UserIndex
        
            If UserIndex > 0 Then
                IsDeadChar = (UserList(UserIndex).flags.Muerto = 1)
                IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
            Else
                IsDeadChar = False
                IsAdminInvisible = False

            End If
        
            If PuedeAgua And PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (Not HayAgua(Map, x, y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                MoveToLegalPos = (.Blocked <> 1) And (UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (HayAgua(Map, x, y))
            Else
                MoveToLegalPos = False

            End If

        End With

    End If

End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, _
                        ByVal Map As Integer, _
                        ByRef x As Integer, _
                        ByRef y As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 26/03/2009
    'Search for a Legal pos for the user who is being teleported.
    '***************************************************

    If MapData(Map, x, y).UserIndex <> 0 Or MapData(Map, x, y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, x, y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace     As Boolean

        Dim tX             As Long

        Dim tY             As Long

        Dim Rango          As Long

        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = y - Rango To y + Rango
                For tX = x - Rango To x + Rango

                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).UserIndex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                        Exit For

                    End If

                Next tX
        
                If FoundPlace Then Exit For
            Next tY
            
            If FoundPlace Then Exit For
        Next Rango
    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            x = tX
            y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(Map, x, y).UserIndex

            If OtherUserIndex <> 0 Then

                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then

                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)

                    End If

                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")
                        Call FlushBuffer(OtherUserIndex)

                    End If

                End If
            
                Call CloseSocket(OtherUserIndex)

            End If

        End If

    End If

End Sub

Function LegalPosNPC(ByVal Map As Integer, _
                     ByVal x As Integer, _
                     ByVal y As Integer, _
                     ByVal AguaValida As Byte, _
                     Optional ByVal IsPet As Boolean = False) As Boolean

    '***************************************************
    'Autor: Unkwnown
    'Last Modification: 09/23/2009
    'Checks if it's a Legal pos for the npc to move to.
    '09/23/2009: Pato - If UserIndex is a AdminInvisible, then is a legal pos.
    '***************************************************
    Dim IsDeadChar       As Boolean

    Dim UserIndex        As Integer

    Dim IsAdminInvisible As Boolean
    
    If (Map <= 0 Or Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function

    End If

    With MapData(Map, x, y)
        UserIndex = .UserIndex

        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).flags.Muerto = 1
            IsAdminInvisible = (UserList(UserIndex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False

        End If
    
        If AguaValida = 0 Then
            LegalPosNPC = (.Blocked <> 1) And (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.POSINVALIDA Or IsPet) And Not HayAgua(Map, x, y)
        Else
            LegalPosNPC = (.Blocked <> 1) And (.UserIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.POSINVALIDA Or IsPet)

        End If

    End With

End Function

Sub SendHelp(ByVal index As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim NumHelpLines As Integer

    Dim LoopC        As Integer

    NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

    For LoopC = 1 To NumHelpLines
        Call WriteConsoleMsg(index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
    Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If Npclist(NpcIndex).NroExpresiones > 0 Then

        Dim randomi

        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))

    End If

End Sub

Sub LookatTile(ByVal UserIndex As Integer, _
               ByVal Map As Integer, _
               ByVal x As Integer, _
               ByVal y As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 26/03/2009
    '13/02/2009: ZaMa - El nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
    '***************************************************

    On Error GoTo errHandler

    'Responde al click del usuario sobre el mapa
    Dim FoundChar      As Byte

    Dim FoundSomething As Byte

    Dim TempCharIndex  As Integer

    Dim Stat           As String

    Dim ft             As FontTypeNames

    With UserList(UserIndex)

        'Rango Vision? (ToxicWaste)
        If (Abs(.Pos.y - y) > RANGO_VISION_Y) Or (Abs(.Pos.x - x) > RANGO_VISION_X) Then
            Exit Sub

        End If
    
        'Posicion valida?
        If InMapBounds(Map, x, y) Then

            With .flags
                .TargetMap = Map
                .TargetX = x
                .TargetY = y

                'Es un obj?
                If MapData(Map, x, y).ObjInfo.ObjIndex > 0 Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = x
                    .TargetObjY = y
                    FoundSomething = 1
                ElseIf MapData(Map, x + 1, y).ObjInfo.ObjIndex > 0 Then

                    'Informa el nombre
                    If ObjData(MapData(Map, x + 1, y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        .TargetObjMap = Map
                        .TargetObjX = x + 1
                        .TargetObjY = y
                        FoundSomething = 1

                    End If

                ElseIf MapData(Map, x + 1, y + 1).ObjInfo.ObjIndex > 0 Then

                    If ObjData(MapData(Map, x + 1, y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = x + 1
                        .TargetObjY = y + 1
                        FoundSomething = 1

                    End If

                ElseIf MapData(Map, x, y + 1).ObjInfo.ObjIndex > 0 Then

                    If ObjData(MapData(Map, x, y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                        'Informa el nombre
                        .TargetObjMap = Map
                        .TargetObjX = x
                        .TargetObjY = y + 1
                        FoundSomething = 1

                    End If

                End If
            
                If FoundSomething = 1 Then
                    .TargetObj = MapData(Map, .TargetObjX, .TargetObjY).ObjInfo.ObjIndex

                    If MostrarCantidad(.TargetObj) Then
                        Call WriteConsoleMsg(UserIndex, ObjData(.TargetObj).Name & " - " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, ObjData(.TargetObj).Name, FontTypeNames.FONTTYPE_INFO)

                    End If
            
                End If

                'Es un personaje?
                If y + 1 <= YMaxMapSize Then
                    If MapData(Map, x, y + 1).UserIndex > 0 Then
                        TempCharIndex = MapData(Map, x, y + 1).UserIndex
                        FoundChar = 1

                    End If

                    If MapData(Map, x, y + 1).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, x, y + 1).NpcIndex
                        FoundChar = 2

                    End If

                End If

                'Es un personaje?
                If FoundChar = 0 Then
                    If MapData(Map, x, y).UserIndex > 0 Then
                        TempCharIndex = MapData(Map, x, y).UserIndex
                        FoundChar = 1

                    End If

                    If MapData(Map, x, y).NpcIndex > 0 Then
                        TempCharIndex = MapData(Map, x, y).NpcIndex
                        FoundChar = 2

                    End If

                End If

            End With
    
            'Reaccion al personaje
            If FoundChar = 1 Then '  Encontro un Usuario?
                If UserList(TempCharIndex).flags.AdminInvisible = 0 Or .flags.Privilegios And PlayerType.Dios Then

                    With UserList(TempCharIndex)

                        If LenB(.DescRM) = 0 And .showName Then 'No tiene descRM y quiere que se vea su nombre.
                            If EsNewbie(TempCharIndex) Then
                                Stat = " <NEWBIE>"

                            End If
                        
                            If .Faccion.ArmadaReal = 1 Then
                                Stat = Stat & " <Ejercito Real> " & "<" & TituloReal(TempCharIndex) & ">"
                            ElseIf .Faccion.FuerzasCaos = 1 Then
                                Stat = Stat & " <Legion Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"

                            End If
                        
                            If .GuildIndex > 0 Then
                                Stat = Stat & " <" & modGuilds.GuildName(.GuildIndex) & ">"

                            End If
                        
                            If Len(.desc) > 0 Then
                                Stat = "Ves a " & .Name & Stat & " - " & .desc
                            Else
                                Stat = "Ves a " & .Name & Stat

                            End If
                                        
                            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                                Stat = Stat & " [CONSEJO DE BANDERBILL]"
                                ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                                Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                                ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                            Else

                                If Not .flags.Privilegios And PlayerType.User Then
                                    Stat = Stat & " <GAME MASTER>"
                                
                                    ' Elijo el color segun el rango del GM:
                                    ' Dios
                                    If .flags.Privilegios = PlayerType.Dios Then
                                        ft = FontTypeNames.fonttype_dios
                                        ' Gm
                                    ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                                        ft = FontTypeNames.FONTTYPE_GM
                                        ' Conse
                                    ElseIf .flags.Privilegios = PlayerType.Consejero Then
                                        ft = FontTypeNames.FONTTYPE_CONSE
                                        ' Rm o Dsrm
                                    ElseIf .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                                        ft = FontTypeNames.FONTTYPE_EJECUCION

                                    End If
                                
                                ElseIf criminal(TempCharIndex) Then
                                    Stat = Stat & " <CRIMINAL>"
                                    ft = FontTypeNames.FONTTYPE_FIGHT
                                Else
                                    Stat = Stat & " <CIUDADANO>"
                                    ft = FontTypeNames.FONTTYPE_CITIZEN

                                End If

                            End If

                        Else  'Si tiene descRM la muestro siempre.
                            Stat = .DescRM
                            ft = FontTypeNames.FONTTYPE_INFOBOLD

                        End If

                    End With
                
                    If LenB(Stat) > 0 Then
                        Call WriteConsoleMsg(UserIndex, Stat, ft)

                    End If
                
                    FoundSomething = 1
                    .flags.TargetUser = TempCharIndex
                    .flags.TargetNPC = 0
                    .flags.TargetNpcTipo = eNPCType.Comun

                End If

            End If
    
            With .flags

                If FoundChar = 2 Then 'Encontro un NPC?

                    Dim estatus            As String
                    Dim MinHp              As Long
                    Dim MaxHp              As Long
                    Dim SupervivenciaSkill As Byte
                    Dim sDesc              As String
                    Dim TimeParalizado     As String
                
                    MinHp = Npclist(TempCharIndex).Stats.MinHp
                    MaxHp = Npclist(TempCharIndex).Stats.MaxHp
                    SupervivenciaSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia)
                
                    If .Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                        estatus = "(" & MinHp & "/" & MaxHp & ") "
                    Else

                        If .Muerto = 0 Then
                    
                            If SupervivenciaSkill <= 10 Then
                                estatus = "(Dudoso) "
                            
                            ElseIf SupervivenciaSkill <= 20 Then

                                If MinHp < (MaxHp / 2) Then
                                    estatus = "(Herido) "
                                Else
                                    estatus = "(Sano) "

                                End If
                            
                            ElseIf SupervivenciaSkill <= 30 Then

                                If MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Malherido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Herido) "
                                Else
                                    estatus = "(Sano) "

                                End If
                            
                            ElseIf SupervivenciaSkill <= 40 Then

                                If MinHp < (MaxHp * 0.25) Then
                                    estatus = "(Muy malherido) "
                                ElseIf MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Herido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Levemente herido) "
                                Else
                                    estatus = "(Sano) "

                                End If
                            
                            ElseIf SupervivenciaSkill < 60 Then

                                If MinHp < (MaxHp * 0.05) Then
                                    estatus = "(Agonizando) "
                                ElseIf MinHp < (MaxHp * 0.1) Then
                                    estatus = "(Casi muerto) "
                                ElseIf MinHp < (MaxHp * 0.25) Then
                                    estatus = "(Muy Malherido) "
                                ElseIf MinHp < (MaxHp * 0.5) Then
                                    estatus = "(Herido) "
                                ElseIf MinHp < (MaxHp * 0.75) Then
                                    estatus = "(Levemente herido) "
                                ElseIf MinHp < (MaxHp) Then
                                    estatus = "(Sano) "
                                Else
                                    estatus = "(Intacto) "

                                End If

                            Else
                                estatus = "(" & MinHp & "/" & MaxHp & ") "

                            End If

                        End If

                    End If
                    
                    'Lorwik> Tiene 100 skills en supervivencia?
                    If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) = 100 Then

                        'Lorwik> Esta paralizado o inmovilizado? Si lo esta miramos el tiempo que le queda.
                        If Npclist(TempCharIndex).flags.Paralizado = 1 Or Npclist(TempCharIndex).flags.Inmovilizado = 1 Then
                            TimeParalizado = " - Tiempo de paralisis: " & Npclist(TempCharIndex).Contadores.Paralisis & " segundos."
                        End If
                        
                    End If
                    
                    If Len(Npclist(TempCharIndex).Desc) > 1 Then
                        Stat = Npclist(TempCharIndex).Desc
                    
                        'Es el rey o el demonio?
                        If Npclist(TempCharIndex).NPCtype = eNPCType.Noble Then
                            If Npclist(TempCharIndex).flags.Faccion = 0 Then 'Es el Rey.

                                'Si es de la Legion Oscura y usuario comun mostramos el mensaje correspondiente y lo ejecutamos:
                                If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                                    Stat = MENSAJE_REY_CAOS

                                    If .Privilegios And PlayerType.User Then
                                        If .Muerto = 0 Then Call UserDie(UserIndex)

                                    End If

                                ElseIf criminal(UserIndex) Then

                                    'Nos fijamos si es criminal enlistable o no enlistable:
                                    If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Or UserList(UserIndex).Faccion.Reenlistadas > 4 Then 'Es criminal no enlistable.
                                        Stat = MENSAJE_REY_CRIMINAL_NOENLISTABLE
                                    Else 'Es criminal enlistable.
                                        Stat = MENSAJE_REY_CRIMINAL_ENLISTABLE

                                    End If

                                End If

                            Else 'Es el demonio

                                'Si es de la Armada Real y usuario comun mostramos el mensaje correspondiente y lo ejecutamos:
                                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                                    Stat = MENSAJE_DEMONIO_REAL

                                    '
                                    If .Privilegios And PlayerType.User Then
                                        If .Muerto = 0 Then Call UserDie(UserIndex)

                                    End If

                                ElseIf Not criminal(UserIndex) Then

                                    'Nos fijamos si es ciudadano enlistable o no enlistable:
                                    If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Or UserList(UserIndex).Faccion.Reenlistadas > 4 Then 'Es ciudadano no enlistable.
                                        Stat = MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE
                                    Else 'Es ciudadano enlistable.
                                        Stat = MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE

                                    End If

                                End If

                            End If

                        End If
                    
                        'Enviamos el mensaje propiamente dicho:
                        Call WriteChatOverHead(UserIndex, Stat, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                    Else

                        If Npclist(TempCharIndex).MaestroUser > 0 Then
                            Call WriteConsoleMsg(Userindex, estatus & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & TimeParalizado, FontTypeNames.FONTTYPE_INFO)
                        
                        Else
                            Call WriteConsoleMsg(Userindex, estatus & Npclist(TempCharIndex).Name & TimeParalizado, FontTypeNames.FONTTYPE_INFO)
                            
                            If UserList(Userindex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                                Call WriteConsoleMsg(Userindex, "Le pego primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                        
                    End If
                
                    FoundSomething = 1
                    .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                    .TargetNPC = TempCharIndex
                    .TargetUser = 0
                    .TargetObj = 0

                End If
            
                If FoundChar = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0

                End If
            
                '*** NO ENCOTRO NADA ***
                If FoundSomething = 0 Then
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0

                End If

            End With

        Else

            If FoundSomething = 0 Then

                With .flags
                    .TargetNPC = 0
                    .TargetNpcTipo = eNPCType.Comun
                    .TargetUser = 0
                    .TargetObj = 0
                    .TargetObjMap = 0
                    .TargetObjX = 0
                    .TargetObjY = 0

                End With

            End If

        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.description)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'***************************************************
'Author: Unknown
'Last Modification: - Kshamenk 23/2/11
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual o hay un bloqueo
'*****************************************************************
 
    Dim x As Integer
    Dim y As Integer
    
    x = Pos.x - Target.x
    y = Pos.y - Target.y
    
    'NE
    If Sgn(x) = -1 And Sgn(y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'NW
    If Sgn(x) = 1 And Sgn(y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'SW
    If Sgn(x) = 1 And Sgn(y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'SE
    If Sgn(x) = -1 And Sgn(y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'Sur
    If Sgn(x) = 0 And Sgn(y) = -1 Then
        FindDirection = eHeading.SOUTH
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'norte
    If Sgn(x) = 0 And Sgn(y) = 1 Then
        FindDirection = eHeading.NORTH
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'oeste
    If Sgn(x) = 1 And Sgn(y) = 0 Then
        FindDirection = eHeading.WEST
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'este
    If Sgn(x) = -1 And Sgn(y) = 0 Then
        FindDirection = eHeading.EAST
        If Not LegalPos(Pos.Map, Pos.x - Sgn(x), Pos.y - Sgn(y)) Then FindDirection = 0
        Exit Function

    End If
    
    'misma
    If Sgn(x) = 0 And Sgn(y) = 0 Then
        FindDirection = 0
        Exit Function

    End If

End Function

Public Function ItemNoEsDeMapa(ByVal index As Integer, _
                               ByVal bIsExit As Boolean) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With ObjData(index)
        ItemNoEsDeMapa = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And Not (.OBJType = eOBJType.otTeleport And bIsExit)
    
    End With

End Function

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With ObjData(index)
        MostrarCantidad = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And .OBJType <> eOBJType.otTeleport

    End With

End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    EsObjetoFijo = OBJType = eOBJType.otForos Or OBJType = eOBJType.otCarteles Or OBJType = eOBJType.otArboles Or OBJType = eOBJType.otYacimiento

End Function

Public Function RestrictStringToByte(ByRef restrict As String) As Byte
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    restrict = UCase$(restrict)

    Select Case restrict

        Case "NEWBIE"
            RestrictStringToByte = 1
        
        Case "ARMADA"
            RestrictStringToByte = 2
        
        Case "CAOS"
            RestrictStringToByte = 3
        
        Case "FACCION"
            RestrictStringToByte = 4
        
        Case Else
            RestrictStringToByte = 0

    End Select

End Function

Public Function RestrictByteToString(ByVal restrict As Byte) As String

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    Select Case restrict

        Case 1
            RestrictByteToString = "NEWBIE"
        
        Case 2
            RestrictByteToString = "ARMADA"
        
        Case 3
            RestrictByteToString = "CAOS"
        
        Case 4
            RestrictByteToString = "FACCION"
        
        Case 0
            RestrictByteToString = "NO"

    End Select

End Function

Public Function TerrainStringToByte(ByRef restrict As String) As Byte
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    restrict = UCase$(restrict)

    Select Case restrict

        Case "NIEVE"
            TerrainStringToByte = 1
        
        Case "DESIERTO"
            TerrainStringToByte = 2
        
        Case "CIUDAD"
            TerrainStringToByte = 3
        
        Case "CAMPO"
            TerrainStringToByte = 4
        
        Case "DUNGEON"
            TerrainStringToByte = 5
        
        Case Else
            TerrainStringToByte = 0

    End Select

End Function

Public Function TerrainByteToString(ByVal restrict As Byte) As String

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 04/18/2011
    '
    '***************************************************
    Select Case restrict

        Case 1
            TerrainByteToString = "NIEVE"
        
        Case 2
            TerrainByteToString = "DESIERTO"
        
        Case 3
            TerrainByteToString = "CIUDAD"
        
        Case 4
            TerrainByteToString = "CAMPO"
        
        Case 5
            TerrainByteToString = "DUNGEON"
        
        Case 0
            TerrainByteToString = "BOSQUE"

    End Select

End Function
