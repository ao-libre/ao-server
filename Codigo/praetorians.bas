Attribute VB_Name = "PraetoriansCoopNPC"
''**************************************************************
'' PraetoriansCoopNPC.bas - Handles the Praeorians NPCs.
''
'' Implemented by Mariano Barrou (El Oso)
''**************************************************************
'
''**************************************************************************
''This program is free software; you can redistribute it and/or modify
''it under the terms of the Affero General Public License;
''either version 1 of the License, or any later version.
''
''This program is distributed in the hope that it will be useful,
''but WITHOUT ANY WARRANTY; without even the implied warranty of
''MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''Affero General Public License for more details.
''
''You should have received a copy of the Affero General Public License
''along with this program; if not, you can find it at http://www.affero.org/oagpl.html
''**************************************************************************
'
'Option Explicit
''''''''''''''''''''''''''''''''''''''''''
''' DECLARACIONES DEL MODULO PRETORIANO ''
''''''''''''''''''''''''''''''''''''''''''
''' Estas constantes definen que valores tienen
''' los NPCs pretorianos en el NPC-HOSTILES.DAT
''' Son FIJAS, pero se podria hacer una rutina que
''' las lea desde el npcshostiles.dat
'Public Const PRCLER_NPC As Integer = 900   ''"Sacerdote Pretoriano"
'Public Const PRGUER_NPC As Integer = 901   ''"Guerrero  Pretoriano"
'Public Const PRMAGO_NPC As Integer = 902   ''"Mago Pretoriano"
'Public Const PRCAZA_NPC As Integer = 903   ''"Cazador Pretoriano"
'Public Const PRKING_NPC As Integer = 904   ''"Rey Pretoriano"
'
'
'' 1 rey.
'' 3 guerres.
'' 1 caza.
'' 1 mago.
'' 2 clerigos.
'Public Const NRO_PRETORIANOS As Integer = 8
'
'' Contiene los index de los miembros del clan.
'Public Pretorianos(1 To NRO_PRETORIANOS) As Integer
'
'
''''''''''''''''''''''''''''''''''''''''''''''
''Esta constante identifica en que mapa esta
''la fortaleza pretoriana (no es lo mismo de
''donde estan los NPCs!).
''Se extrae el dato del server.ini en sub LoadSIni
Public MAPA_PRETORIANO As Integer
''''''''''''''''''''''''''''''''''''''''''''''
''Estos numeros son necesarios por cuestiones de
''sonido. Son los numeros de los wavs del cliente.
Public Const SONIDO_DRAGON_VIVO As Integer = 30
'''ALCOBAS REALES
'''OJO LOS BICHOS TAN HARDCODEADOS, NO CAMBIAR EL MAPA DONDE
'''ESTÁN UBICADOS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'''MUCHO MENOS LA COORDENADA Y DE LAS ALCOBAS YA QUE DEBE SER LA MISMA!!!
'''(HAY FUNCIONES Q CUENTAN CON QUE ES LA MISMA!)
'Public Const ALCOBA1_X As Integer = 35
'Public Const ALCOBA1_Y As Integer = 25
'Public Const ALCOBA2_X As Integer = 67
'Public Const ALCOBA2_Y As Integer = 25

Public Enum ePretorianAI
    King = 1
    Healer
    SpellCaster
    SwordMaster
    Shooter
    Thief
    Last
End Enum

' Contains all the pretorian's combinations, and its the offsets
Public PretorianAIOffset(1 To 7) As Integer
Public PretorianDatNumbers() As Integer
'
''Added by Nacho
''Cuantos pretorianos vivos quedan. Uno por cada alcoba
'Public pretorianosVivos As Integer
'

Public Sub LoadPretorianData()

    Dim PretorianDat As String
    PretorianDat = DatPath & "Pretorianos.dat"

    Dim NroCombinaciones As Integer
    NroCombinaciones = val(GetVar(PretorianDat, "MAIN", "Combinaciones"))

    ReDim PretorianDatNumbers(1 To NroCombinaciones)

    Dim TempInt As Integer
    Dim Counter As Long
    Dim PretorianIndex As Integer

    PretorianIndex = 1

    ' KINGS
    TempInt = val(GetVar(PretorianDat, "KING", "Cantidad"))
    PretorianAIOffset(ePretorianAI.King) = 1
    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "KING", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "KING", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' HEALERS
    TempInt = val(GetVar(PretorianDat, "HEALER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Healer) = PretorianIndex
    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "HEALER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "HEALER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' SPELLCASTER
    TempInt = val(GetVar(PretorianDat, "SPELLCASTER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.SpellCaster) = PretorianIndex
    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SPELLCASTER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SPELLCASTER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' SWORDSWINGER
    TempInt = val(GetVar(PretorianDat, "SWORDSWINGER", "Cantidad"))
    PretorianAIOffset(ePretorianAI.SwordMaster) = PretorianIndex
    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SWORDSWINGER", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "SWORDSWINGER", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' LONGRANGE
    TempInt = val(GetVar(PretorianDat, "LONGRANGE", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Shooter) = PretorianIndex
    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "LONGRANGE", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "LONGRANGE", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' THIEF
    TempInt = val(GetVar(PretorianDat, "THIEF", "Cantidad"))
    PretorianAIOffset(ePretorianAI.Thief) = PretorianIndex
    For Counter = 1 To TempInt

        ' Alto
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "THIEF", "Alto" & Counter))
        PretorianIndex = PretorianIndex + 1
        ' Bajo
        PretorianDatNumbers(PretorianIndex) = val(GetVar(PretorianDat, "THIEF", "Bajo" & Counter))
        PretorianIndex = PretorianIndex + 1

    Next Counter

    ' Last
    PretorianAIOffset(ePretorianAI.Last) = PretorianIndex

    ' Inicializa los clanes pretorianos
    ReDim ClanPretoriano(1 To 2) As clsClanPretoriano
    Set ClanPretoriano(1) = New clsClanPretoriano ' Clan default
    Set ClanPretoriano(2) = New clsClanPretoriano ' Invocable por gms

End Sub

















'
''/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
''/\/\/\/\/\/\/\/\ MODULO DE COMBATE PRETORIANO /\/\/\/\/\/\/\/\/\
''/\/\/\/\/\/\/\/\ (NPCS COOPERATIVOS TIPO CLAN)/\/\/\/\/\/\/\/\/\
''/\/\/\/\/\/\/\/\         por EL OSO           /\/\/\/\/\/\/\/\/\
''/\/\/\/\/\/\/\/\       mbarrou@dc.uba.ar      /\/\/\/\/\/\/\/\/\
''/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'
'
'Public Function esPretoriano(ByVal NpcIndex As Integer) As Boolean
'On Error GoTo ErrHandler
'
'    esPretoriano = (Npclist(NpcIndex).Numero <= 925 And Npclist(NpcIndex).Numero >= 900)
'    Exit Function
'
'    Select Case Npclist(NpcIndex).Numero
'        Case PRCLER_NPC
'            esPretoriano = 1
'        Case PRMAGO_NPC
'            esPretoriano = 2
'        Case PRCAZA_NPC
'            esPretoriano = 3
'        Case PRKING_NPC
'            esPretoriano = 4
'        Case PRGUER_NPC
'            esPretoriano = 5
'    End Select
'Exit Function
'
'ErrHandler:
'    LogError ("Error en NPCAI.EsPretoriano? " & Npclist(NpcIndex).Name)
'End Function
'
'Public Sub CrearClanPretoriano(ByVal X As Integer)
''********************************************************
''Author: EL OSO
''Inicializa el clan Pretoriano.
''Last Modify Date: 22/6/06: (Nacho) Seteamos cuantos NPCs creamos
''********************************************************
'On Error GoTo ErrHandler
'
'    ''------------------------------------------------------
'    ''recibe el X,Y donde EL REY ANTERIOR ESTABA POSICIONADO.
'    ''------------------------------------------------------
'    ''35,25 y 67,25 son las posiciones del rey
'
'    ''Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'    ''Public Const PRCLER_NPC = 900
'    ''Public Const PRGUER_NPC = 901
'    ''Public Const PRMAGO_NPC = 902
'    ''Public Const PRCAZA_NPC = 903
'    ''Public Const PRKING_NPC = 904
'    Dim wp As WorldPos
'    Dim wp2 As WorldPos
'    Dim TeleFrag As Integer
'    Dim PretoIndex As Integer
'    Dim NpcIndex As Integer
'
'    wp.Map = MAPA_PRETORIANO
'    If X < 50 Then ''forma burda de ver que alcoba es
'        wp.X = ALCOBA2_X
'        wp.Y = ALCOBA2_Y
'    Else
'        wp.X = ALCOBA1_X
'        wp.Y = ALCOBA1_Y
'    End If
'
'    pretorianosVivos = 7 'Hay 7 + el Rey.
'    TeleFrag = MapData(wp.Map, wp.X, wp.Y).NpcIndex
'
'    If TeleFrag > 0 Then
'        ''El rey va a pisar a un npc de antiguo rey
'        ''Obtengo en WP2 la mejor posicion cercana
'        Call ClosestLegalPos(wp, wp2)
'        If wp2.X <> 0 And wp2.Y <> 0 Then
'            ''mover al actual
'
'            Call SendData(SendTarget.ToNPCArea, TeleFrag, PrepareMessageCharacterMove(Npclist(TeleFrag).Char.CharIndex, wp2.X, wp2.Y))
'            'Update map and user pos
'            MapData(wp.Map, wp.X, wp.Y).NpcIndex = 0
'            Npclist(TeleFrag).Pos = wp2
'            MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex = TeleFrag
'        Else
'            ''TELEFRAG!!!
'            Call QuitarNPC(TeleFrag)
'        End If
'    End If
'
'    ''ya limpié el lugar para el rey (wp)
'    ''Los otros no necesitan este caso ya que respawnan lejos
'
'    'Busco la posicion legal mas cercana aca, aun que creo que tendría que ir en el crearnpc. (NicoNZ)
'
'    ' REY
'    Call ClosestLegalPos(wp, wp2, False, True)
'    NpcIndex = CrearNPC(PRKING_NPC, MAPA_PRETORIANO, wp2)
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'
'    ' CLERIGO
'    wp.X = wp.X + 3
'    Call ClosestLegalPos(wp, wp2, False, True)
'    NpcIndex = CrearNPC(PRCLER_NPC, MAPA_PRETORIANO, wp2)
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'
'    ' CLERIGO
'    wp.X = wp.X - 6
'    Call ClosestLegalPos(wp, wp2, False, True)
'    NpcIndex = CrearNPC(PRCLER_NPC, MAPA_PRETORIANO, wp2)
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'
'    ' GUERRE
'    wp.Y = wp.Y + 3
'    Call ClosestLegalPos(wp, wp2, False, True)
'    NpcIndex = CrearNPC(PRGUER_NPC, MAPA_PRETORIANO, wp2)
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'
'    ' GUERRE
'    wp.X = wp.X + 3
'    Call ClosestLegalPos(wp, wp2, False, True)
'    NpcIndex = CrearNPC(PRGUER_NPC, MAPA_PRETORIANO, wp2)
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'
'    ' GUERRE
'    wp.X = wp.X + 3
'    Call ClosestLegalPos(wp, wp2, False, True)
'    NpcIndex = CrearNPC(PRGUER_NPC, MAPA_PRETORIANO, wp2)
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'
'    ' KAZA
'    wp.Y = wp.Y - 6
'    wp.X = wp.X - 1
'    Call ClosestLegalPos(wp, wp2, False, True)
'    NpcIndex = CrearNPC(PRCAZA_NPC, MAPA_PRETORIANO, wp2)
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'ErrHandler:
'End Sub
'
'Sub PRREY_AI(ByVal npcind As Integer)
'On Error GoTo errorh
'    'HECHIZOS: NO CAMBIAR ACA
'    'REPRESENTAN LA UBICACION DE LOS SPELLS EN NPC_HOSTILES.DAT y si se los puede cambiar en ese archivo
'    '1- CURAR_LEVES 'NO MODIFICABLE
'    '2- REMOVER PARALISIS 'NO MODIFICABLE
'    '3- CEUGERA - 'NO MODIFICABLE
'    '4- ESTUPIDEZ - 'NO MODIFICABLE
'    '5- CURARVENENO - 'NO MODIFICABLE
'    Dim DAT_CURARLEVES As Integer
'    Dim DAT_REMUEVEPARALISIS As Integer
'    Dim DAT_CEGUERA As Integer
'    Dim DAT_ESTUPIDEZ As Integer
'    Dim DAT_CURARVENENO As Integer
'    DAT_CURARLEVES = 1
'    DAT_REMUEVEPARALISIS = 2
'    DAT_CEGUERA = 3
'    DAT_ESTUPIDEZ = 4
'    DAT_CURARVENENO = 5
'
'
'    Dim UI As Integer
'    Dim X As Integer
'    Dim Y As Integer
'    Dim NPCPosX As Integer
'    Dim NPCPosY As Integer
'    Dim NPCPosM As Integer
'    Dim NPCAlInd As Integer
'    Dim PJEnInd As Integer
'    Dim BestTarget As Integer
'    Dim distBestTarget As Integer
'    Dim dist As Integer
'    Dim e_p As Integer
'    Dim hayPretorianos As Boolean
'    Dim headingloop As Byte
'    Dim nPos As WorldPos
'    ''Dim quehacer As Integer
'        ''1- remueve paralisis con un minimo % de efecto
'        ''2- remueve veneno
'        ''3- cura
'
'    NPCPosM = Npclist(npcind).Pos.Map
'    NPCPosX = Npclist(npcind).Pos.X
'    NPCPosY = Npclist(npcind).Pos.Y
'    BestTarget = 0
'    distBestTarget = 0
'    hayPretorianos = False
'
'    'pick the best target according to the following criteria:
'    'King won't fight. Since praetorians' mission is to keep him alive
'    'he will stay as far as possible from combat environment, but close enought
'    'as to aid his loyal army.
'    'If his army has been annihilated, the king will pick the
'    'closest enemy an chase it using his special 'weapon speedhack' ability
'    For X = NPCPosX - 8 To NPCPosX + 8
'        For Y = NPCPosY - 7 To NPCPosY + 7
'            'scan combat field
'            NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex
'            PJEnInd = MapData(NPCPosM, X, Y).UserIndex
'            If (Npclist(npcind).CanAttack = 1) Then   ''saltea el analisis si no puede atacar para evitar cuentas
'                If (NPCAlInd > 0) Then
'                    e_p = esPretoriano(NPCAlInd)
'                    If e_p > 0 And e_p < 6 And (Not (NPCAlInd = npcind)) Then
'                        hayPretorianos = True
'
'                        'Me curo mientras haya pretorianos (no es lo ideal, debería no dar experiencia tampoco, pero por ahora es lo que hay)
'                        Npclist(npcind).Stats.MinHp = Npclist(npcind).Stats.MaxHp
'                    End If
'
'                    If (Npclist(NPCAlInd).flags.Paralizado = 1 And e_p > 0 And e_p < 6) Then
'                        ''el rey puede desparalizar con una efectividad del 20%
'                        If (RandomNumber(1, 100) < 21) Then
'                            Call NPCRemueveParalisisNPC(npcind, NPCAlInd, DAT_REMUEVEPARALISIS)
'                            Npclist(npcind).CanAttack = 0
'                            Exit Sub
'                        End If
'
'                    ''failed to remove
'                    ElseIf (Npclist(NPCAlInd).flags.Envenenado = 1) Then    ''un chiche :D
'                        If esPretoriano(NPCAlInd) Then
'                            Call NPCRemueveVenenoNPC(npcind, NPCAlInd, DAT_CURARVENENO)
'                            Npclist(npcind).CanAttack = 0
'                            Exit Sub
'                        End If
'                    ElseIf (Npclist(NPCAlInd).Stats.MaxHp > Npclist(NPCAlInd).Stats.MinHp) Then
'                        If esPretoriano(NPCAlInd) And Not (NPCAlInd = npcind) Then
'                            ''cura, salvo q sea yo mismo. Eso lo hace 'despues'
'                            Call NPCCuraLevesNPC(npcind, NPCAlInd, DAT_CURARLEVES)
'                            Npclist(npcind).CanAttack = 0
'                            ''Exit Sub
'                        End If
'                    End If
'                End If
'
'                If PJEnInd > 0 And Not hayPretorianos Then
'                    If Not (UserList(PJEnInd).flags.Muerto = 1 Or UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1 Or UserList(PJEnInd).flags.Ceguera = 1) And UserList(PJEnInd).flags.AdminPerseguible Then
'                        ''si no esta muerto o invisible o ciego... o tiene el /ignorando
'                        dist = Sqr((UserList(PJEnInd).Pos.X - NPCPosX) ^ 2 + (UserList(PJEnInd).Pos.Y - NPCPosY) ^ 2)
'                        If (dist < distBestTarget Or BestTarget = 0) Then
'                            BestTarget = PJEnInd
'                            distBestTarget = dist
'                        End If
'                    End If
'                End If
'            End If  ''canattack = 1
'        Next Y
'    Next X
'
'    If Not hayPretorianos Then
'        ''si estoy aca es porque no hay pretorianos cerca!!!
'        ''Todo mi ejercito fue asesinado
'        ''Salgo a atacar a todos a lo loco a espadazos
'        If BestTarget > 0 Then
'            If EsAlcanzable(npcind, BestTarget) Then
'                Call GreedyWalkTo(npcind, UserList(BestTarget).Pos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
'                'GreedyWalkTo npcind, UserList(BestTarget).Pos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y
'            Else
'                ''el chabon es piola y ataca desde lejos entonces lo castigamos!
'                Call NPCLanzaEstupidezPJ(npcind, BestTarget, DAT_ESTUPIDEZ)
'                Call NPCLanzaCegueraPJ(npcind, BestTarget, DAT_CEGUERA)
'            End If
'
'            ''heading loop de ataque
'            ''teclavolaespada
'            For headingloop = eHeading.NORTH To eHeading.WEST
'                nPos = Npclist(npcind).Pos
'                Call HeadtoPos(headingloop, nPos)
'                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
'                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
'                    If UI > 0 Then
'                        If NpcAtacaUser(npcind, UI) Then
'                            Call ChangeNPCChar(npcind, Npclist(npcind).Char.body, Npclist(npcind).Char.Head, headingloop)
'                        End If
'
'                        ''special speed ability for praetorian king ---------
'                        Npclist(npcind).CanAttack = 1   ''this is NOT a bug!!
'                        '----------------------------------------------------
'
'                    End If
'                End If
'            Next headingloop
'
'        Else    ''no hay targets cerca
'            Call VolverAlCentro(npcind)
'            If (Npclist(npcind).Stats.MinHp < Npclist(npcind).Stats.MaxHp) And (Npclist(npcind).CanAttack = 1) Then
'                ''si no hay ndie y estoy daniado me curo
'                Call NPCCuraLevesNPC(npcind, npcind, DAT_CURARLEVES)
'                Npclist(npcind).CanAttack = 0
'            End If
'
'        End If
'    End If
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.PRREY_AI? ")
'
'End Sub
'
'Sub PRGUER_AI(ByVal npcind As Integer)
'On Error GoTo errorh
'
'    Dim headingloop As Byte
'    Dim nPos As WorldPos
'    Dim X As Integer
'    Dim Y As Integer
'    Dim dist As Integer
'    Dim distBestTarget As Integer
'    Dim NPCPosX As Integer
'    Dim NPCPosY As Integer
'    Dim NPCPosM As Integer
'    Dim UI As Integer
'    Dim PJEnInd As Integer
'    Dim BestTarget As Integer
'    NPCPosM = Npclist(npcind).Pos.Map
'    NPCPosX = Npclist(npcind).Pos.X
'    NPCPosY = Npclist(npcind).Pos.Y
'    BestTarget = 0
'    dist = 0
'    distBestTarget = 0
'
'    For X = NPCPosX - 8 To NPCPosX + 8
'        For Y = NPCPosY - 7 To NPCPosY + 7
'            PJEnInd = MapData(NPCPosM, X, Y).UserIndex
'            If (PJEnInd > 0) Then
'                If (Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1 Or UserList(PJEnInd).flags.Muerto = 1)) And EsAlcanzable(npcind, PJEnInd) And UserList(PJEnInd).flags.AdminPerseguible Then
'                    ''caluclo la distancia al PJ, si esta mas cerca q el actual
'                    ''mejor besttarget entonces ataco a ese.
'                    If (BestTarget > 0) Then
'                        dist = Sqr((UserList(PJEnInd).Pos.X - NPCPosX) ^ 2 + (UserList(PJEnInd).Pos.Y - NPCPosY) ^ 2)
'                        If (dist < distBestTarget) Then
'                            BestTarget = PJEnInd
'                            distBestTarget = dist
'                        End If
'                    Else
'                        distBestTarget = Sqr((UserList(PJEnInd).Pos.X - NPCPosX) ^ 2 + (UserList(PJEnInd).Pos.Y - NPCPosY) ^ 2)
'                        BestTarget = PJEnInd
'                    End If
'                End If
'            End If
'        Next Y
'    Next X
'
'    ''LLamo a esta funcion si lo llevaron muy lejos.
'    ''La idea es que no lo "alejen" del rey y despues queden
'    ''lejos de la batalla cuando matan a un enemigo o este
'    ''sale del area de combate (tipica forma de separar un clan)
'    If Npclist(npcind).flags.Paralizado = 0 Then
'
'        'MEJORA: Si quedan solos, se van con el resto del ejercito
'        If Npclist(npcind).Invent.ArmourEqpSlot <> 0 Then
'            Call CambiarAlcoba(npcind)
'            'si me estoy yendo a alguna alcoba
'        ElseIf BestTarget = 0 Or EstoyMuyLejos(npcind) Then
'            Call VolverAlCentro(npcind)
'        ElseIf BestTarget > 0 Then
'            Call GreedyWalkTo(npcind, UserList(BestTarget).Pos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
'        End If
'    End If
'
'''teclavolaespada
'For headingloop = eHeading.NORTH To eHeading.WEST
'    nPos = Npclist(npcind).Pos
'    Call HeadtoPos(headingloop, nPos)
'    If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
'        UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
'        If UI > 0 Then
'            If Not (UserList(UI).flags.Muerto = 1) Then
'                If NpcAtacaUser(npcind, UI) Then
'                    Call ChangeNPCChar(npcind, Npclist(npcind).Char.body, Npclist(npcind).Char.Head, headingloop)
'                End If
'                Npclist(npcind).CanAttack = 0
'            End If
'        End If
'    End If
'Next headingloop
'
'
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.PRGUER_AI? ")
'
'
'End Sub
'
'Sub PRCLER_AI(ByVal npcind As Integer)
'On Error GoTo errorh
'
'    'HECHIZOS: NO CAMBIAR ACA
'    'REPRESENTAN LA UBICACION DE LOS SPELLS EN NPC_HOSTILES.DAT y si se los puede cambiar en ese archivo
'    '1- PARALIZAR PJS 'MODIFICABLE
'    '2- REMOVER PARALISIS 'NO MODIFICABLE
'    '3- CURARGRAVES - 'NO MODIFICABLE
'    '4- PARALIZAR MASCOTAS - 'NO MODIFICABLE
'    '5- CURARVENENO - 'NO MODIFICABLE
'    Dim DAT_PARALIZARPJ As Integer
'    Dim DAT_REMUEVEPARALISIS As Integer
'    Dim DAT_CURARGRAVES As Integer
'    Dim DAT_PARALIZAR_NPC As Integer
'    Dim DAT_TORMENTAAVANZADA As Integer
'    DAT_PARALIZARPJ = 1
'    DAT_REMUEVEPARALISIS = 2
'    DAT_PARALIZAR_NPC = 3
'    DAT_CURARGRAVES = 4
'    DAT_TORMENTAAVANZADA = 5
'
'    Dim X As Integer
'    Dim Y As Integer
'    Dim NPCPosX As Integer
'    Dim NPCPosY As Integer
'    Dim NPCPosM As Integer
'    Dim NPCAlInd As Integer
'    Dim PJEnInd As Integer
'    Dim centroX As Integer
'    Dim centroY As Integer
'    Dim BestTarget As Integer
'    Dim PJBestTarget As Boolean
'    Dim azar, azar2 As Integer
'    Dim quehacer As Byte
'        ''1- paralizar enemigo,
'        ''2- bombardear enemigo
'        ''3- ataque a mascotas
'        ''4- curar aliado
'    quehacer = 0
'    NPCPosM = Npclist(npcind).Pos.Map
'    NPCPosX = Npclist(npcind).Pos.X
'    NPCPosY = Npclist(npcind).Pos.Y
'    PJBestTarget = False
'    BestTarget = 0
'
'    azar = Sgn(RandomNumber(-1, 1))
'    If azar = 0 Then azar = 1
'    azar2 = Sgn(RandomNumber(-1, 1))
'    If azar2 = 0 Then azar2 = 1
'
'    'pick the best target according to the following criteria:
'    '1) "hoaxed" friends MUST be released
'    '2) enemy shall be annihilated no matter what
'    '3) party healing if no threats
'    For X = NPCPosX + (azar * 8) To NPCPosX + (azar * -8) Step -azar
'        For Y = NPCPosY + (azar2 * 7) To NPCPosY + (azar2 * -7) Step -azar2
'            'scan combat field
'            NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex
'            PJEnInd = MapData(NPCPosM, X, Y).UserIndex
'            If (Npclist(npcind).CanAttack = 1) Then   ''saltea el analisis si no puede atacar para evitar cuentas
'                If (NPCAlInd > 0) Then  ''allie?
'                    If (esPretoriano(NPCAlInd) = 0) Then
'                        If (Npclist(NPCAlInd).MaestroUser > 0) And (Not (Npclist(NPCAlInd).flags.Paralizado > 0)) Then
'                            Call NPCparalizaNPC(npcind, NPCAlInd, DAT_PARALIZAR_NPC)
'                            Npclist(npcind).CanAttack = 0
'                            Exit Sub
'                        End If
'                    Else    'es un PJ aliado en combate
'                        If (Npclist(NPCAlInd).flags.Paralizado = 1) Then
'                            ' amigo paralizado, an hoax vorp YA
'                            Call NPCRemueveParalisisNPC(npcind, NPCAlInd, DAT_REMUEVEPARALISIS)
'                            Npclist(npcind).CanAttack = 0
'                            Exit Sub
'                        ElseIf (BestTarget = 0) Then ''si no tiene nada q hacer..
'                            If (Npclist(NPCAlInd).Stats.MaxHp > Npclist(NPCAlInd).Stats.MinHp) Then
'                                BestTarget = NPCAlInd   ''cura heridas
'                                PJBestTarget = False
'                                quehacer = 4
'                            End If
'                        End If
'                    End If
'                ElseIf (PJEnInd > 0) Then ''aggressor
'                    If Not (UserList(PJEnInd).flags.Muerto = 1) And UserList(PJEnInd).flags.AdminPerseguible Then
'                        If (UserList(PJEnInd).flags.Paralizado = 0) Then
'                            If (Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1)) Then
'                                ''PJ movil y visible, jeje, si o si es target
'                                BestTarget = PJEnInd
'                                PJBestTarget = True
'                                quehacer = 1
'                            End If
'                        Else    ''PJ paralizado, ataca este invisible o no
'                            If Not (BestTarget > 0) Or Not (PJBestTarget) Then ''a menos q tenga algo mejor
'                                BestTarget = PJEnInd
'                                PJBestTarget = True
'                                quehacer = 2
'                            End If
'                        End If  ''endif paralizado
'                    End If  ''end if not muerto
'                End If  ''listo el analisis del tile
'            End If  ''saltea el analisis si no puede atacar, en realidad no es lo "mejor" pero evita cuentas inútiles
'        Next Y
'    Next X
'
'    ''aqui (si llego) tiene el mejor target
'    Select Case quehacer
'    Case 0
'        ''nada que hacer. Buscar mas alla del campo de visión algun aliado, a menos
'        ''que este paralizado pq no puedo ir
'        If Npclist(npcind).flags.Paralizado = 1 Then Exit Sub
'
'        If Not NPCPosM = MAPA_PRETORIANO Then Exit Sub
'
'        If NPCPosX < 50 Then centroX = ALCOBA1_X Else centroX = ALCOBA2_X
'        centroY = ALCOBA1_Y
'        ''aca establecí el lugar de las alcobas
'
'        ''Este doble for busca amigos paralizados lejos para ir a rescatarlos
'        ''Entra aca solo si en el area cercana al rey no hay algo mejor que
'        ''hacer.
'        For X = centroX - 16 To centroX + 16
'            For Y = centroY - 15 To centroY + 15
'                If Not (X < NPCPosX + 8 And X > NPCPosX + 8 And Y < NPCPosY + 7 And Y > NPCPosY - 7) Then
'                ''si no es un tile ya analizado... (evito cuentas)
'                    NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex
'                    If NPCAlInd > 0 Then
'                        If (esPretoriano(NPCAlInd) > 0 And Npclist(NPCAlInd).flags.Paralizado = 1) Then
'                            ''si esta paralizado lo va a rescatar, sino
'                            ''ya va a volver por su cuenta
'                            Call GreedyWalkTo(npcind, NPCPosM, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y)
''                            GreedyWalkTo npcind, NPCPosM, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y
'                            Exit Sub
'                        End If
'                    End If  ''endif npc
'                End If  ''endif tile analizado
'            Next Y
'        Next X
'
'        ''si estoy aca esta totalmente al cuete el clerigo o mal posicionado por rescate anterior
'        If Npclist(npcind).Invent.ArmourEqpSlot = 0 Then
'            Call VolverAlCentro(npcind)
'            Exit Sub
'        End If
'        ''fin quehacer = 0 (npc al cuete)
'
'    Case 1  '' paralizar enemigo PJ
'        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(Npclist(npcind).Spells(DAT_PARALIZARPJ)).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
'        Call NpcLanzaSpellSobreUser(npcind, BestTarget, Npclist(npcind).Spells(DAT_PARALIZARPJ)) ''SPELL 1 de Clerigo es PARALIZAR
'        Exit Sub
'    Case 2  '' ataque a usuarios (invisibles tambien)
'        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(Npclist(npcind).Spells(DAT_TORMENTAAVANZADA)).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
'        Call NpcLanzaSpellSobreUser2(npcind, BestTarget, Npclist(npcind).Spells(DAT_TORMENTAAVANZADA)) ''SPELL 2 de Clerigo es Vax On Tar avanzado
'        Exit Sub
'    Case 3  '' ataque a mascotas
'        If Not (Npclist(BestTarget).flags.Paralizado = 1) Then
'            Call NPCparalizaNPC(npcind, BestTarget, DAT_PARALIZAR_NPC)
'            Npclist(npcind).CanAttack = 0
'        End If  ''TODO: vax on tar sobre mascotas
'    Case 4  '' party healing
'        Call NPCcuraNPC(npcind, BestTarget, DAT_CURARGRAVES)
'        Npclist(npcind).CanAttack = 0
'    End Select
'
'
'
'    ''movimientos
'    ''EL clerigo no tiene un movimiento fijo, pero es esperable
'    ''que no se aleje mucho del rey... y si se aleje de espaderos
'
'    If Npclist(npcind).flags.Paralizado = 1 Then Exit Sub
'
'    If Not (NPCPosM = MAPA_PRETORIANO) Then Exit Sub
'
'    'MEJORA: Si quedan solos, se van con el resto del ejercito
'    If Npclist(npcind).Invent.ArmourEqpSlot <> 0 Then
'        Call CambiarAlcoba(npcind)
'        Exit Sub
'    End If
'
'
'    PJEnInd = MapData(NPCPosM, NPCPosX - 1, NPCPosY).UserIndex
'    If PJEnInd > 0 Then
'        If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
'            ''esta es una forma muy facil de matar 2 pajaros
'            ''de un tiro. Se aleja del usuario pq el centro va a
'            ''estar ocupado, y a la vez se aproxima al rey, manteniendo
'            ''una linea de defensa compacta
'            Call VolverAlCentro(npcind)
'            Exit Sub
'        End If
'    End If
'
'    PJEnInd = MapData(NPCPosM, NPCPosX + 1, NPCPosY).UserIndex
'    If PJEnInd > 0 Then
'        If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
'            Call VolverAlCentro(npcind)
'            Exit Sub
'        End If
'    End If
'
'    PJEnInd = MapData(NPCPosM, NPCPosX, NPCPosY - 1).UserIndex
'    If PJEnInd > 0 Then
'        If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
'            Call VolverAlCentro(npcind)
'            Exit Sub
'        End If
'    End If
'
'    PJEnInd = MapData(NPCPosM, NPCPosX, NPCPosY + 1).UserIndex
'    If PJEnInd > 0 Then
'        If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
'            Call VolverAlCentro(npcind)
'            Exit Sub
'        End If
'    End If
'
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.PRCLER_AI? ")
'
'End Sub
'
'Function EsMagoOClerigo(ByVal PJEnInd As Integer) As Boolean
'On Error GoTo errorh
'
'    EsMagoOClerigo = UserList(PJEnInd).clase = eClass.Mage Or _
'                        UserList(PJEnInd).clase = eClass.Cleric Or _
'                        UserList(PJEnInd).clase = eClass.Druid Or _
'                        UserList(PJEnInd).clase = eClass.Bard
'Exit Function
'
'errorh:
'    LogError ("Error en NPCAI.EsMagoOClerigo? ")
'End Function
'
'Sub NPCRemueveVenenoNPC(ByVal npcind As Integer, ByVal NPCAlInd As Integer, ByVal indice As Integer)
'On Error GoTo errorh
'    Dim indireccion As Integer
'
'    indireccion = Npclist(npcind).Spells(indice)
'    '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
'    Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
'    Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessageCreateFX(Npclist(NPCAlInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
'    Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y))
'    Npclist(NPCAlInd).Veneno = 0
'    Npclist(NPCAlInd).flags.Envenenado = 0
'
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.NPCRemueveVenenoNPC? ")
'
'End Sub
'
'Sub NPCCuraLevesNPC(ByVal npcind As Integer, ByVal NPCAlInd As Integer, ByVal indice As Integer)
'On Error GoTo errorh
'    Dim indireccion As Integer
'
'    indireccion = Npclist(npcind).Spells(indice)
'    '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
'    Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
'    Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y))
'    Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessageCreateFX(Npclist(NPCAlInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
'
'    If (Npclist(NPCAlInd).Stats.MinHp + 5 < Npclist(NPCAlInd).Stats.MaxHp) Then
'        Npclist(NPCAlInd).Stats.MinHp = Npclist(NPCAlInd).Stats.MinHp + 5
'    Else
'        Npclist(NPCAlInd).Stats.MinHp = Npclist(NPCAlInd).Stats.MaxHp
'    End If
'
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.NPCCuraLevesNPC? ")
'
'End Sub
'
'
'Sub NPCRemueveParalisisNPC(ByVal npcind As Integer, ByVal NPCAlInd As Integer, ByVal indice As Integer)
'On Error GoTo errorh
'    Dim indireccion As Integer
'
'    indireccion = Npclist(npcind).Spells(indice)
'    '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
'    Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
'    Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y))
'    Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessageCreateFX(Npclist(NPCAlInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
'    Npclist(NPCAlInd).Contadores.Paralisis = 0
'    Npclist(NPCAlInd).flags.Paralizado = 0
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.NPCRemueveParalisisNPC? ")
'
'End Sub
'
'Sub NPCparalizaNPC(ByVal paralizador As Integer, ByVal Paralizado As Integer, ByVal indice)
'On Error GoTo errorh
'    Dim indireccion As Integer
'
'    indireccion = Npclist(paralizador).Spells(indice)
'    '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
'    Call SendData(SendTarget.ToNPCArea, paralizador, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(paralizador).Char.CharIndex, vbCyan))
'    Call SendData(SendTarget.ToNPCArea, Paralizado, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(Paralizado).Pos.X, Npclist(Paralizado).Pos.Y))
'    Call SendData(SendTarget.ToNPCArea, Paralizado, PrepareMessageCreateFX(Npclist(Paralizado).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
'
'    Npclist(Paralizado).flags.Paralizado = 1
'    Npclist(Paralizado).Contadores.Paralisis = IntervaloParalizado * 2
'
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.NPCParalizaNPC? ")
'
'End Sub
'
'Sub NPCcuraNPC(ByVal curador As Integer, ByVal curado As Integer, ByVal indice As Integer)
'On Error GoTo errorh
'    Dim indireccion As Integer
'
'
'    indireccion = Npclist(curador).Spells(indice)
'    '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
'    Call SendData(SendTarget.ToNPCArea, curador, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(curador).Char.CharIndex, vbCyan))
'    Call SendData(SendTarget.ToNPCArea, curado, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(curado).Pos.X, Npclist(curado).Pos.Y))
'    Call SendData(SendTarget.ToNPCArea, curado, PrepareMessageCreateFX(Npclist(curado).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
'    If Npclist(curado).Stats.MinHp + 30 > Npclist(curado).Stats.MaxHp Then
'        Npclist(curado).Stats.MinHp = Npclist(curado).Stats.MaxHp
'    Else
'        Npclist(curado).Stats.MinHp = Npclist(curado).Stats.MinHp + 30
'    End If
'Exit Sub
'
'errorh:
'    LogError ("Error en NPCAI.NPCcuraNPC? ")
'
'End Sub
'
'    PretoIndex = PretoIndex + 1
'    Pretorianos(PretoIndex) = NpcIndex
'
'Exit Sub
'
'ErrHandler:
'    LogError ("Error en NPCAI.CrearClanPretoriano. Error: " & Err.Number & " - " & Err.description)
'End Sub
'
'Public Sub MuerePretoriano(ByVal NpcIndex As Integer)
''***************************************************
''Author: ZaMa
''Last Modification: 27/06/2010
''Eliminates the pretorian from the current alive ones.
''***************************************************
'
'    Dim PretoIndex As Integer
'
'    For PretoIndex = 1 To NRO_PRETORIANOS
'        If Pretorianos(PretoIndex) = NpcIndex Then
'            Pretorianos(PretoIndex) = 0
'            Exit Sub
'        End If
'    Next PretoIndex
'
'End Sub
'
