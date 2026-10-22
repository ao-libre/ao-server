Attribute VB_Name = "BotAPI"
''
' BotAPI.bas - TCP JSON listener for aobench bot control.
'
' Opens a Winsock server on port 4200 (via frmBotAPI).
' Accepts newline-delimited JSON commands from the gateway process and executes
' them by calling into existing ao-libre game logic (not duplicating it).
'
' Protocol:
'   client -> server:  {"action":"getState","charname":"Bot1"}\n
'   server -> client:  {"ok":true,"state":{...}}\n  or  {"ok":false,"error":"..."}\n
'
' getState fields: charname, pos, heading, hp, maxHp, mana, maxMana, stamina,
'   maxStamina, gold, exp, expNeeded, level, clase, skills, inventory, nearbyNpcs,
'   nearbyPlayers, statusEffects, spells, skillPts, assignedSkills, attributes,
'   bankGold, bank, faction, reputation, party, quests, timestamp
'
' Player actions: walk, attack, castSpell, useItem, useSkill, work, leftClick,
'   say, yell, whisper, pickUp, equipItem, meditate, rest, drop,
'   doubleClick, changeHeading, workLeftClick, useSpellMacro, moveItem,
'   commerceStart, commerceBuy, commerceSell, commerceEnd, commerceChat,
'   userCommerceCreate, userCommerceOffer, userCommerceOk, userCommerceReject,
'   userCommerceConfirm, userCommerceEnd,
'   bankStart, bankDeposit, bankExtractItem, bankDepositGold, bankExtractGold,
'   bankEnd, moveBank,
'   craftBlacksmith, craftCarpenter, initCrafting, craftsmanCreate,
'   train, partyCreate, partyJoin, partyLeave, partyAccept, partyKick,
'   partySetLeader, partyOnline, partyMessage, requestPartyForm,
'   guildMessage, shareNpc, stopSharingNpc,
'   resucitate, heal, home, enlist, leaveFaction, changeDescription,
'   dragSpell, moveSpell, assignSkillPoints, itemUpgrade,
'   petStand, petFollow, releasePet,
'   fightSend, fightAccept,
'   quest, questAccept, questAbandon, questListRequest, questDetailsRequest,
'   forumPost, lookProcess, consultation, online, discord,
'   addAmigo, delAmigo, onAmigos, msgAmigos,
'   guildFundate, guildFundation
'
' GM actions (require GM privileges): gmGoToChar, gmWarpChar, gmWarpMe,
'   gmSummon, gmSpawn, gmKillNpc, gmKillAllNearby, gmCreateItem,
'   gmDestroyItems, gmRevive, gmKick, gmBan, gmSilence, gmJail,
'   gmInvisible, gmMessage, gmServerMessage, gmRain, gmTeleportCreate,
'   gmEditChar, gmSetDesc, gmCreateNpc, gmMakeDumb, gmMakeDumbNoMore,
'   gmRequestCharInfo
'
' Author: aobench (ao-libre fork)
''

Option Explicit

' ---- Port -----------------------------------------------------------------
Public Const BOTAPI_PORT As Integer = 4200

' ---- Heading string -> eHeading -------------------------------------------
Private Function ParseHeading(ByVal s As String) As eHeading
    Select Case LCase$(Trim$(s))
        Case "north": ParseHeading = eHeading.NORTH
        Case "east":  ParseHeading = eHeading.EAST
        Case "south": ParseHeading = eHeading.SOUTH
        Case "west":  ParseHeading = eHeading.WEST
        Case Else:    ParseHeading = 0
    End Select
End Function

' ---- Skill string -> eSkill -----------------------------------------------
Private Function ParseSkill(ByVal s As String) As eSkill
    Select Case s
        Case "Magia":        ParseSkill = eSkill.Magia
        Case "Robar":        ParseSkill = eSkill.Robar
        Case "Tacticas":     ParseSkill = eSkill.Tacticas
        Case "Armas":        ParseSkill = eSkill.Armas
        Case "Meditar":      ParseSkill = eSkill.Meditar
        Case "Apunalar":     ParseSkill = eSkill.Apunalar
        Case "Ocultarse":    ParseSkill = eSkill.Ocultarse
        Case "Supervivencia":ParseSkill = eSkill.Supervivencia
        Case "Talar":        ParseSkill = eSkill.Talar
        Case "Comerciar":    ParseSkill = eSkill.Comerciar
        Case "Defensa":      ParseSkill = eSkill.Defensa
        Case "pesca":        ParseSkill = eSkill.pesca
        Case "Mineria":      ParseSkill = eSkill.Mineria
        Case "Carpinteria":  ParseSkill = eSkill.Carpinteria
        Case "Herreria":     ParseSkill = eSkill.Herreria
        Case "Liderazgo":    ParseSkill = eSkill.Liderazgo
        Case "Domar":        ParseSkill = eSkill.Domar
        Case "Proyectiles":  ParseSkill = eSkill.Proyectiles
        Case "Wrestling":    ParseSkill = eSkill.Wrestling
        Case "Navegacion":   ParseSkill = eSkill.Navegacion
        Case Else:           ParseSkill = 0
    End Select
End Function

' ---- charname -> UserIndex lookup -----------------------------------------
' Returns 0 if not found.
Public Function BotFindUser(ByVal charname As String) As Integer
    Dim i As Integer
    For i = 1 To MaxUsers
        If UserList(i).flags.UserLogged Then
            If LCase$(UserList(i).Name) = LCase$(charname) Then
                BotFindUser = i
                Exit Function
            End If
        End If
    Next i
    BotFindUser = 0
End Function

' ---- Error response helper ------------------------------------------------
Public Function BotError(ByVal msg As String) As String
    BotError = "{""ok"":false,""error"":" & JSON_Encode(msg) & "}"
End Function

' ---- JSON string encoder (escapes quotes and backslashes) -----------------
Public Function JSON_Encode(ByVal s As String) As String
    Dim result As String
    result = s
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, Chr(13), "\r")
    result = Replace(result, Chr(10), "\n")
    result = Replace(result, Chr(9), "\t")
    JSON_Encode = """" & result & """"
End Function

' ===========================================================================
' DISPATCH â€” called by frmBotAPI for each complete JSON message
' Returns a JSON response string (no trailing newline).
' ===========================================================================
Public Function BotDispatch(ByVal jsonMsg As String) As String

    Dim parsed  As Object
    Dim action  As String
    Dim charname As String
    Dim Userindex As Integer

    On Error GoTo ParseError

    ' Parse with the existing JSON.bas module (parse returns a Dictionary)
    Set parsed = parse(jsonMsg)

    If parsed Is Nothing Then
        BotDispatch = BotError("Invalid JSON")
        Exit Function
    End If

    On Error GoTo DispatchError

    ' Required fields
    If Not parsed.Exists("action") Then
        BotDispatch = BotError("action is required")
        Exit Function
    End If
    If Not parsed.Exists("charname") Then
        BotDispatch = BotError("charname is required")
        Exit Function
    End If

    action   = CStr(parsed("action"))
    charname = CStr(parsed("charname"))

    ' getState does not require a logged-in user (returns error if not found)
    Userindex = BotFindUser(charname)

    If action = "getState" Then
        If Userindex = 0 Then
            BotDispatch = BotError("Character not found: " & charname)
        Else
            BotDispatch = BotGetState(Userindex)
        End If
        Exit Function
    End If

    ' All other actions require the user to be online
    If Userindex = 0 Then
        BotDispatch = BotError("Character not found: " & charname)
        Exit Function
    End If

    Select Case action
        Case "walk":        BotDispatch = BotWalk(Userindex, parsed)
        Case "attack":      BotDispatch = BotAttack(Userindex)
        Case "castSpell":   BotDispatch = BotCastSpell(Userindex, parsed)
        Case "useItem":     BotDispatch = BotUseItem(Userindex, parsed)
        Case "useSkill":    BotDispatch = BotUseSkill(Userindex, parsed)
        Case "work":        BotDispatch = BotWork(Userindex, parsed)
        Case "leftClick":   BotDispatch = BotLeftClick(Userindex, parsed)
        Case "say":         BotDispatch = BotSay(Userindex, parsed)
        Case "yell":        BotDispatch = BotYell(Userindex, parsed)
        Case "whisper":     BotDispatch = BotWhisper(Userindex, parsed)
        Case "pickUp":      BotDispatch = BotPickUp(Userindex, parsed)
        Case "equipItem":   BotDispatch = BotEquipItem(Userindex, parsed)
        Case "meditate":    BotDispatch = BotMeditate(Userindex)
        Case "rest":        BotDispatch = BotRest(Userindex)
        Case "drop":        BotDispatch = BotDrop(Userindex, parsed)
        Case Else
            ' Try extended actions (player + GM)
            Dim extResult As String
            extResult = BotDispatchExtended(action, Userindex, parsed)
            If Len(extResult) > 0 Then
                BotDispatch = extResult
            Else
                BotDispatch = BotError("Unknown action: " & action)
            End If
    End Select

    Exit Function

ParseError:
    BotDispatch = BotError("JSON parse error")
    Exit Function

DispatchError:
    BotDispatch = BotError("Internal error in " & action & ": " & Err.Description)
End Function

' ===========================================================================
' getState â€” serialise UserList(i) to JSON
' ===========================================================================
Public Function BotGetState(ByVal i As Integer) As String

    Dim sb As cStringBuilder
    Set sb = New cStringBuilder

    With UserList(i)

        Call sb.Append("{""ok"":true,""state"":{")
        Call sb.Append("""charname"":" & JSON_Encode(.Name) & ",")
        Call sb.Append("""pos"":{""x"":" & .Pos.X & ",""y"":" & .Pos.Y & ",""map"":" & .Pos.Map & "},")
        Call sb.Append("""heading"":" & JSON_Encode(HeadingToString(.Char.heading)) & ",")
        Call sb.Append("""hp"":" & .Stats.MinHp & ",")
        Call sb.Append("""maxHp"":" & .Stats.MaxHp & ",")
        Call sb.Append("""mana"":" & .Stats.MinMAN & ",")
        Call sb.Append("""maxMana"":" & .Stats.MaxMAN & ",")
        Call sb.Append("""stamina"":" & .Stats.MinSta & ",")
        Call sb.Append("""maxStamina"":" & .Stats.MaxSta & ",")
        Call sb.Append("""gold"":" & .Stats.Gld & ",")
        Call sb.Append("""exp"":" & CLng(.Stats.Exp) & ",")
        Call sb.Append("""level"":" & .Stats.ELV & ",")
        Call sb.Append("""clase"":" & CByte(.Clase) & ",")

        ' Skills array
        Dim skillIdx As Integer
        Call sb.Append("""skills"":{")
        For skillIdx = 1 To NUMSKILLS
            If skillIdx > 1 Then Call sb.Append(",")
            Call sb.Append(JSON_Encode(SkillIndexToName(skillIdx)) & _
                ":{""level"":" & .Stats.UserSkills(skillIdx) & _
                ",""exp"":" & .Stats.ExpSkills(skillIdx) & "}")
        Next skillIdx
        Call sb.Append("},")

        ' Inventory array
        Dim slotIdx As Integer
        Call sb.Append("""inventory"":[")
        Dim firstItem As Boolean
        firstItem = True
        For slotIdx = 1 To MAX_INVENTORY_SLOTS
            If .Invent.Object(slotIdx).ObjIndex > 0 Then
                If Not firstItem Then Call sb.Append(",")
                firstItem = False
                Dim equipped As Boolean
                equipped = (.Invent.Object(slotIdx).Equipped <> 0)
                Call sb.Append("{""slot"":" & slotIdx & _
                    ",""itemId"":" & .Invent.Object(slotIdx).ObjIndex & _
                    ",""name"":" & JSON_Encode(ObjData(.Invent.Object(slotIdx).ObjIndex).Name) & _
                    ",""amount"":" & .Invent.Object(slotIdx).Amount & _
                    ",""equipped"":" & IIf(equipped, "true", "false") & "}")
            End If
        Next slotIdx
        Call sb.Append("],")

        ' Nearby NPCs (within 8 tiles)
        Call sb.Append("""nearbyNpcs"":[")
        Dim npcIdx As Integer
        Dim firstNpc As Boolean
        firstNpc = True
        For npcIdx = 1 To MAXNPCS
            If Npclist(npcIdx).Activo Then
                If Npclist(npcIdx).Pos.Map = .Pos.Map Then
                    Dim dx As Integer, dy As Integer
                    dx = Abs(Npclist(npcIdx).Pos.X - .Pos.X)
                    dy = Abs(Npclist(npcIdx).Pos.Y - .Pos.Y)
                    If dx <= 8 And dy <= 8 Then
                        If Not firstNpc Then Call sb.Append(",")
                        firstNpc = False
                        Call sb.Append("{""id"":" & npcIdx & _
                            ",""name"":" & JSON_Encode(Npclist(npcIdx).Name) & _
                            ",""x"":" & Npclist(npcIdx).Pos.X & _
                            ",""y"":" & Npclist(npcIdx).Pos.Y & _
                            ",""hp"":" & Npclist(npcIdx).Stats.MinHp & _
                            ",""maxHp"":" & Npclist(npcIdx).Stats.MaxHp & _
                            ",""type"":" & CByte(Npclist(npcIdx).NPCtype) & "}")
                    End If
                End If
            End If
        Next npcIdx
        Call sb.Append("],")

        ' Nearby players (within 8 tiles)
        Call sb.Append("""nearbyPlayers"":[")
        Dim pIdx As Integer
        Dim firstP As Boolean
        firstP = True
        For pIdx = 1 To MaxUsers
            If pIdx <> i Then
                If UserList(pIdx).flags.UserLogged Then
                    If UserList(pIdx).Pos.Map = .Pos.Map Then
                        Dim pdx As Integer, pdy As Integer
                        pdx = Abs(UserList(pIdx).Pos.X - .Pos.X)
                        pdy = Abs(UserList(pIdx).Pos.Y - .Pos.Y)
                        If pdx <= 8 And pdy <= 8 Then
                            If Not firstP Then Call sb.Append(",")
                            firstP = False
                            Call sb.Append("{""name"":" & JSON_Encode(UserList(pIdx).Name) & _
                                ",""x"":" & UserList(pIdx).Pos.X & _
                                ",""y"":" & UserList(pIdx).Pos.Y & _
                                ",""hp"":" & UserList(pIdx).Stats.MinHp & _
                                ",""maxHp"":" & UserList(pIdx).Stats.MaxHp & "}")
                        End If
                    End If
                End If
            End If
        Next pIdx
        Call sb.Append("],")

        ' Status effects
        Call sb.Append("""statusEffects"":[")
        Dim firstEffect As Boolean
        firstEffect = True
        If .flags.Envenenado <> 0 Then
            Call sb.Append("""poisoned""")
            firstEffect = False
        End If
        If .flags.Paralizado <> 0 Then
            If Not firstEffect Then Call sb.Append(",")
            Call sb.Append("""paralyzed""")
            firstEffect = False
        End If
        If .flags.Meditando Then
            If Not firstEffect Then Call sb.Append(",")
            Call sb.Append("""meditating""")
            firstEffect = False
        End If
        If .flags.Descansar Then
            If Not firstEffect Then Call sb.Append(",")
            Call sb.Append("""resting""")
            firstEffect = False
        End If
        If .flags.Ceguera <> 0 Then
            If Not firstEffect Then Call sb.Append(",")
            Call sb.Append("""blind""")
            firstEffect = False
        End If
        If .flags.Estupidez <> 0 Then
            If Not firstEffect Then Call sb.Append(",")
            Call sb.Append("""dumb""")
            firstEffect = False
        End If
        If .flags.invisible <> 0 Then
            If Not firstEffect Then Call sb.Append(",")
            Call sb.Append("""invisible""")
            firstEffect = False
        End If
        If .flags.Oculto <> 0 Then
            If Not firstEffect Then Call sb.Append(",")
            Call sb.Append("""hidden""")
            firstEffect = False
        End If
        Call sb.Append("],")

        ' â”€â”€ Spell hotbar (35 slots) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Call sb.Append("""spells"":[")
        Dim spIdx As Integer
        For spIdx = 1 To MAXUSERHECHIZOS
            If spIdx > 1 Then Call sb.Append(",")
            Dim spId As Integer
            spId = .Stats.UserHechizos(spIdx)
            If spId > 0 Then
                Call sb.Append("{""slot"":" & spIdx & _
                    ",""spellId"":" & spId & _
                    ",""name"":" & JSON_Encode(Hechizos(spId).Nombre) & _
                    ",""minSkill"":" & Hechizos(spId).MinSkill & _
                    ",""manaCost"":" & Hechizos(spId).ManaCosto & "}")
            Else
                Call sb.Append("null")
            End If
        Next spIdx
        Call sb.Append("],")

        ' â”€â”€ Skill points â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Call sb.Append("""skillPts"":" & .Stats.SkillPts & ",")
        Call sb.Append("""assignedSkills"":" & .Counters.AsignedSkills & ",")

        ' â”€â”€ Character attributes â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Call sb.Append("""attributes"":{")
        Call sb.Append("""Fuerza"":" & .Stats.UserAtributos(eAtributos.Fuerza) & ",")
        Call sb.Append("""Agilidad"":" & .Stats.UserAtributos(eAtributos.Agilidad) & ",")
        Call sb.Append("""Inteligencia"":" & .Stats.UserAtributos(eAtributos.Inteligencia) & ",")
        Call sb.Append("""Carisma"":" & .Stats.UserAtributos(eAtributos.Carisma) & ",")
        Call sb.Append("""Constitucion"":" & .Stats.UserAtributos(eAtributos.Constitucion))
        Call sb.Append("},")

        ' â”€â”€ ELU (XP needed for next level) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Call sb.Append("""expNeeded"":" & CLng(.Stats.ELU) & ",")

        ' â”€â”€ Bank â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Call sb.Append("""bankGold"":" & .Stats.Banco & ",")
        Call sb.Append("""bank"":[")
        Dim bankIdx As Integer
        Dim firstBank As Boolean
        firstBank = True
        For bankIdx = 1 To MAX_BANCOINVENTORY_SLOTS
            If .BancoInvent.Object(bankIdx).ObjIndex > 0 Then
                If Not firstBank Then Call sb.Append(",")
                firstBank = False
                Call sb.Append("{""slot"":" & bankIdx & _
                    ",""itemId"":" & .BancoInvent.Object(bankIdx).ObjIndex & _
                    ",""name"":" & JSON_Encode(ObjData(.BancoInvent.Object(bankIdx).ObjIndex).Name) & _
                    ",""amount"":" & .BancoInvent.Object(bankIdx).Amount & "}")
            End If
        Next bankIdx
        Call sb.Append("],")

        ' â”€â”€ Faction â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Dim factionName As String
        If .flags.ArmadaReal <> 0 Then
            factionName = "real"
        ElseIf .flags.FuerzasCaos <> 0 Then
            factionName = "caos"
        Else
            factionName = "none"
        End If
        Call sb.Append("""faction"":" & JSON_Encode(factionName) & ",")

        ' â”€â”€ Reputation â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Call sb.Append("""reputation"":{")
        Call sb.Append("""npcKills"":" & .Stats.NPCsMuertos & ",")
        Call sb.Append("""playerKills"":" & .Stats.UsuariosMatados & ",")
        Call sb.Append("""citizensKilled"":" & .flags.CiudadanosMatados & ",")
        Call sb.Append("""criminalsKilled"":" & .flags.CriminalesMatados)
        Call sb.Append("},")

        ' â”€â”€ Party â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Dim partyIdx As Integer
        partyIdx = .Counters.PartyIndex
        If partyIdx > 0 Then
            Dim pm As Integer
            Call sb.Append("""party"":{")
            Call sb.Append("""index"":" & partyIdx & ",")
            Call sb.Append("""isLeader"":" & IIf(PartyList(partyIdx).LeaderIndex = i, "true", "false") & ",")
            Call sb.Append("""members"":[")
            Dim pmFirst As Boolean
            pmFirst = True
            For pm = 1 To PartyList(partyIdx).MemberCount
                Dim pmU As Integer
                pmU = PartyList(partyIdx).Members(pm)
                If pmU > 0 And UserList(pmU).flags.UserLogged Then
                    If Not pmFirst Then Call sb.Append(",")
                    pmFirst = False
                    Call sb.Append(JSON_Encode(UserList(pmU).Name))
                End If
            Next pm
            Call sb.Append("]},")
        Else
            Call sb.Append("""party"":null,")
        End If

        ' â”€â”€ Active quests â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        Call sb.Append("""quests"":{")
        Call sb.Append("""active"":[")
        Dim qIdx As Integer
        Dim firstQ As Boolean
        firstQ = True
        For qIdx = 1 To MAXUSERQUESTS
            Dim qI As Integer
            qI = .QuestStats.Quests(qIdx).QuestIndex
            If qI > 0 Then
                If Not firstQ Then Call sb.Append(",")
                firstQ = False
                Call sb.Append("{""slot"":" & qIdx & _
                    ",""questId"":" & qI & _
                    ",""name"":" & JSON_Encode(QuestList(qI).Nombre) & _
                    ",""requiredLevel"":" & QuestList(qI).RequiredLevel & ",")
                ' Kill progress
                Call sb.Append("""npcProgress"":[")
                Dim nkIdx As Integer
                For nkIdx = 1 To QuestList(qI).RequiredNPCs
                    If nkIdx > 1 Then Call sb.Append(",")
                    Call sb.Append("{""npcId"":" & QuestList(qI).RequiredNPC(nkIdx).NpcIndex & _
                        ",""required"":" & QuestList(qI).RequiredNPC(nkIdx).Amount & _
                        ",""killed"":" & .QuestStats.Quests(qIdx).NPCsKilled(nkIdx) & "}")
                Next nkIdx
                Call sb.Append("],")
                Call sb.Append("""rewardXp"":" & QuestList(qI).RewardEXP & _
                    ",""rewardGold"":" & QuestList(qI).RewardGLD & "}")
            End If
        Next qIdx
        Call sb.Append("],")
        Call sb.Append("""doneIds"":[")
        Dim dqIdx As Integer
        For dqIdx = 1 To .QuestStats.NumQuestsDone
            If dqIdx > 1 Then Call sb.Append(",")
            Call sb.Append(.QuestStats.QuestsDone(dqIdx))
        Next dqIdx
        Call sb.Append("]},")

        Call sb.Append("""timestamp"":" & CLng(Timer * 1000) & "}")
        Call sb.Append("}")

    End With

    BotGetState = sb.ToString

End Function

' ---- Heading enum -> string -----------------------------------------------
Private Function HeadingToString(ByVal h As eHeading) As String
    Select Case h
        Case eHeading.NORTH: HeadingToString = "north"
        Case eHeading.EAST:  HeadingToString = "east"
        Case eHeading.SOUTH: HeadingToString = "south"
        Case eHeading.WEST:  HeadingToString = "west"
        Case Else:           HeadingToString = "south"
    End Select
End Function

' ---- Skill index -> name --------------------------------------------------
Private Function SkillIndexToName(ByVal idx As Integer) As String
    Select Case idx
        Case 1:  SkillIndexToName = "Magia"
        Case 2:  SkillIndexToName = "Robar"
        Case 3:  SkillIndexToName = "Tacticas"
        Case 4:  SkillIndexToName = "Armas"
        Case 5:  SkillIndexToName = "Meditar"
        Case 6:  SkillIndexToName = "Apunalar"
        Case 7:  SkillIndexToName = "Ocultarse"
        Case 8:  SkillIndexToName = "Supervivencia"
        Case 9:  SkillIndexToName = "Talar"
        Case 10: SkillIndexToName = "Comerciar"
        Case 11: SkillIndexToName = "Defensa"
        Case 12: SkillIndexToName = "pesca"
        Case 13: SkillIndexToName = "Mineria"
        Case 14: SkillIndexToName = "Carpinteria"
        Case 15: SkillIndexToName = "Herreria"
        Case 16: SkillIndexToName = "Liderazgo"
        Case 17: SkillIndexToName = "Domar"
        Case 18: SkillIndexToName = "Proyectiles"
        Case 19: SkillIndexToName = "Wrestling"
        Case 20: SkillIndexToName = "Navegacion"
        Case Else: SkillIndexToName = "Skill" & idx
    End Select
End Function

' ===========================================================================
' Action handlers
' ===========================================================================

Private Function BotWalk(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("heading") Then
        BotWalk = BotError("heading is required")
        Exit Function
    End If
    Dim h As eHeading
    h = ParseHeading(CStr(p("heading")))
    If h = 0 Then
        BotWalk = BotError("Invalid heading: " & CStr(p("heading")))
        Exit Function
    End If
    Call MoveUserChar(i, h)
    BotWalk = "{""ok"":true,""moved"":true}"
End Function

Private Function BotAttack(ByVal i As Integer) As String
    ' Auto-target nearest NPC if none set
    If UserList(i).flags.TargetNPC = 0 Then
        Dim nearest As Integer, nearDist As Long, nIdx As Integer
        nearest = 0
        nearDist = 999999
        For nIdx = 1 To MAXNPCS
            If Npclist(nIdx).Activo Then
                If Npclist(nIdx).Pos.Map = UserList(i).Pos.Map Then
                    Dim d As Long
                    d = CLng(Abs(Npclist(nIdx).Pos.X - UserList(i).Pos.X)) + _
                        CLng(Abs(Npclist(nIdx).Pos.Y - UserList(i).Pos.Y))
                    If d < nearDist Then
                        nearDist = d
                        nearest = nIdx
                    End If
                End If
            End If
        Next nIdx
        If nearest > 0 Then UserList(i).flags.TargetNPC = nearest
    End If
    If UserList(i).flags.TargetNPC = 0 Then
        BotAttack = BotError("No target in range")
        Exit Function
    End If
    ' UsuarioAtaca handles NPC/player targeting internally
    Call UsuarioAtaca(i)
    BotAttack = "{""ok"":true,""targetNpc"":" & UserList(i).flags.TargetNPC & "}"
End Function

Private Function BotCastSpell(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("slot") Then
        BotCastSpell = BotError("slot is required")
        Exit Function
    End If
    Dim slot As Integer
    slot = CInt(p("slot"))
    If slot < 1 Or slot > MAXUSERHECHIZOS Then
        BotCastSpell = BotError("slot must be 1-" & MAXUSERHECHIZOS)
        Exit Function
    End If
    Dim spellIndex As Integer
    spellIndex = UserList(i).Stats.UserHechizos(slot)
    If spellIndex = 0 Then
        BotCastSpell = BotError("No spell in slot " & slot)
        Exit Function
    End If
    Call LanzarHechizo(spellIndex, i)
    BotCastSpell = "{""ok"":true,""spellSlot"":" & slot & "}"
End Function

Private Function BotUseItem(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("slot") Then
        BotUseItem = BotError("slot is required")
        Exit Function
    End If
    Dim slot As Integer
    slot = CInt(p("slot"))
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then
        BotUseItem = BotError("slot out of range")
        Exit Function
    End If
    If UserList(i).Invent.Object(slot).ObjIndex = 0 Then
        BotUseItem = BotError("No item in slot " & slot)
        Exit Function
    End If
    Call UseInvItem(i, slot)
    BotUseItem = "{""ok"":true,""slot"":" & slot & "}"
End Function

Private Function BotUseSkill(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("skill") Then
        BotUseSkill = BotError("skill is required")
        Exit Function
    End If
    Dim sk As eSkill
    sk = ParseSkill(CStr(p("skill")))
    If sk = 0 Then
        BotUseSkill = BotError("Unknown skill: " & CStr(p("skill")))
        Exit Function
    End If
    Call DoSkillWork(i, sk)
    BotUseSkill = "{""ok"":true,""skill"":" & JSON_Encode(CStr(p("skill"))) & "}"
End Function

Private Function BotWork(ByVal i As Integer, ByVal p As Object) As String
    ' same as useSkill â€” work and useSkill are synonyms in aobench
    BotWork = BotUseSkill(i, p)
End Function

Private Function BotLeftClick(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("x") Or Not p.Exists("y") Then
        BotLeftClick = BotError("x and y are required")
        Exit Function
    End If
    Dim tx As Integer, ty As Integer
    tx = CInt(p("x"))
    ty = CInt(p("y"))
    ' Find NPC at target tile and set as target
    Dim npcIdx As Integer
    For npcIdx = 1 To MAXNPCS
        If Npclist(npcIdx).Activo Then
            If Npclist(npcIdx).Pos.Map = UserList(i).Pos.Map And _
               Npclist(npcIdx).Pos.X = tx And Npclist(npcIdx).Pos.Y = ty Then
                UserList(i).flags.TargetNPC = npcIdx
                BotLeftClick = "{""ok"":true,""targeted"":""npc"",""id"":" & npcIdx & "}"
                Exit Function
            End If
        End If
    Next npcIdx
    ' Find player at target tile
    Dim pIdx As Integer
    For pIdx = 1 To MaxUsers
        If pIdx <> i Then
            If UserList(pIdx).flags.UserLogged Then
                If UserList(pIdx).Pos.Map = UserList(i).Pos.Map And _
                   UserList(pIdx).Pos.X = tx And UserList(pIdx).Pos.Y = ty Then
                    UserList(i).flags.TargetUser = pIdx
                    BotLeftClick = "{""ok"":true,""targeted"":""player"",""name"":" & _
                        JSON_Encode(UserList(pIdx).Name) & "}"
                    Exit Function
                End If
            End If
        End If
    Next pIdx
    BotLeftClick = "{""ok"":true,""targeted"":""tile"",""x"":" & tx & ",""y"":" & ty & "}"
End Function

Private Function BotSay(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("message") Then
        BotSay = BotError("message is required")
        Exit Function
    End If
    Dim msg As String
    msg = CStr(p("message"))
    With UserList(i)
        If .flags.Muerto = 0 Then
            Call SendData(SendTarget.ToPCArea, i, PrepareMessageChatOverHead(msg, .Char.CharIndex, .flags.ChatColor))
        End If
    End With
    BotSay = "{""ok"":true}"
End Function

Private Function BotYell(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("message") Then
        BotYell = BotError("message is required")
        Exit Function
    End If
    Dim msg As String
    msg = CStr(p("message"))
    If UserList(i).flags.Muerto = 0 Then
        Call SendData(SendTarget.ToAll, i, PrepareMessageChatOverHead(msg, UserList(i).Char.CharIndex, UserList(i).flags.ChatColor))
    End If
    BotYell = "{""ok"":true}"
End Function

Private Function BotWhisper(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("message") Or Not p.Exists("target") Then
        BotWhisper = BotError("message and target are required")
        Exit Function
    End If
    Dim msg As String, targetName As String
    msg = CStr(p("message"))
    targetName = CStr(p("target"))
    Dim targetIdx As Integer
    targetIdx = BotFindUser(targetName)
    If targetIdx = 0 Then
        BotWhisper = BotError("Target not found: " & targetName)
        Exit Function
    End If
    ' Send to target and back to sender
    Call WriteConsoleMsg(targetIdx, UserList(i).Name & " te susurra: " & msg, FontTypeNames.FONTTYPE_TALK)
    Call WriteConsoleMsg(i, "Susurraste a " & targetName & ": " & msg, FontTypeNames.FONTTYPE_TALK)
    BotWhisper = "{""ok"":true}"
End Function

Private Function BotPickUp(ByVal i As Integer, ByVal p As Object) As String
    If UserList(i).flags.Muerto = 0 Then
        Call GetObj(i)
    End If
    BotPickUp = "{""ok"":true}"
End Function

Private Function BotEquipItem(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("slot") Then
        BotEquipItem = BotError("slot is required")
        Exit Function
    End If
    Dim slot As Integer
    slot = CInt(p("slot"))
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then
        BotEquipItem = BotError("slot out of range")
        Exit Function
    End If
    If UserList(i).Invent.Object(slot).ObjIndex = 0 Then
        BotEquipItem = BotError("No item in slot " & slot)
        Exit Function
    End If
    Call EquiparInvItem(i, slot)
    BotEquipItem = "{""ok"":true,""slot"":" & slot & "}"
End Function

Private Function BotMeditate(ByVal i As Integer) As String
    UserList(i).flags.Meditando = Not UserList(i).flags.Meditando
    BotMeditate = "{""ok"":true,""meditating"":" & IIf(UserList(i).flags.Meditando, "true", "false") & "}"
End Function

Private Function BotRest(ByVal i As Integer) As String
    UserList(i).flags.Descansar = Not UserList(i).flags.Descansar
    BotRest = "{""ok"":true,""resting"":" & IIf(UserList(i).flags.Descansar, "true", "false") & "}"
End Function

Private Function BotDrop(ByVal i As Integer, ByVal p As Object) As String
    If Not p.Exists("slot") Or Not p.Exists("amount") Then
        BotDrop = BotError("slot and amount are required")
        Exit Function
    End If
    Dim slot As Integer, amount As Integer
    slot = CInt(p("slot"))
    amount = CInt(p("amount"))
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then
        BotDrop = BotError("slot out of range")
        Exit Function
    End If
    Call DropObj(i, slot, amount, UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y)
    BotDrop = "{""ok"":true,""slot"":" & slot & ",""amount"":" & amount & "}"
End Function

' ===========================================================================
' Skill dispatch helper â€” maps eSkill to the correct Trabajo.bas sub
' ===========================================================================
Private Sub DoSkillWork(ByVal i As Integer, ByVal sk As eSkill)
    Select Case sk
        Case eSkill.Talar:        Call DoTalar(i)
        Case eSkill.Mineria:      Call DoMineria(i)
        Case eSkill.pesca:        Call DoPescar(i)
        Case eSkill.Navegacion:   Call DoNavega(i, UserList(i).Char.heading)
        Case eSkill.Ocultarse:    Call DoOcultarse(i)
        Case eSkill.Meditar:
            UserList(i).flags.Meditando = True
        Case eSkill.Carpinteria:  Call CarpinteroConstruirItem(i, UserList(i).flags.TargetObj)
        Case eSkill.Herreria:     Call HerreroConstruirItem(i, UserList(i).flags.TargetObj)
        Case Else
            ' Skills that require a target (Magia, Robar, Domar, etc.) â€” noop;
            ' use castSpell or attack for those
    End Select
End Sub

' ===========================================================================
' aobench: Observer mode (tasks 14.2-14.5)
' ===========================================================================

' Attach socket observerSlot as observer for charname.
' Returns JSON response.
Public Function BotAttachObserver(ByVal observerSlot As Integer, ByVal charname As String) As String
    Dim i As Integer
    i = BotFindUser(charname)
    If i = 0 Then
        BotAttachObserver = BotError("Character not found: " & charname)
        Exit Function
    End If
    ' Reject if already has an observer (task 14.6)
    If UserList(i).ObserverConnID <> -1 Then
        BotAttachObserver = BotError("Already has an observer")
        Exit Function
    End If
    UserList(i).ObserverConnID = observerSlot
    ' Replay current state to observer immediately
    Dim stateJson As String
    stateJson = BotGetState(i)
    Call frmBotAPI.SendToSlot(observerSlot, stateJson)
    BotAttachObserver = "{""ok"":true,""observing"":" & JSON_Encode(charname) & "}"
End Function

' Detach observer from user slot i.
Public Function BotDetachObserver(ByVal i As Integer) As String
    UserList(i).ObserverConnID = -1
    BotDetachObserver = "{""ok"":true}"
End Function

' Called after each packet is written to a user's outgoingData.
' If this user has an observer, mirror the packet to the observer socket.
Public Sub BotMirrorToObserver(ByVal Userindex As Integer, ByVal data As String)
    If UserList(Userindex).ObserverConnID > 0 Then
        Call frmBotAPI.SendToSlot(UserList(Userindex).ObserverConnID, data)
    End If
End Sub

' Returns True if slot is currently observing any bot.
Public Function BotIsObserverSlot(ByVal slot As Integer) As Boolean
    Dim i As Integer
    For i = 1 To MaxUsers
        If UserList(i).flags.UserLogged Then
            If UserList(i).ObserverConnID = slot Then
                BotIsObserverSlot = True
                Exit Function
            End If
        End If
    Next i
    BotIsObserverSlot = False
End Function

' ===========================================================================
' Helper: inject bytes into incomingData then call the handler.
' Handlers read their own params from incomingData â€” this is the cleanest
' way to reuse 100% of existing game logic including all validation.
' ===========================================================================
Private Sub InjectByte(ByVal i As Integer, ByVal v As Byte)
    Call UserList(i).incomingData.WriteByte(v)
End Sub
Private Sub InjectInteger(ByVal i As Integer, ByVal v As Integer)
    Call UserList(i).incomingData.WriteInteger(v)
End Sub
Private Sub InjectLong(ByVal i As Integer, ByVal v As Long)
    Call UserList(i).incomingData.WriteLong(v)
End Sub
Private Sub InjectString(ByVal i As Integer, ByVal v As String)
    Call UserList(i).incomingData.WriteASCIIString(v)
End Sub

' ===========================================================================
' Extended BotDispatch â€” add to the Select Case in BotDispatch
' Call this from BotDispatch for the new action names.
' ===========================================================================
Public Function BotDispatchExtended(ByVal action As String, _
                                     ByVal i As Integer, _
                                     ByVal p As Object) As String
    Select Case action

    ' â”€â”€ Player actions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "doubleClick"
        ' Double-click on tile offset (dx, dy) â€” opens NPC dialogue, doors, objects
        If Not p.Exists("x") Or Not p.Exists("y") Then
            BotDispatchExtended = BotError("x and y are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.DoubleClick
        InjectByte i, CByte(CInt(p("x")) And 255)
        InjectByte i, CByte(CInt(p("y")) And 255)
        Call HandleDoubleClick(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "changeHeading"
        ' Change facing direction without moving
        If Not p.Exists("heading") Then
            BotDispatchExtended = BotError("heading is required")
            Exit Function
        End If
        Dim h As eHeading
        h = ParseHeading(CStr(p("heading")))
        If h = 0 Then
            BotDispatchExtended = BotError("Invalid heading: " & CStr(p("heading")))
            Exit Function
        End If
        InjectByte i, ClientPacketID.ChangeHeading
        InjectByte i, CByte(h)
        Call HandleChangeHeading(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "workLeftClick"
        ' Use work skill on a specific tile (mining node, tree, etc.)
        If Not p.Exists("x") Or Not p.Exists("y") Then
            BotDispatchExtended = BotError("x and y are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.WorkLeftClick
        InjectByte i, CByte(CInt(p("x")) And 255)
        InjectByte i, CByte(CInt(p("y")) And 255)
        Call HandleWorkLeftClick(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "useSpellMacro"
        ' Cast spell by macro slot number
        InjectByte i, ClientPacketID.UseSpellMacro
        Call HandleUseSpellMacro(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "moveItem"
        ' Rearrange inventory: move item from one slot to another
        If Not p.Exists("from") Or Not p.Exists("to") Then
            BotDispatchExtended = BotError("from and to slots are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.moveItem
        InjectByte i, CByte(CInt(p("from")))
        InjectByte i, CByte(CInt(p("to")))
        Call HandleMoveItem(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "commerceStart"
        ' Open NPC commerce window (click on merchant NPC first via leftClick)
        InjectByte i, ClientPacketID.CommerceStart
        Call HandleCommerceStart(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "commerceBuy"
        ' Buy item from NPC shop: slot = NPC inventory slot, amount = quantity
        If Not p.Exists("slot") Or Not p.Exists("amount") Then
            BotDispatchExtended = BotError("slot and amount are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.CommerceBuy
        InjectByte i, CByte(CInt(p("slot")))
        InjectInteger i, CInt(p("amount"))
        Call HandleCommerceBuy(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "commerceSell"
        ' Sell item to NPC: slot = player inventory slot, amount = quantity
        If Not p.Exists("slot") Or Not p.Exists("amount") Then
            BotDispatchExtended = BotError("slot and amount are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.CommerceSell
        InjectByte i, CByte(CInt(p("slot")))
        InjectInteger i, CInt(p("amount"))
        Call HandleCommerceSell(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "commerceEnd"
        InjectByte i, ClientPacketID.CommerceEnd
        Call HandleCommerceEnd(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "userCommerceOffer"
        ' Offer item to another player in a trade: slot, amount, goldAmount
        If Not p.Exists("slot") Then
            BotDispatchExtended = BotError("slot is required")
            Exit Function
        End If
        Dim ucAmount As Integer
        Dim ucGold As Long
        ucAmount = IIf(p.Exists("amount"), CInt(p("amount")), 1)
        ucGold = IIf(p.Exists("gold"), CLng(p("gold")), 0)
        InjectByte i, ClientPacketID.UserCommerceOffer
        InjectByte i, CByte(CInt(p("slot")))
        InjectInteger i, ucAmount
        InjectLong i, ucGold
        Call HandleUserCommerceOffer(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "userCommerceOk"
        InjectByte i, ClientPacketID.UserCommerceOk
        Call HandleUserCommerceOk(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "userCommerceReject"
        InjectByte i, ClientPacketID.UserCommerceReject
        Call HandleUserCommerceReject(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "userCommerceEnd"
        InjectByte i, ClientPacketID.UserCommerceEnd
        Call HandleUserCommerceEnd(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "bankStart"
        InjectByte i, ClientPacketID.BankStart
        Call HandleBankStart(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "bankDeposit"
        If Not p.Exists("slot") Or Not p.Exists("amount") Then
            BotDispatchExtended = BotError("slot and amount are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.BankDeposit
        InjectByte i, CByte(CInt(p("slot")))
        InjectInteger i, CInt(p("amount"))
        Call HandleBankDeposit(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "bankExtractItem"
        If Not p.Exists("slot") Or Not p.Exists("amount") Then
            BotDispatchExtended = BotError("slot and amount are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.BankExtractItem
        InjectByte i, CByte(CInt(p("slot")))
        InjectInteger i, CInt(p("amount"))
        Call HandleBankExtractItem(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "bankDepositGold"
        If Not p.Exists("amount") Then
            BotDispatchExtended = BotError("amount is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.BankDepositGold
        InjectLong i, CLng(p("amount"))
        Call HandleBankDepositGold(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "bankExtractGold"
        If Not p.Exists("amount") Then
            BotDispatchExtended = BotError("amount is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.BankExtractGold
        InjectLong i, CLng(p("amount"))
        Call HandleBankExtractGold(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "bankEnd"
        InjectByte i, ClientPacketID.BankEnd
        Call HandleBankEnd(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "craftBlacksmith"
        ' Craft item at blacksmith: itemIndex = item to craft
        If Not p.Exists("item") Then
            BotDispatchExtended = BotError("item is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.CraftBlacksmith
        InjectInteger i, CInt(p("item"))
        Call HandleCraftBlacksmith(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "craftCarpenter"
        If Not p.Exists("item") Then
            BotDispatchExtended = BotError("item is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.CraftCarpenter
        InjectInteger i, CInt(p("item"))
        Call HandleCraftCarpenter(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "train"
        ' Train skill at NPC trainer: npcIndex = NPC, skill = eSkill
        If Not p.Exists("npc") Or Not p.Exists("skill") Then
            BotDispatchExtended = BotError("npc and skill are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.Train
        InjectInteger i, CInt(p("npc"))
        InjectByte i, CByte(CInt(p("skill")))
        Call HandleTrain(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "partyCreate"
        InjectByte i, ClientPacketID.PartyCreate
        Call HandlePartyCreate(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "partyJoin"
        ' Join a party: leaderName = name of party leader
        If Not p.Exists("leader") Then
            BotDispatchExtended = BotError("leader is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.PartyJoin
        InjectString i, CStr(p("leader"))
        Call HandlePartyJoin(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "partyLeave"
        InjectByte i, ClientPacketID.PartyLeave
        Call HandlePartyLeave(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "partyMessage"
        If Not p.Exists("message") Then
            BotDispatchExtended = BotError("message is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.PartyMessage
        InjectString i, CStr(p("message"))
        Call HandlePartyMessage(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "guildMessage"
        If Not p.Exists("message") Then
            BotDispatchExtended = BotError("message is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GuildMessage
        InjectString i, CStr(p("message"))
        Call HandleGuildMessage(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "shareNpc"
        InjectByte i, ClientPacketID.ShareNpc
        Call HandleShareNpc(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "stopSharingNpc"
        InjectByte i, ClientPacketID.StopSharingNpc
        Call HandleStopSharingNpc(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "resucitate"
        ' Ask NPC to resurrect the (dead) character
        InjectByte i, ClientPacketID.Resucitate
        Call HandleResucitate(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "heal"
        ' Ask NPC healer to heal the character
        InjectByte i, ClientPacketID.Heal
        Call HandleHeal(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "home"
        ' Teleport to home city
        InjectByte i, ClientPacketID.Home
        Call HandleHome(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "enlist"
        ' Join army faction (Royal Army or Chaos Legion)
        InjectByte i, ClientPacketID.Enlist
        Call HandleEnlist(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "changeDescription"
        If Not p.Exists("desc") Then
            BotDispatchExtended = BotError("desc is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.ChangeDescription
        InjectString i, CStr(p("desc"))
        Call HandleChangeDescription(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "dragSpell"
        ' Drag spell from one macro slot to another
        If Not p.Exists("from") Or Not p.Exists("to") Then
            BotDispatchExtended = BotError("from and to are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.DragAndDropHechizos
        InjectByte i, CByte(CInt(p("from")))
        InjectByte i, CByte(CInt(p("to")))
        Call HandleDragAndDropHechizos(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ GM actions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    ' These require the character to have GM privileges (flags.Privilegios)

    Case "gmGoToChar"
        ' Teleport self to target character: target = charname
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.GoToChar)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmWarpChar"
        ' Teleport target character to map/x/y
        If Not p.Exists("target") Or Not p.Exists("map") Or Not p.Exists("x") Or Not p.Exists("y") Then
            BotDispatchExtended = BotError("target, map, x and y are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.WarpChar)
        InjectString i, CStr(p("target"))
        InjectInteger i, CInt(p("map"))
        InjectByte i, CByte(CInt(p("x")))
        InjectByte i, CByte(CInt(p("y")))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmWarpMe"
        ' Teleport self to current target's location
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.WarpMeToTarget)
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmSummon"
        ' Summon target character to self's location
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.SummonChar)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmSpawn"
        ' Spawn NPC: npcIndex = NPC number from NPCs.dat
        If Not p.Exists("npc") Then
            BotDispatchExtended = BotError("npc is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.SpawnCreature)
        InjectInteger i, CInt(p("npc"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmKillNpc"
        ' Kill NPC by its UserList index
        If Not p.Exists("npc") Then
            BotDispatchExtended = BotError("npc is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.KillNPC)
        InjectInteger i, CInt(p("npc"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmKillAllNearby"
        ' Kill all NPCs near self
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.KillAllNearbyNPCs)
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmCreateItem"
        ' Create item in inventory: itemIndex = item number, amount = quantity
        If Not p.Exists("item") Then
            BotDispatchExtended = BotError("item is required")
            Exit Function
        End If
        Dim ciAmount As Integer
        ciAmount = IIf(p.Exists("amount"), CInt(p("amount")), 1)
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.CreateItem)
        InjectInteger i, CInt(p("item"))
        InjectInteger i, ciAmount
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmDestroyItems"
        ' Destroy all items on the ground in area
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.DestroyItems)
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmRevive"
        ' Revive a dead character
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.ReviveChar)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmKick"
        ' Kick a character from the server
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.Kick)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmBan"
        If Not p.Exists("target") Or Not p.Exists("reason") Then
            BotDispatchExtended = BotError("target and reason are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.BanChar)
        InjectString i, CStr(p("target"))
        InjectString i, CStr(p("reason"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmSilence"
        ' Silence a character (can't talk)
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.Silence)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmJail"
        ' Send character to jail
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.Jail)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmInvisible"
        ' Toggle GM invisibility
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.invisible)
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmMessage"
        ' Send GM broadcast message to all players
        If Not p.Exists("message") Then
            BotDispatchExtended = BotError("message is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.GMMessage)
        InjectString i, CStr(p("message"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmServerMessage"
        ' Send server-wide message
        If Not p.Exists("message") Then
            BotDispatchExtended = BotError("message is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.ServerMessage)
        InjectString i, CStr(p("message"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmRain"
        ' Toggle rain weather
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.RainToggle)
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmTeleportCreate"
        ' Create a teleport tile at map/x/y -> destMap/destX/destY
        If Not p.Exists("map") Or Not p.Exists("x") Or Not p.Exists("y") Or _
           Not p.Exists("destMap") Or Not p.Exists("destX") Or Not p.Exists("destY") Then
            BotDispatchExtended = BotError("map, x, y, destMap, destX, destY are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.TeleportCreate)
        InjectInteger i, CInt(p("map"))
        InjectByte i, CByte(CInt(p("x")))
        InjectByte i, CByte(CInt(p("y")))
        InjectInteger i, CInt(p("destMap"))
        InjectByte i, CByte(CInt(p("destX")))
        InjectByte i, CByte(CInt(p("destY")))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmEditChar"
        ' Modify a character stat: target, field (eEditOptions), value
        If Not p.Exists("target") Or Not p.Exists("field") Or Not p.Exists("value") Then
            BotDispatchExtended = BotError("target, field and value are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.EditChar)
        InjectString i, CStr(p("target"))
        InjectByte i, CByte(CInt(p("field")))
        InjectLong i, CLng(p("value"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmSetDesc"
        ' Set another character's description
        If Not p.Exists("target") Or Not p.Exists("desc") Then
            BotDispatchExtended = BotError("target and desc are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.SetCharDescription)
        InjectString i, CStr(p("target"))
        InjectString i, CStr(p("desc"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmCreateNpc"
        ' Create NPC at current position: npcType, heading
        If Not p.Exists("npc") Then
            BotDispatchExtended = BotError("npc is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.CreateNPC)
        InjectInteger i, CInt(p("npc"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmMakeDumb"
        ' Make character unable to cast spells
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.MakeDumb)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmMakeDumbNoMore"
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.MakeDumbNoMore)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "gmRequestCharInfo"
        ' Get character info
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GMCommands
        InjectByte i, CByte(eGMCommands.RequestCharInfo)
        InjectString i, CStr(p("target"))
        Call HandleGMCommands(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Quest actions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "quest"
        ' Talk to quest NPC at (x,y) or the current TargetNPC.
        ' Sends Quest packet: server shows details on first call, finishes on second.
        If p.Exists("x") And p.Exists("y") Then
            ' First set the target NPC by walking onto its tile in the Accion sense.
            ' We inject a DoubleClick so Accion() sets TargetNPC, then immediately
            ' inject Quest so HandleQuest() can use it.
            InjectByte i, ClientPacketID.DoubleClick
            InjectByte i, CByte(CInt(p("x")) And 255)
            InjectByte i, CByte(CInt(p("y")) And 255)
            Call HandleDoubleClick(i)
        End If
        InjectByte i, ClientPacketID.Quest
        Call Quests.HandleQuest(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "questAccept"
        ' Accept the quest that was just offered via quest action.
        InjectByte i, ClientPacketID.QuestAccept
        InjectByte i, 0  ' padding byte read by HandleQuestAccept
        Call Quests.HandleQuestAccept(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "questAbandon"
        ' Abandon an active quest by slot (1-based).
        If Not p.Exists("slot") Then
            BotDispatchExtended = BotError("slot is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.QuestAbandon
        InjectByte i, CByte(CInt(p("slot")))
        Call Quests.HandleQuestAbandon(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "questListRequest"
        ' Request the full quest list for this character (active + done).
        InjectByte i, ClientPacketID.QuestListRequest
        Call Quests.HandleQuestListRequest(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "questDetailsRequest"
        ' Get details for a specific active quest by slot number.
        If Not p.Exists("slot") Then
            BotDispatchExtended = BotError("slot is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.QuestDetailsRequest
        InjectByte i, CByte(CInt(p("slot")))
        Call Quests.HandleQuestDetailsRequest(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Skill assignment â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "assignSkillPoints"
        ' Assign free skill points. Pass a distribution object, e.g.:
        ' {"Armas":3,"Defensa":4,"Tacticas":3}
        ' All values must sum to <= SkillPts. Unmentioned skills get 0.
        If Not p.Exists("distribution") Then
            BotDispatchExtended = BotError("distribution is required")
            Exit Function
        End If
        Dim dist As Object
        Set dist = p("distribution")
        Dim pts(1 To NUMSKILLS) As Byte
        Dim totalPts As Integer
        Dim sName As String
        Dim sk2 As eSkill
        ' Parse distribution
        Dim dKey As Variant
        For Each dKey In dist.Keys()
            sName = CStr(dKey)
            sk2 = ParseSkill(sName)
            If sk2 = 0 Then
                BotDispatchExtended = BotError("Unknown skill: " & sName)
                Exit Function
            End If
            pts(CInt(sk2)) = CByte(CInt(dist(dKey)))
            totalPts = totalPts + CInt(dist(dKey))
        Next dKey
        If totalPts > UserList(i).Stats.SkillPts Then
            BotDispatchExtended = BotError("Not enough skill points (have " & UserList(i).Stats.SkillPts & ", need " & totalPts & ")")
            Exit Function
        End If
        ' Inject ModifySkills packet: 1 byte header + NUMSKILLS bytes
        InjectByte i, ClientPacketID.ModifySkills
        Dim skK As Integer
        For skK = 1 To NUMSKILLS
            InjectByte i, pts(skK)
        Next skK
        Call HandleModifySkills(i)
        BotDispatchExtended = "{""ok"":true,""pointsSpent"":" & totalPts & "}"

    ' â”€â”€ Party actions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "partyAccept"
        ' Accept a pending party join request from player: member = name
        If Not p.Exists("member") Then
            BotDispatchExtended = BotError("member is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.PartyAcceptMember
        InjectString i, CStr(p("member"))
        Call HandlePartyAcceptMember(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "partyKick"
        ' Kick a member from your party: member = name
        If Not p.Exists("member") Then
            BotDispatchExtended = BotError("member is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.PartyKick
        InjectString i, CStr(p("member"))
        Call HandlePartyKick(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "partySetLeader"
        ' Transfer party leadership: member = name of new leader
        If Not p.Exists("member") Then
            BotDispatchExtended = BotError("member is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.PartySetLeader
        InjectString i, CStr(p("member"))
        Call HandlePartySetLeader(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "partyOnline"
        ' List party members online with their XP
        InjectByte i, ClientPacketID.PartyOnline
        Call HandlePartyOnline(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "requestPartyForm"
        ' Request party formation info
        InjectByte i, ClientPacketID.RequestPartyForm
        Call HandlePartyForm(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Faction â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "leaveFaction"
        ' Leave current army faction (/RETIRAR)
        InjectByte i, ClientPacketID.LeaveFaction
        Call HandleLeaveFaction(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Pet actions â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "petStand"
        ' Order pet to stand still (/QUIETO)
        InjectByte i, ClientPacketID.PetStand
        Call HandlePetStand(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "petFollow"
        ' Order pet to follow owner (/ACOMPANAR)
        InjectByte i, ClientPacketID.PetFollow
        Call HandlePetFollow(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "releasePet"
        ' Release / free pet (/LIBERAR)
        InjectByte i, ClientPacketID.ReleasePet
        Call HandleReleasePet(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Item upgrade â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "itemUpgrade"
        ' Upgrade an item: slot = inventory slot of item to upgrade
        If Not p.Exists("slot") Then
            BotDispatchExtended = BotError("slot is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.ItemUpgrade
        InjectByte i, CByte(CInt(p("slot")))
        Call HandleItemUpgrade(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Craftsman (artisan) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "initCrafting"
        ' Open the craftsman's dialogue (doubleClick the artisan NPC first)
        InjectByte i, ClientPacketID.InitCrafting
        Call HandleInitCrafting(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "craftsmanCreate"
        ' Create an item via the craftsman. item = recipe/item index
        If Not p.Exists("item") Then
            BotDispatchExtended = BotError("item is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.CraftsmanCreate
        InjectInteger i, CInt(p("item"))
        Call HandleCraftsmanCreate(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Combat challenges (retos / fights) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "fightSend"
        ' Send a duel challenge to a player: target = name, gold = wager
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.FightSend
        InjectString i, CStr(p("target"))
        InjectLong i, IIf(p.Exists("gold"), CLng(p("gold")), 0)
        Call HandleFightSend(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "fightAccept"
        ' Accept an incoming duel challenge
        InjectByte i, ClientPacketID.FightAccept
        Call HandleFightAccept(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ User commerce (player-to-player trade) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "userCommerceCreate"
        ' Initiate a trade with a nearby player: target = player name
        If Not p.Exists("target") Then
            BotDispatchExtended = BotError("target is required")
            Exit Function
        End If
        ' Find the target user index and set TargetUser
        Dim tradeTarget As Integer
        tradeTarget = BotFindUser(CStr(p("target")))
        If tradeTarget = 0 Then
            BotDispatchExtended = BotError("Player not found: " & CStr(p("target")))
            Exit Function
        End If
        UserList(i).flags.TargetUser = tradeTarget
        ' UserCommerceCreate is handled client-side; server uses TargetUser
        ' when UserCommerceOffer is sent. Call IniciarComercioConUsuario directly.
        Call IniciarComercioConUsuario(i, tradeTarget)
        BotDispatchExtended = "{""ok"":true}"

    Case "userCommerceConfirm"
        ' Confirm the current trade (both sides must confirm)
        InjectByte i, ClientPacketID.UserCommerceConfirm
        Call HandleUserCommerceConfirm(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "commerceChat"
        ' Send a chat message during NPC commerce
        If Not p.Exists("message") Then
            BotDispatchExtended = BotError("message is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.CommerceChat
        InjectString i, CStr(p("message"))
        Call HandleCommerceChat(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Bank â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "moveBank"
        ' Move an item from inventory to bank or bank to inventory
        ' fromBank=true moves bank->inv, false moves inv->bank
        If Not p.Exists("slot") Then
            BotDispatchExtended = BotError("slot is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.MoveBank
        InjectByte i, CByte(CInt(p("slot")))
        InjectByte i, IIf(p.Exists("fromBank") And CBool(p("fromBank")), 1, 0)
        Call HandleMoveBank(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Spells â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "moveSpell"
        ' Move a spell from one hotbar slot to another
        If Not p.Exists("from") Or Not p.Exists("to") Then
            BotDispatchExtended = BotError("from and to are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.MoveSpell
        InjectByte i, CByte(CInt(p("from")))
        InjectByte i, CByte(CInt(p("to")))
        Call HandleMoveSpell(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Forum â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "forumPost"
        ' Post a message to the in-game forum
        ' type: 0=general, 1=real, 2=caos â€” title, body
        If Not p.Exists("title") Or Not p.Exists("body") Then
            BotDispatchExtended = BotError("title and body are required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.ForumPost
        InjectByte i, IIf(p.Exists("type"), CByte(CInt(p("type"))), 0)
        InjectString i, CStr(p("title"))
        InjectString i, CStr(p("body"))
        Call HandleForumPost(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Friends (amigos) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "addAmigo"
        ' Add a player to friends list: name = player name
        If Not p.Exists("name") Then
            BotDispatchExtended = BotError("name is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.AddAmigos
        InjectString i, CStr(p("name"))
        Call Amigos.HandleAddAmigo(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "delAmigo"
        ' Remove a player from friends list: name = player name
        If Not p.Exists("name") Then
            BotDispatchExtended = BotError("name is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.DelAmigos
        InjectString i, CStr(p("name"))
        Call Amigos.HandleDelAmigo(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "onAmigos"
        ' List online friends
        InjectByte i, ClientPacketID.OnAmigos
        Call Amigos.HandleOnAmigo(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "msgAmigos"
        ' Send a message to all online friends: message = text
        If Not p.Exists("message") Then
            BotDispatchExtended = BotError("message is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.MsgAmigos
        InjectString i, CStr(p("message"))
        Call Amigos.HandleMsgAmigo(i)
        BotDispatchExtended = "{""ok"":true}"

    ' â”€â”€ Misc â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    Case "lookProcess"
        ' Look up information about another character: name = player name
        If Not p.Exists("name") Then
            BotDispatchExtended = BotError("name is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.LookProcess
        InjectString i, CStr(p("name"))
        Call HandleLookProcess(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "consultation"
        ' Open the consultation / help panel
        InjectByte i, ClientPacketID.Consultation
        Call HandleConsultation(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "discord"
        ' Link Discord account: code = discord auth code
        If Not p.Exists("code") Then
            BotDispatchExtended = BotError("code is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.Discord
        InjectString i, CStr(p("code"))
        Call HandleDiscord(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "online"
        ' Request online player count / list
        InjectByte i, ClientPacketID.Online
        Call HandleOnline(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "guildFundate"
        ' Initiate guild founding (/FUNDARCLAN): guildName = name
        If Not p.Exists("guildName") Then
            BotDispatchExtended = BotError("guildName is required")
            Exit Function
        End If
        InjectByte i, ClientPacketID.GuildFundate
        InjectString i, CStr(p("guildName"))
        Call HandleGuildFundate(i)
        BotDispatchExtended = "{""ok"":true}"

    Case "guildFundation"
        ' Confirm/complete guild founding (second packet after guildFundate)
        InjectByte i, ClientPacketID.GuildFundation
        Call HandleGuildFundation(i)
        BotDispatchExtended = "{""ok"":true}"

    Case Else
        BotDispatchExtended = ""  ' Not handled here â€” caller will fall through to BotError

    End Select
End Function
