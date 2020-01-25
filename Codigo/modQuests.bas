Attribute VB_Name = "Quests"
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
'along with this program; if not, you can find it at [url=http://www.affero.org/oagpl.html]http://www.affero.org/oagpl.html[/url]
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at [email=aaron@baronsoft.com]aaron@baronsoft.com[/email]
'for more information about ORE please visit [url=http://www.baronsoft.com/]http://www.baronsoft.com/[/url]
Option Explicit
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 15     'Maxima cantidad de quests que puede tener un usuario al mismo tiempo.
 
Public Function TieneQuest(ByVal Userindex As Integer, _
                           ByVal QuestNumber As Integer) As Byte

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
 
    For i = 1 To MAXUSERQUESTS

        If UserList(Userindex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
            TieneQuest = i
            Exit Function

        End If

    Next i
    
    TieneQuest = 0

End Function
 
Public Function FreeQuestSlot(ByVal Userindex As Integer) As Byte

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Devuelve el proximo slot de quest libre.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
 
    For i = 1 To MAXUSERQUESTS

        If UserList(Userindex).QuestStats.Quests(i).QuestIndex = 0 Then
            FreeQuestSlot = i
            Exit Function

        End If

    Next i
    
    FreeQuestSlot = 0

End Function
 
Public Sub HandleQuestAccept(ByVal Userindex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el evento de aceptar una quest.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim NpcIndex  As Integer

    Dim QuestSlot As Byte
 
    Call UserList(Userindex).incomingData.ReadByte
 
    NpcIndex = UserList(Userindex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    QuestSlot = FreeQuestSlot(Userindex)
    
    'Agregamos la quest.
    With UserList(Userindex).QuestStats.Quests(QuestSlot)
        .QuestIndex = Npclist(NpcIndex).QuestNumber
        
        If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
        Call WriteConsoleMsg(Userindex, "Has aceptado la mision " & Chr(34) & QuestList(.QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With

End Sub
 
Public Sub FinishQuest(ByVal Userindex As Integer, _
                       ByVal QuestIndex As Integer, _
                       ByVal QuestSlot As Byte)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el evento de terminar una quest.
    'Last modified: 29/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i              As Integer

    Dim InvSlotsLibres As Byte

    Dim NpcIndex       As Integer
 
    NpcIndex = UserList(Userindex).flags.TargetNPC
    
    With QuestList(QuestIndex)

        'Comprobamos que tenga los objetos.
        If .RequiredOBJs > 0 Then

            For i = 1 To .RequiredOBJs

                If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, Userindex) = False Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub

                End If

            Next i

        End If
        
        'Comprobamos que haya matado todas las criaturas.
        If .RequiredNPCs > 0 Then

            For i = 1 To .RequiredNPCs

                If .RequiredNPC(i).Amount > UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub

                End If

            Next i

        End If
    
        'Comprobamos que el usuario tenga espacio para recibir los items.
        If .RewardOBJs > 0 Then

            'Buscamos la cantidad de slots de inventario libres.
            For i = 1 To MAX_INVENTORY_SLOTS

                If UserList(Userindex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
            Next i
            
            'Nos fijamos si entra
            If InvSlotsLibres < .RewardOBJs Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                Exit Sub

            End If

        End If
    
        'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
        Call WriteConsoleMsg(Userindex, "Has completado la mision " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & "!", FontTypeNames.FONTTYPE_INFO)
        
        'Si la quest pedia objetos, se los saca al personaje.
        If .RequiredOBJs Then

            For i = 1 To .RequiredOBJs
                Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, Userindex)
            Next i

        End If
        
        'Se entrega la experiencia.
        If .RewardEXP Then
            UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + .RewardEXP
            Call WriteConsoleMsg(Userindex, "Has ganado " & .RewardEXP & " puntos de experiencia como recompensa.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'Se entrega el oro.
        If .RewardGLD Then
            UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + .RewardGLD
            Call WriteConsoleMsg(Userindex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'Si hay recompensa de objetos, se entregan.
        If .RewardOBJs > 0 Then

            For i = 1 To .RewardOBJs

                If .RewardOBJ(i).Amount Then
                    Call MeterItemEnInventario(Userindex, .RewardOBJ(i))
                    Call WriteConsoleMsg(Userindex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFO)

                End If

            Next i

        End If
    
        'Actualizamos el personaje
        Call CheckUserLevel(Userindex)
        Call UpdateUserInv(True, Userindex, 0)
    
        'Limpiamos el slot de quest.
        Call CleanQuestSlot(Userindex, QuestSlot)
        
        'Ordenamos las quests
        Call ArrangeUserQuests(Userindex)
    
        'Se agrega que el usuario ya hizo esta quest.
        Call AddDoneQuest(Userindex, QuestIndex)

    End With

End Sub
 
Public Sub AddDoneQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Agrega la quest QuestIndex a la lista de quests hechas.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    With UserList(Userindex).QuestStats
        .NumQuestsDone = .NumQuestsDone + 1
        ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
        .QuestsDone(.NumQuestsDone) = QuestIndex

    End With

End Sub
 
Public Function UserDoneQuest(ByVal Userindex As Integer, _
                              ByVal QuestIndex As Integer) As Boolean

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Verifica si el usuario hizo la quest QuestIndex.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer

    With UserList(Userindex).QuestStats

        If .NumQuestsDone Then

            For i = 1 To .NumQuestsDone

                If .QuestsDone(i) = QuestIndex Then
                    UserDoneQuest = True
                    Exit Function

                End If

            Next i

        End If

    End With
    
    UserDoneQuest = False
        
End Function
 
Public Sub CleanQuestSlot(ByVal Userindex As Integer, ByVal QuestSlot As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Limpia un slot de quest de un usuario.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
 
    With UserList(Userindex).QuestStats.Quests(QuestSlot)

        If .QuestIndex Then
            If QuestList(.QuestIndex).RequiredNPCs Then

                For i = 1 To QuestList(.QuestIndex).RequiredNPCs
                    .NPCsKilled(i) = 0
                Next i

            End If

        End If

        .QuestIndex = 0

    End With

End Sub
 
Public Sub ResetQuestStats(ByVal Userindex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Limpia todos los QuestStats de un usuario
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
 
    For i = 1 To MAXUSERQUESTS
        Call CleanQuestSlot(Userindex, i)
    Next i
    
    With UserList(Userindex).QuestStats
        .NumQuestsDone = 0
        Erase .QuestsDone

    End With

End Sub
 
Public Sub HandleQuest(ByVal Userindex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete Quest.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim NpcIndex As Integer

    Dim tmpByte  As Byte
 
    'Leemos el paquete
    Call UserList(Userindex).incomingData.ReadByte
 
    NpcIndex = UserList(Userindex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    'El NPC hace quests?
    If Npclist(NpcIndex).QuestNumber = 0 Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub

    End If
    
    'El personaje ya hizo la quest?
    If UserDoneQuest(Userindex, Npclist(NpcIndex).QuestNumber) Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Ya has hecho una mision para mi.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub

    End If
 
    'El personaje tiene suficiente nivel?
    If UserList(Userindex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub

    End If
    
    'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho
 
    tmpByte = TieneQuest(Userindex, Npclist(NpcIndex).QuestNumber)
    
    If tmpByte Then
        'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
        Call FinishQuest(Userindex, Npclist(NpcIndex).QuestNumber, tmpByte)
    Else
        'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
        tmpByte = FreeQuestSlot(Userindex)
        
        'El personaje tiene algun slot de quest para la nueva quest?
        If tmpByte = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub

        End If
        
        'Enviamos los detalles de la quest
        Call WriteQuestDetails(Userindex, Npclist(NpcIndex).QuestNumber)

    End If

End Sub
 
Public Sub LoadQuests()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga el archivo QUESTS.DAT en el array QuestList.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    On Error GoTo ErrorHandler

    Dim Reader    As clsIniManager

    Dim NumQuests As Integer

    Dim tmpStr    As String

    Dim i         As Integer

    Dim j         As Integer
    
    'Cargamos el clsIniManager en memoria
    Set Reader = New clsIniManager
    
    'Lo inicializamos para el archivo Quests.DAT
    Call Reader.Initialize(DatPath & "Quests.DAT")
    
    'Redimensionamos el array
    NumQuests = Reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests)
    
    'Cargamos los datos
    For i = 1 To NumQuests

        With QuestList(i)
            .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
            .desc = Reader.GetValue("QUEST" & i, "Desc")
            .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            
            'CARGAMOS OBJETOS REQUERIDOS
            .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))

            If .RequiredOBJs > 0 Then
                ReDim .RequiredOBJ(1 To .RequiredOBJs)

                For j = 1 To .RequiredOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    
                    .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            'CARGAMOS NPCS REQUERIDOS
            .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

            If .RequiredNPCs > 0 Then
                ReDim .RequiredNPC(1 To .RequiredNPCs)

                For j = 1 To .RequiredNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    
                    .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNPC(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
            
            'CARGAMOS OBJETOS DE RECOMPENSA
            .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))

            If .RewardOBJs > 0 Then
                ReDim .RewardOBJ(1 To .RewardOBJs)

                For j = 1 To .RewardOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                    
                    .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If

        End With

    Next i
    
    'Eliminamos la clase
    Set Reader = Nothing
    Exit Sub
                    
ErrorHandler:
    MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical

End Sub
 
Public Sub LoadQuestStats(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga las QuestStats del usuario.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i           As Integer

    Dim j           As Integer

    Dim tmpStr      As String

    Dim Fields()    As String
 
    For i = 1 To MAXUSERQUESTS

        With UserList(Userindex).QuestStats.Quests(i)
            tmpStr = UserFile.GetValue("QUESTS", "Q" & i)
            
            ' Para evitar modificar TODOS los charfiles
            If tmpStr = vbNullString Then
                .QuestIndex = 0

            Else
                Fields = Split(tmpStr, "-")

                .QuestIndex = val(Fields(0))

                If .QuestIndex Then
                    If QuestList(.QuestIndex).RequiredNPCs Then
                        ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)

                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                            .NPCsKilled(j) = val(Fields(j))
                        Next j

                    End If

                End If

            End If

        End With

    Next i
    
    With UserList(Userindex).QuestStats
        tmpStr = UserFile.GetValue("QUESTS", "QuestsDone")
        
        If tmpStr = vbNullString Then
            .NumQuestsDone = 0
        
        Else
            Fields = Split(tmpStr, "-")

            .NumQuestsDone = val(Fields(0))

            If .NumQuestsDone Then
                ReDim .QuestsDone(1 To .NumQuestsDone)

                For i = 1 To .NumQuestsDone
                    .QuestsDone(i) = val(Fields(i))
                Next i

            End If

        End If

    End With
                   
End Sub
 
Public Sub SaveQuestStats(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Guarda las QuestStats del usuario.
    'Last modified: 29/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i      As Integer

    Dim j      As Integer

    Dim tmpStr As String
 
    For i = 1 To MAXUSERQUESTS

        With UserList(Userindex).QuestStats.Quests(i)
            tmpStr = .QuestIndex
            
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then

                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        tmpStr = tmpStr & "-" & .NPCsKilled(j)
                    Next j

                End If

            End If
        
            Call UserFile.ChangeValue("QUESTS", "Q" & i, tmpStr)

        End With

    Next i
    
    With UserList(Userindex).QuestStats
        tmpStr = .NumQuestsDone
        
        If .NumQuestsDone Then

            For i = 1 To .NumQuestsDone
                tmpStr = tmpStr & "-" & .QuestsDone(i)
            Next i

        End If
        
        Call UserFile.ChangeValue("QUESTS", "QuestsDone", tmpStr)

    End With

End Sub
 
Public Sub HandleQuestListRequest(ByVal Userindex As Integer)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete QuestListRequest.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
 
    'Leemos el paquete
    Call UserList(Userindex).incomingData.ReadByte
    
    Call WriteQuestListSend(Userindex)

End Sub
 
Public Sub ArrangeUserQuests(ByVal Userindex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Ordena las quests del usuario de manera que queden todas al principio del arreglo.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer

    Dim j As Integer
 
    With UserList(Userindex).QuestStats

        For i = 1 To MAXUSERQUESTS - 1

            If .Quests(i).QuestIndex = 0 Then

                For j = i + 1 To MAXUSERQUESTS

                    If .Quests(j).QuestIndex Then
                        .Quests(i) = .Quests(j)
                        Call CleanQuestSlot(Userindex, j)
                        Exit For

                    End If

                Next j

            End If

        Next i

    End With

End Sub
 
Public Sub HandleQuestDetailsRequest(ByVal Userindex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete QuestInfoRequest.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim QuestSlot As Byte
 
    'Leemos el paquete
    Call UserList(Userindex).incomingData.ReadByte
    
    QuestSlot = UserList(Userindex).incomingData.ReadByte
    
    Call WriteQuestDetails(Userindex, UserList(Userindex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)

End Sub
 
Public Sub HandleQuestAbandon(ByVal Userindex As Integer)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete QuestAbandon.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Leemos el paquete.
    Call UserList(Userindex).incomingData.ReadByte
    
    'Borramos la quest.
    Call CleanQuestSlot(Userindex, UserList(Userindex).incomingData.ReadByte)
    
    'Ordenamos la lista de quests del usuario.
    Call ArrangeUserQuests(Userindex)
    
    'Enviamos la lista de quests actualizada.
    Call WriteQuestListSend(Userindex)

End Sub
