Attribute VB_Name = "modSistemaComercio"
'*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio

    Compra = 1
    Venta = 2

End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, _
                    ByVal Userindex As Integer, _
                    ByVal NpcIndex As Integer, _
                    ByVal Slot As Integer, _
                    ByVal Cantidad As Integer)

    '*************************************************
    'Author: Nacho (Integer)
    'Last Modification: 06/04/2020
    '27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
    '  - 06/13/08 (NicoNZ)
    '07/06/2010: ZaMa - Los objetos se loguean si superan la cantidad de 1k (antes era solo si eran 1k).
    '24/01/2020: WyroX = Reduzco la cantidad de paquetes que se envian, actualizo solo los slots necesarios y solo el oro, no todos los stats.
    '06/04/2020: FrankoH298 - No podemos vender monturas en uso.
    '*************************************************
    Dim Precio As Long

    Dim Objeto As obj

    If Cantidad < 1 Or Slot < 1 Then Exit Sub
    With UserList(Userindex)
        If .flags.Equitando = 1 Then
            If .Invent.MonturaEqpSlot = Slot Then
                Call WriteConsoleMsg(Userindex, "No podes vender tu montura mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
        End If
    End With
    
    If Modo = eModoComercio.Compra Then
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Sub
        ElseIf Cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(Userindex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
            Call Ban(UserList(Userindex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
            UserList(Userindex).flags.Ban = 1
            Call WriteErrorMsg(Userindex, "Has sido baneado por el Sistema AntiCheat.")
            Call CloseSocket(Userindex)
            Exit Sub
        ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then
            Exit Sub
        End If

        If Cantidad > Npclist(NpcIndex).Invent.Object(Slot).Amount Then Cantidad = Npclist(NpcIndex).Invent.Object(Slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
        
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        
        Precio = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor / Descuento(Userindex) * Cantidad) + 0.5)

        If UserList(Userindex).Stats.Gld < Precio Then
            Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not MeterItemEnInventario(Userindex, Objeto) Then Exit Sub

        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - Precio
        Call WriteUpdateGold(Userindex)

        Call QuitarNpcInvItem(NpcIndex, Slot, Cantidad)
        Call UpdateNpcInvToAll(False, NpcIndex, Slot)

        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(Userindex).Name & " compro del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.Amount >= 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(Userindex).Name & " compro del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If

        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(Userindex).Name & " compro " & ObjData(Objeto.ObjIndex).Name)
        End If
        
    ElseIf Modo = eModoComercio.Venta Then

        If Cantidad > UserList(Userindex).Invent.Object(Slot).Amount Then Cantidad = UserList(Userindex).Invent.Object(Slot).Amount

        Objeto.Amount = Cantidad
        Objeto.ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex

        If Objeto.ObjIndex = 0 Then
            Exit Sub
        ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
            Call WriteConsoleMsg(Userindex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then

            If Npclist(NpcIndex).Name <> "SR" Then
                Call WriteConsoleMsg(Userindex, "Las armaduras del ejercito real solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

        ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then

            If Npclist(NpcIndex).Name <> "SC" Then
                Call WriteConsoleMsg(Userindex, "Las armaduras de la legion oscura solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

        ElseIf UserList(Userindex).Invent.Object(Slot).Amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf Slot < LBound(UserList(Userindex).Invent.Object()) Or Slot > UBound(UserList(Userindex).Invent.Object()) Then
            Exit Sub
        ElseIf UserList(Userindex).flags.Privilegios And PlayerType.Consejero Then
            Call WriteConsoleMsg(Userindex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If

        Call QuitarUserInvItem(Userindex, Slot, Cantidad)
        Call UpdateUserInv(False, Userindex, Slot)

        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + Precio
        
        If UserList(Userindex).Stats.Gld > MAXORO Then UserList(Userindex).Stats.Gld = MAXORO

        Call WriteUpdateGold(Userindex)

        Dim NpcSlot As Integer

        NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)

        If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
            Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount

            If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS
            End If

            Call UpdateNpcInvToAll(False, NpcIndex, NpcSlot)
        End If

        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(Userindex).Name & " vendio al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
        ElseIf Objeto.Amount >= 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(Userindex).Name & " vendio al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).Name)
            End If
        End If
    End If

    Call SubirSkill(Userindex, eSkill.Comerciar, True)

End Sub

Public Sub IniciarComercioNPC(ByVal Userindex As Integer)
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    
    '*************************************************
    Call UpdateNpcInv(True, Userindex, UserList(Userindex).flags.TargetNPC, 0)
    UserList(Userindex).flags.Comerciando = True
    Call WriteCommerceInit(Userindex)

End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, _
                              ByVal Objeto As Integer, _
                              ByVal Cantidad As Integer) As Integer
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    SlotEnNPCInv = 1

    Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        SlotEnNPCInv = SlotEnNPCInv + 1

        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
    Loop
    
    If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
        SlotEnNPCInv = 1
        
        Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
            SlotEnNPCInv = SlotEnNPCInv + 1

            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
        Loop
        
        If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    
    End If
    
End Function

Private Function Descuento(ByVal Userindex As Integer) As Single
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    Descuento = 1 + UserList(Userindex).Stats.UserSkills(eSkill.Comerciar) / 100

End Function

''
' Update the inventory of the Npc to the user
'
' @param updateAll if is needed to update all
' @param userIndex The index of the User
' @param npcIndex The index of the NPC
' @param slot The slot to update

Private Sub UpdateNpcInv(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Byte)
'***************************************************
    Dim obj As obj
    Dim LoopC As Byte
    Dim Desc As Single
    Dim val As Single
    
    Desc = Descuento(Userindex)
    
    'Actualiza un solo slot
    If Not UpdateAll Then
        With Npclist(NpcIndex).Invent.Object(Slot)
            obj.ObjIndex = .ObjIndex
            obj.Amount = .Amount
            
            If .ObjIndex > 0 Then
                val = (ObjData(.ObjIndex).Valor) / Desc
            End If
            
            Call WriteChangeNPCInventorySlot(Userindex, Slot, obj, val)
        End With
    Else
    'Actualiza todos los slots
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            With Npclist(NpcIndex).Invent.Object(LoopC)
                obj.ObjIndex = .ObjIndex
                obj.Amount = .Amount
                
                If .ObjIndex > 0 Then
                    val = (ObjData(.ObjIndex).Valor) / Desc
                End If
                
                Call WriteChangeNPCInventorySlot(Userindex, LoopC, obj, val)
            End With
        Next LoopC
    End If
End Sub

''
' Update the inventory of the Npc to all users trading with him
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC
' @param slot The slot to update

Public Sub UpdateNpcInvToAll(ByVal UpdateAll As Boolean, ByVal NpcIndex As Integer, ByVal Slot As Byte)
'***************************************************
    Dim LoopC As Byte
    
    ' Recorremos todos los usuarios
    For LoopC = 1 To LastUser
        With UserList(LoopC)
            ' Si esta comerciando
            If .flags.Comerciando Then
                ' Si el ultimo NPC que cliqueo es el que hay que actualizar
                If .flags.TargetNPC = NpcIndex Then
                    ' Actualizamos el inventario del NPC
                    Call UpdateNpcInv(UpdateAll, LoopC, NpcIndex, Slot)
                End If
            End If
        End With
    Next
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param ObjIndex  El numero de objeto al cual le calculamos el precio de venta

Public Function SalePrice(ByVal ObjIndex As Integer) As Single

    '*************************************************
    'Author: Nicolas (NicoNZ)
    '
    '*************************************************
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
    If ItemNewbie(ObjIndex) Then Exit Function
    
    SalePrice = ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA

End Function
