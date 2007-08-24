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

Public Const REDUCTOR_PRECIOVENTA = 3

Public Sub Comercio(Modo As eModoComercio, UserIndex As Integer, NpcIndex As Integer, Slot As Integer, Cantidad As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
Dim Precio As Long
Dim Objeto As Obj

If Cantidad < 1 Or Slot < 1 Then Exit Sub

If Modo = eModoComercio.Compra Then
    
    Objeto.amount = Cantidad
    Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
    
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Sub
    ElseIf Cantidad > MAX_INVENTORY_OBJS Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
        Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
        UserList(UserIndex).flags.Ban = 1
        Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).amount > 0 Then
        Exit Sub
    End If
    
    If Cantidad > Npclist(NpcIndex).Invent.Object(Slot).amount Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(Slot).amount
    
    Precio = ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor / Descuento(UserIndex) * Cantidad
    
    If UserList(UserIndex).Stats.GLD < Precio Then
        Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    
    If MeterItemEnInventario(UserIndex, Objeto) = False Then
        Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
    
    
    Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(Slot), Cantidad)
    
ElseIf Modo = eModoComercio.Venta Then
    
    If Cantidad > UserList(UserIndex).Invent.Object(Slot).amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).amount
    
    Objeto.amount = Cantidad
    Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
    If Objeto.ObjIndex = 0 Then
        Exit Sub
    ElseIf ObjData(Objeto.ObjIndex).Newbie = 1 Then
        Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio objetos para newbies.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
        Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then
        If Npclist(NpcIndex).name <> "SR" Then
            Call WriteConsoleMsg(UserIndex, "Las armaduras de la Armada solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then
        If Npclist(NpcIndex).name <> "SC" Then
            Call WriteConsoleMsg(UserIndex, "Las armaduras de la Legión solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    ElseIf UserList(UserIndex).Invent.Object(Slot).Equipped <> 0 Then
        Call WriteConsoleMsg(UserIndex, "No puedes vender el item si lo tienes equipado.", FontTypeNames.FONTTYPE_INFO) 'TODO: En vez de hacer esto, que lo desequipe.
        Exit Sub
    ElseIf UserList(UserIndex).Invent.Object(Slot).amount < 0 Or Cantidad = 0 Then
        Exit Sub
    ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then
        Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
        Exit Sub
    ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
        Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
    
    Call QuitarUserInvItem(UserIndex, Slot, Cantidad)
    
    Precio = ObjData(Objeto.ObjIndex).Valor \ REDUCTOR_PRECIOVENTA * Cantidad
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
    
    If UserList(UserIndex).Stats.GLD > MAXORO Then _
        UserList(UserIndex).Stats.GLD = MAXORO
    
    Dim NpcSlot As Integer
    NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.amount)
    
    If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
        'Mete el obj en el slot
        Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
        Npclist(NpcIndex).Invent.Object(NpcSlot).amount = Npclist(NpcIndex).Invent.Object(NpcSlot).amount + Objeto.amount
        If Npclist(NpcIndex).Invent.Object(NpcSlot).amount > MAX_INVENTORY_OBJS Then
            Npclist(NpcIndex).Invent.Object(NpcSlot).amount = MAX_INVENTORY_OBJS
        End If
    End If
    
End If
        
    Call UpdateUserInv(True, UserIndex, 0)
    Call WriteUpdateUserStats(UserIndex)
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
    Call WriteTradeOK(UserIndex)
        
    Call SubirSkill(UserIndex, eSkill.Comerciar)
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
    UserList(UserIndex).flags.Comerciando = True
    Call WriteCommerceInit(UserIndex)
End Sub

Private Function SlotEnNPCInv(NpcIndex As Integer, Objeto As Integer, Cantidad As Integer) As Integer
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    SlotEnNPCInv = 1
    Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto _
      And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).amount + Cantidad <= MAX_INVENTORY_OBJS
        
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

Private Function Descuento(ByVal UserIndex As Integer) As Single
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100
End Function

Private Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    Dim Slot As Byte
    Dim val As Long
    
    For Slot = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then
            Dim thisObj As Obj
            thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
            thisObj.amount = Npclist(NpcIndex).Invent.Object(Slot).amount
            val = (ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor) / Descuento(UserIndex)
            
            Call WriteChangeNPCInventorySlot(UserIndex, thisObj, val)
        Else
            Dim DummyObj As Obj
            Call WriteChangeNPCInventorySlot(UserIndex, DummyObj, 0)
        End If
    Next Slot
End Sub

