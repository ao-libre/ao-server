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

Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 06/13/08 (NicoNZ)
'*************************************************
    Dim Precio As Long
    Dim Objeto As Obj
    
    If Cantidad < 1 Or Slot < 1 Then Exit Sub
    
    If Modo = eModoComercio.Compra Then
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
        
        Objeto.amount = Cantidad
        Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
        
        Precio = Round(ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor / Descuento(UserIndex) * Cantidad, 0)
            
        If UserList(UserIndex).Stats.GLD < Precio Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If MeterItemEnInventario(UserIndex, Objeto) = False Then
            'Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
        
        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(Slot), Cantidad)
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).name & " compró del NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).name)
        ElseIf Objeto.amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " compró del NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).name)
            End If
        End If
        
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
        End If
        
    ElseIf Modo = eModoComercio.Venta Then
        
        If Cantidad > UserList(UserIndex).Invent.Object(Slot).amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).amount
        
        Objeto.amount = Cantidad
        Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        If Objeto.ObjIndex = 0 Then
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Newbie = 1 Then
            Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio objetos para newbies.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
            Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then
            If Npclist(NpcIndex).name <> "SR" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras de la Armada solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub
            End If
        ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then
            If Npclist(NpcIndex).name <> "SC" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras de la Legión solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub
            End If
        ElseIf UserList(UserIndex).Invent.Object(Slot).amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
            Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        End If
        
        Call QuitarUserInvItem(UserIndex, Slot, Cantidad)
        
        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        Precio = Round(SalePrice(ObjData(Objeto.ObjIndex).Valor) * Cantidad, 0)
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
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Objeto.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).name & " vendió al NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).name)
        ElseIf Objeto.amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " vendió al NPC " & Objeto.amount & " " & ObjData(Objeto.ObjIndex).name)
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

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
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

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last Modified: 06/14/08
'Last Modified By: Nicolás Ezequiel Bouhid (NicoNZ)
'*************************************************
    Dim Slot As Byte
    Dim val As Long
    
    For Slot = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then
            Dim thisObj As Obj
            thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
            thisObj.amount = Npclist(NpcIndex).Invent.Object(Slot).amount
            val = (ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor) / Descuento(UserIndex)
            
            Call WriteChangeNPCInventorySlot(UserIndex, Slot, thisObj, val)
        Else
            Dim DummyObj As Obj
            Call WriteChangeNPCInventorySlot(UserIndex, Slot, DummyObj, 0)
        End If
    Next Slot
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param valor  El valor de compra de objeto

Public Function SalePrice(ByVal Valor As Long) As Single
'*************************************************
'Author: Nicolás (NicoNZ)
'
'*************************************************

    SalePrice = Valor / REDUCTOR_PRECIOVENTA

End Function
