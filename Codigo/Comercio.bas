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
    'Last modified: 07/06/2010
    '27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
    '  - 06/13/08 (NicoNZ)
    '07/06/2010: ZaMa - Los objetos se loguean si superan la cantidad de 1k (antes era solo si eran 1k).
    '*************************************************
    Dim Precio As Long

    Dim Objeto As obj
    
    If Cantidad < 1 Or Slot < 1 Then Exit Sub
    
    If Modo = eModoComercio.Compra Then
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Sub
        ElseIf Cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(Userindex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
            Call Ban(UserList(Userindex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
            UserList(Userindex).flags.Ban = 1
            Call WriteErrorMsg(Userindex, "Has sido baneado por el Sistema AntiCheat.")
            Call FlushBuffer(Userindex)
            Call CloseSocket(Userindex)
            Exit Sub
        ElseIf Not Npclist(NpcIndex).Invent.Object(Slot).Amount > 0 Then
            Exit Sub

        End If
        
        If Cantidad > Npclist(NpcIndex).Invent.Object(Slot).Amount Then Cantidad = Npclist(UserList(Userindex).flags.TargetNPC).Invent.Object(Slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
        
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        
        Precio = CLng((ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Valor / Descuento(Userindex) * Cantidad) + 0.5)

        If UserList(Userindex).Stats.Gld < Precio Then
            Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If MeterItemEnInventario(Userindex, Objeto) = False Then
            'Call WriteConsoleMsg(UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
            Call WriteTradeOK(Userindex)
            Exit Sub

        End If
        
        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - Precio
        
        Call QuitarNpcInvItem(UserList(Userindex).flags.TargetNPC, CByte(Slot), Cantidad)
        
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
            Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
            Call WriteTradeOK(Userindex)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then

            If Npclist(NpcIndex).Name <> "SR" Then
                Call WriteConsoleMsg(Userindex, "Las armaduras del ejercito real solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
                Call WriteTradeOK(Userindex)
                Exit Sub

            End If

        ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then

            If Npclist(NpcIndex).Name <> "SC" Then
                Call WriteConsoleMsg(Userindex, "Las armaduras de la legion oscura solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
                Call WriteTradeOK(Userindex)
                Exit Sub

            End If

        ElseIf UserList(Userindex).Invent.Object(Slot).Amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf Slot < LBound(UserList(Userindex).Invent.Object()) Or Slot > UBound(UserList(Userindex).Invent.Object()) Then
            Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
            Exit Sub
        ElseIf UserList(Userindex).flags.Privilegios And PlayerType.Consejero Then
            Call WriteConsoleMsg(Userindex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
            Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
            Call WriteTradeOK(Userindex)
            Exit Sub

        End If
        
        Call QuitarUserInvItem(Userindex, Slot, Cantidad)
        
        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
        UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld + Precio
        
        If UserList(Userindex).Stats.Gld > MAXORO Then UserList(Userindex).Stats.Gld = MAXORO
        
        Dim NpcSlot As Integer

        NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
        
        If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
            Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount

            If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS

            End If

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
    
    Call UpdateUserInv(True, Userindex, 0)
    Call WriteUpdateUserStats(Userindex)
    Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
    Call WriteTradeOK(Userindex)
        
    Call SubirSkill(Userindex, eSkill.Comerciar, True)

End Sub

Public Sub IniciarComercioNPC(ByVal Userindex As Integer)
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    Call EnviarNpcInv(Userindex, UserList(Userindex).flags.TargetNPC)
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
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal Userindex As Integer, ByVal NpcIndex As Integer)

    '*************************************************
    'Author: Nacho (Integer)
    'Last Modified: 06/14/08
    'Last Modified By: Nicolas Ezequiel Bouhid (NicoNZ)
    '*************************************************
    Dim Slot As Byte

    Dim val  As Single
    
    For Slot = 1 To MAX_NORMAL_INVENTORY_SLOTS

        If Npclist(NpcIndex).Invent.Object(Slot).ObjIndex > 0 Then

            Dim thisObj As obj
            
            thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex
            thisObj.Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount
            
            val = (ObjData(thisObj.ObjIndex).Valor) / Descuento(Userindex)
            
            Call WriteChangeNPCInventorySlot(Userindex, Slot, thisObj, val)
        Else

            Dim DummyObj As obj

            Call WriteChangeNPCInventorySlot(Userindex, Slot, DummyObj, 0)

        End If

    Next Slot

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
